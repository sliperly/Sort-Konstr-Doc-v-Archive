#!/usr/bin/env python3
r"""
Автоматическая сортировка архивов номенклатуры по категориям 1С УНФ.

Логика:
  1. Читает Структура.xlsx (выгрузка из 1С УНФ), строит соответствие
     код номенклатуры -> путь категории, используя ГРУППИРОВКУ (Outline)
     Excel-листа (тот самый +/- слева от номеров строк).
  2. Для каждого файла в SOURCE ищет коды НФ-XXXXXXXX в ЛЮБОМ месте имени
     (например, Тележка (НФ-00034720).rar, Копия НФ-00000001.dwg).
  3. Находит категорию по коду, перемещает файл в DEST/<категория>/.
     Если папки категории ещё нет — создаёт её (вместе с родительскими).
  4. Если в целевой папке уже есть файл с таким именем — заменяет его.
  5. Если в имени несколько кодов:
       - все коды ведут в одну категорию -> файл переносится;
       - коды из разных категорий (или какого-то кода нет в прайс-листе)
         -> файл НЕ ТРОГАЕТСЯ, причина пишется в unprocessed_report.csv.
  6. Файлы без кода в имени, без совпадения в прайс-листе или при ошибке
     создания папки/переноса НЕ ТРОГАЮТСЯ, только логируются.

Использование:
  python nomenclature_sort.py                 # обычный запуск, реально переносит файлы
  python nomenclature_sort.py --dry-run        # только показать, что было бы сделано
  python nomenclature_sort.py --dump-map       # вывести карту код->категория и выйти
                                                  (для проверки, что прайс-лист читается верно)

После обычного запуска (не --dry-run и не --dump-map) результат дополнительно
отправляется POST-запросом на N8N_WEBHOOK_URL (см. константу ниже), если она
задана. Так n8n может слать уведомления, не имея прямого доступа к Python
и файловой системе Windows.

Запускается Планировщиком заданий Windows по расписанию (9:00, 13:00, 16:00).
Все файлы программы, прайс-лист и логи лежат в C:\NomenklaturaSort\.

Результат каждого запуска (успешно/с ошибками, время выполнения, список
файлов) пишется в logs\app_run.log — там же дублируется вывод, даже если
exe запущен без консоли (например, из-под Планировщика заданий).
"""

import re
import csv
import shutil
import sys
import json
import argparse
import platform
import time
import traceback
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
except ImportError:
    _msg = json.dumps({"status": "error", "message": "Не установлен openpyxl. Установите: pip install openpyxl"}, ensure_ascii=False)
    if sys.stdout is not None:
        print(_msg)
    sys.exit(1)


# --- Конфигурация путей ---
# Программа, прайс-лист и логи — локально, без зависимости от сетевых
# ссылок и от Hermes agent.
WORKSPACE = Path(r"C:\NomenklaturaSort")
PRICE_LIST = WORKSPACE / "PriceList" / "Структура.xlsx"  # обновляйте этот файл вручную
LOG_DIR = WORKSPACE / "logs"

# Источник и назначение — реальные сетевые папки на KometaNAS.
SOURCE = Path(r"\\KometaNAS\Производство\Номенклатура\Номенклатурщик")
DEST = Path(r"\\KometaNAS\Производство\Номенклатура\Готовая номенклатура")

TRANSFER_LOG = LOG_DIR / "transfer_log.csv"
ISSUES_LOG = LOG_DIR / "unprocessed_report.csv"
APP_LOG = LOG_DIR / "app_run.log"

# URL webhook-узла n8n. Оставьте пустой строкой, если уведомления пока не нужны.
# Порт 5678 у вас уже проброшен наружу контейнера n8n (видно в docker ps),
# поэтому localhost с хоста Windows должен достучаться без проблем.
N8N_WEBHOOK_URL = "http://localhost:5678/webhook/nomenclature-sort"

# Код в имени файла: НФ-XXXXXXXX в любом месте имени.
# (?!\d) — чтобы НФ-123456789 (9 цифр) не принимался за код из 8 цифр.
# Групп в паттерне нет, поэтому findall() возвращает сами коды целиком.
FILENAME_CODE_PATTERN = re.compile(r"НФ-\d{8}(?!\d)")

# Код в ячейке прайс-листа: только в начале ячейки. Якорь нужен, чтобы
# отличать строки номенклатуры от заголовков категорий.
CODE_PATTERN = re.compile(r"^(НФ-\d{8})")

# Столбец с кодом номенклатуры (индекс с 1). ПРОВЕРЬТЕ по вашему файлу
# через --dump-map перед первым реальным запуском.
CODE_COLUMN = 2
# Столбец, из которого брать текст заголовка категории, если строка —
# заголовок, а не код (обычно название лежит в том же или соседнем столбце).
NAME_COLUMN = 5


def safe_print(*args, file=None, **kwargs):
    """
    Обёртка вокруг print(), которая не роняет процесс, если stdout/stderr
    отсутствуют (сборка --windowed/--noconsole, запуск из Планировщика заданий).
    Дополнительно дублирует вывод в APP_LOG, чтобы результат был виден
    даже без консоли.
    """
    target = file if file is not None else sys.stdout
    text = " ".join(str(a) for a in args)

    if target is not None:
        try:
            print(text, file=target, **kwargs)
        except Exception:
            pass

    try:
        APP_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(APP_LOG, "a", encoding="utf-8-sig") as f:
            f.write(f"[{datetime.now().isoformat(timespec='seconds')}] {text}\n")
    except Exception:
        pass


def long_path(p: Path) -> str:
    """
    Обходит ограничение Windows в 260 символов на длину пути.
    Без этого shutil.move падает с FileNotFoundError на глубоко
    вложенных категориях, хотя папка реально существует.
    """
    if platform.system() != "Windows":
        return str(p)
    resolved = str(p.resolve())
    if resolved.startswith("\\\\?\\"):
        return resolved
    if resolved.startswith("\\\\"):
        return "\\\\?\\UNC\\" + resolved.lstrip("\\")
    return "\\\\?\\" + resolved


def notify_n8n(payload: dict) -> None:
    """Отправляет результат в n8n. Не роняет скрипт, если n8n недоступен."""
    if not N8N_WEBHOOK_URL:
        return
    try:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            N8N_WEBHOOK_URL, data=data,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        urllib.request.urlopen(req, timeout=10)
    except (urllib.error.URLError, OSError) as e:
        safe_print(f"Предупреждение: не удалось отправить уведомление в n8n: {e}", file=sys.stderr)


def build_category_map(xlsx_path: Path) -> dict[str, str]:
    """Строит {код: относительный_путь_категории} по группировке Excel."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    mapping: dict[str, str] = {}
    stack: dict[int, str] = {}  # уровень вложенности -> название категории

    for row in ws.iter_rows(min_row=1):
        code_cell = row[CODE_COLUMN - 1]
        level = ws.row_dimensions[code_cell.row].outline_level
        value = str(code_cell.value) if code_cell.value else ""
        match = CODE_PATTERN.match(value)

        if match:
            # строка с кодом номенклатуры — собрать путь из всех родительских
            # заголовков (уровни строго меньше текущего), КРОМЕ самого верхнего —
            # он на диске не отражён как папка (например, "Товары" в 1С — это
            # верхнеуровневая категория, а не имя реальной директории)
            path_parts = [stack[lvl] for lvl in sorted(stack) if lvl < level]
            path_parts = path_parts[1:]
            mapping[match.group(1)] = "/".join(path_parts) if path_parts else "Без категории"
        else:
            # заголовок категории — обновить стек, сбросить более глубокие уровни
            name_cell = row[NAME_COLUMN - 1].value or code_cell.value
            if name_cell and str(name_cell).strip():
                stack[level] = str(name_cell).strip()
                for lvl in [l for l in stack if l > level]:
                    del stack[lvl]

    return mapping


def resolve_category(filename: str, category_map: dict[str, str]):
    """
    Определяет категорию файла по кодам в его имени.
    Возвращает (категория, None) при успехе или (None, причина) при отказе.
    """
    # dict.fromkeys убирает дубли, сохраняя порядок появления в имени
    codes = list(dict.fromkeys(FILENAME_CODE_PATTERN.findall(filename)))
    if not codes:
        return None, "в имени не найден код НФ-XXXXXXXX"

    missing = [c for c in codes if c not in category_map]
    if missing:
        if len(codes) == 1:
            return None, f"код {missing[0]} не найден в прайс-листе"
        return None, f"коды не найдены в прайс-листе: {', '.join(missing)}"

    categories = {category_map[c] for c in codes}
    if len(categories) > 1:
        details = "; ".join(f"{c} -> {category_map[c]}" for c in codes)
        return None, f"коды относятся к разным категориям: {details}"

    return categories.pop(), None


def main():
    start_time = time.monotonic()
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="ничего не перемещать, только показать план")
    parser.add_argument("--dump-map", action="store_true", help="вывести карту код->категория и выйти")
    args = parser.parse_args()

    if not PRICE_LIST.exists():
        safe_print(json.dumps({"status": "error", "message": f"Прайс-лист недоступен: {PRICE_LIST}"}, ensure_ascii=False))
        sys.exit(1)

    category_map = build_category_map(PRICE_LIST)
    if not category_map:
        safe_print(json.dumps({"status": "error", "message": "Прайс-лист прочитан, но карта категорий пуста — проверьте CODE_COLUMN/NAME_COLUMN"}, ensure_ascii=False))
        sys.exit(1)

    if args.dump_map:
        for code, path in sorted(category_map.items()):
            safe_print(f"{code}\t{path}")
        safe_print(f"\nВсего кодов в прайс-листе: {len(category_map)}", file=sys.stderr)
        return

    if not SOURCE.exists():
        safe_print(json.dumps({"status": "error", "message": f"Папка-источник недоступна: {SOURCE}"}, ensure_ascii=False))
        sys.exit(1)

    transferred = []
    issues = []

    for file in sorted(SOURCE.iterdir()):
        if not file.is_file():
            continue
        if file.name.lower() == "desktop.ini":
            continue  # служебный файл Windows, не номенклатура — молча пропустить

        category_path, error = resolve_category(file.name, category_map)
        if error:
            issues.append((file.name, error))
            continue

        target_dir = DEST / category_path
        target_file = target_dir / file.name
        target_file_lp = long_path(target_file)
        replaced = Path(target_file_lp).exists()
        action = "ЗАМЕНА" if replaced else "новый"
        relative_target = str(target_file.relative_to(DEST))

        if args.dry_run:
            transferred.append((file.name, relative_target, action))
            continue

        try:
            # Создать папку категории (и родительские), если её ещё нет
            Path(long_path(target_dir)).mkdir(parents=True, exist_ok=True)
            if replaced:
                Path(target_file_lp).unlink()
            shutil.move(long_path(file), target_file_lp)
        except OSError as e:
            # Ошибка одного файла (нет прав, файл занят и т.п.) не должна
            # останавливать обработку остальных
            issues.append((file.name, f"ошибка переноса в {category_path}: {e}"))
            continue

        transferred.append((file.name, relative_target, action))

    if args.dry_run:
        safe_print(json.dumps({
            "status": "dry-run",
            "would_transfer": len(transferred),
            "would_skip": len(issues),
            "transfers": transferred,
            "issues": issues,
        }, ensure_ascii=False, indent=2))
        return

    # --- Логи ---
    timestamp = datetime.now().isoformat(timespec="seconds")
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    write_header = not TRANSFER_LOG.exists()
    with open(TRANSFER_LOG, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if write_header:
            writer.writerow(["Timestamp", "FileName", "TargetPath", "Action"])
        for name, path, action in transferred:
            writer.writerow([timestamp, name, path, action])

    with open(ISSUES_LOG, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["FileName", "Reason"])
        for name, reason in issues:
            writer.writerow([name, reason])

    duration_seconds = round(time.monotonic() - start_time, 1)

    result = {
        "status": "ok",
        "timestamp": timestamp,
        "duration_seconds": duration_seconds,
        "transferred_count": len(transferred),
        "transferred": [
            {"file": name, "target": path, "action": action}
            for name, path, action in transferred
        ],
        "issues_count": len(issues),
        "issues": [{"file": name, "reason": reason} for name, reason in issues],
    }

    # Человекочитаемая сводка — для app_run.log и консоли.
    summary_lines = [
        f"Готово за {duration_seconds} сек. Перенесено файлов: {len(transferred)}, проблем: {len(issues)}.",
    ]
    if transferred:
        summary_lines.append("Успешно обработаны:")
        for name, path, action in transferred:
            summary_lines.append(f"  [{action}] {name} -> {path}")
    if issues:
        summary_lines.append("Не обработаны:")
        for name, reason in issues:
            summary_lines.append(f"  {name}: {reason}")
    safe_print("\n".join(summary_lines))

    # JSON — отдельной строкой, для парсинга n8n/скриптами.
    safe_print(json.dumps(result, ensure_ascii=False))
    notify_n8n(result)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        # Без этого при --noconsole необработанное исключение исчезает бесследно
        safe_print("КРИТИЧЕСКАЯ ОШИБКА:\n" + traceback.format_exc())
        sys.exit(1)

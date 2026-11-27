#!/usr/bin/env python3
r"""
ЭТАП 1. Чтение структуры из универсального отчёта УНФ и проверка названий.

Скрипт ТОЛЬКО ЧИТАЕТ выгрузку и создаёт один Excel-отчёт. Диск архива не затрагивается,
никакие файлы не переносятся и не переименовываются.

Что делает:
  1. Строит дерево групп и изделий по кодам родителей (столбец «Номенклатура, Код»).
  2. Проверяет целостность дерева: дубли кодов, потерянные родители, циклы.
  3. Работает только с группой «Товары» (ROOT_GROUP_NAME в kd_common.py) и только с
     действующими изделиями: позиции с пометкой «Недействительна» игнорируются.
  4. Проверяет названия групп (будущие папки) и изделий (будущие имена файлов):
     запрещённые символы, пробелы, латиница/кириллица, размеры с латинской x,
     нестандартные символы, длина пути, одинаковые и «неотличимые» названия.
  5. Показывает известные проблемные места структуры («Со светом» и др.).

Использование:
  python kd_stage1.py
  python kd_stage1.py --unf "D:\Выгрузки\Структура_1.xlsx"
  python kd_stage1.py --probe "Со светом" --probe "Мойка сварная"

Отчёт: reports\unf_structure_<дата_время>.xlsx. Журнал запусков: logs\kd_run.log.
"""

import argparse
import statistics
import sys
import time
import traceback
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from kd_common import (ARCHIVE_ROOT, DEFAULT_PROBES, FILE_NAME_OVERHEAD, MAX_FILE_PATH, MAX_FOLDER_PATH,
                       REPORT_DIR, ROOT_GROUP_NAME, UNF_REPORT, log)
from kd_names import ERR, WARN, check_name, cluster_by_visual, diff_chars, visual_key
from kd_unf import INFO, UnfFormatError, read_unf

NO_CATEGORY = "(корень архива)"

PROBLEM_HEADERS = ["Серьёзность", "Объект", "Вид проблемы", "Код", "Строка Excel", "Название в УНФ",
                   "С пометками ⟦…⟧", "Папка / путь", "Подробности", "Предложение"]


def path_str(parts) -> str:
    return "/".join(parts) if parts else NO_CATEGORY


# ---------------------------------------------------------------------------
# Область работы: только «Товары»
# ---------------------------------------------------------------------------

def build_scope(tree):
    """
    Возвращает (корневые коды, {код: (узел, путь)}) для группы «Товары».
    Недействительные изделия в область работы не входят; группы остаются.
    """
    key = visual_key(ROOT_GROUP_NAME)
    roots = [n for n in tree.roots() if visual_key(n.name) == key]
    if not roots:
        raise UnfFormatError(f"в выгрузке нет корневой группы «{ROOT_GROUP_NAME}»; "
                             f"корневые группы: {', '.join(n.name.strip() for n in tree.roots())}")
    scope = {}
    for root in roots:
        for node, parts in tree.walk(root.code):
            if node.invalid and not node.is_group:
                continue
            scope[node.code] = (node, parts)
    return {r.code for r in roots}, scope


def invalid_in_goods(tree, root_codes) -> int:
    """Сколько недействительных изделий лежит в «Товарах» (они игнорируются)."""
    return sum(1 for root in root_codes for node, _p in tree.walk(root)
               if node.invalid and not node.is_group)


def group_metrics(tree, scope) -> dict:
    direct, total, subgroups = Counter(), Counter(), Counter()
    over, maxlen = Counter(), Counter()
    file_lengths = []
    for code, (node, parts) in scope.items():
        if node.is_group:
            if node.parent in scope:
                subgroups[node.parent] += 1
        elif node.parent in scope:
            direct[node.parent] += 1
            file_len = (len(ARCHIVE_ROOT) + 1 + len("\\".join(parts)) + 1
                        + FILE_NAME_OVERHEAD + len(node.name.strip()))
            file_lengths.append(file_len)
            maxlen[node.parent] = max(maxlen[node.parent], file_len)
            if file_len > MAX_FILE_PATH:
                over[node.parent] += 1
            p = node.parent
            while p is not None and p in scope:
                total[p] += 1
                p = tree.nodes[p].parent
    return {"direct": direct, "total": total, "subgroups": subgroups,
            "over": over, "maxlen": maxlen, "file_lengths": file_lengths}


# ---------------------------------------------------------------------------
# Проблемы названий
# ---------------------------------------------------------------------------

def collect_name_problems(tree, scope, root_codes):
    rows = []
    stats = Counter()

    def add(severity, obj, kind, node, marked, path, detail, suggestion=""):
        rows.append([severity, obj, kind, node.code, node.row, node.name, marked, path, detail, suggestion])

    for code, (node, parts) in scope.items():
        if code in root_codes:
            continue
        obj = "папка" if node.is_group else "изделие"
        path = path_str(parts)
        for p in check_name(node.name, obj):
            add(p.severity, obj, p.kind, node, p.marked, path, p.detail, p.suggestion)

        if node.is_group:
            base = len(ARCHIVE_ROOT) + 1 + len("\\".join(parts))
            if base > MAX_FOLDER_PATH:
                add(WARN, obj, "длинный путь папки", node, "", path,
                    f"путь вместе с корнем архива: {base} символов при пороге {MAX_FOLDER_PATH}; "
                    f"для имени файла остаётся {MAX_FILE_PATH - base - 1} символов")

    # Одинаковые и «неотличимые» названия папок у одного родителя
    for parent_code, children in tree.children.items():
        if parent_code not in scope:
            continue
        named = [(c, tree.nodes[c].name) for c in children if tree.nodes[c].is_group]
        for cluster in cluster_by_visual(named):
            same = cluster.kind == "одинаковые для Windows"
            kind = ("одинаковые имена папок у одного родителя" if same
                    else "похожие имена папок у одного родителя")
            for code, name in cluster.members:
                others = [(c, n) for c, n in cluster.members if c != code]
                detail = ("на диске получится одна папка из разных групп УНФ: " if same
                          else "разные папки с почти неотличимыми именами: ")
                detail += "; ".join(f"{c} «{n}»" for c, n in others)
                if not same:
                    detail += ". Отличия: " + diff_chars(name, others[0][1])
                node = tree.nodes[code]
                add(ERR if same else WARN, "папка", kind, node, "", path_str(scope[code][1]), detail)

    # Похожие названия изделий (разные написания, одинаковый вид)
    items = [(c, n.name) for c, (n, _p) in scope.items() if not n.is_group]
    same_name_items = 0
    for cluster in cluster_by_visual(items):
        if len({n for _, n in cluster.members}) == 1:
            same_name_items += len(cluster.members)
            continue
        for code, name in cluster.members:
            others = [(c, n) for c, n in cluster.members if c != code and n != name]
            if not others:
                continue
            node = tree.nodes[code]
            detail = "выглядит как: " + "; ".join(f"{c} «{n}»" for c, n in others[:3])
            detail += ". Отличия: " + diff_chars(name, others[0][1])
            add(WARN, "изделие", "похожие названия изделий", node, "", path_str(scope[code][1]), detail)
    stats["same_name_items"] = same_name_items
    return rows, stats


# ---------------------------------------------------------------------------
# Остальные разделы отчёта
# ---------------------------------------------------------------------------

def tree_rows(tree, scope, root_codes, metrics) -> list:
    rows = []
    for code, (node, parts) in scope.items():
        if not node.is_group or code in root_codes:
            continue
        total = metrics["total"][code]
        rows.append([
            "/".join(parts), len(parts), node.name.strip(), code, node.parent or "",
            metrics["direct"][code], total, metrics["subgroups"][code],
            "да" if total == 0 else "", node.row,
            len(ARCHIVE_ROOT) + 1 + len("\\".join(parts)),
            metrics["over"][code] or "", metrics["maxlen"][code] or "",
        ])
    rows.sort(key=lambda r: r[0].casefold())
    return rows


def integrity_rows(tree, scope, root_codes) -> list:
    rows = [[i.severity, i.kind, i.code, i.row, i.detail] for i in tree.issues]
    for root in tree.roots():
        rows.append([INFO, "корневая группа", root.code, root.row,
                     f"«{root.name.strip()}»" + (" — в работе" if root.code in root_codes else " — не относится к архиву КД")])
    orphans = [n for n in tree.orphan_items() if not n.invalid]
    rows.append([INFO, "действующие изделия без группы", "", "",
                 f"{len(orphans)} (вне «{ROOT_GROUP_NAME}», в архив КД не попадают)"])
    return rows


def probe_rows(tree, scope, metrics, probes) -> list:
    rows = []
    groups = [n for n in tree.nodes.values() if n.is_group]
    for probe in probes:
        key = visual_key(probe)
        found = [n for n in groups if visual_key(n.name) == key]
        if not found:
            rows.append([probe, "не найдена", "", "", "", "", "", "", ""])
            continue
        for n in found:
            in_scope = n.code in scope
            parent = tree.nodes.get(n.parent)
            kids = [tree.nodes[c].name.strip() for c in tree.children.get(n.code, []) if tree.nodes[c].is_group]
            rows.append([
                probe, n.name.strip(), n.code,
                f"{parent.name.strip()} ({parent.code})" if parent else "",
                "/".join(scope[n.code][1]) if in_scope else f"вне «{ROOT_GROUP_NAME}»",
                metrics["direct"][n.code] if in_scope else "",
                metrics["total"][n.code] if in_scope else "",
                len(kids), "; ".join(kids[:8]) + (" …" if len(kids) > 8 else ""),
            ])
    return rows


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="DDE6F0")
SEVERITY_FILL = {ERR: PatternFill("solid", fgColor="F8D7DA"), WARN: PatternFill("solid", fgColor="FFF3CD")}


def xl(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)[:32000]
    return value


def write_table(wb, title, headers, rows, widths, severity_col=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in rows:
        ws.append([xl(c) for c in row])
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if severity_col is not None:
        for row in ws.iter_rows(min_row=2):
            fill = SEVERITY_FILL.get(row[severity_col].value)
            if fill:
                row[severity_col].fill = fill
    return ws


def write_summary(wb, lines):
    ws = wb.active
    ws.title = "Сводка"
    for kind, a, b in lines:
        ws.append([xl(a), xl(b)] if kind != "h" else [xl(a)])
        if kind == "h":
            ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 60


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Этап 1: структура УНФ и проверка названий (только чтение)")
    parser.add_argument("--unf", type=Path, default=UNF_REPORT, help="выгрузка универсального отчёта УНФ")
    parser.add_argument("--out", type=Path, default=REPORT_DIR, help="папка для отчёта")
    parser.add_argument("--probe", action="append", default=[], help="название группы для раздела «Известные случаи»")
    args = parser.parse_args(argv)

    start = time.monotonic()
    if not args.unf.exists():
        log(f"Ошибка: выгрузка не найдена: {args.unf}")
        return 1

    log(f"=== Этап 1. Читаю структуру УНФ: {args.unf}")
    tree = read_unf(args.unf)
    root_codes, scope = build_scope(tree)
    metrics = group_metrics(tree, scope)
    invalid_count = invalid_in_goods(tree, root_codes)

    groups_in = [n for c, (n, _p) in scope.items() if n.is_group and c not in root_codes]
    items_in = [n for c, (n, _p) in scope.items() if not n.is_group]
    log(f"Записей в выгрузке: {len(tree.nodes)} (групп {sum(1 for n in tree.nodes.values() if n.is_group)}); "
        f"в работе «{ROOT_GROUP_NAME}»: групп {len(groups_in)}, действующих изделий {len(items_in)} "
        f"(недействительных пропущено: {invalid_count})")

    problems, pstats = collect_name_problems(tree, scope, root_codes)
    sev_order = {ERR: 0, WARN: 1}
    problems.sort(key=lambda r: (sev_order.get(r[0], 2), r[1], r[2], r[7].casefold(), r[5].casefold()))
    by_kind = Counter((r[0], r[1], r[2]) for r in problems)

    probes = DEFAULT_PROBES + args.probe
    probe_data = probe_rows(tree, scope, metrics, probes)
    integrity = integrity_rows(tree, scope, root_codes)
    errors_in_tree = [i for i in tree.issues if i.severity == ERR]

    # --- Сводка ---
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    lines = [
        ("h", "Источник", ""),
        ("r", "Файл выгрузки", str(args.unf)),
        ("r", "Изменён", datetime.fromtimestamp(args.unf.stat().st_mtime).strftime("%Y-%m-%d %H:%M")),
        ("r", "Отчёт создан", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("h", "Дерево", ""),
        ("r", "Записей в выгрузке", len(tree.nodes)),
        ("r", "Корневые группы", "; ".join(n.name.strip() for n in tree.roots())),
        ("r", f"В работе: группа «{ROOT_GROUP_NAME}»", f"групп {len(groups_in)}, действующих изделий {len(items_in)}"),
        ("r", "Правило", "в дерево, подсчёты и проверки входят только действующие изделия"),
        ("r", "Недействительных изделий в «Товарах» (проигнорированы)", invalid_count),
        ("r", "Групп без действующих изделий", sum(1 for r in tree_rows(tree, scope, root_codes, metrics) if r[8])),
        ("r", "Изделий прямо в корне «Товары» (без подгруппы)", metrics["direct"][next(iter(root_codes))]),
        ("r", "Действующих изделий без группы (вне «Товары»)", sum(1 for n in tree.orphan_items() if not n.invalid)),
        ("r", "Максимальная глубина", max((len(p) for _n, p in scope.values()), default=0)),
        ("r", "Ошибки целостности дерева", len(errors_in_tree)),
        ("h", f"Длина путей (корень архива {len(ARCHIVE_ROOT)} симв., имя файла «НФ-… Название.расш»)", ""),
        ("r", "Медианная длина полного пути файла", int(statistics.median(metrics["file_lengths"])) if metrics["file_lengths"] else 0),
        ("r", f"Изделий с путём файла длиннее {MAX_FILE_PATH}",
         f"{sum(1 for x in metrics['file_lengths'] if x > MAX_FILE_PATH)} из {len(metrics['file_lengths'])}"),
        ("r", "Изделий с путём файла длиннее 300", sum(1 for x in metrics["file_lengths"] if x > 300)),
        ("r", "Папок с путём длиннее порога", sum(1 for r in problems if r[2] == "длинный путь папки")),
        ("h", "Проблемы названий (подробности — на листе «Проблемы названий»)", ""),
    ]
    for (sev, obj, kind), n in sorted(by_kind.items(), key=lambda kv: (sev_order.get(kv[0][0], 2), kv[0][1], -kv[1])):
        lines.append(("r", f"[{sev}] {obj}: {kind}", n))
    lines.append(("r", "Изделия с точно одинаковым названием (не считаются проблемой)", pstats["same_name_items"]))
    lines += [
        ("h", "Листы отчёта", ""),
        ("r", "Дерево папок", "все группы «Товары», путь, число изделий, длина пути"),
        ("r", "Проблемы названий", "что исправить в УНФ или учесть при создании папок и файлов"),
        ("r", "Целостность дерева", "ошибки выгрузки и общая картина по корням"),
        ("r", "Известные случаи", "проверка мест, где прежнее чтение ошибалось"),
    ]

    wb = openpyxl.Workbook()
    write_summary(wb, lines)
    write_table(wb, "Дерево папок",
                ["Путь", "Уровень", "Название", "Код", "Код родителя", "Изделий прямо", "Изделий всего",
                 "Подгрупп", "Нет действующих изделий", "Строка Excel", "Длина пути папки с корнем архива",
                 f"Изделий, у которых путь файла длиннее {MAX_FILE_PATH}", "Макс. длина пути файла"],
                tree_rows(tree, scope, root_codes, metrics), [70, 8, 40, 14, 14, 12, 12, 10, 12, 10, 14, 18, 14])
    write_table(wb, "Проблемы названий", PROBLEM_HEADERS, problems,
                [14, 10, 34, 14, 10, 44, 44, 50, 70, 40], severity_col=0)
    write_table(wb, "Целостность дерева", ["Серьёзность", "Вид", "Код", "Строка Excel", "Подробности"],
                integrity, [14, 34, 14, 12, 100], severity_col=0)
    write_table(wb, "Известные случаи",
                ["Образец", "Группа", "Код", "Родитель", "Путь", "Изделий прямо", "Изделий всего", "Подгрупп",
                 "Подгруппы (первые 8)"], probe_data, [30, 36, 14, 50, 70, 12, 12, 10, 80])

    args.out.mkdir(parents=True, exist_ok=True)
    out_path = args.out / f"unf_structure_{stamp}.xlsx"
    wb.save(out_path)

    errs = sum(1 for r in problems if r[0] == ERR)
    warns = sum(1 for r in problems if r[0] == WARN)
    log(f"Проблем названий: ошибок {errs}, предупреждений {warns}; ошибок целостности дерева: {len(errors_in_tree)}")
    log(f"Отчёт: {out_path}")
    log(f"Готово за {round(time.monotonic() - start, 1)} сек.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except UnfFormatError as e:
        log(f"Ошибка: {e}")
        sys.exit(1)
    except Exception:
        log("КРИТИЧЕСКАЯ ОШИБКА:\n" + traceback.format_exc())
        sys.exit(1)

"""
Общие настройки и вспомогательные функции проекта «Архив КД».

Все пути и пороги, которые может понадобиться менять, собраны здесь.
"""

import sys
from datetime import datetime
from pathlib import Path

# --- Папки и файлы ---
WORKSPACE = Path(r"C:\NomenklaturaSort")
# Выгрузка универсального отчёта УНФ (справочник «Номенклатура» с кодом родителя).
# Чтобы обновить структуру, положите новую выгрузку на место старой или
# укажите путь ключом --unf.
UNF_REPORT = WORKSPACE / "PriceList" / "Структура_1.xlsx"
# Прайс-лист, по которому работал старый сортировщик. Нужен только для сравнения (--compare).
OLD_PRICE_LIST = WORKSPACE / "PriceList" / "Структура.xlsx"
REPORT_DIR = WORKSPACE / "reports"
LOG_DIR = WORKSPACE / "logs"
RUN_LOG = LOG_DIR / "kd_run.log"      # журнал запусков; только дописывается, не перезаписывается

# Корень архива на диске (нужен для оценки длины путей)
ARCHIVE_ROOT = r"\\KometaNAS\Производство\Номенклатура\Готовая номенклатура"

# Работаем только с этой корневой группой УНФ; остальные корни (Услуги, Материалы и т.д.)
# к архиву КД не относятся.
ROOT_GROUP_NAME = "Товары"

# --- Пороги проверок названий ---
MAX_FOLDER_NAME = 100     # длиннее — предупреждение о длинном имени папки
MAX_ITEM_NAME = 150       # длиннее — предупреждение о длинном названии изделия
MAX_FOLDER_PATH = 200     # путь папки вместе с корнем архива; длиннее — предупреждение
MAX_FILE_PATH = 259       # полный путь файла; классический предел Windows (MAX_PATH = 260 с завершающим символом)
FILE_NAME_OVERHEAD = 18   # «НФ-00000000 » + расширение (например, «.7z»)

# Темы для раздела «Известные случаи» (названия групп, проверяются на любое написание
# латиницей/кириллицей). Добавляйте свои.
DEFAULT_PROBES = [
    "Со светом",
    "На уголке 40х40 AISI 430",
    "Полка-решетка",
    "Столешница 0.8 AISI 430 (базовый)",
]


def log(*args) -> None:
    """Печатает сообщение и дописывает его в журнал запусков. Не падает без консоли."""
    text = " ".join(str(a) for a in args)
    if sys.stdout is not None:
        try:
            print(text)
        except Exception:
            pass
    try:
        RUN_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(RUN_LOG, "a", encoding="utf-8-sig") as f:
            f.write(f"[{datetime.now().isoformat(timespec='seconds')}] {text}\n")
    except Exception:
        pass

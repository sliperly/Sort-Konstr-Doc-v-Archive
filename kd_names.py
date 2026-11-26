"""
Проверки названий папок и изделий.

Главное правило: нормализация (замена латиницы на кириллицу и т.п.) используется
ТОЛЬКО для поиска похожих названий и подсказок. Сами названия модуль никогда не меняет.
"""

import difflib
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass

from kd_common import MAX_FOLDER_NAME, MAX_ITEM_NAME

ERR, WARN = "ошибка", "предупреждение"

# Символы, запрещённые в имени файла или папки Windows. «/» тоже: в пути он разделитель.
WIN_ILLEGAL = set('<>:"\\|?*/')

# Латинские буквы, неотличимые от кириллических (строгие двойники).
LAT_TO_CYR = {
    "A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М", "O": "О",
    "P": "Р", "T": "Т", "X": "Х", "Y": "У",
    "a": "а", "c": "с", "e": "е", "o": "о", "p": "р", "x": "х", "y": "у",
}
CYR_TO_LAT = {v: k for k, v in LAT_TO_CYR.items()}

# Более широкая таблица только для поиска «визуально одинаковых» названий.
_KEY_LOOKALIKES = str.maketrans({
    "a": "а", "b": "в", "c": "с", "e": "е", "h": "н", "k": "к", "m": "м",
    "o": "о", "p": "р", "t": "т", "x": "х", "y": "у",
})

# Допустимые, но не ASCII и не кириллические символы (не считаются «нестандартными»).
EXTRA_ALLOWED = set("«»№°×±²³Ø⌀–—’‘“”″′…·•µ€")

_VISIBLE = {" ": "␣", "\u00a0": "⍽", "\n": "↵", "\r": "↵", "\t": "⇥"}


@dataclass
class Problem:
    kind: str
    severity: str
    detail: str
    marked: str = ""          # название с выделенными ⟦символами⟧
    suggestion: str = ""


# ---------------------------------------------------------------------------
# Вспомогательное
# ---------------------------------------------------------------------------

def script_of(ch: str):
    if "A" <= ch <= "Z" or "a" <= ch <= "z":
        return "lat"
    if "\u0400" <= ch <= "\u04ff":
        return "cyr"
    return None


def show(ch: str) -> str:
    """Символ с Unicode-кодом, например «C (U+0043)»."""
    shown = _VISIBLE.get(ch, ch if ord(ch) >= 32 else f"\\x{ord(ch):02x}")
    return f"{shown} (U+{ord(ch):04X})"


def mark(name: str, indexes) -> str:
    """Оборачивает выбранные символы в ⟦ ⟧, пробелы показывает видимыми значками."""
    chosen = set(indexes)
    out, opened = [], False
    for i, ch in enumerate(name):
        if i in chosen:
            if not opened:
                out.append("⟦")
                opened = True
            out.append(_VISIBLE.get(ch, ch))
        else:
            if opened:
                out.append("⟧")
                opened = False
            out.append(ch)
    if opened:
        out.append("⟧")
    return "".join(out)


def win_key(name: str) -> str:
    """Ключ «одинаково для Windows»: без регистра, краевых пробелов и точек, с одним пробелом."""
    s = unicodedata.normalize("NFC", name)
    return re.sub(r"\s+", " ", s).strip(" .").casefold()


def visual_key(name: str) -> str:
    """Ключ «выглядит одинаково»: ещё и ё=е, × = х, латинские двойники = кириллица."""
    s = unicodedata.normalize("NFC", name).casefold().replace("ё", "е").replace("×", "х")
    s = s.translate(_KEY_LOOKALIKES)
    return re.sub(r"\s+", " ", s).strip(" .")


def diff_chars(a: str, b: str, limit: int = 4) -> str:
    """Какие именно символы отличаются между двумя названиями (с Unicode-кодами)."""
    parts = []
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        left = "".join(show(c) for c in a[i1:i2]) or "—"
        right = "".join(show(c) for c in b[j1:j2]) or "—"
        parts.append(f"{left} ↔ {right}")
        if len(parts) >= limit:
            parts.append("…")
            break
    return "; ".join(parts) if parts else "различий в символах нет"


def _is_allowed(ch: str) -> bool:
    o = ord(ch)
    if 32 <= o < 127:
        return True
    if "\u0410" <= ch <= "\u044f" or ch in "Ёё":
        return True
    return ch in EXTRA_ALLOWED


# ---------------------------------------------------------------------------
# Проверка одного названия
# ---------------------------------------------------------------------------

def check_name(name: str, kind: str) -> list:
    """
    Проверяет название. kind: «папка» (имя папки на диске) или «изделие» (из названия
    будет формироваться имя файла). Возвращает список Problem.
    """
    out: list = []
    stripped = name.strip()
    if not stripped:
        return [Problem("пустое название", ERR, "название пустое", repr(name))]

    target = "папку" if kind == "папка" else "файл"

    # 1. Запрещённые символы
    bad = [(i, ch) for i, ch in enumerate(name) if ch in WIN_ILLEGAL or ord(ch) < 32]
    if bad:
        chars = []
        for _, ch in bad:
            if ch not in chars:
                chars.append(ch)
        note = ""
        if "/" in chars:
            note = " («/» в УНФ воспринимается как часть названия, а в пути это разделитель)"
        out.append(Problem(
            "запрещённые символы", ERR,
            f"{target} с таким именем в Windows создать нельзя: " + ", ".join(show(c) for c in chars) + note,
            mark(name, [i for i, _ in bad])))

    # 2. Пробелы по краям
    lead = len(name) - len(name.lstrip())
    trail = len(name) - len(name.rstrip())
    if lead or trail:
        out.append(Problem(
            "пробелы по краям", ERR if kind == "папка" else WARN,
            f"в начале: {lead}, в конце: {trail}"
            + ("; Windows обрезает такие пробелы, папка получит другое имя" if kind == "папка" else ""),
            mark(name, list(range(lead)) + list(range(len(name) - trail, len(name))))))

    # 3. Точка в конце (для папки)
    if kind == "папка" and stripped.endswith("."):
        out.append(Problem("точка в конце", ERR, "Windows не допускает имя папки, оканчивающееся точкой",
                           mark(name, [len(name.rstrip()) - 1])))

    # 4. Несколько пробелов подряд
    multi = [m for m in re.finditer(r"(?<=\S)[ \u00a0]{2,}(?=\S)", name)]
    if multi:
        idx = [i for m in multi for i in range(m.start(), m.end())]
        out.append(Problem("несколько пробелов подряд", WARN,
                           f"мест: {len(multi)}; такие названия трудно сравнивать на глаз", mark(name, idx)))

    # 5. Необычные пробелы
    odd_space = [i for i, ch in enumerate(name)
                 if ch != " " and unicodedata.category(ch) == "Zs"]
    if odd_space:
        chars = sorted({name[i] for i in odd_space})
        out.append(Problem("необычный пробел", WARN,
                           "вместо обычного пробела: " + ", ".join(show(c) for c in chars)
                           + "; визуально неотличим от обычного, сравнение строк не совпадёт",
                           mark(name, odd_space)))

    # 6. Нестандартные символы
    handled = set(WIN_ILLEGAL)
    weird = [(i, ch) for i, ch in enumerate(name)
             if ch not in handled and ord(ch) >= 32 and not ch.isspace() and not _is_allowed(ch)]
    if weird:
        chars = []
        for _, ch in weird:
            if ch not in chars:
                chars.append(ch)
        descr = []
        for ch in chars:
            nm = unicodedata.name(ch, "?")
            hint = ""
            if unicodedata.category(ch).startswith("M"):
                hint = " — комбинируемый знак: буква, скорее всего, набрана «в разобранном виде» (например, «й»)"
            descr.append(f"{show(ch)} {nm}{hint}")
        out.append(Problem("нестандартный символ", WARN, "; ".join(descr), mark(name, [i for i, _ in weird])))

    # 7. Латиница и кириллица в одном слове
    for m in re.finditer(r"[^\W\d_]+", name):
        token = m.group(0)
        lat = [i for i, ch in enumerate(token) if script_of(ch) == "lat"]
        cyr = [i for i, ch in enumerate(token) if script_of(ch) == "cyr"]
        if not lat or not cyr:
            continue
        if len(cyr) > len(lat):
            minority, to_map, label = lat, LAT_TO_CYR, "латинские буквы в кириллическом слове"
        elif len(lat) > len(cyr):
            minority, to_map, label = cyr, CYR_TO_LAT, "кириллические буквы в латинском слове"
        else:
            minority, to_map, label = None, None, "слово написано и латиницей, и кириллицей поровну"
        idx = [m.start() + i for i in (minority if minority is not None else lat + cyr)]
        if minority is not None:
            chars = [token[i] for i in minority]
            codes = ", ".join(show(c) for c in dict.fromkeys(chars))
            fixed = list(token)
            ok = True
            for i in minority:
                repl = to_map.get(token[i])
                if repl is None:
                    ok = False
                    break
                fixed[i] = repl
            suggestion = "".join(fixed) if ok else ""
            detail = f"{label}: {codes} в слове «{token}»"
            if not ok:
                detail += "; для части букв нет визуального двойника"
        else:
            detail = f"{label}: «{token}»"
            suggestion = ""
            if all(c in LAT_TO_CYR or c in CYR_TO_LAT for c in token):
                cyr_variant = "".join(LAT_TO_CYR.get(c, c) for c in token)
                lat_variant = "".join(CYR_TO_LAT.get(c, c) for c in token)
                suggestion = f"«{cyr_variant}» (кириллица) или «{lat_variant}» (латиница); выбрать по смыслу"
        out.append(Problem("смесь латиницы и кириллицы в слове", WARN, detail, mark(name, idx), suggestion))

    # 8. Латинская x и знак × между цифрами в размерах
    for m in re.finditer(r"(?<=\d)\s?([xX])\s?(?=\d)", name):
        i = m.start(1)
        out.append(Problem(
            "латинская x между цифрами", WARN,
            f"{show(name[i])} в размере; в названиях принята кириллическая «х» (U+0445). "
            "Если это артикул поставщика, исправлять не нужно",
            mark(name, [i]), "заменить на «х» (U+0445)"))
    for m in re.finditer(r"(?<=\d)\s?(×)\s?(?=\d)", name):
        i = m.start(1)
        out.append(Problem("знак × между цифрами", WARN,
                           f"{show('×')} вместо кириллической «х» (U+0445) в размере",
                           mark(name, [i]), "заменить на «х» (U+0445)"))

    # 9. Длина
    limit = MAX_FOLDER_NAME if kind == "папка" else MAX_ITEM_NAME
    if len(stripped) > limit:
        out.append(Problem("слишком длинное название", WARN,
                           f"{len(stripped)} символов при пороге {limit}", ""))

    return out


# ---------------------------------------------------------------------------
# Сравнение названий между собой
# ---------------------------------------------------------------------------

@dataclass
class Cluster:
    kind: str       # «одинаковые для Windows» | «похожие»
    members: list   # [(код, название)]


def cluster_by_visual(named) -> list:
    """
    Группирует названия, которые выглядят одинаково. named — список (код, название).
    «одинаковые для Windows» — получатся одна папка; «похожие» — разные папки с почти
    неотличимыми именами (отличаются ё/е, латиницей, пробелами, регистром у ×/х и т.п.).
    """
    buckets = defaultdict(list)
    for code, name in named:
        key = visual_key(name)
        if key:
            buckets[key].append((code, name))
    out = []
    for members in buckets.values():
        if len(members) < 2:
            continue
        same = len({win_key(n) for _, n in members}) == 1
        out.append(Cluster("одинаковые для Windows" if same else "похожие", members))
    return out

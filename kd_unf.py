"""
Чтение выгрузки универсального отчёта УНФ (справочник «Номенклатура») и
построение дерева групп и изделий.

Иерархия строится ТОЛЬКО по кодам: в столбце A записано «Название, КодРодителя»,
собственный код — в столбце «Код». Уровни группировки Excel и названия групп для
построения дерева не используются: названия в разных ветках повторяются
(например, «Со светом»), уровни группировки в прайс-листе противоречивы.

Модуль ничего не пишет на диск и ничего не знает о папках архива.
"""

import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl

HEADER_NAME = "Номенклатура, Код"
HEADER_GROUP = "Это группа"
HEADER_CODE = "Код"
HEADER_INVALID = "Недействительна"
HEADER_NO_PRICE = "Не использовать в прайс-листах"

CODE_RE = re.compile(r"^НФ-\d{8}$")
# «Название, КодРодителя»; у корневых записей после запятой пусто. Название может
# содержать запятые и переносы строк, поэтому берётся всё до последней запятой.
NAME_PARENT_RE = re.compile(
    r"^(?P<name>.*),[ \t\u00a0]*(?P<parent>НФ-\d{8})?[ \t\u00a0]*$", re.S)

ERR, WARN, INFO = "ошибка", "предупреждение", "инфо"


class UnfFormatError(Exception):
    """Файл не похож на ожидаемую выгрузку универсального отчёта."""


@dataclass
class Node:
    code: str
    name: str            # как записано в УНФ, без обрезки пробелов
    parent: str          # код родителя или None
    is_group: bool
    invalid: bool        # «Недействительна»
    no_price: bool       # «Не использовать в прайс-листах»
    row: int             # номер строки в Excel


@dataclass
class Issue:
    severity: str
    kind: str
    code: str
    row: int
    detail: str


def _yes(value) -> bool:
    return value is not None and str(value).strip().lower() == "да"


def _cell(cells, index):
    return cells[index] if index is not None and index < len(cells) else None


class Tree:
    def __init__(self, source: Path = None):
        self.source = source
        self.nodes: dict = {}                 # код -> Node, в порядке строк файла
        self.children = defaultdict(list)     # код группы -> коды детей, в порядке строк
        self.issues: list = []
        self.cycle_nodes: set = set()
        self.header_row = None
        self.rows_read = 0

    # --- служебное ---
    def add_issue(self, severity, kind, code, row, detail):
        self.issues.append(Issue(severity, kind, code or "", row or 0, detail))

    def title(self, code: str) -> str:
        return self.nodes[code].name.strip()

    # --- проверка целостности после чтения ---
    def finalize(self) -> None:
        for code, node in self.nodes.items():
            if node.parent is None:
                continue
            parent = self.nodes.get(node.parent)
            if parent is None:
                self.add_issue(ERR, "родитель не найден", code, node.row,
                               f"код родителя {node.parent} отсутствует в выгрузке")
            elif not parent.is_group:
                self.add_issue(ERR, "родитель не является группой", code, node.row,
                               f"родитель {node.parent} помечен как не группа")
            else:
                self.children[node.parent].append(code)
        self._find_cycles()

    def _find_cycles(self) -> None:
        state: dict = {}   # 1 — в обходе, 2 — проверено
        for start in self.nodes:
            if state.get(start) == 2:
                continue
            path, cur = [], start
            while cur and cur in self.nodes and state.get(cur) is None:
                state[cur] = 1
                path.append(cur)
                cur = self.nodes[cur].parent
            if cur and state.get(cur) == 1:
                members = path[path.index(cur):]
                self.cycle_nodes.update(members)
                self.add_issue(ERR, "цикл в иерархии", cur, self.nodes[cur].row,
                               "группы ссылаются друг на друга по кругу: " + " → ".join(members))
            for c in path:
                state[c] = 2

    # --- обход ---
    def roots(self) -> list:
        """Группы без родителя."""
        return [n for n in self.nodes.values() if n.parent is None and n.is_group]

    def orphan_items(self) -> list:
        """Изделия без группы."""
        return [n for n in self.nodes.values() if n.parent is None and not n.is_group]

    def walk(self, root_code: str):
        """
        Обходит поддерево корня сверху вниз и выдаёт (узел, путь), где путь — список названий
        групп ниже корня: для группы — включая её саму, для изделия — группа, в которой оно лежит.
        """
        seen = set()
        stack = [(root_code, [])]
        while stack:
            code, parts = stack.pop()
            if code in seen:
                continue
            seen.add(code)
            node = self.nodes[code]
            yield node, parts
            if node.is_group:
                for child in reversed(self.children.get(code, [])):
                    cn = self.nodes[child]
                    stack.append((child, parts + [cn.name.strip()] if cn.is_group else parts))


def read_unf(path: Path) -> Tree:
    """Читает выгрузку и возвращает дерево. Не изменяет файл."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        tree = Tree(source=path)
        cols: dict = {}

        for row_no, row in enumerate(ws.iter_rows(values_only=True), start=1):
            cells = list(row)

            if tree.header_row is None:
                first = cells[0] if cells else None
                if first is not None and str(first).strip() == HEADER_NAME:
                    tree.header_row = row_no
                    cols = {str(c).strip(): i for i, c in enumerate(cells) if c is not None}
                    missing = [h for h in (HEADER_GROUP, HEADER_CODE) if h not in cols]
                    if missing:
                        raise UnfFormatError(f"в шапке нет столбцов: {', '.join(missing)}")
                continue

            raw = _cell(cells, 0)
            code_value = _cell(cells, cols[HEADER_CODE])
            if raw is None and code_value is None:
                continue
            tree.rows_read += 1

            code = str(code_value).strip() if code_value is not None else ""
            raw_text = "" if raw is None else str(raw)
            match = NAME_PARENT_RE.match(raw_text)
            if match:
                name, parent = match.group("name"), match.group("parent")
            else:
                name, parent = raw_text, None
                tree.add_issue(WARN, "не выделен код родителя", code, row_no,
                               f"в ячейке нет запятой перед кодом родителя: «{raw_text[:60]}»")

            if not code:
                tree.add_issue(ERR, "нет кода", "", row_no, f"строка без кода: «{raw_text[:60]}»")
                continue
            if not CODE_RE.match(code):
                tree.add_issue(WARN, "код не по формату НФ-XXXXXXXX", code, row_no,
                               f"код «{code}» не соответствует формату")
            if code in tree.nodes:
                tree.add_issue(ERR, "дубль кода", code, row_no,
                               f"код уже встречался в строке {tree.nodes[code].row}; эта строка пропущена")
                continue

            tree.nodes[code] = Node(
                code=code, name=name, parent=parent,
                is_group=_yes(_cell(cells, cols[HEADER_GROUP])),
                invalid=_yes(_cell(cells, cols.get(HEADER_INVALID))),
                no_price=_yes(_cell(cells, cols.get(HEADER_NO_PRICE))),
                row=row_no,
            )

        if tree.header_row is None:
            raise UnfFormatError(f"не найдена шапка «{HEADER_NAME}» — это не выгрузка универсального отчёта")
        if not tree.nodes:
            raise UnfFormatError("в выгрузке нет ни одной записи")
    finally:
        wb.close()

    tree.finalize()
    return tree

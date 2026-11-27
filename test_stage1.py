"""
Тесты этапа 1. Запуск:  python test_stage1.py
Ничего не читают с вашего диска: выгрузка создаётся во временной папке.
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl

from kd_names import check_name, cluster_by_visual
from kd_stage1 import build_scope, collect_name_problems, group_metrics, invalid_in_goods
from kd_unf import UnfFormatError, read_unf

HEADER = ["Номенклатура, Код", None, None, "Не использовать в прайс-листах", None, "Недействительна", "Это группа", "Код"]


def make_report(rows, with_header=True):
    """rows: (название, код родителя, группа?, код, недействительна?)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Параметры: Тип объекта: Справочник"])   # шапка отчёта, как в настоящей выгрузке
    ws.append([])
    if with_header:
        ws.append(HEADER)
    for name, parent, is_group, code, invalid in rows:
        text = f"{name}, {parent}" if parent else f"{name}, "
        ws.append([text, None, None, None if is_group else "Нет", None,
                   None if is_group else ("Да" if invalid else "Нет"), "Да" if is_group else "Нет", code])
    tmp = Path(tempfile.mkdtemp()) / "unf.xlsx"
    wb.save(tmp)
    return tmp


BASE = [
    ("Товары", None, True, "НФ-00000001", False),
    ("Услуги", None, True, "НФ-00000002", False),
    ("Мойки", "НФ-00000001", True, "НФ-00000010", False),
    ("На уголке 40х40", "НФ-00000010", True, "НФ-00000011", False),
    ("Со светом", "НФ-00000011", True, "НФ-00000012", False),
    ("Столы", "НФ-00000001", True, "НФ-00000020", False),
    ("Со светом", "НФ-00000020", True, "НФ-00000021", False),      # то же имя в другой ветке
    ("Мойка, большая, 600х600", "НФ-00000012", False, "НФ-00000100", False),
    ("Стол светлый", "НФ-00000021", False, "НФ-00000101", False),
    ("Стол тёмный", "НФ-00000021", False, "НФ-00000102", True),        # недействительное изделие
    ("Устаревшее", "НФ-00000001", True, "НФ-00000030", False),
    ("Старьё", "НФ-00000030", False, "НФ-00000103", True),             # единственное изделие группы недействительно
    ("Работа сварщика", "НФ-00000002", False, "НФ-00000200", False),   # вне «Товары»
    ("Изделие без группы", None, False, "НФ-00000300", False),
]


class TreeTests(unittest.TestCase):
    def test_parent_codes_and_commas(self):
        tree = read_unf(make_report(BASE))
        self.assertEqual(tree.nodes["НФ-00000100"].name, "Мойка, большая, 600х600")
        self.assertEqual(tree.nodes["НФ-00000100"].parent, "НФ-00000012")
        self.assertIsNone(tree.nodes["НФ-00000001"].parent)
        self.assertEqual(tree.issues, [])

    def test_same_group_name_in_different_branches(self):
        tree = read_unf(make_report(BASE))
        _roots, scope = build_scope(tree)
        self.assertEqual(scope["НФ-00000012"][1], ["Мойки", "На уголке 40х40", "Со светом"])
        self.assertEqual(scope["НФ-00000021"][1], ["Столы", "Со светом"])
        # изделие лежит в той «Со светом», на которую указывает код родителя
        self.assertEqual(scope["НФ-00000100"][1], ["Мойки", "На уголке 40х40", "Со светом"])
        self.assertEqual(scope["НФ-00000101"][1], ["Столы", "Со светом"])

    def test_scope_is_only_goods(self):
        tree = read_unf(make_report(BASE))
        roots, scope = build_scope(tree)
        self.assertEqual(roots, {"НФ-00000001"})
        self.assertNotIn("НФ-00000200", scope)   # «Услуги»
        self.assertNotIn("НФ-00000300", scope)   # без группы
        self.assertEqual(len(tree.orphan_items()), 1)

    def test_counts(self):
        tree = read_unf(make_report(BASE))
        _roots, scope = build_scope(tree)
        m = group_metrics(tree, scope)
        self.assertEqual(m["direct"]["НФ-00000012"], 1)
        self.assertEqual(m["total"]["НФ-00000010"], 1)   # через две вложенные группы
        self.assertEqual(m["total"]["НФ-00000001"], 2)

    def test_invalid_items_are_ignored(self):
        tree = read_unf(make_report(BASE))
        roots, scope = build_scope(tree)
        self.assertIn("НФ-00000101", scope)
        self.assertNotIn("НФ-00000102", scope)
        self.assertNotIn("НФ-00000103", scope)
        self.assertEqual(invalid_in_goods(tree, roots), 2)
        m = group_metrics(tree, scope)
        self.assertEqual(m["total"]["НФ-00000021"], 1)    # недействительный «Стол тёмный» не считается
        self.assertEqual(m["total"]["НФ-00000030"], 0)    # в группе остались только недействительные
        self.assertIn("НФ-00000030", scope)               # сама группа остаётся в дереве

    def test_invalid_items_are_not_checked_for_names(self):
        rows = BASE + [
            ("Ванна 400x400", "НФ-00000021", False, "НФ-00000110", True),    # недействительная, латинская x
            ("Ванна 500x500", "НФ-00000021", False, "НФ-00000111", False),   # действующая, латинская x
        ]
        tree = read_unf(make_report(rows))
        roots, scope = build_scope(tree)
        problems, _stats = collect_name_problems(tree, scope, roots)
        codes = {r[3] for r in problems}
        self.assertIn("НФ-00000111", codes)
        self.assertNotIn("НФ-00000110", codes)

    def test_missing_parent_duplicate_and_item_as_parent(self):
        rows = BASE + [
            ("Сирота", "НФ-00009999", False, "НФ-00000400", False),                 # родителя нет
            ("Дубль", "НФ-00000001", False, "НФ-00000100", False),                  # код уже был
            ("Ребёнок изделия", "НФ-00000100", False, "НФ-00000401", False),       # родитель — не группа
        ]
        kinds = {i.kind for i in read_unf(make_report(rows)).issues}
        self.assertIn("родитель не найден", kinds)
        self.assertIn("дубль кода", kinds)
        self.assertIn("родитель не является группой", kinds)

    def test_cycle(self):
        rows = BASE + [
            ("Цикл A", "НФ-00000501", True, "НФ-00000500", False),
            ("Цикл B", "НФ-00000500", True, "НФ-00000501", False),
        ]
        tree = read_unf(make_report(rows))
        self.assertIn("цикл в иерархии", {i.kind for i in tree.issues})
        self.assertEqual(tree.cycle_nodes, {"НФ-00000500", "НФ-00000501"})
        _roots, scope = build_scope(tree)      # обход не должен зависнуть
        self.assertNotIn("НФ-00000500", scope)

    def test_not_a_report(self):
        with self.assertRaises(UnfFormatError):
            read_unf(make_report(BASE, with_header=False))


class NameTests(unittest.TestCase):
    def kinds(self, name, obj="папка"):
        return {p.kind for p in check_name(name, obj)}

    def test_latin_letter_inside_cyrillic_word(self):
        problems = check_name("1. C" + "толы", "папка")      # C — латинская
        mixed = [p for p in problems if p.kind == "смесь латиницы и кириллицы в слове"]
        self.assertEqual(len(mixed), 1)
        self.assertIn("U+0043", mixed[0].detail)
        self.assertEqual(mixed[0].suggestion, "Столы")        # С — кириллическая
        self.assertIn("⟦C⟧", mixed[0].marked)

    def test_clean_name_has_no_problems(self):
        self.assertEqual(check_name("Ванна моечная С-КТ430 400х400х250", "изделие"), [])
        self.assertEqual(check_name("Столы AISI 304", "папка"), [])

    def test_latin_x_between_digits(self):
        self.assertIn("латинская x между цифрами", self.kinds("Ванна 400x400", "изделие"))
        self.assertNotIn("латинская x между цифрами", self.kinds("Ванна 400х400", "изделие"))  # кириллическая х
        self.assertIn("знак × между цифрами", self.kinds("Ванна 400×400", "изделие"))

    def test_illegal_and_edge_spaces(self):
        self.assertIn("запрещённые символы", self.kinds("Электрика / сантехника"))
        self.assertIn("запрещённые символы", self.kinds('Столы "Пицца-бар"'))
        self.assertIn("пробелы по краям", self.kinds("Архив "))
        self.assertIn("точка в конце", self.kinds("Прочее."))
        self.assertIn("пробелы по краям", self.kinds(" Прочее"))

    def test_spaces_and_odd_chars(self):
        self.assertIn("несколько пробелов подряд", self.kinds("Двери  распашные"))
        self.assertIn("необычный пробел", self.kinds("Двери\u00a0распашные"))
        self.assertIn("нестандартный символ", self.kinds("Двери\u0306 распашные"))   # комбинируемый знак

    def test_tie_between_scripts_gets_both_variants(self):
        problems = [p for p in check_name("Ванна C" + "М", "изделие")        # C латинская + М кириллическая
                    if p.kind == "смесь латиницы и кириллицы в слове"]
        self.assertEqual(len(problems), 1)
        self.assertIn("кириллица", problems[0].suggestion)
        self.assertIn("латиница", problems[0].suggestion)

    def test_article_with_latin_only_is_not_mixed(self):
        self.assertNotIn("смесь латиницы и кириллицы в слове", self.kinds("Профиль ET01MKT75X25F", "изделие"))

    def test_clusters(self):
        windows_same = cluster_by_visual([("1", "Архив"), ("2", "Архив"), ("3", "Прочее")])
        self.assertEqual(len(windows_same), 1)
        self.assertEqual(windows_same[0].kind, "одинаковые для Windows")
        similar = cluster_by_visual([("1", "Полка-решетка"), ("2", "Полка-решётка")])
        self.assertEqual(len(similar), 1)
        self.assertEqual(similar[0].kind, "похожие")
        latin = cluster_by_visual([("1", "Столы"), ("2", "C" + "толы")])      # латинская C
        self.assertEqual(latin[0].kind, "похожие")
        self.assertEqual(cluster_by_visual([("1", "Столы"), ("2", "Стулья")]), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)

import re
import unittest

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

import build
import test_build as fixtures


class CategoriesTests(unittest.TestCase):
    setUp = fixtures.BuilderTests.setUp
    tearDown = fixtures.BuilderTests.tearDown
    edit = fixtures.BuilderTests.edit

    def test_new_category_gets_navigation_and_same_card_template(self):
        def update(book):
            book['Разделы'].append(['Новый раздел'])
            book['Меню'].append(['Новый раздел', 'Новое блюдо', '250 г', 350,
                                  'Текст блюда', 'test.jpg', 'Да', ''])
        self.edit(update)
        report = build.build(self.root, self.out)
        self.assertEqual(report['sections'], 2)
        self.assertEqual(report['cards'], 2)
        page = (self.out / 'index.html').read_text()
        self.assertEqual(page.count('class="dish-card"'), 2)
        self.assertIn('Новый раздел', page)
        self.assertIn('data-category-target="category-c-', page)
        self.assertIn('Новое блюдо', page)
        self.assertIn('350 ₽', page)
        ids = re.findall(r'\bid="([^"]+)"', page)
        self.assertEqual(len(ids), len(set(ids)))

    def test_empty_category_waits_for_a_visible_dish(self):
        self.edit(lambda w:w['Разделы'].append(['Пока пустой раздел']))
        report = build.build(self.root, self.out)
        self.assertEqual(report['sections'], 1)
        self.assertNotIn('Пока пустой раздел', (self.out / 'index.html').read_text())

    def test_released_dropdown_uses_an_expandable_excel_table(self):
        book = load_workbook(build.ROOT / 'menu.xlsx')
        try:
            sections = book['Разделы']
            self.assertIn('MenuCategories', sections.tables)
            table = sections.tables['MenuCategories']
            self.assertEqual(book.defined_names['MenuSections'].attr_text,
                             'MenuCategories[Раздел]')
            self.assertEqual(table.tableColumns[0].name, 'Раздел')
            first_col, first_row, last_col, last_row = range_boundaries(table.ref)
            self.assertEqual((first_col, first_row, last_col), (1, 1, 1))
            last_used = max(r for r in range(2, sections.max_row + 1) if sections.cell(r, 1).value)
            self.assertGreaterEqual(last_row, last_used)
            self.assertTrue(any(rule.formula1 == 'MenuSections'
                                for rule in book['Меню'].data_validations.dataValidation))
        finally:
            book.close()


if __name__ == '__main__':
    unittest.main()

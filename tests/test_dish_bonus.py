import unittest

import build
import test_build as fixtures


class DishBonusTests(unittest.TestCase):
    setUp = fixtures.BuilderTests.setUp
    tearDown = fixtures.BuilderTests.tearDown
    edit = fixtures.BuilderTests.edit

    def test_bonus_belongs_only_to_selected_dishes_and_is_escaped(self):
        bonus = 'Салат <Свежий>\nКомпот — 1 л\nЛепёшки — 2 шт'

        def update(w):
            menu = w['Меню']
            menu['I1'] = 'Бонус к блюду'
            menu['I2'] = bonus
            menu.append(['Завтраки', 'Второе блюдо', '1 шт', 100, '', '', 'Да', 'second', bonus])
            menu.append(['Завтраки', 'Без бонуса', '1 шт', 150, '', '', 'Да', 'third', ''])
            w['Настройки'].append(['Бонус к блюду', 'Устаревший общий бонус'])
            w['Настройки'].append(['Условия бонуса', 'Условия бонуса уточняются.'])

        self.edit(update)
        model = build.read_menu(self.root / 'menu.xlsx', self.root)
        self.assertEqual([d['id'] for d in model['dishes'] if d['bonus']], ['dish-test', 'second'])
        build.build(self.root, self.out)
        page = (self.out / 'index.html').read_text()
        self.assertEqual(page.count('class="dish-bonus"'), 2)
        self.assertIn('Салат &lt;Свежий&gt;', page)
        self.assertNotIn('Устаревший общий бонус', page)
        self.assertNotIn('Условия бонуса уточняются.', page)
        self.assertNotIn('class="bonus"', page)

    def test_bonus_cannot_be_attached_to_an_unnamed_portion_row(self):
        self.edit(lambda w: (setattr(w['Меню']['I1'], 'value', 'Бонус к блюду'),
                             setattr(w['Меню']['I3'], 'value', 'Бонус')))
        with self.assertRaisesRegex(build.MenuError, 'строка 3'):
            build.read_menu(self.root / 'menu.xlsx', self.root)


if __name__ == '__main__':
    unittest.main()

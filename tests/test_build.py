import importlib.util,json,shutil,tempfile,unittest
from pathlib import Path
from decimal import Decimal
from openpyxl import Workbook,load_workbook
from PIL import Image
import sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT));import build

class BuilderTests(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)/'project';self.root.mkdir()
  (self.root/'photos').mkdir();shutil.copytree(ROOT/'assets',self.root/'assets');shutil.copytree(ROOT/'templates',self.root/'templates')
  Image.new('RGB',(1600,1000),'#bc7841').save(self.root/'photos/test.jpg')
  wb=Workbook();ws=wb.active;ws.title='Меню';ws.append(build.HEADERS);ws.append(['Завтраки','Тест <script>','150 г',220,'Описание & состав','test.jpg','Да','dish-test']);ws.append(['','','300 г',350,'','','',''])
  wb.create_sheet('Разделы').append(['Раздел']);wb['Разделы'].append(['Завтраки'])
  wb.create_sheet('Настройки').append(['Параметр','Значение']);wb['Настройки'].append(['Название','Кафе']);wb['Настройки'].append(['Предварительная версия','Да']);wb.create_sheet('Ссылки').append(['Название ссылки','Адрес HTTPS'])
  wb.save(self.root/'menu.xlsx');self.out=Path(self.tmp.name)/'site'
 def tearDown(self):self.tmp.cleanup()
 def edit(self,callback):
  w=load_workbook(self.root/'menu.xlsx');callback(w);w.save(self.root/'menu.xlsx');w.close()
 def test_static_build_and_photo_versions(self):
  report=build.build(self.root,self.out);self.assertEqual(report['cards'],1);self.assertEqual(report['portions'],2)
  page=(self.out/'index.html').read_text();self.assertIn('&lt;script&gt;',page);self.assertNotIn('Тест <script>',page);self.assertIn('350 ₽',page);self.assertIn('noindex',page);self.assertEqual(len(list((self.out/'photos').glob('*.webp'))),2)
  for image in (self.out/'photos').glob('*.webp'):
   with Image.open(image) as im:self.assertEqual(im.format,'WEBP');self.assertLessEqual(max(im.size),1200)
  self.assertIn('Disallow: /',(self.out/'robots.txt').read_text())
 def test_bad_price_does_not_replace_site(self):
  build.build(self.root,self.out);before=(self.out/'index.html').read_bytes();self.edit(lambda w:setattr(w['Меню']['D3'],'value','не цена'))
  with self.assertRaisesRegex(build.MenuError,'строка 3'):build.build(self.root,self.out)
  self.assertEqual((self.out/'index.html').read_bytes(),before)
 def test_no_formula_or_traversal(self):
  self.edit(lambda w:setattr(w['Меню']['D2'],'value','=100+20'))
  with self.assertRaisesRegex(build.MenuError,'формулы'):build.read_menu(self.root/'menu.xlsx',self.root)
  self.edit(lambda w:(setattr(w['Меню']['D2'],'value',220),setattr(w['Меню']['F2'],'value','../menu.xlsx')))
  with self.assertRaisesRegex(build.MenuError,'имя файла'):build.read_menu(self.root/'menu.xlsx',self.root)
 def test_unknown_category_and_orphan_portion(self):
  self.edit(lambda w:setattr(w['Меню']['A2'],'value','Несуществующий раздел'))
  with self.assertRaisesRegex(build.MenuError,'раздел'):build.read_menu(self.root/'menu.xlsx',self.root)
  self.edit(lambda w:w['Меню'].delete_rows(2))
  with self.assertRaisesRegex(build.MenuError,'нет блюда'):build.read_menu(self.root/'menu.xlsx',self.root)
 def test_hide_and_missing_photo(self):
  self.edit(lambda w:setattr(w['Меню']['G2'],'value','Нет'))
  with self.assertRaisesRegex(build.MenuError,'нет видимых'):build.read_menu(self.root/'menu.xlsx',self.root)
  self.edit(lambda w:(setattr(w['Меню']['G2'],'value','Да'),setattr(w['Меню']['F2'],'value','missing.jpg')))
  with self.assertRaisesRegex(build.MenuError,'не найден'):build.read_menu(self.root/'menu.xlsx',self.root)
 def test_duplicate_codes_new_empty_code_and_decimal_price(self):
  self.edit(lambda w:w['Меню'].append(['Завтраки','Другое блюдо','1 шт',100,'','','Да','dish-test']))
  with self.assertRaisesRegex(build.MenuError,'повторяется код'):build.read_menu(self.root/'menu.xlsx',self.root)
  self.edit(lambda w:(setattr(w['Меню']['H4'],'value',None),setattr(w['Меню']['D4'],'value','100,50')))
  m=build.read_menu(self.root/'menu.xlsx',self.root);self.assertEqual(len(m['dishes']),2);self.assertEqual(m['dishes'][1]['portions'][0]['price'],Decimal('100.50'));self.assertTrue(m['dishes'][1]['id'].startswith('dish-'))
 def test_production_metadata_requires_valid_url(self):
  self.edit(lambda w:(w['Настройки'].append(['Предварительная версия 2','Нет']),w['Настройки'].append(['Сайт URL','javascript:bad'])))
  with self.assertRaisesRegex(build.MenuError,'HTTPS'):build.read_menu(self.root/'menu.xlsx',self.root)
  self.edit(lambda w:(setattr(w['Настройки']['B3'],'value','Нет'),setattr(w['Настройки']['B5'],'value','https://example.com/cafe')))
  # Row 3 is the original draft flag; row 5 is the URL.
  build.build(self.root,self.out);page=(self.out/'index.html').read_text();self.assertIn('rel="canonical"',page);self.assertNotIn('noindex,nofollow',page);self.assertIn('https://example.com/cafe/',(self.out/'sitemap.xml').read_text())

class HeaderContactsTests(unittest.TestCase):
 setUp=BuilderTests.setUp
 tearDown=BuilderTests.tearDown
 edit=BuilderTests.edit
 # Reuse the workbook fixture, not the base test methods.
 def test_header_contacts_and_hidden_draft_notice(self):
  self.edit(lambda w:w['Настройки'].append(['Телефон','+7 993 278-07-07']))
  build.build(self.root,self.out);page=(self.out/'index.html').read_text()
  self.assertNotIn('class="draft-notice"',page);self.assertIn('noindex,nofollow',page)
  self.assertEqual(page.count('class="contact-action phone-action"'),2)
  self.assertIn('href="tel:+79932780707"',page);self.assertIn('telegram-action is-pending" disabled',page)
  self.assertNotIn('href="https://t.me/',page)
  self.edit(lambda w:w['Настройки'].append(['Telegram URL','https://t.me/example_test_cafe']))
  build.build(self.root,self.out);page=(self.out/'index.html').read_text()
  self.assertEqual(page.count('href="https://t.me/example_test_cafe"'),2)
  self.assertNotIn('telegram-action is-pending',page)
  self.edit(lambda w:setattr(w['Настройки']['B5'],'value','javascript:bad'))
  with self.assertRaisesRegex(build.MenuError,'HTTPS'):build.read_menu(self.root/'menu.xlsx',self.root)

if __name__=='__main__':unittest.main()

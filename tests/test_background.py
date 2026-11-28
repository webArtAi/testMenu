import hashlib
import unittest
from unittest.mock import patch

from PIL import Image

import build
import test_build as fixtures


class BackgroundTests(unittest.TestCase):
    setUp = fixtures.BuilderTests.setUp
    tearDown = fixtures.BuilderTests.tearDown
    edit = fixtures.BuilderTests.edit

    def settings(self, **values):
        def update(book):
            sheet = book['Настройки']
            rows = {sheet.cell(r, 1).value:r for r in range(2, sheet.max_row + 1)}
            for key, value in values.items():
                if key in rows:
                    sheet.cell(rows[key], 2, value)
                else:
                    sheet.append([key, value])
        self.edit(update)

    def css(self):
        return (self.out / 'assets/site.css').read_text()

    def test_default_keeps_published_background_images_and_css(self):
        # Use a controlled pair: replacing the real cafe background later must
        # not make the workflow's tests depend on its old image bytes.
        for name, size in [('background.webp',(180,225)), ('background-mobile.webp',(90,112))]:
            Image.new('RGB',size,'#775533').save(self.root / 'assets' / name)
        hashes=tuple(hashlib.sha256((self.root / 'assets' / name).read_bytes()).hexdigest()
                     for name in ['background.webp','background-mobile.webp'])
        self.settings(**{'Фон':'background.webp'})
        with patch.object(build,'ORIGINAL_BACKGROUND_HASHES',hashes):
            build.build(self.root, self.out)
        for name in ['background.webp', 'background-mobile.webp']:
            self.assertEqual((self.out / 'assets' / name).read_bytes(),
                             (self.root / 'assets' / name).read_bytes())
        self.assertTrue(self.css().endswith("\nbody{background-image:url('./background.webp')}@media(max-width:640px){body{background-image:url('./background-mobile.webp')}}\n"))

    def test_solid_color_and_custom_hex_override_both_screen_backgrounds(self):
        self.settings(**{'Тип фона':'Цвет', 'Цвет фона':'Бежевый', 'Фон':'not-uploaded.jpg'})
        build.build(self.root, self.out)
        self.assertTrue(self.css().endswith('\nbody{background:#e9dfcd}@media(max-width:640px){body{background:#e9dfcd}}\n'))
        self.settings(**{'Цвет фона':'#AbC'})
        build.build(self.root, self.out)
        self.assertTrue(self.css().endswith('\nbody{background:#aabbcc}@media(max-width:640px){body{background:#aabbcc}}\n'))

    def test_gradient_uses_selected_colors_and_direction(self):
        self.settings(**{'Тип фона':'Градиент', 'Градиент: цвет 1':'Тёмно-коричневый',
                         'Градиент: цвет 2':'#336699', 'Градиент: направление':'Слева направо',
                         'Фон':'not-uploaded.jpg'})
        build.build(self.root, self.out)
        self.assertTrue(self.css().endswith('\nbody{background:linear-gradient(to right,#21160e,#336699);background-attachment:fixed}@media(max-width:640px){body{background:linear-gradient(to right,#21160e,#336699);background-attachment:fixed}}\n'))

    def test_one_jpeg_prepares_desktop_and_mobile_and_keeps_original(self):
        source = self.root / 'assets/новый фон.jpg'
        orientation = Image.Exif()
        orientation[274] = 6
        Image.new('RGB', (1200, 1800), '#bb6633').save(source, exif=orientation)
        original = source.read_bytes()
        self.settings(**{'Тип фона':'Картинка', 'Фон':source.name})
        build.build(self.root, self.out)
        digest = hashlib.sha256(original).hexdigest()[:18]
        for width, expected in [(1800, (1800, 1200)), (900, (900, 600))]:
            name = f'background-{digest}-{width}.webp'
            with Image.open(self.out / 'assets' / name) as image:
                self.assertEqual(image.format, 'WEBP')
                self.assertEqual(image.size, expected)
            self.assertIn(name, self.css())
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse((self.out / 'assets' / source.name).exists())

    def test_replacing_default_file_regenerates_mobile_from_new_image(self):
        source = self.root / 'assets/background.webp'
        Image.new('RGB', (1600, 1000), '#dd2211').save(source)
        original = source.read_bytes()
        self.settings(**{'Фон':source.name})
        build.build(self.root, self.out)
        digest = hashlib.sha256(original).hexdigest()[:18]
        with Image.open(self.out / f'assets/background-{digest}-900.webp') as image:
            self.assertEqual(image.size, (900, 562))
            red, green, blue = image.convert('RGB').getpixel((450, 280))
            self.assertGreater(red, green * 3)
            self.assertGreater(red, blue * 3)
        self.assertNotIn("background-image:url('./background-mobile.webp')", self.css().splitlines()[-1])

    def test_background_can_also_be_logo_without_removing_logo_file(self):
        Image.new('RGB', (300, 200), '#557744').save(self.root / 'assets/shared.png')
        self.settings(**{'Фон':'shared.png', 'Логотип':'shared.png'})
        build.build(self.root, self.out)
        self.assertTrue((self.out / 'assets/shared.png').is_file())
        self.assertIn('./assets/shared.png', (self.out / 'index.html').read_text())

    def test_invalid_background_never_replaces_previous_site(self):
        self.settings(**{'Тип фона':'Цвет', 'Цвет фона':'Чёрный'})
        build.build(self.root, self.out)
        original = (self.out / 'index.html').read_bytes()
        (self.root / 'assets/broken.jpg').write_bytes(b'not a picture')
        invalid = [
            {'Тип фона':'Цвет', 'Цвет фона':'red; background:url(https://example.com)'},
            {'Тип фона':'Градиент', 'Градиент: цвет 2':'not a color'},
            {'Тип фона':'Градиент', 'Градиент: цвет 2':'Белый', 'Градиент: направление':'url(bad)'},
            {'Тип фона':'Картинка', 'Фон':'missing.jpg'},
            {'Тип фона':'Картинка', 'Фон':'broken.jpg'},
            {'Тип фона':'Unexpected'},
        ]
        for values in invalid:
            with self.subTest(values=values):
                self.settings(**values)
                with self.assertRaises(build.MenuError):
                    build.build(self.root, self.out)
                self.assertEqual((self.out / 'index.html').read_bytes(), original)


if __name__ == '__main__':
    unittest.main()

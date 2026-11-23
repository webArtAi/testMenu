import contextlib
import hashlib
import io
import json
import unittest

from PIL import Image

import build
import test_build as fixtures


class PhotoFramingTests(unittest.TestCase):
    setUp = fixtures.BuilderTests.setUp
    tearDown = fixtures.BuilderTests.tearDown

    def config(self, **values):
        value = {"version": 1, "photos": {"test.jpg": values}}
        (self.root / "photo-framing.json").write_text(json.dumps(value))

    def test_individual_crop_keeps_photo_and_requests_sharper_source(self):
        photo = self.root / "photos/test.jpg"
        original = photo.read_bytes()
        self.config(source_sha256=hashlib.sha256(original).hexdigest(), zoom=1.6, x=10, y=-5)
        build.build(self.root, self.out)
        page = (self.out / "index.html").read_text()
        self.assertIn('--photo-zoom:1.6;--photo-x:10%;--photo-y:-5%', page)
        self.assertIn('calc(160vw - 44.8px)', page)
        self.assertIn('576px', page)
        self.assertEqual(photo.read_bytes(), original)
        self.assertFalse((self.out / "photo-framing.json").exists())

    def test_bad_crop_does_not_replace_previous_site(self):
        build.build(self.root, self.out)
        original = (self.out / "index.html").read_bytes()
        invalid = [
            {"zoom": 1.2, "x": 30},
            {"zoom": 0.5},
            {"zoom": float('nan')},
            {"zoom": True},
            {"source_sha256": "invalid"},
        ]
        for value in invalid:
            with self.subTest(value=value):
                self.config(**value)
                with self.assertRaises(build.MenuError):
                    build.build(self.root, self.out)
                self.assertEqual((self.out / "index.html").read_bytes(), original)

    def test_replaced_photo_does_not_inherit_an_old_crop(self):
        photo = self.root / "photos/test.jpg"
        self.config(source_sha256=hashlib.sha256(photo.read_bytes()).hexdigest(), zoom=1.6)
        Image.new('RGB', (600, 1600), '#578a6c').save(photo)
        with contextlib.redirect_stderr(io.StringIO()) as warnings:
            build.build(self.root, self.out)
        self.assertIn('обычный кадр 4:3', warnings.getvalue())
        page = (self.out / "index.html").read_text()
        self.assertNotIn('--photo-zoom:', page)
        self.assertIn('calc(100vw - 28px)', page)


if __name__ == '__main__':
    unittest.main()

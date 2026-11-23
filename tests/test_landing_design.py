"""Очистка примера лендинга, фолбек-токены и параметрический CSS (задачи 1–2)."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))

from core import landing_design

EXAMPLE = """<!DOCTYPE html><html><head><style>
body { background: #101214; color: #f5f2ee; font-family: "Lab Grotesque", Arial, sans-serif; max-width: 1200px; }
h1 { font-family: "Lab Grotesque"; font-size: 54px; font-weight: 600; }
.btn { background: #e3b7b2; color: #211b1b; border-radius: 200px; max-width: 354px; }
header { background: #121317; } footer { background: #131418; }
</style></head><body><script>alert(1)</script>
<img src="data:image/png;base64,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" onerror="hack()">
<p onmouseover="x()">Текст</p></body></html>"""


class SanitizeTest(unittest.TestCase):
    def test_removes_scripts_handlers_and_data_urls(self):
        cleaned = landing_design.sanitize_example(EXAMPLE.encode("utf-8"))
        self.assertNotIn("<script", cleaned)
        self.assertNotIn("alert(1)", cleaned)
        self.assertNotIn("onerror", cleaned)
        self.assertNotIn("onmouseover", cleaned)
        self.assertNotIn("AAAA", cleaned)          # payload data:-URL вырезан
        self.assertIn("data:image/png", cleaned)   # сама заглушка осталась
        self.assertIn("Текст", cleaned)
        self.assertLess(len(cleaned), len(EXAMPLE))

    def test_rejects_non_html(self):
        with self.assertRaises(ValueError):
            landing_design.sanitize_example("это не html".encode("utf-8"))
        with self.assertRaises(ValueError):
            landing_design.sanitize_example(b"\xff\xd8\xff\xe0\x00\x10JFIFbinary")


class TokensTest(unittest.TestCase):
    def test_from_example(self):
        tokens = landing_design.extract_tokens(EXAMPLE)
        self.assertEqual(tokens["ground"].upper(), "#101214")
        self.assertEqual(tokens["ink"].upper(), "#F5F2EE")
        self.assertIn("Lab Grotesque", tokens["font"])
        self.assertIn("Lab Grotesque", tokens["heading_font"])
        self.assertEqual(tokens["h1_size"], "54px")
        self.assertEqual(tokens["h1_weight"], "600")
        self.assertEqual(tokens["button_bg"].upper(), "#E3B7B2")
        self.assertEqual(tokens["button_radius"], "200px")
        self.assertEqual(tokens["header_bg"].upper(), "#121317")
        self.assertEqual(tokens["footer_bg"].upper(), "#131418")
        self.assertEqual(tokens["width"], "1200px")
        self.assertEqual(tokens["accent"].upper(), "#E3B7B2")

    def test_defaults_without_styles(self):
        tokens = landing_design.extract_tokens("<html><body><p>Привет</p></body></html>")
        expected = {"ground", "ink", "font", "heading_font", "h1_size", "h1_weight",
                    "accent", "button_bg", "button_color", "button_radius",
                    "header_bg", "footer_bg", "width", "radius"}
        self.assertEqual(set(tokens), expected)
        for value in tokens.values():
            self.assertTrue(str(value).strip(), f"пустой токен в {tokens}")


class FallbackCssTest(unittest.TestCase):
    def test_carries_tokens(self):
        tokens = landing_design.extract_tokens(EXAMPLE)
        css = landing_design.generate_css(tokens)
        self.assertIn("#101214", css.upper())
        self.assertIn("Lab Grotesque", css)
        self.assertIn("#E3B7B2", css.upper())
        self.assertIn(".hero", css)
        self.assertIn("@media", css)

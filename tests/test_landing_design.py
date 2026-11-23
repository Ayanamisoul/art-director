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
        self.assertEqual(landing_design.validate_profile_css(css), [])


GOOD_CSS = (":root { --ink: #111; --ground: #fff; }\n"
            ".hero { background: #f0f0f0; }\n"
            "@media (max-width: 740px) { .hero { padding: 20px; } }\n")


class ValidatorTest(unittest.TestCase):
    def test_accepts_good_css(self):
        self.assertEqual(landing_design.validate_profile_css(GOOD_CSS), [])

    def test_rejects_danger(self):
        root = ":root { --ink: #111; --ground: #fff; } "
        cases = (
            (".hero { background: #fff; }", "нет :root с --ink/--ground"),
            ("@import url('x.css');", "импорт"),
            (root + ".hero { background: url(https://evil/x.png); }", "внешний url"),
            (root + "@font-face { font-family: X; src: url(https://f.woff2); }", "внешний шрифт"),
            (root + ".hero { color: red;", "скобки"),
            (".weird-unknown { color: red; } " + root, "чужой селектор"),
            (root + ".hero { background: url('javascript:alert(1)'); }", "javascript: url"),
            ("#" * 50000, "лимит 48 КБ"),
        )
        for bad, why in cases:
            with self.subTest(why=why):
                self.assertTrue(landing_design.validate_profile_css(bad), why)

    def test_allows_data_font_face_and_known_pseudos(self):
        css = (root_ := ":root { --ink: #111; --ground: #fff; } "
               "@font-face { font-family: F; src: url(data:font/woff2;base64,AA); }"
               " .button:hover, .items li::before { color: red; }"
               " .chapter:nth-of-type(3n) { background: #eee; }"
               " .hero-copy > .lead + .button { color: #333; }")
        self.assertEqual(landing_design.validate_profile_css(css), [])


class ReadDesignTest(unittest.TestCase):
    def _chat(self, *answers):
        calls = []

        def chat(system, user):
            calls.append((system, user))
            return answers[len(calls) - 1]
        return chat, calls

    def test_retries_then_succeeds_and_strips_fences(self):
        chat, calls = self._chat(".hero {", "```css\n" + GOOD_CSS + "\n```")
        css = landing_design.read_design(EXAMPLE, chat)
        self.assertIn(":root", css)
        self.assertNotIn("```", css)
        self.assertEqual(len(calls), 2)
        self.assertIn("СКЕЛЕТ", calls[0][1].upper())   # модель получила подготовленный ввод

    def test_raises_after_two_failures(self):
        chat, _ = self._chat(".hero {", "@import url('x.css');")
        with self.assertRaises(landing_design.DesignError):
            landing_design.read_design(EXAMPLE, chat)

    def test_prepends_missing_root(self):
        chat, _ = self._chat(".hero { background: #f0f0f0; }")
        css = landing_design.read_design(EXAMPLE, chat)
        self.assertIn("--ink:", css)
        self.assertIn("--ground:", css)

    def test_model_designer_is_callable(self):
        self.assertTrue(callable(landing_design.model_designer()))


class PrepareInputTest(unittest.TestCase):
    def test_bounded_and_labelled(self):
        rules = " ".join(f".c{i} {{ color: #12345{i % 10}; }}" for i in range(1500))
        big = ('<html><head><style>' + rules + '</style></head><body>'
               + '<section class="hero"><h1 class="title">Заголовок</h1></section>'
               + '<div class="chapter">Глава</div></body></html>')
        payload = landing_design.prepare_model_input(big)
        self.assertLess(len(payload), 60000)
        self.assertIn("СТИЛИ", payload.upper())
        self.assertIn("СКЕЛЕТ", payload.upper())
        self.assertIn(".c0", payload)
        self.assertIn("hero", payload)

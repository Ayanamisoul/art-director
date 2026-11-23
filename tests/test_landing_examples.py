"""Реестр примеров лендингов: хранение, генерация CSS, кэш (задача 3)."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))

from core import landing_examples

EXAMPLE_HTML_TEXT = """<html><head><style>
body { background: #101214; color: #f5f2ee; font-family: "Lab Grotesque", Arial, sans-serif; }
.btn { background: #e3b7b2; color: #211b1b; border-radius: 200px; }
</style></head><body><h1>Заголовок</h1><p>Текст</p></body></html>"""
EXAMPLE_HTML = EXAMPLE_HTML_TEXT.encode("utf-8")

GOOD_CSS = ":root { --ink: #111; --ground: #fff; } .hero { background: #f0f0f0; }"


def fixed_designer(css):
    return lambda system, user: css


class RegistryTest(unittest.TestCase):
    def setUp(self):
        import tempfile
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.enterContext(patch.object(landing_examples, "EXAMPLES_DIR", self.root))


class UploadTest(RegistryTest):
    def test_saves_cleaned_html_and_generated_css(self):
        payload = EXAMPLE_HTML.replace(b"<h1>", b"<script>x()</script><h1>")
        result = landing_examples.accept_upload_bytes("vp.html", payload,
                                                      designer=fixed_designer(GOOD_CSS))
        self.assertEqual(result, {"name": "vp.html", "replaced": False,
                                  "css_source": "model", "notes": []})
        self.assertNotIn("<script", (self.root / "vp.html").read_text(encoding="utf-8"))
        self.assertIn(":root", (self.root / "vp.css").read_text(encoding="utf-8"))
        self.assertEqual(landing_examples.list_names(), ["vp.html"])
        self.assertEqual(landing_examples.newest(), "vp.html")

    def test_resolve_accepts_name_with_and_without_extension(self):
        landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML,
                                             designer=fixed_designer(GOOD_CSS))
        self.assertEqual(landing_examples.resolve("vp"), "vp.html")
        self.assertEqual(landing_examples.resolve("vp.html"), "vp.html")
        self.assertIsNone(landing_examples.resolve("нет-такого"))

    def test_falls_back_without_model(self):
        def broken(system, user):
            raise OSError("сеть недоступна")

        result = landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML,
                                                      designer=broken)
        self.assertEqual(result["css_source"], "fallback")
        self.assertTrue(result["notes"])
        css = (self.root / "vp.css").read_text(encoding="utf-8")
        self.assertIn("#101214", css.upper())      # токены примера попали в фолбек

    def test_rejects_garbage_and_oversize(self):
        with self.assertRaises(ValueError):
            landing_examples.accept_upload_bytes("x.html", b"not html at all")
        with self.assertRaises(ValueError):
            landing_examples.accept_upload_bytes("x.pdf", EXAMPLE_HTML)
        with patch.object(landing_examples, "MAX_SIZE", 100), self.assertRaises(ValueError):
            landing_examples.accept_upload_bytes("big.html", b"<html>" + b" " * 200)

    def test_builtin_profile_wins_over_same_name_example(self):
        landing_examples.accept_upload_bytes("glavbuh.html", EXAMPLE_HTML,
                                             designer=fixed_designer(GOOD_CSS))
        self.assertIsNone(landing_examples.resolve("glavbuh"))
        self.assertIsNone(landing_examples.resolve("glavbuh.html"))
        self.assertTrue(landing_examples.is_available_profile("glavbuh"))
        self.assertFalse(landing_examples.is_available_profile("glavbuh.html"))

    def test_reupload_regenerates_css(self):
        calls = []

        def designer(system, user):
            calls.append(user)
            return f":root {{ --ink: #{len(calls):03}; --ground: #fff; }} .hero {{}}" \
                   if len(calls) > 1 else GOOD_CSS

        landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML, designer=designer)
        landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML, designer=designer)
        self.assertEqual(len(calls), 2)
        self.assertIn("--ink: #002", (self.root / "vp.css").read_text(encoding="utf-8"))


class ProfileCssTest(RegistryTest):
    def test_cache_and_sidecar_fallback(self):
        landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML,
                                             designer=fixed_designer(GOOD_CSS))
        css = landing_examples.profile_css_for("vp.html")
        self.assertIn(":root", css)

        sidecar = self.root / "saved.profile.css"
        sidecar.write_text(GOOD_CSS, encoding="utf-8")
        (self.root / "vp.html").unlink()
        (self.root / "vp.css").unlink()
        self.assertIsNone(landing_examples.profile_css_for("vp.html"))
        restored = landing_examples.profile_css_for("vp.html", fallback=sidecar)
        self.assertIn(":root", restored)

    def test_available_profiles(self):
        landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML,
                                             designer=fixed_designer(GOOD_CSS))
        self.assertTrue(landing_examples.is_available_profile("vp"))
        self.assertTrue(landing_examples.is_available_profile("vp.html"))
        self.assertFalse(landing_examples.is_available_profile("ghost.html"))

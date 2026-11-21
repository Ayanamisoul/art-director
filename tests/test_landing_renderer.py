import sys
import tempfile
import unittest
from html.parser import HTMLParser
from pathlib import Path
from unittest.mock import patch

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from core import assets
from core.landings import PROFILES, render_landing


class Document(HTMLParser):
    def __init__(self, html):
        super().__init__()
        self.elements = []
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        self.elements.append((tag, dict(attrs)))

    def attributes(self, tag, name):
        return [attrs[name] for elem, attrs in self.elements if elem == tag and name in attrs]


class LandingRendererTest(unittest.TestCase):
    def test_headings_nest_under_their_module_in_every_profile(self):
        for profile in PROFILES:
            html = render_landing("Ассемблер", "", [{"title": "Память", "subsections": [
                {"title": "Адресация", "text": "Базовый адрес", "items": ["Индекс"]},
                {"title": "Чтение", "text": "Загрузка значения"}]}], profile)
            self.assertIn("<h3>Адресация</h3>", html)
            self.assertLess(html.index("<li>Индекс</li>"), html.index("<h3>Чтение</h3>"))
            self.assertIn("Загрузка значения", html)

    sections = [{"title": "План занятий", "text": "Разберите кейсы.",
                 "items": ["Подготовка", "Практика"]}]

    def test_profiles_are_deterministic_standalone_documents(self):
        self.assertEqual(set(PROFILES), {"glavbuh", "school", "law"})
        for profile in PROFILES:
            with self.subTest(profile=profile):
                args = ("Управление проектами", "Курс <для> менеджеров", self.sections, profile)
                html = render_landing(*args, cta="Посмотреть план")
                self.assertEqual(html, render_landing(*args, cta="Посмотреть план"))
                doc = Document(html)
                self.assertEqual(doc.attributes("html", "data-profile"), [profile])
                self.assertTrue(any(tag == "h1" for tag, _ in doc.elements))
                self.assertIn("Курс &lt;для&gt; менеджеров", html)
                ids = {attrs["id"] for _, attrs in doc.elements if "id" in attrs}
                self.assertTrue(all(href.startswith("#") and href[1:] in ids
                                    for href in doc.attributes("a", "href")))
                self.assertFalse(doc.attributes("link", "href"))
                self.assertFalse(doc.attributes("script", "src"))
                self.assertIn("Практика", html)
                self.assertIn("Посмотреть план", html)
                if profile == "school":
                    self.assertIn("data:font/woff2;base64,", html)

    def test_cta_without_a_destination_is_omitted(self):
        html = render_landing("Заголовок", "", [], "law", cta="Содержание")
        self.assertNotIn("Содержание", html)

    def test_inputs_are_literal_and_wrong_profile_is_reported(self):
        html = render_landing("<script>1</script>", "", [], "glavbuh", brand="<img>")
        self.assertNotIn("<script>1</script>", html)
        self.assertIn("&lt;script&gt;1&lt;/script&gt;", html)
        self.assertIn("&lt;img&gt;", html)
        with self.assertRaises(ValueError):
            render_landing("Тема", "", [], "other")

    def test_only_an_explicit_local_image_is_embedded(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(assets, "ASSETS_DIR", Path(folder)):
            Image.new("RGB", (2, 2)).save(Path(folder) / "фото.png")
            args = ("Тема", "", self.sections, "school")
            self.assertFalse(Document(render_landing(*args)).attributes("img", "src"))
            image = Document(render_landing(*args, image="фото.png")).attributes("img", "src")
            self.assertEqual(len(image), 1)
            self.assertTrue(image[0].startswith("data:image/png;base64,"))
            for invalid in ("../фото.png", "missing.png", "https://example.com/a.png"):
                with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                    render_landing(*args, image=invalid)


if __name__ == "__main__":
    unittest.main()

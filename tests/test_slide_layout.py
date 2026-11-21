import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation
from pptx.util import Inches
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from core import decks
from pptx_builder import normalize_slide


class SlideLayoutTests(unittest.TestCase):
    def test_brief_and_two_cards_share_a_page(self):
        from slide_content import paginate_slide
        slide = normalize_slide({"title": "Адресация", "bullets": ["Сравните инструкции"],
                                 "cards": [{"title": "mov", "text": "Чтение"}, {"title": "lea", "text": "Адрес"}]})
        self.assertEqual(paginate_slide(slide), [slide])

    def test_long_table_keeps_rows_and_repeats_headers(self):
        rows = [[str(i), str(i*i)] for i in range(12)]
        slides = [{"kind": "title", "title": "Практикум"}, {"title": "Результаты",
                   "table": {"headers": ["Шаг", "Квадрат"], "rows": rows}}]
        with tempfile.TemporaryDirectory() as tmp, patch.object(decks, "DATA_DIR", Path(tmp)), patch.object(decks, "render_pptx_previews"):
            self.assertEqual(decks.save_edited("steps", "Практикум", "", slides), [])
            tables = [sh.table for page in Presentation(decks.pptx_path("steps")).slides
                      for sh in page.shapes if sh.has_table]
            self.assertGreater(len(tables), 1)
            self.assertTrue(all(t.cell(0, 0).text == "Шаг" for t in tables))
            self.assertEqual([[cell.text for cell in row.cells] for t in tables for row in list(t.rows)[1:]], rows)

    def test_named_subsections_remain_grouped_when_split(self):
        from slide_content import split_for_space
        slide = normalize_slide({"title": "Память", "subsections": [
            {"title": "Адресация", "text": "Базовый адрес", "items": ["Индекс", "Масштаб"]}]})
        group = {"text": "Адресация", "sub": ["Базовый адрес", "Индекс", "Масштаб"]}
        self.assertEqual(slide["bullets"], [group])
        self.assertEqual(normalize_slide(slide), slide)
        pages = split_for_space(slide)
        self.assertEqual([p["bullets"][0]["text"] for p in pages], ["Адресация"] * 2)
        self.assertEqual([v for p in pages for v in p["bullets"][0]["sub"]], group["sub"])

    def test_normalization_retains_overlong_card_and_subpoints(self):
        prose = "Сверьте текст макета с брифом. " * 15 + "Последняя деталь."
        slide = normalize_slide({"kind": "text", "title": "Проверка",
                                 "bullets": [{"text": "Этап", "sub": [prose]}],
                                 "cards": [{"title": "Итог", "text": prose}]})
        self.assertEqual(slide["bullets"][0]["sub"], [prose])
        self.assertEqual(slide["cards"][0]["text"], prose)

    def test_long_input_splits_without_losing_words_and_reopens(self):
        prose = " ".join(f"Условие {i} определяет требования к результату."
                         for i in range(1, 41))
        slides = [{"kind": "title", "title": "Макет"},
                  {"kind": "content", "title": "Условия", "bullets": [prose]}]
        with tempfile.TemporaryDirectory() as tmp, patch.object(decks, "DATA_DIR", Path(tmp)), \
                patch.object(decks, "render_pptx_previews"):
            self.assertEqual(decks.save_edited("test", "Макет", "", slides), [])
            saved = decks.read_presentation("test")
            self.assertGreater(len(saved["slides"]), 2)
            content = [b if isinstance(b, str) else b["text"]
                       for s in saved["slides"][1:] for b in s["bullets"]]
            self.assertEqual(" ".join(content).split(), prose.split())
            self.assertEqual(decks.save_edited("test", **saved), [])
            self.assertEqual(saved["slides"], decks.read_presentation("test")["slides"])
            self.assertEqual(len(Presentation(decks.pptx_path("test")).slides), len(saved["slides"]))

    def test_font_is_larger_in_a_spacious_box(self):
        from slide_layout import fit_body
        rows = ["Обсудите результат и зафиксируйте критерии готовности."] * 3
        self.assertGreater(fit_body(rows, 800, 300), fit_body(rows, 315, 200))
        with self.assertRaises(ValueError):
            fit_body(rows, 70, 28)

    def test_template_field_geometry_can_require_an_extra_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            source = Presentation()
            source.slides.add_slide(source.slide_layouts[0])
            body = source.slides.add_slide(source.slide_layouts[1]).placeholders[1]
            body.width, body.height = Inches(7), Inches(.62)
            template = directory / "source.pptx"
            source.save(template)
            slides = [{"kind": "title", "title": "Программа"},
                      {"kind": "content", "title": "Этапы", "bullets": ["Начало работы", "Приёмка"]}]
            with patch.object(decks, "DATA_DIR", directory), patch.object(decks, "template_path", return_value=template), \
                    patch("pptx_builder._template_analysis", return_value={}), patch.object(decks, "render_pptx_previews"):
                self.assertEqual(decks.save_edited("narrow", "Программа", "", slides, template=template.name), [])
                pages = decks.read_presentation("narrow")["slides"]
                self.assertEqual([s["bullets"] for s in pages[1:]], [["Начало работы"], ["Приёмка"]])
                self.assertEqual(len(Presentation(decks.pptx_path("narrow")).slides), 3)

    def test_existing_closing_slide_preserves_its_bullets(self):
        points = ["Проверили макет", "Согласовали условия"]
        slides = [{"kind": "title", "title": "Практикум"},
                  {"kind": "closing", "title": "Итог", "bullets": points}]
        with tempfile.TemporaryDirectory() as tmp, patch.object(decks, "DATA_DIR", Path(tmp)), \
                patch.object(decks, "render_pptx_previews"):
            self.assertEqual(decks.save_edited("existing", "Практикум", "", slides), [])
            page = Presentation(decks.pptx_path("existing")).slides[-1]
            text = "\n".join(s.text for s in page.shapes if s.has_text_frame)
            for value in points:
                self.assertIn(value, text)
            self.assertEqual(decks.read_presentation("existing")["slides"][-1]["kind"], "closing")

    def test_rejected_book_length_block_preserves_saved_deck(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(decks, "DATA_DIR", Path(tmp)), \
                patch.object(decks, "render_pptx_previews"):
            slides = [{"kind": "title", "title": "Памятка"}]
            self.assertEqual(decks.save_edited("guide", "Памятка", "", slides), [])
            pptx = decks.pptx_path("guide").read_bytes()
            side = decks.read_presentation("guide")
            slides.append({"kind": "content", "title": "Весь материал", "bullets": ["Целая глава. " * 1000]})
            issues = decks.save_edited("guide", "Памятка", "", slides)
            self.assertIn("разделите", issues[0])
            self.assertEqual(decks.read_presentation("guide"), side)
            self.assertEqual(decks.pptx_path("guide").read_bytes(), pptx)

    def test_transparent_vector_webp_and_rotated_photo(self):
        from slide_layout import add_fitted_picture
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "logo.svg").write_text('<svg xmlns="http://www.w3.org/2000/svg" width="60" height="20">'
                                             '<circle cx="10" cy="10" r="5" fill="red"/></svg>', encoding="utf-8")
            Image.new("RGBA", (30, 10), (20, 80, 160, 90)).save(folder / "mark.webp")
            image = Image.new("RGB", (60, 20))
            exif = image.getexif()
            exif[274] = 6
            image.save(folder / "phone.jpg", exif=exif)
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            for name, ratio in [("logo.svg", 3), ("mark.webp", 3), ("phone.jpg", 1/3)]:
                with self.subTest(name=name):
                    shape = add_fitted_picture(slide, folder / name, 0, 0, Inches(4), Inches(4))
                    self.assertAlmostEqual(shape.width / shape.height, ratio, delta=.01)
                    self.assertAlmostEqual(shape.left + shape.width/2, Inches(2), delta=1)
                    self.assertAlmostEqual(shape.top + shape.height/2, Inches(2), delta=1)
                    if name != "phone.jpg":
                        from io import BytesIO
                        with Image.open(BytesIO(shape.image.blob)) as pixels:
                            self.assertLess(pixels.getchannel("A").getextrema()[0], 255)

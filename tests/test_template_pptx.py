import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from pptx_builder import build_presentation, validate_presentation


class TemplatePresentationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.output = self.directory / "output.pptx"
        self.template = self.directory / "source.pptx"
        self.picture = self.directory / "logo.png"
        Image.new("RGB", (48, 32), "blue").save(self.picture)
        source = Presentation()
        for title in ("First", "Middle", "Last"):
            slide = source.slides.add_slide(source.slide_layouts[6])
            for y, text, size in ((1, title, 24), (2, "Subtitle", 16)):
                box = slide.shapes.add_textbox(Inches(1), Inches(y), Inches(4), Inches(1))
                run = box.text_frame.paragraphs[0].add_run()
                run.text = text
                run.font.name = "Georgia"
                run.font.size = Pt(size)
            slide.shapes.add_picture(str(self.picture), Inches(7), Inches(1))
        source.save(self.template)

    def build(self, slides, template=None):
        payload = {"title": "Deck", "slides": slides}
        self.assertEqual(validate_presentation(payload), [])
        with patch("pptx_builder._template_analysis", return_value={}):
            build_presentation(self.output, payload, template=template)
        return Presentation(self.output)

    def test_new_presentation_has_one_slide_per_item(self):
        slides = [{"kind": "title", "title": "Intro"},
                  {"kind": "content", "title": "Facts", "bullets": ["One"]},
                  {"kind": "closing", "title": "End"}]
        result = self.build(slides)
        self.assertEqual(len(result.slides), len(slides))
        self.assertEqual(result.slides[0].shapes[0].text, "Intro")

    def test_exact_fields_retain_geometry_other_text_and_cloned_media(self):
        def item(text):
            return {"source_slide": 2, "template_fields": {"2": text}}
        result = self.build([item("First replacement"), item("Other replacement")], self.template)
        for idx, expected in enumerate(("First replacement", "Other replacement")):
            slide = result.slides[idx]
            self.assertEqual(slide.shapes[0].text, expected)
            self.assertEqual(slide.shapes[0].width, Inches(4))
            self.assertEqual(slide.shapes[0].text_frame.paragraphs[0].runs[0].font.name, "Georgia")
            self.assertEqual(slide.shapes[1].text, "Subtitle")
            self.assertEqual(slide.shapes[2].image.blob, self.picture.read_bytes())

    def test_freeform_preserves_subtitle_slot_and_nested_content(self):
        result = self.build([{"kind": "title", "title": "Intro", "subtitle": "Author"},
                             {"kind": "content", "title": "Facts", "bullets": [
                                 {"text": "One", "sub": ["Nested"]}]}], self.template)
        self.assertEqual(result.slides[0].shapes[0].width, Inches(4))
        self.assertEqual(result.slides[0].shapes[1].text, "Author")
        self.assertIn("Nested", result.slides[1].shapes[1].text)

    def test_single_slide_design_can_be_reused_as_a_template(self):
        source = Presentation()
        page = source.slides.add_slide(source.slide_layouts[6])
        for i, value in enumerate(("Old title", "Old subtitle")):
            page.shapes.add_textbox(Inches(1), Inches(i + 1), Inches(5), Inches(1)).text = value
        page.shapes.add_picture(str(self.picture), Inches(7), Inches(1))
        source.save(self.template)
        result = self.build([{"kind": "title", "title": "Intro", "subtitle": "New topic"},
                             {"kind": "content", "title": "Scope", "bullets": ["Details"]}], self.template)
        for page in result.slides:
            self.assertTrue(any(getattr(shape, "image", None) for shape in page.shapes))
        self.assertEqual(result.slides[1].shapes[1].text, "Details")

    def test_unknown_source_or_field_fails_before_output_is_written(self):
        for slide in ({"source_slide": 8, "template_fields": {}},
                      {"source_slide": 1, "template_fields": {"999": "text"}}):
            with self.subTest(slide=slide), self.assertRaises(ValueError):
                self.build([slide], self.template)
            self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from pptx_builder import build_presentation


class NativeLayoutTests(unittest.TestCase):
    def test_blank_branded_slide_styles_new_table_from_its_master(self):
        from pptx.dml.color import RGBColor
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        from pptx.enum.shapes import MSO_SHAPE
        source = Presentation()
        page = source.slides.add_slide(source.slide_layouts[6])
        page.background.fill.solid()
        page.background.fill.fore_color.rgb = RGBColor.from_string("102030")
        page.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8), 0, Inches(1), Inches(.2)).name = "Brand stripe"
        theme = page.slide_layout.slide_master.part.part_related_by(RT.THEME)
        theme._blob = theme.blob.replace(b'typeface="Calibri"', b'typeface="Georgia"')
        template = self.directory / "brand.pptx"
        source.save(template)
        out = self.directory / "result.pptx"
        with patch("pptx_builder._template_analysis", return_value={}):
            build_presentation(out, {"slides": [{"kind": "content", "title": "Demo",
                "table": {"headers": ["One","Two"], "rows": [["A","B"]]}}]}, template=template)
        page = Presentation(out).slides[0]
        self.assertIn("Brand stripe", [s.name for s in page.shapes])
        cell = next(s.table for s in page.shapes if s.has_table).cell(1, 0)
        self.assertEqual(str(cell.fill.fore_color.rgb), "102030")
        self.assertEqual(cell.text_frame.paragraphs[0].runs[0].font.name, "Georgia")

    def test_replaced_template_is_analyzed_again_even_with_same_timestamp(self):
        import fitz
        import os
        from pptx_builder import _template_analysis
        template = self.directory / "same.pptx"
        template.write_bytes(b"first")
        stamp = template.stat().st_mtime_ns
        def render(source, path):
            with fitz.open() as doc:
                doc.new_page().insert_text((50, 100), "Brand", fontname="cour" if source.read_bytes() == b"first" else "tiro")
                doc.save(path)
        with patch("pptx_builder._pptx_to_pdf", side_effect=render) as convert:
            first = _template_analysis(template)
            self.assertEqual(first, _template_analysis(template))
            template.write_bytes(b"other")
            os.utime(template, ns=(stamp, stamp))
            self.assertNotEqual(first["font"], _template_analysis(template)["font"])
            self.assertEqual(convert.call_count, 2)

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.directory = Path(tmp.name)
        Image.new("RGB", (90, 720), "teal").save(self.directory / "portrait.png")

    def render(self, slide):
        path = self.directory / "out.pptx"
        with patch("pptx_builder.ASSETS_DIR", self.directory):
            build_presentation(path, {"slides": [dict(kind="content", title="Report", **slide)]})
        return Presentation(path)

    def test_short_content_uses_available_space_and_disables_autoshrink(self):
        slide = self.render({"bullets": ["Concise statement"]}).slides[0]
        body = next(s for s in slide.shapes if s.has_text_frame and "Concise" in s.text)
        self.assertGreaterEqual(body.text_frame.paragraphs[0].runs[0].font.size.pt, 28)
        self.assertEqual(body.text_frame.auto_size, MSO_AUTO_SIZE.NONE)

    def test_portrait_fits_inside_slide_and_outside_body_and_statistic(self):
        deck = self.render({"bullets": ["Statement"], "image": "portrait.png",
                            "stats": [{"value": "90%", "caption": "Completion"}]})
        shapes = deck.slides[0].shapes
        image = next(s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
        self.assertAlmostEqual(image.width / image.height, 90 / 720, places=4)
        self.assertLessEqual(image.top + image.height, deck.slide_height)
        body = next(s for s in shapes if s.has_text_frame and "Statement" in s.text)
        statistic = next(s for s in shapes if s.has_text_frame and s.text == "90%")
        self.assertLessEqual(body.left + body.width, image.left)
        self.assertLessEqual(image.top + image.height, statistic.top)

    def test_cards_and_stat_captions_have_readable_explicit_sizes(self):
        deck = self.render({"cards": [{"title": "One", "text": "First option"},
                                      {"title": "Two", "text": "Second option"}],
                            "stats": [{"value": "87%", "caption": "Completion"}]})
        for shape in deck.slides[0].shapes:
            if shape.has_text_frame and shape.text:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        self.assertGreaterEqual(run.font.size.pt, 18, run.text)

    def test_oversized_source_field_gives_actionable_error_without_resizing(self):
        source = Presentation()
        page = source.slides.add_slide(source.slide_layouts[6])
        box = page.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(.6))
        box.text = "ACME"
        box.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
        template = self.directory / "source.pptx"
        source.save(template)
        out = self.directory / "edited.pptx"
        fields = {str(box.shape_id): "Очень длинное название новой компании " * 15}
        with self.assertRaisesRegex(ValueError, "поле"):
            build_presentation(out, {"slides": [{"source_slide": 1, "template_fields": fields}]},
                               template=template)
        self.assertFalse(out.exists())

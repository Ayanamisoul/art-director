import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from pptx_builder import build_presentation


class PdfTemplateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.folder = Path(self.tmp.name)
        self.template = self.folder / "brand.pdf"
        with fitz.open() as doc:
            for title, y, size in [("Old cover", 240, 64), ("Old heading", 110, 38)]:
                page = doc.new_page(width=960, height=540)
                page.draw_rect(page.rect, fill=(0.96, 0.94, 0.9), color=None)
                page.draw_circle((45, 29), 10, fill=(0.84, 0.06, 0.09), color=None)
                page.insert_text((48, y), title, fontname="tibo", fontsize=size,
                                 color=(0.15, 0.12, 0.3))
                page.insert_text((48, y + 80), "Old body", fontname="tiro", fontsize=20)
            doc.save(self.template)

    def render(self, slides):
        target = self.folder / "new.pptx"
        build_presentation(target, {"slides": slides}, template=self.template)
        return Presentation(target)

    def test_cover_uses_source_type_scale_position_and_color(self):
        deck = self.render([{"kind": "title", "title": "New cover", "subtitle": "New body"}])
        title = next(s for s in deck.slides[0].shapes if s.has_text_frame and s.text == "New cover")
        run = title.text_frame.paragraphs[0].runs[0]
        self.assertAlmostEqual(title.left / deck.slide_width, .05, places=2)
        self.assertAlmostEqual(run.font.size.pt, 64, delta=1)
        self.assertEqual(run.font.name, "Times New Roman")
        self.assertTrue(run.font.bold)
        self.assertEqual(str(run.font.color.rgb), "261F4D")

    def test_table_uses_reference_background_and_body_font(self):
        deck = self.render([{"kind":"content","title":"New heading",
                             "table":{"headers":["First","Second"],"rows":[["A","B"]]}}])
        table = next(s.table for s in deck.slides[0].shapes if s.has_table)
        cell = table.cell(1, 0)
        self.assertEqual(cell.text_frame.paragraphs[0].runs[0].font.name, "Times New Roman")
        color = cell.fill.fore_color.rgb
        self.assertTrue(all(abs(a-b) <= 2 for a,b in zip(color, (245,240,230))), color)

    def test_vertical_layout_keeps_heading_in_its_column(self):
        from types import SimpleNamespace
        from pptx.enum.text import PP_ALIGN
        from pdf_template import _title_area

        for alignment in (PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT):
            source = SimpleNamespace(box=fitz.Rect(400, 110, 800, 150), align=alignment)
            layout = SimpleNamespace(title=source, width=960, vertical=348, margin=50)
            area = _title_area(layout, 480)
            self.assertGreater(area.width, 200)
            self.assertLess(area.x1, layout.vertical)

    def test_content_stays_editable_and_preserves_structured_fields(self):
        Image.new("RGB", (40, 80), "#34ab98").save(self.folder / "asset.png")
        slide = {"kind": "content", "title": "New heading",
                 "bullets": [{"text": "First point", "sub": ["More detail"]}],
                 "cards": [{"title": "One", "text": "Before"}, {"title": "Two", "text": "After"}],
                 "stats": [{"value": "18%", "caption": "Increase"}], "image": "asset.png"}
        with patch("pptx_builder.ASSETS_DIR", self.folder):
            deck = self.render([slide])
        shapes = deck.slides[0].shapes
        text = "\n".join(s.text for s in shapes if s.has_text_frame)
        for item in ("First point", "More detail", "One", "Before", "Two", "After", "18%", "Increase"):
            self.assertIn(item, text)
        self.assertNotIn("'text'", text)
        self.assertNotIn("Old", text)
        self.assertTrue(any(s.shape_type == MSO_SHAPE_TYPE.PICTURE and
                            s.image.size == (40, 80) for s in shapes))

    def test_remove_photos_retains_repeated_corner_logo(self):
        from io import BytesIO
        logo = self.folder / "logo.png"
        Image.new("RGB", (80, 25), "#e62314").save(logo)
        source = self.folder / "logo.pdf"
        with fitz.open() as doc:
            for title in ("Cover", "Closing"):
                page = doc.new_page(width=600, height=360)
                page.insert_image(fitz.Rect(20, 18, 100, 43), filename=str(logo))
                page.insert_text((20, 160), title, fontsize=58)
            doc.save(source)
        path = self.folder / "output.pptx"
        build_presentation(path, {"slides": [{"kind": "title", "title": "New title"}]},
                           template=source, template_images="remove")
        background = Presentation(path).slides[0].shapes[0].image.blob
        with Image.open(BytesIO(background)) as image:
            color = image.getpixel((int(.10 * image.width), int(.08 * image.height)))
            self.assertEqual(color[:3], (230, 35, 20))


if __name__ == "__main__":
    unittest.main()

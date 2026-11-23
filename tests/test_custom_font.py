import struct
import base64
import sys
import tempfile
import unittest
import zipfile
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
from core import artifacts, decks, fonts


def sample_font(flags=0):
    builder = FontBuilder(1000, isTTF=True)
    characters = {ord(c): f"u{ord(c):04X}" for c in " CustomDeckFontТест"}
    glyphs = {}
    for name in [".notdef", *characters.values()]:
        pen = TTGlyphPen(None)
        if name != characters[32]:
            pen.moveTo((40, 0))
            pen.lineTo((300, 700))
            pen.lineTo((560, 0))
            pen.closePath()
        glyphs[name] = pen.glyph()
    builder.setupGlyphOrder(list(glyphs))
    builder.setupCharacterMap(characters)
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (600, 0) for name in glyphs})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({"familyName": "CustomDeckFont", "styleName": "Regular",
                            "fullName": "CustomDeckFont Regular", "psName": "CustomDeckFont-Regular",
                            "version": "Version 1.0"})
    builder.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200, fsType=flags)
    builder.setupPost()
    out = BytesIO()
    builder.save(out)
    return out.getvalue()


class CustomFontTests(unittest.TestCase):
    def test_save_and_edit_embed_the_full_font(self):
        root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.enterContext(patch.object(fonts, "FONT_DIR", root / "fonts"))
        self.enterContext(patch.object(decks, "DATA_DIR", root))
        self.enterContext(patch.object(decks, "render_pptx_previews"))
        data = sample_font()
        key, family = fonts.accept_font(data)
        slides = [{"kind": "title", "title": "Тест"}, {"kind": "text", "title": "Тест", "bullets": ["Тест"]},
                  {"kind": "closing", "title": "Тест"}]
        for theme in ({"font_id": key}, {}):
            self.assertEqual(decks.save_edited("test", "Тест", "", slides, theme), [])
            deck = Presentation(decks.pptx_path("test"))
            for slide in deck.slides:
                self.assertEqual({r.get("typeface") for r in slide._element.xpath(".//a:r/a:rPr/a:latin")}, {family})
            with zipfile.ZipFile(decks.pptx_path("test")) as archive:
                part = archive.read("ppt/fonts/font1.fntdata")
                self.assertEqual(struct.unpack_from("<II", part), (len(part), len(data)))
                self.assertTrue(part.endswith(data))
                self.assertIn(b"embeddedFontLst", archive.read("ppt/presentation.xml"))

    def test_reject_damaged_restricted_and_preview_only_fonts(self):
        for raw in (b"broken", sample_font(2), sample_font(4), sample_font(0x200)):
            with self.subTest(size=len(raw)), self.assertRaises(ValueError):
                fonts.accept_font(raw)

    def test_landing_keeps_embedded_font_after_profile_change(self):
        root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.enterContext(patch.object(fonts, "FONT_DIR", root / "fonts"))
        self.enterContext(patch.object(artifacts, "DATA_DIR", root))
        data = sample_font()
        key, _ = fonts.accept_font(data)
        for profile, selected in (("law", key), ("school", "")):
            artifacts.save_landing("page", "Тест", "", [], profile, font_id=selected)
            self.assertEqual(artifacts.read_landing("page")["font_id"], key)
            self.assertIn(base64.b64encode(data).decode(), (root / "page.html").read_text(encoding="utf-8"))

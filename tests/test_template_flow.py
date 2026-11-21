import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))

from core import artifacts, decks, templates
from pptx import Presentation
from pptx.util import Inches


class TemplateFlowTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        for module, name in ((templates, "TEMPLATES_DIR"),
                             (decks, "DATA_DIR"), (artifacts, "DATA_DIR")):
            self.enterContext(patch.object(module, name, root))
        self.enterContext(patch.object(decks, "render_pptx_previews", return_value=0))
        self.root = root

    def test_selected_template_wins_and_stays_on_edit(self):
        for name in ("a.pptx", "b.pptx"):
            Presentation().save(self.root / name)
        templates.set_preferred("b.pptx")
        slides = [{"kind": "title", "title": "Титул"}]
        with patch.object(decks, "build_presentation"):
            decks.agent_save_presentation("deck", "Титул", slides, template="a.pptx")
            self.assertEqual(decks.read_presentation("deck")["theme"]["template"], "a.pptx")
            decks.agent_save_presentation("deck", "Правка", slides)
            self.assertEqual(decks.read_presentation("deck")["theme"]["template"], "a.pptx")

    def test_no_template_stays_selected_on_edit(self):
        Presentation().save(self.root / "preferred.pptx")
        templates.set_preferred("preferred.pptx")
        slides = [{"kind": "title", "title": "Титул"}]
        with patch.object(decks, "build_presentation"):
            decks.agent_save_presentation("deck", "Титул", slides, template="__none__")
            decks.agent_save_presentation("deck", "Титул", slides, template="__auto__")
            self.assertNotIn("template", decks.read_presentation("deck")["theme"])

    def test_judge_sees_confirmed_filename(self):
        from web import _last_artifact_summary
        slides = [{"kind": "title", "title": "Первый слайд"}]
        result = decks.agent_save_presentation("имя_файла", "Титул", slides, template="__none__")
        messages = [{"role": "assistant", "tool_calls": [
            {"id": "generate", "function": {"name": "save_presentation"}}]},
            {"role": "tool", "tool_call_id": "generate", "content": result}]
        summary = _last_artifact_summary(artifacts.saved_results(messages))
        self.assertIn("имя_файла.pptx", summary)
        self.assertIn("Первый слайд", summary)

    def test_read_template_returns_actual_shape_id(self):
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        box = slide.shapes.add_textbox(0, 0, Inches(3), Inches(1))
        box.text = "Название компании"
        deck.save(self.root / "source.pptx")
        outline = templates.read_template("source.pptx")
        self.assertEqual(outline["slides"][0]["fields"][str(box.shape_id)], box.text)

    def test_broken_upload_preserves_existing_template(self):
        path = self.root / "source.pptx"
        Presentation().save(path)
        before = path.read_bytes()
        with self.assertRaises(ValueError):
            templates.accept_upload_bytes(path.name, b"invalid")
        self.assertEqual(path.read_bytes(), before)

    def test_landing_round_trip_keeps_profile_and_escapes_embedded_brief(self):
        sections = [{"title": "Практикум", "text": "Материалы занятия"}]
        title = "Тема </script> & заметки"
        artifacts.save_landing("курс", title, "Описание", sections, profile="school")
        brief = artifacts.read_landing("курс")
        self.assertEqual(brief["title"], title)
        artifacts.save_landing("курс", "Новая тема", "Описание", sections)
        self.assertEqual(artifacts.read_landing("курс")["profile"], "school")
        self.assertIn("Новая тема", (self.root / "курс.html").read_text(encoding="utf-8"))
        self.assertEqual(artifacts.list_files(), ["курс.html"])

import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import AsyncMock, patch

from fastapi.testclient import TestClient
from PIL import Image
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
import web
from core import assets, sessions, templates
from core import decks, fonts
import pptx_builder
from test_custom_font import sample_font


class PresentationFormTests(unittest.TestCase):
    def setUp(self):
        directory = self.enterContext(tempfile.TemporaryDirectory())
        root = Path(directory)
        self.enterContext(patch.object(assets, "ASSETS_DIR", root))
        self.enterContext(patch.object(web, "ASSETS_DIR", root))
        self.enterContext(patch.object(templates, "TEMPLATES_DIR", root))
        self.enterContext(patch.object(sessions, "SESSIONS_DIR", root))
        self.enterContext(patch.object(sessions, "sessions", {}))
        self.client = TestClient(web.app)

    def test_attached_images_keep_distinct_content_and_reach_generation(self):
        uploaded = []
        for color in ("red", "blue"):
            stream = BytesIO()
            Image.new("RGB", (8, 6), color).save(stream, format="PNG")
            response = self.client.post("/api/assets", data={"slide_attachment": "true"},
                                        files={"file": ("фото.png", stream.getvalue())})
            self.assertEqual(response.status_code, 200)
            name = response.json()["name"]
            self.assertEqual((assets.ASSETS_DIR / name).read_bytes(), stream.getvalue())
            uploaded.append(name)
        self.assertNotEqual(*uploaded)
        with patch.object(web, "agent_reply", new_callable=AsyncMock, return_value="") as generate:
            self.client.post("/api/create-presentation", data={"brief": "Тема", "image_names": uploaded})
        messages = generate.await_args.args[1]
        for name in uploaded:
            self.assertIn(name, messages[0]["content"])

    def test_corrupt_image_and_missing_attachment_are_rejected(self):
        response = self.client.post("/api/assets", data={"slide_attachment": "true"},
                                    files={"file": ("bad.jpg", b"broken")})
        self.assertEqual(response.status_code, 400)
        self.assertFalse(assets.list_names())
        with patch.object(web, "agent_reply", new_callable=AsyncMock) as generate:
            response = self.client.post("/api/create-presentation",
                                        data={"brief": "Тема", "image_names": "missing.jpg"})
        self.assertEqual(response.status_code, 400)
        generate.assert_not_awaited()

    def test_font_upload_is_validated_and_pinned_for_generation(self):
        self.enterContext(patch.object(fonts, "FONT_DIR", assets.ASSETS_DIR / "fonts"))
        self.assertEqual(self.client.post("/api/fonts", files={"file": ("bad.ttf", b"bad")}).status_code, 400)
        font = self.client.post("/api/fonts", files={"file": ("brand.ttf", sample_font())}).json()
        self.assertEqual(font["family"], "CustomDeckFont")
        with patch.object(web, "agent_reply", new_callable=AsyncMock, return_value="") as generate:
            response = self.client.post("/api/create-presentation", data={"brief": "Тема", "font_id": "missing"})
            self.assertEqual(response.status_code, 400)
            generate.assert_not_awaited()
            self.client.post("/api/create-presentation", data={"brief": "Тема", "font_id": font["id"]})
            self.assertEqual(generate.await_args.kwargs["selections"]["save_presentation"]["font_id"], font["id"])

    def test_missing_picture_in_actual_slide_triggers_repair(self):
        self.enterContext(patch.object(decks, "DATA_DIR", assets.ASSETS_DIR))
        self.enterContext(patch.object(pptx_builder, "ASSETS_DIR", assets.ASSETS_DIR))
        self.enterContext(patch.object(decks, "render_pptx_previews"))
        self.enterContext(patch.object(web, "semantic_check", side_effect=lambda *args: {"ok": True, "issues": []}))
        image = "board.png"
        Image.new("RGB", (10, 10), "red").save(assets.ASSETS_DIR / image)
        calls = []

        async def reply(backend, messages, toolbox, *, selections, **kwargs):
            name = selections["save_presentation"]["name"]
            slides = [{"kind": "title", "title": "Тема"},
                      {"kind": "text", "title": "Плата", "bullets": ["Детали платы"]},
                      {"kind": "closing", "title": "Итоги"}]
            if calls:
                self.assertIn(image, messages[-1]["content"])
                slides[1]["image"] = image
            calls.append(name)
            result = decks.agent_save_presentation(name, "Тема", slides, template="__none__")
            call_id = str(len(calls))
            messages.extend([{"role": "assistant", "tool_calls": [{"id": call_id, "function": {"name": "save_presentation"}}]},
                             {"role": "tool", "tool_call_id": call_id, "content": result}])
            return "Готово"

        with patch.object(web, "agent_reply", reply):
            response = self.client.post("/api/create-presentation", data={"brief": "Плата", "image_names": image})
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(len(calls), 2)
        self.assertTrue(response.json()["validation"]["ok"])

    def test_invalid_template_stops_before_model_call(self):
        with patch.object(web, "agent_reply", new_callable=AsyncMock) as generate:
            response = self.client.post("/api/create-presentation", data={"brief": "Тема"},
                                        files={"template": ("макет.pptx", b"invalid")})
        self.assertEqual(response.status_code, 400)
        generate.assert_not_awaited()

    def test_no_saved_file_retries_once_and_returns_no_download(self):
        with patch.object(web, "agent_reply", new_callable=AsyncMock, return_value="Готово") as generate:
            response = self.client.post("/api/create-presentation", data={"brief": "Тема"})
        self.assertEqual(response.status_code, 422)
        self.assertEqual(generate.await_count, 2)
        options = generate.await_args.kwargs
        self.assertEqual(options["output_format"], "presentation")
        self.assertEqual(options["selections"]["save_presentation"]["template"], "__none__")
        self.assertNotIn("file", response.json())

    def test_different_uploads_with_same_filename_keep_their_templates(self):
        chosen = []
        with patch.object(web, "agent_reply", new_callable=AsyncMock, return_value="") as generate:
            for title in ("Первый", "Второй"):
                file = BytesIO()
                deck = Presentation()
                deck.core_properties.title = title
                deck.save(file)
                self.client.post("/api/create-presentation", data={"brief": title},
                                 files={"template": ("шаблон.pptx", file.getvalue())})
                selection = generate.await_args.kwargs["selections"]["save_presentation"]
                chosen.append((selection["template"], selection["name"]))
        self.assertNotEqual(chosen[0][0], chosen[1][0])
        self.assertNotEqual(chosen[0][1], chosen[1][1])
        for (filename, _), title in zip(chosen, ("Первый", "Второй")):
            self.assertEqual(Presentation(templates.template_path(filename)).core_properties.title, title)

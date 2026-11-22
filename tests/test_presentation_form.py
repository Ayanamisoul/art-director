import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import AsyncMock, patch

from fastapi.testclient import TestClient
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "agent"))
import web
from core import sessions, templates


class PresentationFormTests(unittest.TestCase):
    def setUp(self):
        directory = self.enterContext(tempfile.TemporaryDirectory())
        root = Path(directory)
        self.enterContext(patch.object(templates, "TEMPLATES_DIR", root))
        self.enterContext(patch.object(sessions, "SESSIONS_DIR", root))
        self.enterContext(patch.object(sessions, "sessions", {}))
        self.client = TestClient(web.app)

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

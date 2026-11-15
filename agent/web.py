"""Веб-клиент агента: FastAPI поверх того же ядра, что и CLI (main.py).

  cd agent && py -m uvicorn web:app --port 8000
  дальше http://localhost:8000

Сессии чатов — в памяти процесса (учебный стенд; local-first файлы — в data/).
"""
import sys
import uuid
from contextlib import asynccontextmanager
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # импорты agent/*

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from config import BACKENDS, DEFAULT_BACKEND
from judge import semantic_check
from mcp_client import MCPToolbox
from prompts import SKILLS
from pptx_builder import build_presentation, normalize_slide, validate_presentation
from validation import _visible_text, validate_html
from main import agent_reply
import json as jsonlib

ARTIFACTS_SERVER = Path(__file__).parent / "mcp_servers" / "artifacts.py"
DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "artifacts"
DATA_DIR.mkdir(parents=True, exist_ok=True)
WEB_DIR = Path(__file__).parent / "web"

toolbox = MCPToolbox(ARTIFACTS_SERVER)
sessions: dict[str, dict] = {}  # session_id -> {"backend": str, "skill": str, "messages": list}


@asynccontextmanager
async def lifespan(app: FastAPI):
    await toolbox.connect()
    print("MCP подключён:", ", ".join(t.name for t in toolbox.tools))
    yield
    await toolbox.close()


app = FastAPI(title="Арт-директор", lifespan=lifespan)
app.mount("/files", StaticFiles(directory=DATA_DIR), name="files")


class ChatRequest(BaseModel):
    message: str
    session_id: str | None = None
    backend: str | None = None  # ollama | xkiro — переключатель из UI


@app.get("/")
async def index():
    return FileResponse(WEB_DIR / "index.html")


@app.post("/api/chat")
async def chat_endpoint(req: ChatRequest):
    session_id = req.session_id or uuid.uuid4().hex
    state = sessions.setdefault(session_id, {
        "backend": DEFAULT_BACKEND, "skill": "chat", "messages": [],
    })
    if req.backend in BACKENDS:
        state["backend"] = req.backend
    if not state["messages"]:
        state["messages"].append({"role": "system", "content": SKILLS[state["skill"]]})
    state["messages"].append({"role": "user", "content": req.message})

    tool_log: list[str] = []
    backend = BACKENDS[state["backend"]]
    try:
        answer = await agent_reply(backend, state["messages"], toolbox, log=tool_log)
    except Exception as exc:
        state["messages"].pop()
        return {"session_id": session_id, "error": str(exc), "tools": tool_log}

    # --- семантическая валидация: судья сверяет результат с запросом ---
    validation = {"ok": True, "issues": []}
    artifact = _last_artifact_summary(state["messages"])
    if artifact:
        validation = semantic_check(backend, req.message, artifact)
        if not validation["ok"]:
            feedback = ("Проверка качества нашла несоответствия запросу:\n- "
                        + "\n- ".join(validation["issues"])
                        + "\nИсправь: пересоздай результат инструментами с учётом замечаний.")
            state["messages"].append({"role": "user", "content": feedback})
            try:
                answer = await agent_reply(backend, state["messages"], toolbox, log=tool_log)
            except Exception as exc:
                answer = f"(ошибка при исправлении: {exc})"
            artifact2 = _last_artifact_summary(state["messages"])
            validation = semantic_check(backend, req.message, artifact2) if artifact2 else validation

    return {"session_id": session_id, "answer": answer, "tools": tool_log,
            "backend": state["backend"], "validation": validation}


def _last_artifact_summary(messages: list[dict]) -> str | None:
    """Содержимое последнего сохранённого артефакта — для судьи."""
    for message in reversed(messages):
        for call in (message.get("tool_calls") or []):
            fname = call.get("function", {}).get("name", "")
            if fname not in ("save_presentation", "save_html"):
                continue
            try:
                args = jsonlib.loads(call["function"].get("arguments") or "{}")
            except jsonlib.JSONDecodeError:
                continue
            if fname == "save_presentation":
                lines = [f"Презентация «{args.get('title', '')}» / {args.get('subtitle', '')}"]
                for s in args.get("slides", []):
                    bullets = s.get("bullets") or []
                    if isinstance(bullets, str):
                        bullets = [bullets]
                    lines.append(f"Слайд: {s.get('title', '')} — " + "; ".join(map(str, bullets))[:300])
                return "\n".join(lines)[:3000]
            text = _visible_text(args.get("html", ""))
            return "Лендинг, видимый текст:\n" + " ".join(text.split())[:2000]
    return None


# --- редактор презентаций ---

@app.get("/editor/{name}")
async def editor_page(name: str):
    return FileResponse(WEB_DIR / "editor.html")


@app.get("/api/presentation/{name}")
async def get_presentation(name: str):
    sidecar = DATA_DIR / f"{name}.json"
    if not sidecar.exists():
        raise HTTPException(404, f"нет структуры {name}.json — презентация не создана агентом?")
    return jsonlib.loads(sidecar.read_text(encoding="utf-8"))


class PresentationSave(BaseModel):
    title: str = ""
    subtitle: str = ""
    slides: list[dict]
    theme: dict = {}


@app.post("/api/presentation/{name}")
async def save_presentation_endpoint(name: str, data: PresentationSave):
    if not all(c in "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-\u0400-\u04FF" for c in name):
        raise HTTPException(400, "недопустимое имя")
    slides = [normalize_slide(s) for s in data.slides]
    payload = {"title": data.title, "subtitle": data.subtitle, "slides": slides}
    issues = validate_presentation(payload)
    if issues:
        return {"ok": False, "issues": issues}
    path = DATA_DIR / f"{name}.pptx"
    build_presentation(path, payload, data.theme or None)
    path.with_suffix(".json").write_text(
        jsonlib.dumps({**payload, "theme": data.theme}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    return {"ok": True, "file": f"/files/{name}.pptx"}


@app.get("/api/files")
async def list_files():
    files = sorted(p.name for p in DATA_DIR.iterdir()
                   if p.is_file() and p.suffix != ".json")  # .json — служебные сайдкары
    return {"files": files}

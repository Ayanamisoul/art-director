"""Веб-клиент агента: FastAPI поверх того же ядра, что и CLI (main.py).

  cd agent && py -m uvicorn web:app --port 8000
  дальше http://localhost:8000

Сессии чатов — в памяти процесса (учебный стенд; local-first файлы — в data/).
"""
import sys
import uuid
from contextlib import asynccontextmanager
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")  # cp1251-консоль vs эмодзи в логах
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent))  # импорты agent/*

from fastapi import FastAPI, HTTPException, UploadFile
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
import re

ARTIFACTS_SERVER = Path(__file__).parent / "mcp_servers" / "artifacts.py"
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" / "artifacts"
ASSETS_DIR = ROOT / "data" / "assets"
TEMPLATES_DIR = ROOT / "data" / "templates"
for d in (DATA_DIR, ASSETS_DIR, TEMPLATES_DIR):
    d.mkdir(parents=True, exist_ok=True)
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
app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")


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

    # --- что сохранили: для мгновенного предпросмотра в чате ---
    saved = {"presentation": None, "html": None}
    for message in state["messages"]:
        if message.get("role") != "tool":
            continue
        content = message.get("content") or ""
        m = re.search(r"Презентация сохранена и прошла проверку: .+?([^\\/]+)\.pptx", content)
        if m:
            saved["presentation"] = m.group(1)
        m = re.search(r"Сохранено и прошло проверку: .+?([^\\/]+)\.html", content)
        if m:
            saved["html"] = m.group(1)

    return {"session_id": session_id, "answer": answer, "tools": tool_log,
            "backend": state["backend"], "validation": validation, "saved": saved}


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
                if args.get("template"):
                    lines.append(f"Использован шаблон: {args['template']}")
                if args.get("accent_color"):
                    lines.append(f"Акцентный цвет: {args['accent_color']}")
                for s in args.get("slides", []):
                    bullets = s.get("bullets") or []
                    if isinstance(bullets, str):
                        bullets = [bullets]
                    image_note = f" [картинка: {s['image']}]" if s.get("image") else ""
                    lines.append(f"Слайд ({s.get('kind', 'content')}): {s.get('title', '')}{image_note} — "
                                 + "; ".join(map(str, bullets))[:300])
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


# --- загрузки пользователя: изображения и шаблоны (local-first) ---

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg", ".bmp"}


def _safe_filename(raw: str) -> str:
    name = Path(raw).name
    if not name or name.startswith("."):
        raise HTTPException(400, "недопустимое имя файла")
    return name


@app.post("/api/assets")
async def upload_asset(file: UploadFile):
    name = _safe_filename(file.filename or "image.png")
    if Path(name).suffix.lower() not in IMAGE_EXTS:
        raise HTTPException(400, f"«{name}» — не изображение. Форматы: "
                                 f"{', '.join(sorted(e.lstrip('.') for e in IMAGE_EXTS))}")
    for p in ASSETS_DIR.iterdir():  # инвариант: имена без регистра
        if p.name.lower() == name.lower():
            name = p.name
            break
    replaced = (ASSETS_DIR / name).exists()
    (ASSETS_DIR / name).write_bytes(await file.read())
    return {"ok": True, "name": name, "replaced": replaced}


@app.post("/api/templates")
async def upload_template(file: UploadFile):
    name = _safe_filename(file.filename or "template.pptx")
    if Path(name).suffix.lower() != ".pptx":
        raise HTTPException(400, f"«{name}» — не шаблон. Нужен файл .pptx (PowerPoint)")
    (TEMPLATES_DIR / name).write_bytes(await file.read())
    return {"ok": True, "name": name, "replaced": (TEMPLATES_DIR / name).exists()}


@app.get("/api/assets")
async def list_assets():
    return {"files": sorted(p.name for p in ASSETS_DIR.iterdir() if p.is_file())}


@app.get("/api/templates")
async def list_templates():
    return {"files": sorted(p.name for p in TEMPLATES_DIR.iterdir() if p.is_file())}

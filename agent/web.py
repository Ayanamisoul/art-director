"""Веб-клиент агента: FastAPI поверх того же ядра, что и CLI (main.py).

  cd agent && py -m uvicorn web:app --port 8000
  дальше http://localhost:8000

Сессии чатов хранятся на диске пользователя (data/sessions/*.json, local-first):
можно вернуться к чату по конкретной презентации и продолжить правки.
"""
import sys
import uuid
from contextlib import asynccontextmanager
from datetime import datetime
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
from previews import render_pptx_previews
from validation import _visible_text, validate_html
from main import agent_reply
import json as jsonlib
import re

ARTIFACTS_SERVER = Path(__file__).parent / "mcp_servers" / "artifacts.py"
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" / "artifacts"
ASSETS_DIR = ROOT / "data" / "assets"
TEMPLATES_DIR = ROOT / "data" / "templates"
SESSIONS_DIR = ROOT / "data" / "sessions"
PREVIEW_DIR = DATA_DIR / "previews"
for d in (DATA_DIR, ASSETS_DIR, TEMPLATES_DIR, SESSIONS_DIR, PREVIEW_DIR):
    d.mkdir(parents=True, exist_ok=True)
WEB_DIR = Path(__file__).parent / "web"

toolbox = MCPToolbox(ARTIFACTS_SERVER)
sessions: dict[str, dict] = {}  # session_id -> {id, backend, skill, messages, presentation, ...}


def _load_sessions() -> None:
    for f in SESSIONS_DIR.glob("*.json"):
        try:
            state = jsonlib.loads(f.read_text(encoding="utf-8"))
            sessions[state["id"]] = state
        except Exception as exc:  # битый файл не должен ронять сервер
            print(f"сессия {f.name} пропущена: {exc}")


def _save_session(state: dict) -> None:
    state["updated"] = datetime.now().isoformat(timespec="seconds")
    (SESSIONS_DIR / f"{state['id']}.json").write_text(
        jsonlib.dumps(state, ensure_ascii=False), encoding="utf-8")


_load_sessions()


@asynccontextmanager
async def lifespan(app: FastAPI):
    await toolbox.connect()
    print("MCP подключён:", ", ".join(t.name for t in toolbox.tools))
    yield
    await toolbox.close()


app = FastAPI(title="Арт-директор", lifespan=lifespan)
app.mount("/files", StaticFiles(directory=DATA_DIR), name="files")
app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")
app.mount("/previews", StaticFiles(directory=PREVIEW_DIR), name="previews")


@app.get("/api/presentation/{name}/previews")
async def presentation_previews(name: str):
    """Есть ли настоящие PNG-превью слайдов (рендер из файла) и сколько."""
    out_dir = DATA_DIR / "previews" / name
    if not out_dir.exists():
        return {"ready": False, "count": 0}
    count = len(list(out_dir.glob("p*.png")))
    return {"ready": count > 0, "count": count}


class ChatRequest(BaseModel):
    message: str
    session_id: str | None = None
    backend: str | None = None  # ollama | xkiro | zai — переключатель из UI
    template: str | None = None  # шаблон, выбранный пользователем в UI
    output_format: str | None = None  # presentation | landing | None = решает модель


@app.get("/")
async def index():
    return FileResponse(WEB_DIR / "index.html")


@app.post("/api/chat")
async def chat_endpoint(req: ChatRequest):
    session_id = req.session_id or uuid.uuid4().hex
    state = sessions.get(session_id)
    if state is None:
        state = {"id": session_id, "backend": DEFAULT_BACKEND, "skill": "chat",
                 "messages": [], "presentation": None}
        sessions[session_id] = state
    if req.backend in BACKENDS:
        state["backend"] = req.backend
    if req.template is not None:
        state["chosen_template"] = req.template  # пустая строка = «авто»
    if req.output_format is not None:
        state["chosen_format"] = req.output_format  # пустая строка = решает модель
    # системный промпт пересобираем каждый ход: подмешиваем актуальные файлы
    # пользователя (шаблоны, картинки, готовые презентации) и фокус сессии
    system = _system_prompt(state)
    if state["messages"] and state["messages"][0].get("role") == "system":
        state["messages"][0]["content"] = system
    else:
        state["messages"].insert(0, {"role": "system", "content": system})
    state["messages"].append({"role": "user", "content": req.message})

    tool_log: list[str] = []
    backend = BACKENDS[state["backend"]]
    try:
        answer = await agent_reply(backend, state["messages"], toolbox, log=tool_log)
    except Exception as exc:
        state["messages"].pop()
        _save_session(state)
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
        m = re.search(r"Презентация сохранена и прошла проверку: .+?([^\\/()\s]+)\.pptx", content)
        if m:
            saved["presentation"] = m.group(1)
        m = re.search(r"Сохранено и прошло проверку: .+?([^\\/()\s]+)\.html", content)
        if m:
            saved["html"] = m.group(1)
    # привязываем сессию к презентации, чтобы можно было вернуться к чату по ней
    if saved["presentation"] and not state.get("presentation"):
        state["presentation"] = saved["presentation"]
    _save_session(state)

    return {"session_id": session_id, "answer": answer, "tools": tool_log,
            "backend": state["backend"], "validation": validation, "saved": saved,
            "presentation": state.get("presentation")}


def _system_prompt(state: dict) -> str:
    """Скилл + текущая обстановка: что пользователь загрузил и что уже создано.
    Обновляется каждый ход — модель всегда видит актуальные файлы."""
    templates = sorted(TEMPLATES_DIR.glob("*.pptx"), key=lambda p: p.stat().st_mtime)
    images = sorted(p.name for p in ASSETS_DIR.iterdir() if p.is_file())
    decks = sorted(p.stem for p in DATA_DIR.glob("*.pptx"))
    lines = [SKILLS[state["skill"]] if state.get("skill") in SKILLS else SKILLS["chat"],
             "", "ФАЙЛЫ ПОЛЬЗОВАТЕЛЯ НА ЭТОЙ МАШИНЕ (актуально на сейчас):"]
    if templates:
        lines.append("- Шаблоны .pptx (свежий — последний): "
                     + ", ".join(p.name for p in templates)
                     + "\n  СОЗДАВАЙ ПРЕЗЕНТАЦИИ ПО САМОМУ СВЕЖЕМУ ШАБЛОНУ (template=...), "
                       "если пользователь не попросил другой стиль или без шаблона.")
    else:
        lines.append("- Шаблонов нет. Если пользователь просит «по моему шаблону» — "
                     "попроси загрузить .pptx в панели файлов.")
    if images:
        lines.append("- Изображения: " + ", ".join(images))
    if decks:
        lines.append("- Сохранённые презентации: " + ", ".join(decks)
                     + "\n  Для правок: get_presentation(имя) -> изменить нужные слайды -> "
                       "save_presentation с тем же именем.")
    if state.get("presentation"):
        lines.append(f"- ФОКУС ЭТОЙ СЕССИИ: презентация «{state['presentation']}» — "
                     f"пользователь обсуждает и правит её. Правки сохраняй с этим же именем.")
    chosen = state.get("chosen_template")
    if chosen:
        lines.append(f"- ПОЛЬЗОВАТЕЛЬ ВЫБРАЛ ШАБЛОН: «{chosen}» — во ВСЕХ вызовах save_presentation "
                     f"передавай template=«{chosen}», даже если пользователь в тексте не повторяет.")
    fmt = state.get("chosen_format")
    if fmt == "presentation":
        lines.append("- ФОРМАТ РЕЗУЛЬТАТА — ТОЛЬКО ПРЕЗЕНТАЦИЯ: сохраняй через save_presentation "
                     "(.pptx со слайдами). save_html и лендинги НЕ ИСПОЛЬЗОВАТЬ, даже если "
                     "запрос звучит как справочник — делай слайды.")
    elif fmt == "landing":
        lines.append("- ФОРМАТ РЕЗУЛЬТАТА — ТОЛЬКО ЛЕНДИНГ: сохраняй через save_html "
                     "(одностраничный HTML). save_presentation НЕ ИСПОЛЬЗОВАТЬ.")
    return "\n".join(lines)


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


# --- сессии чатов: список, история, привязка к презентации ---

@app.get("/api/sessions")
async def list_sessions():
    items = []
    for s in sessions.values():
        first_user = next((m["content"] for m in s["messages"]
                           if m.get("role") == "user" and isinstance(m.get("content"), str)), "")
        items.append({"id": s["id"], "presentation": s.get("presentation"),
                      "updated": s.get("updated", ""), "turns": sum(
                          1 for m in s["messages"] if m.get("role") == "user"),
                      "preview": first_user[:120]})
    items.sort(key=lambda x: x["updated"], reverse=True)
    return {"sessions": items}


@app.get("/api/session/{session_id}")
async def get_session(session_id: str):
    state = sessions.get(session_id)
    if state is None:
        raise HTTPException(404, "сессия не найдена")
    return {"id": state["id"], "presentation": state.get("presentation"),
            "backend": state.get("backend"), "messages": state["messages"]}


class ForPresentation(BaseModel):
    name: str


@app.post("/api/session/for-presentation")
async def session_for_presentation(body: ForPresentation):
    """Найти свежий чат по презентации или создать новый — для кнопки «чат» в панели."""
    candidates = [s for s in sessions.values() if s.get("presentation") == body.name]
    if candidates:
        return {"id": max(candidates, key=lambda s: s.get("updated", ""))["id"], "created": False}
    sid = uuid.uuid4().hex
    sessions[sid] = {"id": sid, "backend": DEFAULT_BACKEND, "skill": "chat",
                     "messages": [], "presentation": body.name}
    _save_session(sessions[sid])
    return {"id": sid, "created": True}


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
    template: str = ""


@app.post("/api/presentation/{name}")
async def save_presentation_endpoint(name: str, data: PresentationSave):
    if not all(c in "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-а-яА-ЯёЁ" for c in name):
        raise HTTPException(400, "недопустимое имя")
    slides = [normalize_slide(s) for s in data.slides]
    payload = {"title": data.title, "subtitle": data.subtitle, "slides": slides}
    issues = validate_presentation(payload)
    if issues:
        return {"ok": False, "issues": issues}
    # шаблон: из запроса, иначе из прошлой сборки — чтобы правки в редакторе
    # не сбрасывали фирменный стиль
    template_name = data.template or (data.theme or {}).get("template", "")
    if not template_name:
        old = DATA_DIR / f"{name}.json"
        if old.exists():
            try:
                template_name = (jsonlib.loads(old.read_text(encoding="utf-8")).get("theme") or {}).get("template", "")
            except Exception:
                pass
    template_path = TEMPLATES_DIR / template_name if template_name and (TEMPLATES_DIR / template_name).exists() else None
    if template_name:
        data.theme["template"] = template_name
    path = DATA_DIR / f"{name}.pptx"
    build_presentation(path, payload, data.theme or None, template=template_path)
    path.with_suffix(".json").write_text(
        jsonlib.dumps({**payload, "theme": data.theme}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    render_pptx_previews(path)  # обновить настоящие превью после правки
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
    # новый шаблон становится выбранным — юзер кидает файл, чтобы работать по нему
    (TEMPLATES_DIR / ".preferred").write_text(name, encoding="utf-8")
    return {"ok": True, "name": name, "replaced": (TEMPLATES_DIR / name).exists()}


class PreferredTemplate(BaseModel):
    name: str = ""


@app.post("/api/templates/preferred")
async def set_preferred_template(body: PreferredTemplate):
    """Выбор шаблона из выпадающего списка в чате (пустое имя = авто/самый свежий)."""
    hint = TEMPLATES_DIR / ".preferred"
    name = body.name.strip()
    if name:
        if not (TEMPLATES_DIR / Path(name).name).exists():
            raise HTTPException(404, f"шаблон «{name}» не найден")
        name = Path(name).name
    hint.write_text(name, encoding="utf-8")
    return {"ok": True, "preferred": name}


@app.get("/api/assets")
async def list_assets():
    return {"files": sorted(p.name for p in ASSETS_DIR.iterdir() if p.is_file())}


@app.get("/api/templates")
async def list_templates():
    return {"files": sorted(p.name for p in TEMPLATES_DIR.iterdir() if p.is_file())}

"""Веб-клиент агента. Тонкая сборка: роутеры + чат-эндпоинт + статика.

cd agent && py -m uvicorn web:app --port 8000  →  http://localhost:8000

Рабочая логика — в core/ (decks, templates, assets, sessions, previews),
HTTP-роутеры — в routers/. Здесь только чат-цикл агента (ему нужны toolbox
и judge) и монтирование статики.
"""
import json
import sys
import re
from contextlib import asynccontextmanager
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")  # cp1251-консоль vs эмодзи в логах
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from config import BACKENDS, DEFAULT_BACKEND
from core import sessions as core_sessions
from core import templates as core_templates
from core.artifacts import DATA_DIR
from core.assets import ASSETS_DIR
from core.previews import PREVIEW_DIR
from judge import semantic_check
from mcp_client import MCPToolbox
from prompts import SKILLS
from routers import assets as assets_router
from routers import files as files_router
from routers import sessions as sessions_router
from routers import templates as templates_router
from main import agent_reply

ARTIFACTS_SERVER = Path(__file__).parent / "mcp_servers" / "artifacts.py"
WEB_DIR = Path(__file__).parent / "web"

toolbox = MCPToolbox(ARTIFACTS_SERVER)


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
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")

app.include_router(sessions_router.router)
app.include_router(templates_router.router)
app.include_router(assets_router.router)
app.include_router(files_router.router)

core_sessions.load_all()


class ChatRequest(BaseModel):
    message: str
    session_id: str | None = None
    backend: str | None = None      # ollama | xkiro | zai — переключатель из UI
    template: str | None = None     # шаблон, выбранный пользователем в UI
    output_format: str | None = None  # presentation | landing | None = решает модель
    template_images: str | None = None  # remove = чекбокс «без фото из шаблона»


@app.get("/")
async def index():
    return FileResponse(WEB_DIR / "index.html")


@app.get("/editor/{name}")
async def editor_page(name: str):
    return FileResponse(WEB_DIR / "editor.html")


def _system_prompt(state: dict) -> str:
    """Скилл + текущая обстановка: файлы пользователя и фокус сессии.
    Пересобирается каждый ход — модель всегда видит актуальные списки."""
    from core.assets import text_list as assets_text
    from core.artifacts import DATA_DIR as ARTIFACTS
    from core.templates import TEMPLATES_DIR, preferred

    templates = sorted(TEMPLATES_DIR.glob("*.ppt*"), key=lambda p: p.stat().st_mtime) \
        + sorted(TEMPLATES_DIR.glob("*.pdf"), key=lambda p: p.stat().st_mtime)
    templates.sort(key=lambda p: p.stat().st_mtime)
    images = [p.name for p in ASSETS_DIR.iterdir() if p.is_file()]
    decks = sorted(p.stem for p in ARTIFACTS.glob("*.pptx"))
    lines = [SKILLS.get(state.get("skill", "chat"), SKILLS["chat"]),
             "", "ФАЙЛЫ ПОЛЬЗОВАТЕЛЯ НА ЭТОЙ МАШИНЕ (актуально на сейчас):"]
    if templates:
        lines.append("- Шаблоны (свежий — последний): "
                     + ", ".join(p.name for p in templates)
                     + "\n  СОЗДАВАЙ ПРЕЗЕНТАЦИИ ПО САМОМУ СВЕЖЕМУ ШАБЛОНУ (template=...), "
                       "если пользователь не попросил другой стиль или без шаблона.")
    else:
        lines.append("- Шаблонов нет. Если пользователь просит «по моему шаблону» — "
                     "попроси загрузить .pptx или .pdf в панели файлов.")
    if images:
        lines.append("- " + assets_text().replace("\n", "\n  "))
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
    if state.get("chosen_template_images"):
        lines.append("- ШАБЛОН БЕЗ ФОТО: пользователь выключил чужие картинки — во всех "
                     "save_presentation по PDF-шаблону передавай template_images=\"remove\".")
    fmt = state.get("chosen_format")
    if fmt == "presentation":
        lines.append("- ФОРМАТ РЕЗУЛЬТАТА — ТОЛЬКО ПРЕЗЕНТАЦИЯ: сохраняй через save_presentation "
                     "(.pptx со слайдами). save_html и лендинги НЕ ИСПОЛЬЗОВАТЬ, даже если "
                     "запрос звучит как справочник — делай слайды.")
    elif fmt == "landing":
        lines.append("- ФОРМАТ РЕЗУЛЬТАТА — ТОЛЬКО ЛЕНДИНГ: сохраняй через save_html "
                     "(одностраничный HTML). save_presentation НЕ ИСПОЛЬЗОВАТЬ.")
    return "\n".join(lines)


@app.post("/api/chat")
async def chat_endpoint(req: ChatRequest):
    state = core_sessions.ensure(req.session_id)
    if req.backend in BACKENDS:
        state["backend"] = req.backend
    if req.template is not None:
        state["chosen_template"] = req.template  # пустая строка = «авто»
    if req.output_format is not None:
        state["chosen_format"] = req.output_format  # пустая строка = решает модель
    if req.template_images is not None and req.template_images:
        state["chosen_template_images"] = req.template_images

    # системный промпт пересобираем каждый ход: актуальные файлы + фокус сессии
    system = _system_prompt(state)
    if state["messages"] and state["messages"][0].get("role") == "system":
        state["messages"][0]["content"] = system
    else:
        state["messages"].insert(0, {"role": "system", "content": system})
    state["messages"].append({"role": "user", "content": req.message})

    tool_log: list[str] = []
    backend = BACKENDS[state["backend"]]
    # выбор шаблона пользователем важнее догадок модели — фиксируем на каждый ход
    core_templates.set_preferred(state.get("chosen_template")
                                 or core_templates.newest())
    try:
        answer = await agent_reply(backend, state["messages"], toolbox, log=tool_log)
    except Exception as exc:
        state["messages"].pop()
        core_sessions.save(state)
        return {"session_id": state["id"], "error": str(exc), "tools": tool_log}

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
    if saved["presentation"] and not state.get("presentation"):
        state["presentation"] = saved["presentation"]
    core_sessions.save(state)

    return {"session_id": state["id"], "answer": answer, "tools": tool_log,
            "backend": state["backend"], "validation": validation, "saved": saved,
            "presentation": state.get("presentation")}


def _last_artifact_summary(messages: list[dict]) -> str | None:
    """Содержимое последнего сохранённого артефакта — для судьи."""
    from validation import _visible_text
    for message in reversed(messages):
        for call in (message.get("tool_calls") or []):
            fname = call.get("function", {}).get("name", "")
            if fname not in ("save_presentation", "save_html"):
                continue
            try:
                args = json.loads(call["function"].get("arguments") or "{}")
            except json.JSONDecodeError:
                continue
            if fname == "save_presentation":
                lines = [f"Презентация «{args.get('title', '')}» / {args.get('subtitle', '')}"]
                if args.get("template"):
                    lines.append(f"Использован шаблон: {args['template']}")
                if args.get("accent_color"):
                    lines.append(f"Акцентный цвет: {args['accent_color']}")
                for s in args.get("slides", []):
                    blines = []
                    for b in s.get("bullets") or []:
                        if isinstance(b, dict):
                            blines.append(str(b.get("text") or ""))
                            blines += ["  - " + str(x) for x in (b.get("sub") or [])]
                        else:
                            blines.append(str(b))
                    st = s.get("stats") or []
                    stats_note = (" [цифры: " + ", ".join(
                        (x.get("value", "") + " " + x.get("caption", "")).strip()
                        for x in st if isinstance(x, dict)) + "]") if st else ""
                    image_note = f" [картинка: {s['image']}]" if s.get("image") else ""
                    cards = s.get("cards") or []
                    cards_note = (" [карточки: " + "; ".join(
                        (x.get("title", "") + " — " + x.get("text", "")).strip(" —")
                        for x in cards if isinstance(x, dict)) + "]") if cards else ""
                    lines.append(f"Слайд ({s.get('kind', 'content')}): {s.get('title', '')}{image_note}{stats_note}{cards_note} — "
                                 + "; ".join(blines)[:300])
                return "\n".join(lines)[:3000]
            text = _visible_text(args.get("html", ""))
            return "Лендинг, видимый текст:\n" + " ".join(text.split())[:2000]
    return None

"""Веб-клиент агента. Тонкая сборка: роутеры + чат-эндпоинт + статика.

cd agent && py -m uvicorn web:app --port 8000  →  http://localhost:8000

Рабочая логика — в core/ (decks, templates, assets, sessions, previews),
HTTP-роутеры — в routers/. Здесь только чат-цикл агента (ему нужны toolbox
и judge) и монтирование статики.
"""
import asyncio
import hashlib
import sys
from contextlib import asynccontextmanager
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")  # cp1251-консоль vs эмодзи в логах
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from config import BACKENDS
from core import sessions as core_sessions
from core import decks
from core import templates as core_templates
from core.artifacts import DATA_DIR, saved_results
from core.landings import PROFILES
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
_active_chats: set[str] = set()


class ChatRequest(BaseModel):
    message: str
    session_id: str | None = None
    backend: str | None = None
    template: str | None = None     # шаблон, выбранный пользователем в UI
    output_format: str | None = None  # presentation | landing | None = решает модель
    landing_profile: str | None = None
    template_images: str | None = None  # remove = чекбокс «без фото из шаблона»


@app.get("/")
async def index(session: str | None = None):
    return FileResponse(WEB_DIR / ("index.html" if session else "create.html"))


@app.get("/workspace")
async def workspace():
    return FileResponse(WEB_DIR / "index.html")


@app.post("/api/create-presentation")
async def create_presentation(brief: str = Form(...), template: UploadFile | None = File(None)):
    brief = brief.strip()
    if not brief:
        raise HTTPException(400, "Опишите тему презентации.")
    chosen = "__none__"
    if template and template.filename:
        source = Path(template.filename).name
        content = await template.read()
        # Разные файлы с одинаковым названием должны оставаться доступны своим презентациям.
        stored = f"{hashlib.sha256(content).hexdigest()[:16]}_{source}"
        try:
            chosen, _ = core_templates.accept_upload_bytes(stored, content, prefer=False)
        except ValueError as exc:
            raise HTTPException(400, f"Проверьте шаблон «{source}». Нужен читаемый PPTX или PDF без пароля.") from exc
    state = core_sessions.ensure(None)
    name = f"presentation_{state['id']}"
    request = ChatRequest(message=brief, template=chosen, output_format="presentation")
    result = await _chat_turn(request, state, create_name=name)
    deck = decks.read_presentation(name) if (result.get("saved") or {}).get("presentation") == name else None
    if deck is None or not decks.pptx_path(name).is_file():
        if result.get("error"):
            raise HTTPException(502, "Сервис генерации не ответил. Повторите попытку через минуту.")
        raise HTTPException(422, "Собрать презентацию не удалось. Попробуйте уточнить тему или сократить исходный текст.")
    return {"name": name, "title": deck["title"], "slides": len(deck["slides"]),
            "session_id": state["id"], "file": f"/files/{name}.pptx", "validation": result["validation"]}


@app.get("/api/landing/profiles")
async def landing_profiles():
    return {"profiles": PROFILES}


@app.get("/editor/{name}")
async def editor_page(name: str):
    return FileResponse(WEB_DIR / "editor.html")


def _system_prompt(state: dict) -> str:
    """Скилл + текущая обстановка: файлы пользователя и фокус сессии.
    Пересобирается каждый ход — модель всегда видит актуальные списки."""
    from core.assets import text_list as assets_text

    lines = [SKILLS.get(state.get("skill"), SKILLS["chat"])]
    fmt = state.get("chosen_format")
    if fmt != "landing":
        lines.append(core_templates.text_list())
        chosen = state.get("chosen_template")
        if chosen == "__none__":
            lines.append("Собирай презентацию без шаблона.")
        elif chosen:
            lines.append(f"Выбран шаблон «{chosen}». Его оформление применяется при сборке.")
        if state.get("presentation"):
            lines.append(f"Фокус чата — презентация «{state['presentation']}». Прочитай её перед правкой.")
    if fmt != "presentation":
        if state.get("chosen_landing_profile"):
            lines.append(f"Стиль лендинга выбран в интерфейсе: {state['chosen_landing_profile']}.")
        if state.get("landing"):
            lines.append(f"Последний лендинг — «{state['landing']}». Для правки вызови get_landing.")
    if any(ASSETS_DIR.iterdir()):
        lines.append(assets_text())
    if fmt in ("presentation", "landing"):
        lines.append("Пользователь выбрал формат " + ("презентации." if fmt == "presentation" else "лендинга."))
    return "\n".join(lines)


@app.post("/api/chat")
async def chat_endpoint(req: ChatRequest):
    state = core_sessions.ensure(req.session_id)
    session_id = state["id"]
    if session_id in _active_chats:
        raise HTTPException(409, "Подождите завершения предыдущего сообщения в этом чате.")
    _active_chats.add(session_id)
    try:
        return await _chat_turn(req, state)
    finally:
        _active_chats.discard(session_id)


async def _chat_turn(req: ChatRequest, state: dict, *, create_name: str = ""):
    fmt = state.get("chosen_format") if req.output_format is None else req.output_format
    chosen = state.get("chosen_template", "") if req.template is None else req.template
    profile = state.get("chosen_landing_profile", "") if req.landing_profile is None else req.landing_profile
    if fmt != "landing" and chosen and chosen != "__none__" and chosen not in core_templates.list_names():
        raise HTTPException(400, f"Шаблон «{chosen}» недоступен. Выберите другой шаблон.")
    if fmt != "presentation" and profile and profile not in PROFILES:
        raise HTTPException(400, "Неизвестный стиль лендинга")
    if req.backend in BACKENDS:
        state["backend"] = req.backend
    if req.template is not None:
        state["chosen_template"] = req.template  # пустая строка = «авто»
    if req.output_format is not None:
        state["chosen_format"] = req.output_format  # пустая строка = решает модель
    if req.template_images is not None:
        state["chosen_template_images"] = req.template_images
    if req.landing_profile is not None:
        state["chosen_landing_profile"] = req.landing_profile

    # системный промпт пересобираем каждый ход: актуальные файлы + фокус сессии
    system = _system_prompt(state)
    if create_name:
        system += (f"\nСоздай готовую презентацию сейчас и сохрани её под именем {create_name}. "
                   "По короткой теме самостоятельно составь план, подбери понятные примеры и вывод. "
                   "Если аудитория и объём не указаны, ориентируйся на широкую аудиторию и 6–10 слайдов. "
                   "Не запрашивай согласование плана. Сохрани числовые данные брифа и не выдумывай статистику.")
    if state["messages"] and state["messages"][0].get("role") == "system":
        state["messages"][0]["content"] = system
    else:
        state["messages"].insert(0, {"role": "system", "content": system})
    state["messages"].append({"role": "user", "content": req.message})

    tool_log: list[str] = []
    backend = BACKENDS[state["backend"]]
    # В «Авто» постоянный профиль задаёт сборщик. Явная ссылка или название
    # образца в брифе меняет его и без переключения выпадающего списка.
    if not profile:
        brief = req.message.casefold()
        for key, host in (("glavbuh", "about.1gl.ru/ai"), ("school", "school.law.ru/neyro"),
                          ("law", "neyro-law.ru")):
            if host in brief or PROFILES[key].casefold() in brief:
                profile = key
                break
    if not chosen:
        chosen = next((name for name in core_templates.list_names()
                       if name.casefold() in req.message.casefold()), "")
    selections = {"save_presentation": {"template": chosen or "__auto__"},
                  "save_landing": {"profile": profile or ""}}
    if create_name:
        selections["save_presentation"]["name"] = create_name
        selections["get_presentation"] = {"name": create_name}
    if chosen and chosen != "__none__":
        selections["read_template"] = {"name": chosen}
    if req.template_images is not None or state.get("chosen_template_images"):
        selections["save_presentation"]["template_images"] = state.get("chosen_template_images") or "auto"
    start = len(state["messages"]) - 1
    async def reply():
        return await agent_reply(backend, state["messages"], toolbox, log=tool_log,
                                 output_format=state.get("chosen_format", ""), selections=selections)

    try:
        answer = await reply()
        if create_name and not saved_results(state["messages"][start:])["presentation"]:
            state["messages"].append({"role": "user", "content":
                "Файл ещё не сохранён. Исправь замечания сборщика и вызови save_presentation с готовыми слайдами."})
            answer = await reply()
    except Exception as exc:
        del state["messages"][start:]
        core_sessions.save(state)
        return {"session_id": state["id"], "error": str(exc), "tools": tool_log}

    # --- семантическая валидация: судья сверяет результат с запросом ---
    validation = {"ok": True, "issues": []}
    artifact = _last_artifact_summary(saved_results(state["messages"][start:]))
    if artifact:
        validation = await asyncio.to_thread(semantic_check, backend, req.message, artifact)
        if not validation["ok"]:
            feedback = ("Проверка качества нашла несоответствия запросу:\n- "
                        + "\n- ".join(validation["issues"])
                        + "\nИсправь: пересоздай результат инструментами с учётом замечаний.")
            state["messages"].append({"role": "user", "content": feedback})
            try:
                answer = await reply()
            except Exception as exc:
                answer = f"(ошибка при исправлении: {exc})"
            artifact2 = _last_artifact_summary(saved_results(state["messages"][start:]))
            if artifact2:
                validation = await asyncio.to_thread(semantic_check, backend, req.message, artifact2)

    # --- что сохранили: для мгновенного предпросмотра в чате ---
    saved = saved_results(state["messages"][start:])
    if saved["presentation"]:
        state["presentation"] = saved["presentation"]
    if saved["html"]:
        state["landing"] = saved["html"]
    core_sessions.save(state)

    return {"session_id": state["id"], "answer": answer, "tools": tool_log,
            "backend": state["backend"], "validation": validation, "saved": saved,
            "presentation": state.get("presentation")}


def _last_artifact_summary(saved: dict) -> str | None:
    """Судья читает сохранённый результат текущего хода."""
    from pptx import Presentation
    from pptx_builder import template_text_shapes
    from core.decks import pptx_path, read_presentation
    from core.artifacts import read_landing
    from validation import _visible_text

    if saved["presentation"]:
        name = saved["presentation"]
        deck = Presentation(pptx_path(name))
        theme = (read_presentation(name) or {}).get("theme", {})
        lines = [f"Файл {name}.pptx сохранён, {len(deck.slides)} слайдов. "
                 f"Шаблон: {theme.get('template', 'базовый')}."]
        for i, slide in enumerate(deck.slides, 1):
            text = [s.text for s in template_text_shapes(slide.shapes)]
            for shape in slide.shapes:
                if shape.has_table:
                    text.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
            lines.append(f"{i}. " + " ".join(text)[:450])
        return "\n".join(lines)[:5000]
    if saved["html"]:
        name = saved["html"]
        brief = read_landing(name) or {}
        html = (DATA_DIR / f"{name}.html").read_text(encoding="utf-8")
        return (f"Файл {name}.html сохранён, стиль {brief.get('profile', '')}. "
                + " ".join(_visible_text(html).split())[:3500])
    return None

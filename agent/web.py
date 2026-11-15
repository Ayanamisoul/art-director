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

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from config import BACKENDS, DEFAULT_BACKEND
from mcp_client import MCPToolbox
from prompts import SKILLS
from main import agent_reply

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
    try:
        answer = await agent_reply(BACKENDS[state["backend"]], state["messages"],
                                   toolbox, log=tool_log)
    except Exception as exc:
        state["messages"].pop()
        return {"session_id": session_id, "error": str(exc), "tools": tool_log}

    return {"session_id": session_id, "answer": answer, "tools": tool_log,
            "backend": state["backend"]}


@app.get("/api/files")
async def list_files():
    files = sorted(p.name for p in DATA_DIR.iterdir() if p.is_file())
    return {"files": files}

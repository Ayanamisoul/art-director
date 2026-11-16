"""Роутер сессий чатов: список, чтение, удаление, привязка к презентации."""
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from core import sessions

router = APIRouter()


class ForPresentation(BaseModel):
    name: str


@router.get("/api/sessions")
async def list_sessions():
    return {"sessions": sessions.summaries()}


@router.get("/api/session/{session_id}")
async def get_session(session_id: str):
    state = sessions.get(session_id)
    if state is None:
        raise HTTPException(404, "сессия не найдена")
    return {"id": state["id"], "presentation": state.get("presentation"),
            "backend": state.get("backend"), "messages": state["messages"]}


@router.delete("/api/session/{session_id}")
async def delete_session(session_id: str):
    if not sessions.delete(session_id):
        raise HTTPException(404, "сессия не найдена")
    return {"ok": True}


@router.post("/api/session/for-presentation")
async def session_for_presentation(body: ForPresentation):
    """Найти свежий чат по презентации или создать новый (кнопка «💬 чат»)."""
    state = sessions.for_presentation(body.name)
    return {"id": state["id"], "created": state["updated"] == ""}

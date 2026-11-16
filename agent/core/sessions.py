"""Чат-сессии на диске (data/sessions/*.json). Логика без HTTP и MCP.

Одна сессия: {id, backend, skill, messages, presentation, chosen_template,
chosen_format, created, updated}. Данные — только локально (local-first).
"""
import json
import uuid
from datetime import datetime
from pathlib import Path

from config import DEFAULT_BACKEND

ROOT = Path(__file__).resolve().parents[2]
SESSIONS_DIR = ROOT / "data" / "sessions"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)

sessions: dict[str, dict] = {}


def load_all() -> None:
    """Прочитать сессии с диска (вызывается при старте приложения)."""
    for f in SESSIONS_DIR.glob("*.json"):
        try:
            state = json.loads(f.read_text(encoding="utf-8"))
            sessions[state["id"]] = state
        except Exception as exc:  # битый файл не должен ронять сервер
            print(f"сессия {f.name} пропущена: {exc}")


def get(session_id: str) -> dict | None:
    return sessions.get(session_id)


def ensure(session_id: str | None = None) -> dict:
    """Вернуть сессию по id или создать новую."""
    session_id = session_id or uuid.uuid4().hex
    state = sessions.get(session_id)
    if state is None:
        state = {"id": session_id, "backend": DEFAULT_BACKEND, "skill": "chat",
                 "messages": [], "presentation": None}
        sessions[session_id] = state
    return state


def save(state: dict) -> None:
    state["updated"] = datetime.now().isoformat(timespec="seconds")
    (SESSIONS_DIR / f"{state['id']}.json").write_text(
        json.dumps(state, ensure_ascii=False), encoding="utf-8")


def delete(session_id: str) -> bool:
    """Удалить сессию из памяти и с диска. False — если её не было."""
    if sessions.pop(session_id, None) is None:
        return False
    f = SESSIONS_DIR / f"{session_id}.json"
    if f.exists():
        f.unlink()
    return True


def summaries() -> list[dict]:
    """Список сессий для панели: свежие сверху."""
    items = []
    for s in sessions.values():
        first_user = next((m["content"] for m in s["messages"]
                           if m.get("role") == "user" and isinstance(m.get("content"), str)), "")
        items.append({"id": s["id"], "presentation": s.get("presentation"),
                      "updated": s.get("updated", ""),
                      "turns": sum(1 for m in s["messages"] if m.get("role") == "user"),
                      "preview": first_user[:120]})
    items.sort(key=lambda x: x["updated"], reverse=True)
    return items


def for_presentation(name: str, backend: str = "") -> dict:
    """Найти свежий чат по презентации или создать новый (кнопка «💬 чат»)."""
    candidates = [s for s in sessions.values() if s.get("presentation") == name]
    if candidates:
        return max(candidates, key=lambda s: s.get("updated", ""))
    sid = uuid.uuid4().hex
    state = {"id": sid, "backend": backend or DEFAULT_BACKEND, "skill": "chat",
             "messages": [], "presentation": name}
    sessions[sid] = state
    save(state)
    return state

"""Роутер файлов и презентаций: список, чтение, правки, экспорт в PDF."""
from pathlib import Path

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from core import decks
from core.artifacts import DATA_DIR, list_files
from core.previews import render_pptx_previews

router = APIRouter()


class PresentationSave(BaseModel):
    title: str = ""
    subtitle: str = ""
    slides: list[dict]
    theme: dict = {}
    template: str = ""


@router.get("/api/files")
async def api_list_files():
    return {"files": list_files()}


@router.get("/api/presentation/{name}/previews")
async def presentation_previews(name: str):
    """Есть ли настоящие PNG-превью слайдов (рендер из файла) и сколько."""
    out_dir = DATA_DIR / "previews" / decks._safe_name(name)
    count = len(list(out_dir.glob("p*.png"))) if out_dir.exists() else 0
    return {"ready": count > 0, "count": count}


@router.get("/api/presentation/{name}")
async def get_presentation(name: str):
    deck = decks.read_presentation(name)
    if deck is None:
        raise HTTPException(404, f"нет структуры {name}.json — презентация не создана агентом?")
    return deck


@router.post("/api/presentation/{name}")
async def save_presentation(name: str, data: PresentationSave):
    if not all(c in "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-а-яА-ЯёЁ"
               for c in name):
        raise HTTPException(400, "недопустимое имя")
    issues = decks.save_edited(name, data.title, data.subtitle, data.slides,
                               data.theme, data.template)
    if issues:
        return {"ok": False, "issues": issues}
    return {"ok": True, "file": f"/files/{decks._safe_name(name)}.pptx"}


@router.post("/api/presentation/{name}/pdf")
async def export_presentation_pdf(name: str):
    """Экспорт презентации в PDF: data/artifacts/<name>.pdf (PowerPoint COM)."""
    try:
        pdf = decks.export_pdf(name)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc))
    except RuntimeError as exc:
        raise HTTPException(500, str(exc))
    return {"ok": True, "file": f"/files/{pdf.name}"}

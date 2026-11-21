"""Роутер файлов и презентаций: список, чтение, правки, экспорт в PDF."""
import subprocess

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from core import decks
from core.artifacts import DATA_DIR, html_to_pdf, list_files, safe_name
from core.templates import template_path
from pptx_builder import theme_accent_from_template

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
    pptx = decks.pptx_path(name)
    stamp = pptx.stat().st_mtime_ns if pptx.is_file() else 0
    pages = list(out_dir.glob("p*.png"))
    ready = bool(stamp and pages and min(p.stat().st_mtime_ns for p in pages) >= stamp)
    return {"ready": ready, "count": len(pages) if ready else 0, "version": str(stamp)}


@router.get("/api/presentation/{name}")
async def get_presentation(name: str):
    deck = decks.read_presentation(name)
    if deck is None:
        raise HTTPException(404, f"нет структуры {name}.json — презентация не создана агентом?")
    template = (deck.get("theme") or {}).get("template", "")
    if template.lower().endswith(".pptx"):
        try:
            deck["template_accent"] = theme_accent_from_template(template_path(template))
        except FileNotFoundError:
            pass
    return deck


@router.post("/api/presentation/{name}")
async def save_presentation(name: str, data: PresentationSave):
    if name != decks._safe_name(name):
        raise HTTPException(400, "недопустимое имя")
    issues = decks.save_edited(name, data.title, data.subtitle, data.slides,
                               data.theme, data.template)
    if issues:
        return {"ok": False, "issues": issues}
    return {"ok": True, "file": f"/files/{decks._safe_name(name)}.pptx",
            "slides": decks.read_presentation(name)["slides"]}


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


@router.post("/api/landing/{name}/pdf")
def export_landing_pdf(name: str):
    if not (DATA_DIR / f"{safe_name(name)}.html").is_file():
        raise HTTPException(404, "Лендинг не найден")
    try:
        result = html_to_pdf(name)
    except (OSError, subprocess.SubprocessError) as exc:
        raise HTTPException(500, str(exc)) from exc
    if not result.startswith("PDF готов:"):
        raise HTTPException(500, result)
    return {"ok": True, "file": f"/files/{safe_name(name)}.pdf"}

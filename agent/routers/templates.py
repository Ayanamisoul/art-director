"""Роутер шаблонов: список, загрузка (.pptx/.pdf), выбранный шаблон."""
import pathlib

from fastapi import APIRouter, HTTPException, UploadFile
from pydantic import BaseModel

from core import templates

router = APIRouter()


class PreferredTemplate(BaseModel):
    name: str = ""


@router.get("/api/templates/{name}/outline")
async def template_outline(name: str, start: int = 1, count: int = 5):
    try:
        return templates.read_template(name, start, count)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc


@router.get("/api/templates")
async def list_templates():
    return {"files": templates.list_names(), "newest": templates.newest()}


@router.post("/api/templates")
async def upload_template(file: UploadFile):
    name = pathlib.Path(file.filename or "template.pptx").name
    if not name:
        raise HTTPException(400, "нет имени файла")
    try:
        saved, replaced = templates.accept_upload_bytes(name, await file.read())
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    return {"ok": True, "name": saved, "replaced": replaced}


@router.get("/api/templates/preferred")
async def get_preferred_template():
    return {"preferred": templates.preferred()}


@router.post("/api/templates/preferred")
async def set_preferred_template(body: PreferredTemplate):
    """Выбор шаблона из выпадающего списка (пустое имя = авто/самый свежий)."""
    try:
        name = templates.set_preferred(body.name)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc))
    return {"ok": True, "preferred": name}

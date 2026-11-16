"""Роутер картинок пользователя: список, загрузка, удаление."""
from fastapi import APIRouter, HTTPException, UploadFile

from core import assets

router = APIRouter()


@router.get("/api/assets")
async def list_assets():
    return {"files": assets.list_names()}


@router.post("/api/assets")
async def upload_asset(file: UploadFile):
    name = (file.filename or "image.png").replace("/", "_").replace("\\", "_")
    try:
        saved, replaced = assets.accept_upload_bytes(name, await file.read())
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    return {"ok": True, "name": saved, "replaced": replaced}


@router.delete("/api/assets/{name}")
async def delete_asset(name: str):
    if not assets.delete(name):
        raise HTTPException(404, "картинка не найдена")
    return {"ok": True}

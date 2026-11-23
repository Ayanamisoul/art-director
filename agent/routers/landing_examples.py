"""Роутер примеров лендингов: список, загрузка HTML-образца оформления."""
import asyncio
import pathlib

from fastapi import APIRouter, HTTPException, UploadFile

from core import landing_design, landing_examples

router = APIRouter()


@router.get("/api/landing/examples")
async def list_landing_examples():
    return {"files": landing_examples.list_names(), "newest": landing_examples.newest()}


@router.post("/api/landing/examples")
async def upload_landing_example(file: UploadFile):
    name = pathlib.Path(file.filename or "example.html").name
    if not name:
        raise HTTPException(400, "нет имени файла")
    data = await file.read()
    try:
        result = await asyncio.to_thread(
            landing_examples.accept_upload_bytes, name, data,
            landing_design.model_designer())
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    return {"ok": True, **result}

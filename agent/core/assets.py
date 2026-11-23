"""Картинки пользователя (data/assets/): загрузка, список, удаление."""
import hashlib
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import fitz
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
ASSETS_DIR = ROOT / "data" / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg", ".bmp"}


def list_names() -> list[str]:
    return sorted(p.name for p in ASSETS_DIR.iterdir() if p.is_file())


def save_slide_upload(filename: str, data: bytes) -> str:
    """Привязать изображение к брифу, сохранив отдельные версии одноимённых файлов."""
    source = Path(filename.replace("\\", "/")).name
    suffix = Path(source).suffix.lower()
    if suffix not in IMAGE_EXTS:
        raise ValueError(f"«{source}»: выберите файл изображения.")
    digest = hashlib.sha256(data).hexdigest()[:16]
    try:
        if suffix == ".svg":
            with fitz.open(stream=data, filetype="svg") as doc:
                page = doc[0]
                scale = 1600 / max(page.rect.width, page.rect.height)
                data = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=True).tobytes("png")
            source += ".png"
        else:
            with Image.open(BytesIO(data)) as image:
                image.verify()
    except (OSError, ValueError, RuntimeError, ZeroDivisionError, Image.DecompressionBombError) as exc:
        raise ValueError(f"Не удалось прочитать изображение «{source}».") from exc
    target = ASSETS_DIR / f"{digest}_{source}"
    if not target.exists():
        with TemporaryDirectory(dir=ASSETS_DIR) as temp:
            staged = Path(temp) / "image"
            staged.write_bytes(data)
            staged.replace(target)
    return target.name


def accept_upload_bytes(name: str, data: bytes) -> tuple[str, bool]:
    """Сохранить картинку (инвариант имён без регистра). (имя, перезаписана?)"""
    if Path(name).suffix.lower() not in IMAGE_EXTS:
        raise ValueError(f"«{name}» — не изображение. Форматы: "
                         f"{', '.join(sorted(e.lstrip('.') for e in IMAGE_EXTS))}")
    for p in ASSETS_DIR.iterdir():
        if p.name.lower() == name.lower():
            name = p.name
            break
    replaced = (ASSETS_DIR / name).exists()
    (ASSETS_DIR / name).write_bytes(data)
    return name, replaced


def delete(name: str) -> bool:
    """Удалить картинку по имени. False — если её не было."""
    p = ASSETS_DIR / Path(name).name
    if not p.exists():
        return False
    p.unlink()
    return True


def text_list() -> str:
    """Текстовый список для MCP-инструмента list_images."""
    files = list_names()
    if not files:
        return "Пользователь пока не загрузил изображений."
    return ("Изображения пользователя (СТАВЬ НА СЛАЙДЫ ТОЛЬКО ЕСЛИ ПОЛЬЗОВАТЕЛЬ "
            "ЯВНО ПРОСИЛ ИМЕННО ИХ ИСПОЛЬЗОВАТЬ):\n" + "\n".join(files))

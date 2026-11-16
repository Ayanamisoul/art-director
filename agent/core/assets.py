"""Картинки пользователя (data/assets/): загрузка, список, удаление."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
ASSETS_DIR = ROOT / "data" / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg", ".bmp"}


def list_names() -> list[str]:
    return sorted(p.name for p in ASSETS_DIR.iterdir() if p.is_file())


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

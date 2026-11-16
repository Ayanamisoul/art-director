"""Шаблоны презентаций пользователя (.pptx и .pdf): реестр, выбор, загрузка."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TEMPLATES_DIR = ROOT / "data" / "templates"
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

TEMPLATE_EXTS = {".pptx", ".pdf"}
_PREFERRED = ".preferred"


def list_names() -> list[str]:
    return sorted(p.name for p in TEMPLATES_DIR.iterdir()
                  if p.is_file() and p.suffix.lower() in TEMPLATE_EXTS)


def newest() -> str:
    """Имя самого свежего шаблона (по времени изменения файла)."""
    files = [p for p in TEMPLATES_DIR.iterdir()
             if p.is_file() and p.suffix.lower() in TEMPLATE_EXTS]
    files.sort(key=lambda p: p.stat().st_mtime)
    return files[-1].name if files else ""


def preferred() -> str:
    """Шаблон, выбранный пользователем (hint-файл), иначе самый свежий.
    Гарант: даже если модель забыла template, применён именно этот."""
    try:
        name = (TEMPLATES_DIR / _PREFERRED).read_text(encoding="utf-8").strip()
        if name and (TEMPLATES_DIR / name).exists():
            return name
    except OSError:
        pass
    return newest()


def set_preferred(name: str) -> str:
    name = name.strip()
    if name:
        if not (TEMPLATES_DIR / Path(name).name).exists():
            raise FileNotFoundError(f"шаблон «{name}» не найден")
        name = Path(name).name
    (TEMPLATES_DIR / _PREFERRED).write_text(name, encoding="utf-8")
    return name


def accept_upload_bytes(name: str, data: bytes) -> tuple[str, bool]:
    """Сохранить загруженный шаблон. Возвращает (имя, был_ли_перезаписан).
    Новый шаблон автоматически становится выбранным."""
    validate_upload_name(name)
    path = TEMPLATES_DIR / name
    replaced = path.exists()
    path.write_bytes(data)
    set_preferred(name)
    return name, replaced


def validate_upload_name(name: str) -> None:
    if Path(name).suffix.lower() not in TEMPLATE_EXTS:
        raise ValueError(f"«{name}» — не шаблон. Поддерживаются .pptx и .pdf")


def text_list() -> str:
    """Текстовый список для MCP-инструмента list_templates."""
    files = [p for p in sorted(TEMPLATES_DIR.iterdir(), key=lambda p: p.stat().st_mtime)
             if p.suffix.lower() in TEMPLATE_EXTS]
    if not files:
        return "Пользователь пока не загрузил шаблонов."
    return ("Шаблоны пользователя (самый свежий — последний, используй его "
            "в save_presentation параметром template; .pdf — страницы станут "
            "фоном слайдов):\n" + "\n".join(p.name for p in files))

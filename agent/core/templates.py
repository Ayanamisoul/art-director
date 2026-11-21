"""Шаблоны презентаций пользователя (.pptx и .pdf): реестр, выбор, загрузка."""
import tempfile
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile

import fitz
from pptx import Presentation
from pptx.exc import PackageNotFoundError

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
        if name == "__none__" or name in list_names():
            return name
    except OSError:
        pass
    return newest()


def set_preferred(name: str) -> str:
    name = name.strip()
    if name and name != "__none__":
        if name not in list_names():
            raise FileNotFoundError(f"шаблон «{name}» не найден")
        name = Path(name).name
    (TEMPLATES_DIR / _PREFERRED).write_text(name, encoding="utf-8")
    return name


def accept_upload_bytes(name: str, data: bytes) -> tuple[str, bool]:
    """Сохранить загруженный шаблон. Возвращает (имя, был_ли_перезаписан).
    Новый шаблон автоматически становится выбранным."""
    validate_upload_name(name)
    path = TEMPLATES_DIR / Path(name).name
    try:
        if path.suffix.lower() == ".pdf":
            with fitz.open(stream=data, filetype="pdf") as pdf:
                if pdf.needs_pass or not len(pdf):
                    raise ValueError("Для шаблона нужен PDF со страницами, доступный без пароля")
        else:
            Presentation(BytesIO(data))
    except (BadZipFile, KeyError, PackageNotFoundError, fitz.FileDataError, SyntaxError) as exc:
        raise ValueError(f"Не удалось прочитать «{path.name}». Проверьте формат файла.") from exc
    replaced = path.exists()
    with tempfile.TemporaryDirectory(dir=TEMPLATES_DIR) as folder:
        staged = Path(folder) / path.name
        staged.write_bytes(data)
        staged.replace(path)
    set_preferred(path.name)
    return path.name, replaced


def validate_upload_name(name: str) -> None:
    if Path(name).suffix.lower() not in TEMPLATE_EXTS:
        raise ValueError(f"«{name}» — не шаблон. Поддерживаются .pptx и .pdf")


def template_path(name: str) -> Path:
    validate_upload_name(name)
    path = TEMPLATES_DIR / Path(name).name
    if not path.is_file():
        raise FileNotFoundError(f"Шаблон «{path.name}» не найден")
    return path


def read_template(name: str, start: int = 1, count: int = 5) -> dict:
    """Текст и адреса полей исходника. Номера слайдов начинаются с единицы."""
    from pptx_builder import template_text_shapes

    path = template_path(name)
    if start < 1 or not 1 <= count <= 20:
        raise ValueError("start должен быть ≥ 1, count — от 1 до 20")

    slides = []
    if path.suffix.lower() == ".pdf":
        with fitz.open(path) as pdf:
            total = len(pdf)
            for i in range(start - 1, min(total, start - 1 + count)):
                slides.append({"source_slide": i + 1, "text": pdf[i].get_text()[:6000]})
    else:
        deck = Presentation(path)
        total = len(deck.slides)
        for i in range(start - 1, min(total, start - 1 + count)):
            fields = {str(s.shape_id): s.text for s in template_text_shapes(deck.slides[i].shapes)
                      if s.text or s.is_placeholder}
            slides.append({"source_slide": i + 1, "fields": fields})
    return {"name": path.name, "total": total, "slides": slides}


def text_list() -> str:
    """Текстовый список для MCP-инструмента list_templates."""
    files = [p for p in sorted(TEMPLATES_DIR.iterdir(), key=lambda p: p.stat().st_mtime)
             if p.suffix.lower() in TEMPLATE_EXTS]
    if not files:
        return "Пользователь пока не загрузил шаблонов."
    return ("Шаблоны пользователя (самый свежий — последний, используй его "
            "в save_presentation параметром template; .pdf — страницы станут "
            "фоном слайдов):\n" + "\n".join(p.name for p in files))

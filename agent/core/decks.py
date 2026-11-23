"""Презентации: сборка, правки, чтение, экспорт в PDF.

Единая точка входа для агента (через MCP) и веб-редактора — раньше логика
дублировалась в mcp_servers/artifacts.py и web.py.
"""
import json
import subprocess
import tempfile
from pathlib import Path

from core.artifacts import DATA_DIR
from core.assets import ASSETS_DIR
from core.fonts import font_path
from core.previews import render_pptx_previews
from core.templates import newest, preferred as preferred_template, template_path
from pptx_builder import build_presentation, normalize_slide, validate_presentation
from slide_content import paginate_slide, split_for_space
from slide_layout import SlideFitError, measure_with

SCRIPT = Path(__file__).resolve().parents[2] / "scripts" / "pptx2pdf.ps1"




def _safe_name(name: str) -> str:
    import re
    name = re.sub(r"\.(pptx|json|pdf)$", "", name.strip(), flags=re.IGNORECASE)
    slug = re.sub(r"[^\w\-]", "_", name)
    return slug or "presentation"


def sidecar_path(name: str) -> Path:
    return DATA_DIR / f"{_safe_name(name)}.json"


def pptx_path(name: str) -> Path:
    return DATA_DIR / f"{_safe_name(name)}.pptx"


def picture_names(name: str) -> set[str]:
    from pptx import Presentation
    deck = Presentation(pptx_path(name))
    return {node.get("name") for slide in deck.slides for node in
            slide._element.xpath(".//p:pic/p:nvPicPr/p:cNvPr")}


def read_presentation(name: str) -> dict | None:
    """Структура презентации из сайдкара (None — если её нет)."""
    sidecar = sidecar_path(name)
    if not sidecar.exists():
        return None
    return json.loads(sidecar.read_text(encoding="utf-8"))


def get_presentation_text(name: str) -> str:
    """Строка для MCP-инструмента get_presentation."""
    sidecar = sidecar_path(name)
    if not sidecar.exists():
        return (f"Презентация «{name}» не найдена. Вызови list_artifacts, "
                f"чтобы увидеть, что уже сохранено.")
    return sidecar.read_text(encoding="utf-8")


def save_edited(name: str, title: str, subtitle: str, slides: list[dict],
                theme: dict | None = None, template: str = "",
                template_images: str | None = None) -> list[str]:
    """Сохранить презентацию из веб-редактора. Возвращает список замечаний."""
    try:
        pages = [page for s in slides for page in paginate_slide(normalize_slide(s))]
    except ValueError as exc:
        return [str(exc)]
    for page in pages:
        if page.get("kind") == "title" and subtitle and not page.get("subtitle"):
            page["subtitle"] = subtitle
    payload = {"title": title, "subtitle": subtitle, "slides": pages}
    old = read_presentation(name) or {}
    theme = {**(old.get("theme") or {}), **(theme or {})}
    try:
        custom_font = font_path(theme["font_id"]) if theme.get("font_id") else None
    except ValueError as exc:
        return [str(exc)]
    template_name = template or theme.get("template")
    if template_name == "__none__":
        theme.pop("template", None)
        template_name = None
    source = None
    if template_name:
        try:
            source = template_path(template_name)
        except (ValueError, FileNotFoundError) as exc:
            return [str(exc)]
        theme["template"] = source.name
    if source and source.suffix.lower() == ".pdf":
        theme["template_images"] = template_images or theme.get("template_images") or "auto"
    else:
        theme.pop("template_images", None)
    path = pptx_path(name)
    while True:
        issues = validate_presentation(payload)
        if issues:
            return issues
        try:
            with measure_with(custom_font):
                build_presentation(path, payload, theme or None, template=source,
                                   template_images=theme.get("template_images", "auto"))
            break
        except SlideFitError as exc:
            idx = exc.index
            replacement = split_for_space(pages[idx]) if idx is not None else []
            if not replacement:
                where = f"Слайд {idx+1} «{pages[idx].get('title','')}». " if idx is not None else ""
                return [where + str(exc)]
            pages[idx:idx+1] = replacement
        except ValueError as exc:
            return [str(exc)]
    path.with_suffix(".json").write_text(
        json.dumps({**payload, "theme": theme}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    render_pptx_previews(path)
    return []


def agent_save_presentation(name: str, title: str, slides: list[dict],
                            subtitle: str = "", accent_color: str = "",
                            template: str = "", template_images: str | None = None,
                            font_id: str = "") -> str:
    """Сборка по структуре от модели. Строка-ответ для модели (как в MCP-инструменте).
    Если template не передан — применяется выбранный пользователем (или свежий)."""
    if template == "__auto__":
        template = "" if sidecar_path(name).exists() else newest() or "__none__"
    elif not template and not sidecar_path(name).exists():
        template = preferred_template()
    theme = {"accent": accent_color.lstrip("#")} if accent_color else {}
    if font_id:
        theme["font_id"] = font_id
    issues = save_edited(name, title, subtitle, slides, theme, template, template_images)
    if issues:
        return ("Презентация НЕ сохранена. Исправь замечания и вызови save_presentation снова:\n- "
                + "\n- ".join(issues))
    saved = read_presentation(name)
    theme = saved["theme"]
    template_name = theme.get("template")
    path = pptx_path(name)
    note = f", по шаблону {template_name}" if template_name else ""
    if template_name and Path(template_name).suffix.lower() == ".pdf":
        note += " (без фото)" if theme.get("template_images") in ("remove", "all") else ""
    return f"Презентация сохранена и прошла проверку: {path} ({len(saved['slides'])} слайдов{note})"


def export_pdf(name: str) -> Path:
    """Конвертировать .pptx в PDF (PowerPoint COM). Путь к PDF."""
    pptx = pptx_path(name)
    if not pptx.exists():
        raise FileNotFoundError(f"презентация {pptx} не найдена")
    pdf_path = DATA_DIR / f"{_safe_name(name)}.pdf"
    with tempfile.TemporaryDirectory(dir=DATA_DIR) as folder:
        staged_pdf = Path(folder) / "slides.pdf"
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
             "-in", str(pptx.resolve()), "-out", str(staged_pdf.resolve())],
            capture_output=True, timeout=180)
        if result.returncode or not staged_pdf.is_file():
            err = result.stderr.decode("utf-8", errors="replace")[:300]
            raise RuntimeError(f"не удалось конвертировать в PDF: {err}")
        staged_pdf.replace(pdf_path)
    return pdf_path

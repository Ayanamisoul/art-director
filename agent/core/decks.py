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
from core.previews import render_pptx_previews
from core.templates import TEMPLATES_DIR, preferred as preferred_template
from pptx_builder import build_presentation, normalize_slide, validate_presentation

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


def _carry_over_template(name: str, theme: dict | None, template: str) -> tuple[Path | None, str]:
    """Шаблон для пересборки: явный > из theme > из прошлой сборки."""
    template_name = template or (theme or {}).get("template", "")
    if not template_name:
        old = sidecar_path(name)
        if old.exists():
            try:
                template_name = (json.loads(old.read_text(encoding="utf-8"))
                                 .get("theme") or {}).get("template", "")
            except Exception:
                pass
    path = None
    if template_name:
        candidate = TEMPLATES_DIR / Path(template_name).name
        if candidate.exists():
            path = candidate
            template_name = candidate.name
    return path, template_name


def save_edited(name: str, title: str, subtitle: str, slides: list[dict],
                theme: dict | None = None, template: str = "",
                template_images: str = "auto") -> list[str]:
    """Сохранить презентацию из веб-редактора. Возвращает список замечаний."""
    payload = {"title": title, "subtitle": subtitle,
               "slides": [normalize_slide(s) for s in slides]}
    issues = validate_presentation(payload)
    if issues:
        return issues
    theme = theme or {}
    template_path, template_name = _carry_over_template(name, theme, template)
    if template_name:
        theme["template"] = template_name
    path = pptx_path(name)
    build_presentation(path, payload, theme or None, template=template_path,
                       template_images=template_images)
    path.with_suffix(".json").write_text(
        json.dumps({**payload, "theme": theme}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    render_pptx_previews(path)
    return []


def agent_save_presentation(name: str, title: str, slides: list[dict],
                            subtitle: str = "", accent_color: str = "",
                            template: str = "",  # игнорируется: шаблон выбирает пользователь
                            template_images: str = "auto") -> str:
    """Сборка по структуре от модели. Строка-ответ для модели (как в MCP-инструменте).
    Если template не передан — применяется выбранный пользователем (или свежий)."""
    data = {"title": title, "subtitle": subtitle,
            "slides": [normalize_slide(s) for s in slides]}
    issues = validate_presentation(data)
    # шаблон выбирает ПОЛЬЗОВАТЕЛЬ (.preferred в UI); параметр template от модели
    # сознательно игнорируется — иначе модель угадывает шаблон по имени файла
    template_name = preferred_template()
    template_path = TEMPLATES_DIR / template_name if template_name else None
    if issues:
        return ("Презентация НЕ сохранена. Исправь замечания и вызови save_presentation снова:\n- "
                + "\n- ".join(issues))
    theme = {"accent": accent_color.lstrip("#")} if accent_color else {}
    if template_name:
        theme["template"] = template_name
    path = pptx_path(name)
    build_presentation(path, data, theme or None, template=template_path,
                       template_images=template_images)
    path.with_suffix(".json").write_text(
        json.dumps({**data, "theme": theme}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    note = f", по шаблону {template_name}" if template_name else ""
    if template_name and Path(template_name).suffix.lower() == ".pdf":
        note += " (без фото)" if template_images in ("remove", "all") else ""
    previews = render_pptx_previews(path)
    if previews:
        note += f", готовы превью ({previews} слайдов)"
    return f"Презентация сохранена и прошла проверку: {path} ({len(data['slides'])} слайдов{note})"


def export_pdf(name: str) -> Path:
    """Конвертировать .pptx в PDF (PowerPoint COM). Путь к PDF."""
    pptx = pptx_path(name)
    if not pptx.exists():
        raise FileNotFoundError(f"презентация {pptx} не найдена")
    pdf_path = DATA_DIR / f"{_safe_name(name)}.pdf"
    with tempfile.TemporaryDirectory() as profile:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
             "-in", str(pptx.resolve()), "-out", str(pdf_path.resolve())],
            capture_output=True, timeout=180)
    if not pdf_path.exists():
        err = result.stderr.decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"не удалось конвертировать в PDF: {err}")
    return pdf_path

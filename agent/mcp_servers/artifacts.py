"""MCP-сервер «artifacts» — первые инструменты агента.

Инструменты (их увидит LLM через MCP-клиент и сможет вызывать):
  save_html(name, html)   — сохранить сгенерированный HTML на диск пользователя
  html_to_pdf(name)       — сконвертировать сохранённый HTML в PDF (headless Chrome)
  list_artifacts()        — показать, что уже сохранено

Local-first: всё пишется в <репо>/data/artifacts/ на машине пользователя.
Сервер ничего не отправляет наружу и ничего не хранит «у себя».

Запуск (stdio-транспорт, сервер поднимается клиентом на каждую сессию):
  py agent/mcp_servers/artifacts.py
"""
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from mcp.server.mcpserver import MCPServer

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # доступ к agent/
from pptx_builder import build_presentation, normalize_slide, validate_presentation
from validation import validate_html

# <репо>/data/artifacts/ — вне git (см. .gitignore)
DATA_DIR = Path(__file__).resolve().parents[2] / "data" / "artifacts"
DATA_DIR.mkdir(parents=True, exist_ok=True)

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

mcp = MCPServer("artifacts")


def _safe_name(name: str) -> str:
    """Только буквы/цифры/дефис/подчёркивание — защита от путей вида ../..
    Плюс срезаем расширение, если модель передала его вопреки docstring'у."""
    name = re.sub(r"\.(html?|pdf)$", "", name.strip(), flags=re.IGNORECASE)
    slug = re.sub(r"[^\w\-]", "_", name)
    return slug or "artifact"


@mcp.tool()
def save_html(name: str, html: str) -> str:
    """Сохранить HTML-страницу на диск. name — короткое имя без расширения,
    html — полный код страницы. Перед сохранением страница проходит проверку
    качества: если есть замечания, они возвращаются в ответе — исправь их
    и вызови save_html ещё раз с исправленным кодом."""
    path = DATA_DIR / f"{_safe_name(name)}.html"
    path.write_text(html, encoding="utf-8")
    issues = validate_html(html)
    if issues:
        return (f"Сохранено: {path}. НО есть замечания валидации:\n- "
                + "\n- ".join(issues)
                + "\nИсправь их и сохрани заново тем же именем.")
    return f"Сохранено и прошло проверку: {path}"


@mcp.tool()
def html_to_pdf(name: str) -> str:
    """Конвертировать ранее сохранённый HTML (save_html) в PDF.
    name — то же имя, что передавали в save_html. Возвращает путь к PDF."""
    html_path = DATA_DIR / f"{_safe_name(name)}.html"
    if not html_path.exists():
        return f"Ошибка: файл {html_path} не найден. Сначала вызови save_html."
    pdf_path = html_path.with_suffix(".pdf")
    # Отдельный user-data-dir, чтобы не конфликтовать с открытым Chrome пользователя
    with tempfile.TemporaryDirectory() as profile:
        result = subprocess.run(
            [
                CHROME,
                "--headless=new",
                "--disable-gpu",
                f"--user-data-dir={profile}",
                f"--print-to-pdf={pdf_path}",
                "--no-pdf-header-footer",
                html_path.as_uri(),
            ],
            capture_output=True,
            timeout=120,
        )
    if not pdf_path.exists():
        return f"Ошибка конвертации: {result.stderr.decode('utf-8', errors='replace')[:300]}"
    return f"PDF готов: {pdf_path}"


@mcp.tool()
def save_presentation(name: str, title: str, slides: list[dict], subtitle: str = "",
                      accent_color: str = "") -> str:
    """Собрать и сохранить презентацию .pptx (открывается и редактируется в PowerPoint).
    name — короткое имя файла; title/subtitle — заголовок презентации;
    slides — список слайдов, каждый: {"kind": "title|text|section|closing",
    "title": "заголовок слайда", "bullets": ["пункт 1", "пункт 2"]}.
    Поле с пунктами называется СТРОГО "bullets" (список строк).
    Правила: первый слайд kind=title, последний kind=closing; на текстовом слайде
    до 6 пунктов, каждый до 140 символов. accent_color — hex бренда, напр. "2E5E4E".
    Перед сохранением структура проверяется: при замечаниях файл НЕ создаётся —
    исправь их и вызови инструмент снова."""
    data = {"title": title, "subtitle": subtitle, "slides": slides}
    data["slides"] = [normalize_slide(s) for s in slides]
    issues = validate_presentation(data)
    if issues:
        return ("Презентация НЕ сохранена. Исправь замечания и вызови save_presentation снова:\n- "
                + "\n- ".join(issues))
    theme = {"accent": accent_color.lstrip("#")} if accent_color else None
    path = DATA_DIR / f"{_safe_name(name)}.pptx"
    build_presentation(path, data, theme)
    # sidecar JSON: структура презентации для веб-редактора и правок через чат
    import json as _json
    path.with_suffix(".json").write_text(
        _json.dumps({**data, "theme": theme or {}}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    return f"Презентация сохранена и прошла проверку: {path} ({len(slides)} слайдов)"


@mcp.tool()
def list_artifacts() -> str:
    """Показать список сохранённых файлов (HTML и PDF)."""
    files = sorted(p.name for p in DATA_DIR.iterdir() if p.is_file())
    return "Сохранённые файлы:\n" + "\n".join(files) if files else "Пока ничего не сохранено."


if __name__ == "__main__":
    mcp.run()  # stdio: общается с MCP-клиентом через stdin/stdout

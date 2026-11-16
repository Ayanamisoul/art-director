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
ASSETS_DIR = Path(__file__).resolve().parents[2] / "data" / "assets"
TEMPLATES_DIR = Path(__file__).resolve().parents[2] / "data" / "templates"
for d in (DATA_DIR, ASSETS_DIR, TEMPLATES_DIR):
    d.mkdir(parents=True, exist_ok=True)

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
                      accent_color: str = "", template: str = "") -> str:
    """Собрать и сохранить презентацию .pptx (открывается и редактируется в PowerPoint).
    name — короткое имя файла; title/subtitle — заголовок презентации;
    slides — список слайдов, каждый: {"kind": "title|text|section|closing",
    "title": "заголовок слайда", "bullets": ["пункт 1", "пункт 2"],
    "image": "имя файла картинки из list_images" (необязательно, для text-слайда)}.
    Поле с пунктами называется СТРОГО "bullets" (список строк).
    Правила: первый слайд kind=title, последний kind=closing; на текстовом слайде
    до 6 пунктов, каждый до 140 символов. accent_color — hex бренда, напр. "2E5E4E".
    template — имя .pptx-шаблона из list_templates: тема/цвета возьмутся из него.
    Если у пользователя загружены шаблоны — ОБЯЗАТЕЛЬНО передай в template самый
    свежий из них (последний в списке list_templates), если пользователь не попросил другой.
    Для правки существующей презентации: сначала get_presentation с её именем,
    затем save_presentation с ТЕМ ЖЕ именем и исправленной структурой.
    Перед сохранением структура проверяется: при замечаниях файл НЕ создаётся —
    исправь их и вызови инструмент снова."""
    data = {"title": title, "subtitle": subtitle, "slides": slides}
    data["slides"] = [normalize_slide(s) for s in slides]
    issues = validate_presentation(data)
    template_name = ""
    template_path = None
    if template:
        template_name = Path(template).name
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
            issues.append(f"шаблон «{template}» не найден — вызови list_templates")
    elif template_name := _newest_template():
        # у пользователя есть шаблоны, но модель не передала template —
        # применяем самый свежий автоматически (brand-стиль важнее «дефолтного» зелёного)
        template_path = TEMPLATES_DIR / template_name
    if issues:
        return ("Презентация НЕ сохранена. Исправь замечания и вызови save_presentation снова:\n- "
                + "\n- ".join(issues))
    theme = {"accent": accent_color.lstrip("#")} if accent_color else {}
    if template_name:
        theme["template"] = template_name
    path = DATA_DIR / f"{_safe_name(name)}.pptx"
    build_presentation(path, data, theme or None, template=template_path)
    # sidecar JSON: структура презентации для веб-редактора и правок через чат
    import json as _json
    path.with_suffix(".json").write_text(
        _json.dumps({**data, "theme": theme}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    note = f", по шаблону {template_name}" if template_name else ""
    return f"Презентация сохранена и прошла проверку: {path} ({len(slides)} слайдов{note})"


def _newest_template() -> str:
    """Имя самого свежего .pptx-шаблона пользователя (или пустая строка)."""
    candidates = sorted(TEMPLATES_DIR.glob("*.pptx"), key=lambda p: p.stat().st_mtime)
    return candidates[-1].name if candidates else ""


@mcp.tool()
def get_presentation(name: str) -> str:
    """Прочитать текущую структуру презентации (JSON: title, subtitle, slides,
    theme). Вызови ПЕРЕД правками существующей презентации, чтобы не потерять
    слайды: измени нужные слайды в полученной структуре и сохрани всё целиком
    через save_presentation с тем же именем."""
    sidecar = DATA_DIR / f"{_safe_name(name)}.json"
    if not sidecar.exists():
        return (f"Презентация «{name}» не найдена. Вызови list_artifacts, чтобы увидеть, "
                f"что уже сохранено.")
    return sidecar.read_text(encoding="utf-8")


@mcp.tool()
def list_images() -> str:
    """Список изображений, загруженных пользователем (логотипы, фото).
    Их имена можно ставить в поле "image" текстового слайда."""
    files = sorted(p.name for p in ASSETS_DIR.iterdir() if p.is_file())
    return "Изображения пользователя:\n" + "\n".join(files) if files else \
        "Пользователь пока не загрузил изображений."


@mcp.tool()
def list_templates() -> str:
    """Список .pptx-шаблонов, загруженных пользователем (свежий — последним).
    Имя шаблона нужно передать в save_presentation параметром template.
    Если шаблоны есть — создавай презентации по самому свежему."""
    files = sorted(TEMPLATES_DIR.glob("*.pptx"), key=lambda p: p.stat().st_mtime)
    names = [p.name for p in files]
    if not names:
        return "Пользователь пока не загрузил шаблонов."
    return ("Шаблоны пользователя (самый свежий — последний, используй его "
            "в save_presentation параметром template):\n" + "\n".join(names))


@mcp.tool()
def list_artifacts() -> str:
    """Показать список сохранённых файлов (HTML и PDF)."""
    files = sorted(p.name for p in DATA_DIR.iterdir() if p.is_file())
    return "Сохранённые файлы:\n" + "\n".join(files) if files else "Пока ничего не сохранено."


if __name__ == "__main__":
    mcp.run()  # stdio: общается с MCP-клиентом через stdin/stdout

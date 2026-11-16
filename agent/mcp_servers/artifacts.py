"""MCP-сервер «artifacts» — ТОНКИЙ адаптер над core/.

Раньше здесь жила вся логика; теперь инструменты только пробрасывают вызовы
в core/ (decks, templates, assets, artifacts). MCP нужен как отдельный канал
для CLI и как требование заказчика; веб-приложение от него не зависит.

Запуск (stdio-транспорт, сервер поднимается клиентом на каждую сессию):
  py agent/mcp_servers/artifacts.py
"""
import sys
from pathlib import Path

from mcp.server.mcpserver import MCPServer

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # доступ к agent/
from core import assets, artifacts, decks, templates
from core.previews import render_pptx_previews  # noqa: F401 (используется через decks)

mcp = MCPServer("artifacts")


@mcp.tool()
def save_html(name: str, html: str) -> str:
    """Сохранить HTML-страницу на диск. name — короткое имя без расширения,
    html — полный код страницы. Перед сохранением страница проходит проверку
    качества: если есть замечания, они возвращаются в ответе — исправь их
    и вызови save_html ещё раз с исправленным кодом."""
    return artifacts.save_html(name, html)


@mcp.tool()
def html_to_pdf(name: str) -> str:
    """Конвертировать ранее сохранённый HTML (save_html) в PDF.
    name — то же имя, что передавали в save_html. Возвращает путь к PDF."""
    return artifacts.html_to_pdf(name)


@mcp.tool()
def save_presentation(name: str, title: str, slides: list[dict], subtitle: str = "",
                      accent_color: str = "", template_images: str = "auto") -> str:
    """Собрать и сохранить презентацию .pptx (открывается и редактируется в PowerPoint).
    name — короткое имя файла; title/subtitle — заголовок презентации;
    slides — список слайдов, каждый: {"kind": "title|text|section|closing",
    "title": "заголовок слайда", "bullets": ["пункт 1", "пункт 2"],
    "image": "имя файла картинки из list_images" (необязательно, для text-слайда,
    СТАВЬ ТОЛЬКО ЕСЛИ ПОЛЬЗОВАТЕЛЬ ЯВНО ПРОСИЛ КАРТИНКУ)}.
    Поле с пунктами называется СТРОГО "bullets" (список строк).
    Правила: первый слайд kind=title, последний kind=closing; на текстовом слайде
    до 6 пунктов, каждый до 140 символов; один смысловой блок = один слайд,
    длинный текст дели по темам, не вываливай всё в один слайд.
    accent_color — hex бренда, напр. "2E5E4E".
    Шаблон выбирает пользователь в интерфейсе — параметр template передавать НЕ НУЖНО:
    система применит выбранный шаблон автоматически.
    Для правки существующей презентации: сначала get_presentation с её именем,
    затем save_presentation с ТЕМ ЖЕ именем и исправленной структурой.
    Перед сохранением структура проверяется: при замечаниях файл НЕ создаётся —
    исправь их и вызови инструмент снова."""
    return decks.agent_save_presentation(name, title, slides, subtitle,
                                         accent_color, template_images=template_images)


@mcp.tool()
def get_presentation(name: str) -> str:
    """Прочитать текущую структуру презентации (JSON: title, subtitle, slides,
    theme). Вызови ПЕРЕД правками существующей презентации, чтобы не потерять
    слайды: измени нужные слайды в полученной структуре и сохрани всё целиком
    через save_presentation с тем же именем."""
    return decks.get_presentation_text(name)


@mcp.tool()
def list_images() -> str:
    """Список изображений, загруженных пользователем (логотипы, фото).
    Их имена можно ставить в поле "image" текстового слайда — но только если
    пользователь явно попросил использовать картинку."""
    return assets.text_list()


@mcp.tool()
def list_templates() -> str:
    """Список .pptx-шаблонов, загруженных пользователем (свежий — последним). Имя шаблона нужно передать в save_presentation параметром
    template. Если шаблоны есть — создавай презентации по самому свежему
    (или выбранному пользователем, см. контекст)."""
    return templates.text_list()


@mcp.tool()
def list_artifacts() -> str:
    """Показать список сохранённых файлов (HTML, PDF, PPTX)."""
    return artifacts.list_text()


if __name__ == "__main__":
    mcp.run()  # stdio: общается с MCP-клиентом через stdin/stdout

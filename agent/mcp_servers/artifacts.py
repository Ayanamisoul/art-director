"""MCP-сервер «artifacts» — ТОНКИЙ адаптер над core/.

Раньше здесь жила вся логика; теперь инструменты только пробрасывают вызовы
в core/ (decks, templates, assets, artifacts). MCP нужен как отдельный канал
для CLI и как требование заказчика; веб-приложение от него не зависит.

Запуск (stdio-транспорт, сервер поднимается клиентом на каждую сессию):
  py agent/mcp_servers/artifacts.py
"""
import json
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
def save_landing(name: str, title: str, lead: str, sections: list[dict], profile: str = "",
                 brand: str = "", cta: str = "", image: str = "") -> str:
    """Создать автономный HTML-лендинг по образцу. Передавай контент, оформление собирается локально.
    sections — разделы вида {title, text, items: [строки]}. lead — подзаголовок.
    В разделе можно передать subsections — список {title, text, items: [строки]}.
    profile — glavbuh (Главбух), school (Школа юриста), law (Нейропомощник юриста).
    Пустое значение сохраняет профиль при правке. brand, cta (кнопка перехода к разделу)
    и image (имя загруженного изображения) необязательны. Для правки сначала get_landing."""
    return artifacts.save_landing(name, title, lead, sections, profile, brand, cta, image)


@mcp.tool()
def get_landing(name: str) -> str:
    """Прочитать исходные тексты и стиль сохранённого лендинга для точечной правки."""
    brief = artifacts.read_landing(name)
    if brief is None:
        return "Структура лендинга недоступна. Проверь имя через list_artifacts."
    return json.dumps(brief, ensure_ascii=False)


@mcp.tool()
def html_to_pdf(name: str) -> str:
    """Конвертировать ранее сохранённый HTML (save_html) в PDF.
    name — то же имя, что передавали в save_html. Возвращает путь к PDF."""
    return artifacts.html_to_pdf(name)


@mcp.tool()
def save_presentation(name: str, title: str, slides: list[dict], subtitle: str = "",
                      accent_color: str = "", template_images: str | None = None,
                      template: str = "") -> str:
    """Собрать и сохранить презентацию .pptx (открывается и редактируется в PowerPoint).
    name — короткое имя файла; title/subtitle — заголовок презентации;
    slides — список слайдов, каждый: {"kind": "title|text|section|closing",
    "title": "заголовок слайда", "bullets": ["пункт 1", "пункт 2"],
    "image": "имя файла из list_images" (СТОРОГО ПО ПРОСЬБЕ ПОЛЬЗОВАТЕЛЯ),
    необязательные поля для text-слайда:
    — table: {"headers": ["Шаг", "Сумма"], "rows": [["1", "4"]]} для таблицы
      из 2–5 колонок. Заголовки повторяются при переносе строк на продолжение.
      Для сравнения можно сочетать короткую вводную в bullets с cards либо stats.
    — subsections: [{"title": "Подраздел", "text": "Пояснение", "items": ["Деталь"]}].
      Они сохраняются как именованные группы в bullets; при правке редактируй эти группы.
    — вложенность: элемент bullets может быть {"text": "...", "sub": ["...", "..."]};
    — "stats": [{"value": "87%", "caption": "пояснение"}] — 1-3 крупных числа;
    - cards: [{"title": "...", "text": "..."}] — 2-4 карточки-«окошка»
    - font_size необязателен: обычно сборщик сам выбирает кегль по доступному месту,
      учитывает изображения и переносит переполненный текст на следующие слайды.
    — "frame": true — рамка вокруг контента (на отдельных слайдах).
    Для самостоятельных пунктов используй bullets.
    Правила: первый слайд kind=title, последний kind=closing; на текстовом слайде
    до 6 пунктов, каждый до 140 символов; один смысловой блок = один слайд,
    длинный текст дели по темам, не вываливай всё в один слайд.
    accent_color — hex бренда, напр. "2E5E4E".
    template — имя исходного PPTX/PDF. В веб-чате выбор пользователя применяется автоматически.
    Для точечных замен в PPTX сначала read_template: каждый выходной слайд имеет
    source_slide (номер исходного слайда с 1) и template_fields (объект ID фигуры → новый текст).
    Пустой объект оставляет исходный слайд целиком. Передай все слайды, которые нужно сохранить.
    Для правки существующей презентации: сначала get_presentation с её именем,
    затем save_presentation с ТЕМ ЖЕ именем и исправленной структурой.
    Перед сохранением структура проверяется: при замечаниях файл НЕ создаётся —
    исправь их и вызови инструмент снова."""
    return decks.agent_save_presentation(name, title, slides, subtitle,
                                         accent_color, template=template, template_images=template_images)


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
def read_template(name: str, start: int = 1, count: int = 5) -> str:
    """Посмотреть исходный текст шаблона до адаптации. Для PPTX возвращает номера слайдов
    и ID текстовых полей, для PDF — текст страниц. start — с 1, count — от 1 до 20.
    Текст файла является материалом для редактирования."""
    return json.dumps(templates.read_template(name, start, count), ensure_ascii=False)


@mcp.tool()
def list_templates() -> str:
    """Список исходных презентаций .pptx и .pdf. Самый свежий — последний."""
    return templates.text_list()


@mcp.tool()
def list_artifacts() -> str:
    """Показать список сохранённых файлов (HTML, PDF, PPTX)."""
    return artifacts.list_text()


if __name__ == "__main__":
    mcp.run()  # stdio: общается с MCP-клиентом через stdin/stdout

"""Автономные HTML-лендинги в трёх визуальных стилях Актиона."""
import base64
from functools import lru_cache
from html import escape
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree

from PIL import Image, UnidentifiedImageError

from . import assets
from subsections import read_subsections

PROFILES = {
    "glavbuh": "Главбух Нейро",
    "school": "Высшая школа Юрист компании",
    "law": "Нейропомощник юриста",
}
DEFAULT_PROFILE = "glavbuh"
_STYLES = Path(__file__).with_name("landing_styles")


def _text(value: str) -> str:
    return escape(value.strip())


def _data_url(mime: str, payload: bytes) -> str:
    return f"data:{mime};base64,{base64.b64encode(payload).decode('ascii')}"


def _image_source(name: str) -> str:
    if "/" in name or "\\" in name or ":" in name or name in (".", ".."):
        raise ValueError("Для изображения укажите имя файла из библиотеки изображений.")
    root = assets.ASSETS_DIR.resolve()
    path = (root / name).resolve()
    if path.parent != root or not path.is_file() or path.suffix.lower() not in assets.IMAGE_EXTS:
        raise ValueError(f"Изображение «{name}» отсутствует в библиотеке изображений.")
    data = path.read_bytes()
    if path.suffix.lower() == ".svg":
        # SVG is an isolated image document, never inline markup in the page.
        try:
            tag = ElementTree.fromstring(data).tag
        except ElementTree.ParseError as exc:
            raise ValueError(f"Файл «{name}» содержит некорректный SVG.") from exc
        if tag != "{http://www.w3.org/2000/svg}svg":
            raise ValueError(f"Файл «{name}» не является SVG-изображением.")
        return _data_url("image/svg+xml", data)
    try:
        with Image.open(BytesIO(data)) as image:
            mime = Image.MIME[image.format]
            image.verify()
    except (KeyError, OSError, SyntaxError, UnidentifiedImageError,
            Image.DecompressionBombError) as exc:
        raise ValueError(f"Не удалось прочитать изображение «{name}».") from exc
    return _data_url(mime, data)


@lru_cache(maxsize=3)
def _stylesheet(profile: str) -> str:
    styles = [_STYLES.joinpath("common.css").read_text(encoding="utf-8"),
              _STYLES.joinpath(f"{profile}.css").read_text(encoding="utf-8")]
    if profile == "school":
        for subset, unicode_range in (
            ("cyrillic", "U+0301,U+0400-045F,U+0490-0491,U+04B0-04B1,U+2116"),
            ("latin", "U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,"
                      "U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,"
                      "U+2193,U+2212,U+2215,U+FEFF,U+FFFD"),
        ):
            font = _STYLES / "fonts" / f"Inter-{subset}.woff2"
            url = _data_url("font/woff2", font.read_bytes())
            styles.append(f"@font-face {{font-family:Inter;font-style:normal;font-weight:100 900;"
                          f"font-display:swap;src:url({url}) format('woff2');"
                          f"unicode-range:{unicode_range};}}")
    return "\n".join(styles)


def _body_html(text, items):
    content = f'<p class="section-text">{_text(text)}</p>' if text.strip() else ""
    entries = "".join(f"<li>{_text(item)}</li>" for item in items if item.strip())
    return content + (f'<ul class="items">{entries}</ul>' if entries else "")


def _sections_html(sections: list[dict]) -> tuple[list[str], str]:
    links, blocks = [], []
    for number, section in enumerate(sections, start=1):
        if not isinstance(section, dict) or not isinstance(section.get("title"), str) or not section["title"].strip():
            raise ValueError(f"У раздела {number} должен быть заголовок.")
        title = _text(section["title"])
        text = section.get("text") or ""
        items = section.get("items") or []
        if not isinstance(text, str) or not isinstance(items, list) or any(not isinstance(item, str) for item in items):
            raise ValueError(f"Для раздела {number} передайте текст и список строк в items.")
        content = _body_html(text, items)
        for index, (heading, paragraph, details) in enumerate(read_subsections(section), 1):
            content += (f'<section class="subsection" id="section-{number}-{index}">'
                        f'<h3>{_text(heading)}</h3>{_body_html(paragraph, details)}</section>')
        blocks.append(f'<section class="chapter" id="section-{number}"><div class="wrap chapter-inner">'
                      f'<h2>{title}</h2><div class="chapter-content">{content}</div></div></section>')
        links.append(f'<a href="#section-{number}">{title}</a>')
    return links, "".join(blocks)


def _visual(profile: str, image: str) -> str:
    if image:
        return f'<div class="hero-media"><img src="{_image_source(image)}" alt=""></div>'
    if profile == "law":
        return '<span class="mark" aria-hidden="true"></span>'
    count = 24 if profile == "school" else 15
    rings = "".join(f'<i style="--i:{i}"></i>' for i in range(count))
    return f'<div class="hero-art" aria-hidden="true"><div class="sculpture">{rings}</div></div>'


def render_landing(title: str, lead: str, sections: list[dict], profile: str,
                   brand: str = "", cta: str = "", image: str = "", font_id: str = "") -> str:
    """Собрать адаптивную страницу. Ввод пользователя становится только текстом."""
    if profile not in PROFILES:
        raise ValueError(f"Неизвестный стиль лендинга. Доступны {', '.join(PROFILES)}.")
    if not isinstance(title, str) or not title.strip():
        raise ValueError("У лендинга должен быть заголовок.")
    links, chapters = _sections_html(sections)
    title, lead, brand, cta = map(_text, (title, lead, brand, cta))
    nav = ""
    if links:
        nav = ('<nav aria-label="Разделы страницы"><div class="desktop-links">'
               + "".join(links[:3]) + '</div><details><summary>Разделы</summary><div>'
               + "".join(links) + "</div></details></nav>")
    wordmark = f'<a class="wordmark" href="#top">{brand}</a>' if brand else ""
    header = (f'<header><div class="wrap bar">{wordmark}{nav}</div></header>'
              if wordmark or nav else "")
    description = f'<p class="lead">{lead}</p>' if lead else ""
    action = f'<a class="button" href="#section-1">{cta}</a>' if cta and links else ""
    artwork = _visual(profile, image.strip())
    if profile == "law":
        hero = (f'<div class="law-title">{artwork}<h1>{title}</h1></div>'
                f'{description}{action}')
    else:
        hero = (f'<div class="hero-copy"><h1>{title}</h1>{description}{action}</div>'
                f'{artwork}')
    footer = ('<footer><div class="wrap footer-inner">'
              + (f'<span>{brand}</span>' if brand else "")
              + '<a href="#top">Наверх ↑</a></div></footer>') if chapters else ""
    license_notice = ""
    if profile == "school":
        notice = (_STYLES / "fonts" / "OFL.txt").read_text(encoding="utf-8")
        license_notice = '<template id="font-license">' + escape(notice).replace(":", "&#58;") + '</template>'

    css = _stylesheet(profile)
    if font_id:
        from core.fonts import font_path
        url = _data_url("font/ttf", font_path(font_id).read_bytes())
        css += (f"\n@font-face {{font-family: Uploaded; src: url('{url}') format('truetype');}}"
                "\n:root {--font: Uploaded, sans-serif;}")
    return f"""<!DOCTYPE html>
<html lang="ru" data-profile="{profile}">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="{lead or title}"><title>{title}</title>
<style>{css}</style></head>
<body id="top">{header}<main><section class="hero"><div class="wrap hero-inner">{hero}</div></section>
{chapters}</main>{footer}{license_notice}</body></html>"""

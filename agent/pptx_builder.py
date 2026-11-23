"""Сборка PPTX из структурированных данных.

Модель отдаёт JSON-структуру (заголовок, слайды, буллеты), а здесь код
собирает из неё настоящий .pptx с единым дизайном. Дизайн — в коде,
поэтому результат предсказуемый, а файл редактируется в PowerPoint.

Структура на входе:
{
  "title": "Название презентации",
  "subtitle": "подзаголовок / компания",
  "slides": [
    {"kind": "title",    "title": "...", "subtitle": "..."},
    {"kind": "content",  "title": "...", "bullets": ["...", "..."]},
    {"kind": "section",  "title": "..."},
    {"kind": "closing",  "title": "Спасибо", "subtitle": "контакты"}
  ]
}
"""
import copy
from difflib import SequenceMatcher
import re
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from template_palette import readable, slide_theme, tint
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SLIDE_W = Inches(13.333)  # 16:9
SLIDE_H = Inches(7.5)

# папки локальных ресурсов пользователя (local-first)
ASSETS_DIR = Path(__file__).resolve().parent.parent / "data" / "assets"
TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "data" / "templates"

DEFAULT_THEME = {
    "bg": "FFFFFF",        # фон контентных слайдов
    "accent": "2E5E4E",    # акцентный (заголовки, плашки)
    "dark": "1A2E28",      # тёмный фон (титул/разделители/финал)
    "text": "2B2B2B",      # основной текст
    "light": "FFFFFF",     # текст на тёмном
    "font": "Calibri",
}

MAX_BULLETS = 6        # больше — перегруз слайда
MAX_BULLET_LEN = 140   # символов в буллете
MAX_SUB = 4            # подпунктов у одного пункта
MAX_SUB_LEN = 100      # символов в подпункте
MAX_STATS = 3          # блоков с крупными цифрами на слайде
MAX_TOTAL_LINES = 8    # строк на слайде (пункты + подпункты)

# Модели часто используют свои имена полей — принимаем синонимы.
_KIND_ALIASES = {
    "text": "content", "bullets": "content", "list": "content",
    "cover": "title", "intro": "title",
    "end": "closing", "final": "closing", "contacts": "closing",
}
_BULLET_ALIASES = ("bullets", "content", "points", "items", "text")


def _plain_bullets(bullets) -> list:
    """Пункты в едином виде: [{"text": str, "sub": [str, ...]}]."""
    out = []
    for b in bullets or []:
        if isinstance(b, dict):
            text = str(b.get("text") or "").strip()
            subs = [str(x).strip() for x in (b.get("sub") or []) if str(x).strip()]
            out.append({"text": text, "sub": subs})
        else:
            text = str(b).strip()
            if text:
                out.append({"text": text, "sub": []})
    return out


_NO_EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FAFF\u2600-\u27BF\uFE00-\uFE0F\u2B00-\u2BFF\u2049\u203C]+"
)


def _no_emoji(text: str) -> str:
    """Вырезать эмодзи и декоративные символы, схлопнуть лишние пробелы."""
    import re as _re
    out = _NO_EMOJI_RE.sub(" ", str(text or ""))
    return _re.sub(r"\s{2,}", " ", out).strip()


def normalize_slide(slide: dict) -> dict:
    """Приводит слайд от модели к каноническому виду.
    kind: type/cover/bullets/... -> title|content|section|closing
    буллеты: ищем в bullets/content/points/items; строку режем на предложения."""
    s = dict(slide)
    if "source_slide" in s or "template_fields" in s:
        return s
    kind = str(s.get("kind") or s.get("type") or "content").lower()
    s["kind"] = _KIND_ALIASES.get(kind, kind)
    s["title"] = _no_emoji(str(s.get("title") or "")).strip()
    if s.get("subtitle") is not None:
        s["subtitle"] = _no_emoji(str(s.get("subtitle") or "")).strip()
    bullets = None
    for key in _BULLET_ALIASES:
        if s.get(key):
            bullets = s[key]
            break
    if isinstance(bullets, str):
        # модели часто пишут пункты одной строкой с переводами строк или
        # нумерацией — сначала режем по строкам, и только одиночный длинный
        # абзац без переносов — по предложениям
        lines = [ln.strip() for ln in bullets.splitlines() if ln.strip()]
        if len(lines) <= 1 and len(bullets) > MAX_BULLET_LEN:
            lines = [p.strip() for p in re.split(r"(?<=[.!?])\s+", bullets) if p.strip()]
        bullets = lines
    # нормализация пунктов: режем по строкам, снимаем нумерацию/маркеры
    # («1.», «•», «-»), в т.ч. внутри уже разбитых пунктов ("пункт\n2. пункт")
    strip_re = re.compile(r"^\s*(?:\d+[.)]|[-•*–])\s+")
    def clean(value):
        return _no_emoji(strip_re.sub("", str(value).strip())).strip()

    items = []
    for item in bullets or []:
        if isinstance(item, dict):
            heading = clean(item.get("text") or item.get("title") or "")
            subs = item.get("sub") or item.get("items") or []
            if isinstance(subs, str):
                subs = subs.splitlines()
            subpoints = [v for x in subs if (v := clean(x))]
            if heading:
                items.append({"text": heading, "sub": subpoints})
            else:
                items.extend(subpoints)
        else:
            items.extend(v for line in str(item).splitlines() if (v := clean(line)))
    from subsections import read_subsections
    for heading, paragraph, details in read_subsections(s):
        lines = [paragraph, *details] if paragraph else details
        items.append({"text": clean(heading), "sub": [v for line in lines if (v := clean(line))]})
    s.pop("subsections", None)
    if s.get("table") is not None:
        from slide_table import checked_table
        s["table"] = checked_table(s["table"])
    s["bullets"] = items
    for alias in _BULLET_ALIASES[1:]:
        s.pop(alias, None)
    # Сохраняем весь исходный материал. Лимиты проверяются после разбиения.
    stats = []
    for st in s.get("stats") or []:
        if isinstance(st, dict):
            raw = st.get("value")
            if raw is None or raw == "":
                raw = st.get("number")
            value = _no_emoji(str(raw)) if raw is not None else ""
            caption = _no_emoji(str(st.get("caption") or st.get("label") or ""))
            if value:
                stats.append({"value": value, "caption": caption})
    s["stats"] = stats
    cards = []
    for c in s.get("cards") or []:
        if isinstance(c, dict):
            title = _no_emoji(str(c.get("title") or ""))
            text = _no_emoji(str(c.get("text") or c.get("body") or ""))
            if title or text:
                cards.append({"title": title, "text": text})
    s["cards"] = cards
    s["frame"] = bool(s.get("frame"))
    fs = s.get("font_size") or s.get("body_size") or s.get("font")
    try:
        s["font_size"] = min(40, max(24, int(fs)))
    except (TypeError, ValueError):
        s.pop("font_size", None)
    return s


def _rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.lstrip("#"))


def validate_presentation(data: dict) -> list[str]:
    """Проверки структуры ДО сборки файла. Пустой список = всё ок.
    Ожидает уже нормализованные слайды (normalize_slide)."""
    issues: list[str] = []
    slides = data.get("slides")
    if not isinstance(slides, list) or not slides:
        return ["нет слайдов (slides — пустой или не список)"]
    kinds = {"title", "content", "section", "closing"}
    for i, slide in enumerate(slides, start=1):
        if "source_slide" in slide or "template_fields" in slide:
            index = slide.get("source_slide")
            fields = slide.get("template_fields", {})
            if type(index) is not int or index < 1:
                issues.append(f"слайд {i}: source_slide — номер исходного слайда от 1")
            if not isinstance(fields, dict) or any(
                not str(key).isascii() or not str(key).isdigit() or int(key) < 1
                or not isinstance(text, str) for key, text in fields.items()
            ):
                issues.append(f"слайд {i}: template_fields — объект с ID текстовых полей и строками")
            continue
        kind = slide.get("kind", "content")
        if kind not in kinds:
            issues.append(f"слайд {i}: неизвестный kind «{kind}» (есть: {', '.join(sorted(kinds))})")
        if not slide.get("title"):
            issues.append(f"слайд {i}: нет title")
        if len((slide.get("title") or "").removesuffix(" (продолжение)")) > 110:
            issues.append(f"слайд {i}: заголовок длиннее 110 символов — перенесите пояснение в основной текст")
        if kind in ("title", "section", "closing"):
            if len(slide.get("title") or "") > 70:
                issues.append(f"слайд {i}: заголовок титула длиннее 70 символов — сократи")
            if len(slide.get("subtitle") or "") > 90:
                issues.append(f"слайд {i}: подзаголовок длиннее 90 символов — сократи")
        bullets = _plain_bullets(slide.get("bullets"))
        cards = slide.get("cards") or []
        if kind == "content" and not any((bullets, cards, slide.get("image"), slide.get("stats"), slide.get("table"))):
            got = ", ".join(sorted(slide.keys()))
            issues.append(
                f"слайд {i}: content-слайд без буллетов "
                f"(пришли поля: {got}; нужно поле \"bullets\": [\"строка\", ...])"
            )
        if len(bullets) > MAX_BULLETS:
            issues.append(f"слайд {i}: буллетов {len(bullets)} > {MAX_BULLETS} — сократи")
        for b in bullets:
            if len(b["text"]) > MAX_BULLET_LEN:
                issues.append(f"слайд {i}: буллет длиннее {MAX_BULLET_LEN} символов — сократи")
                break
            if len(b["sub"]) > MAX_SUB:
                issues.append(f"слайд {i}: подпунктов {len(b['sub'])} > {MAX_SUB} — сократи")
                break
            if any(len(x) > MAX_SUB_LEN for x in b["sub"]):
                issues.append(f"слайд {i}: подпункт длиннее {MAX_SUB_LEN} символов — сократи")
                break
        total_lines = sum(1 + len(b["sub"]) for b in bullets)
        if total_lines > MAX_TOTAL_LINES:
            issues.append(f"слайд {i}: строк {total_lines} > {MAX_TOTAL_LINES} — раздели материал на два слайда")
        if slide.get("font_size") is not None and \
                not isinstance(slide.get("font_size"), int):
            issues.append(f"слайд {i}: font_size — целое число 24–40")
        cards = slide.get("cards") or []
        if cards and (not isinstance(cards, list) or len(cards) < 2 or len(cards) > 4):
            issues.append(f"слайд {i}: cards — список из 2-4 карточек (title/text)")
        stats = slide.get("stats") or []
        if not isinstance(stats, list) or len(stats) > MAX_STATS:
            issues.append(f"слайд {i}: stats — список до {MAX_STATS} блоков {{value, caption}}")
        else:
            for st in stats:
                if not isinstance(st, dict) or not str(st.get("value") or "").strip():
                    issues.append(f"слайд {i}: блок stats без value — {{\"value\": \"100%\", \"caption\": \"...\"}}")
                    break
        image_name = slide.get("image")
        if image_name and not (ASSETS_DIR / Path(str(image_name)).name).exists():
            issues.append(f"слайд {i}: изображение «{image_name}» не найдено — "
                          f"вызови list_images и выбери из списка")
    if "source_slide" not in slides[0] and slides[0].get("kind", "content") != "title":
        issues.append("первым должен быть титульный слайд (kind=title)")
    return issues


def theme_accent_from_template(template_path: Path) -> str | None:
    """Достаёт accent1 из темы .pptx-шаблона (ppt/theme/theme1.xml)."""
    try:
        with zipfile.ZipFile(template_path) as z:
            xml = z.read("ppt/theme/theme1.xml").decode("utf-8", errors="replace")
        m = re.search(r'<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', xml)
        return m.group(1) if m else None
    except Exception:
        return None


def build_presentation(path: Path, data: dict, theme: dict | None = None,
                       template: Path | None = None,
                       template_images: str = "auto") -> None:
    """Собрать .pptx по структуре. Путь должен заканчиваться на .pptx.

    Если задан template — работаем «по скелету»: слайды шаблона остаются как есть
    (с их дизайном), мы заменяем только текст в плейсхолдерах. Чужая структура
    уходит, наша — накладывается. Неиспользованные слайды шаблона удаляются,
    недостающие клонируются с сохранением дизайна. Если в шаблоне нет слайдов
    с нужными плейсхолдерами — фолбэк на лейауты, а без шаблона — собственный
    дизайн по умолчанию."""
    t = {**DEFAULT_THEME, **(theme or {})}
    def save(prs):
        if t.get("font_id"):
            from core.fonts import embed, font_path
            embed(prs, font_path(t["font_id"]))
        prs.save(path)
    if any("source_slide" in sd or "template_fields" in sd for sd in data["slides"]):
        if not template or not template.is_file() or template.suffix.lower() != ".pptx":
            raise ValueError("Для редактирования полей нужен исходный шаблон .pptx")
    # верхнеуровневый подзаголовок (компания) — на титульный слайд, если у него нет своего
    for sd in data.get("slides", []):
        if sd.get("kind", sd.get("type", "")) in ("title", "cover", "intro") \
                and not sd.get("subtitle") and data.get("subtitle"):
            sd["subtitle"] = data["subtitle"]
    if template and template.exists() and template.suffix.lower() == ".pdf":
        from pdf_template import render_pdf_template
        prs = Presentation()
        render_pdf_template(prs, data["slides"], template, template_images,
                            SLIDE_W, _plain_bullets, _slide_image)
        save(prs)
        return
    if template and template.exists():
        prs = Presentation(str(template))
        t.update(user_accent=(theme or {}).get("accent"), user_font=(theme or {}).get("font"))
        if all("source_slide" in sd for sd in data["slides"]):
            if not _build_on_skeleton(prs, data, t):
                raise ValueError("В исходной презентации нет слайдов")
            save(prs)
            return
        if not (theme or {}).get("accent"):
            accent = theme_accent_from_template(template)
            if accent:
                t["accent"] = accent
        # карта дизайна: фирменный шрифт + положение логотипа на листах
        try:
            analysis = _template_analysis(template)
        except Exception as exc:
            print(f"analysis of template failed: {exc}")
            analysis = {}
        if not (theme or {}).get("font"):
            t["font"] = _template_font_pptx(prs) or analysis.get("font") or t["font"]
        if _build_on_skeleton(prs, data, t,
                              logo_title=analysis.get("logo_first"),
                              logo_content=analysis.get("logo_content"),
                              logo_last=analysis.get("logo_last")):
            save(prs)
            return
        # фолбэк: у шаблона нет слайдов с плейсхолдерами — собираем на лейаутах
        _wipe_template_slides(prs)
        title_layout = (_layout_with(prs, PP_PLACEHOLDER.CENTER_TITLE)
                        or _layout_with(prs, PP_PLACEHOLDER.TITLE)
                        or prs.slide_layouts[0])
        content_layout = (_layout_with(prs, PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.OBJECT)
                          or _layout_with(prs, PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.BODY)
                          or title_layout)
        section_layout = (_layout_with(prs, PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.BODY)
                          or title_layout)
        _emit_slides(prs, data, t, title_layout, content_layout, section_layout,
                     use_template_layouts=True)
    else:
        prs = Presentation()
        prs.slide_width = SLIDE_W
        prs.slide_height = SLIDE_H
        blank = prs.slide_layouts[6]
        _emit_slides(prs, data, t, blank, blank, blank, use_template_layouts=False)
    save(prs)


def _emit_slides(prs, data, t, title_layout, content_layout, section_layout,
                 use_template_layouts: bool) -> None:
    from slide_layout import fitting_slide
    for index, sd in enumerate(data["slides"]):
        with fitting_slide(index):
            kind = sd.get("kind", "content")
            layout = (title_layout if kind in ("title", "closing")
                      else section_layout if kind == "section" else content_layout)
            slide = prs.slides.add_slide(layout)
            if use_template_layouts:
                _retext_slide(prs, slide, sd, slide_theme(slide, t))
                continue
            sw, sh = prs.slide_width, prs.slide_height
            accent = t["accent"]
            if uses_body(sd):
                kind = "content"
            color = t["dark"] if kind in ("title", "closing") else accent if kind == "section" else t["bg"]
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = _rgb(color)
            if kind in ("title", "closing", "section"):
                cover = slide.shapes.add_textbox(int(sw*.07), int(sh*.28), int(sw*.86), int(sh*.26))
                _native_heading(cover, sd.get("title", ""), t["font"], t["light"], 28, 48, centered=True)
                if sd.get("subtitle"):
                    subtitle = slide.shapes.add_textbox(int(sw*.1), int(sh*.56), int(sw*.8), int(sh*.23))
                    _native_heading(subtitle, sd["subtitle"], t["font"], t["light"], 18, 28,
                                    centered=True, bold=False)
                continue

            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, int(sw*.014), sh)
            bar.fill.solid()
            bar.fill.fore_color.rgb = _rgb(accent)
            bar.line.fill.background()
            title = slide.shapes.add_textbox(int(sw*.055), int(sh*.045), int(sw*.89), int(sh*.18))
            _native_heading(title, sd.get("title", ""), t["font"], accent, 24, 38)
            if sd.get("underline", True):
                _add_title_underline(slide, title, accent)
            body = slide.shapes.add_textbox(int(sw*.065), int(sh*.26), int(sw*.87), int(sh*.65))
            _native_content(slide, prs, body, sd, t, markers=True)
            if sd.get("frame"):
                _add_frame(slide, prs, int(sw*.025), int(sh*.025), int(sw*.95), int(sh*.95), accent)


# ---------- сборка «по скелету» шаблона ----------

_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _slide_role(slide) -> str | None:
    """Роль слайда по плейсхолдерам: title / content / section (или None)."""
    types = {ph.placeholder_format.type for ph in slide.placeholders}
    titleish = bool(types & {PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE})
    body = bool(types & {PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY})
    sub = PP_PLACEHOLDER.SUBTITLE in types
    if titleish and sub and not body:
        return "title"
    if titleish and body:
        return "content"
    if titleish:
        return "section"
    return None


def _text_shapes(slide) -> list:
    """Текстовые шейпы слайда, отсортированные по кеглю (крупный — первым)."""
    out = []
    for sp in slide.shapes:
        if sp.has_text_frame and sp.text_frame.text.strip():
            sizes = [r.font.size.pt for p in sp.text_frame.paragraphs
                     for r in p.runs if r.font.size]
            out.append((max(sizes) if sizes else 0, sp))
    out.sort(key=lambda pair: -pair[0])
    return [sp for _, sp in out]


def _freeform_role(slide) -> str | None:
    """Роль слайда «дизайнерского» шаблона, где весь текст — текстбоксы."""
    texts = _text_shapes(slide)
    if not texts:
        return None
    sizes = []
    for sp in texts:
        sizes.extend(r.font.size.pt for p in sp.text_frame.paragraphs
                     for r in p.runs if r.font.size)
    if len(texts) == 1:
        return "title" if sizes and max(sizes) >= 24 else None
    body = texts[1]
    paras = body.text_frame.paragraphs
    if len(paras) >= 2 or len(body.text_frame.text) >= 80:
        return "content"
    return "title"  # крупный заголовок + короткий подзаголовок


def _build_on_skeleton(prs, data: dict, t: dict,
                       logo_title=None, logo_content=None, logo_last=None) -> bool:
    """Наложить нашу структуру на слайды шаблона, сохранив их дизайн.
    Раскладка по позиции: ПЕРВЫЙ лист шаблона — титул, ПОСЛЕДНИЙ — финал
    (контакты), средние — контент. Каждая роль при нехватке клонируется.
    Одностраничный образец с пригодными текстовыми полями используется повторно."""
    slides = list(prs.slides)
    if not slides:
        return False
    pools: dict[str, list] = {
        "title": [slides[0]],
        "closing": [slides[-1]],
        "content": slides[1:-1] or [slides[-1]],
    }
    modes: dict[int, str] = {}
    for pool in pools.values():
        for s in pool:
            modes[id(s)] = "ph" if _slide_role(s) else "ff"

    used: set[int] = set()
    ordered: list = []

    # Allocate every page before changing any source slide. Repeated pages
    # must inherit the original text, images and decorations.
    for sd in data["slides"]:
        kind = sd.get("kind", "content")
        if "source_slide" in sd:
            index = sd["source_slide"]
            if type(index) is not int or not 1 <= index <= len(slides):
                raise ValueError(f"Исходный слайд {index} отсутствует в шаблоне ({len(slides)} слайдов)")
            source = slides[index - 1]
        else:  # content и section — средние листы
            pool = pools.get(kind, pools["content"])
            source = next((s for s in pool if id(s) not in used), pool[-1])
        base = _duplicate_slide(prs, source) if id(source) in used else source
        modes[id(base)] = modes[id(source)]
        used.add(id(base))
        ordered.append(base)

    from slide_layout import fitting_slide
    for index, (sd, base) in enumerate(zip(data["slides"], ordered)):
        if "source_slide" in sd:
            fields = {str(sh.shape_id): sh for sh in template_text_shapes(base.shapes)}
            for key, value in sd.get("template_fields", {}).items():
                sh = fields.get(str(key))
                if sh is None:
                    raise ValueError(f"На исходном слайде {sd['source_slide']} нет текстового поля {key}")
                if not isinstance(value, str):
                    raise ValueError(f"Поле {key}: значение должно быть строкой")
                _check_source_text(base, sh, value, sd["source_slide"])
                _set_shape_lines(sh, value.replace("\r\n", "\n").split("\n"))
            continue
        with fitting_slide(index):
            kind = sd.get("kind", "content")
            logo = (logo_title if kind == "title"
                    else logo_last if kind == "closing" else logo_content)
            if modes[id(base)] == "ph":
                _retext_slide(prs, base, sd, slide_theme(base, t), logo=logo)
            else:
                _retext_freeform(prs, base, sd, slide_theme(base, t), logo=logo)

    # чужие слайды, не понадобившиеся ни под один наш, — удалить
    for s in list(prs.slides):
        if id(s) not in used:
            _delete_slide(prs, s)
    _reorder(prs, ordered)
    return True


def template_text_shapes(shapes):
    """Редактируемые текстовые поля слайда, включая текст внутри групп."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from template_text_shapes(shape.shapes)
        elif shape.has_text_frame and not (
            shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER
        ):
            yield shape


def _has_title_decor(slide, prs):
    """У шаблона своё оформление зоны заголовка (линии в верхней трети)?"""
    limit = int(prs.slide_height * 0.35)
    for sh in slide.shapes:
        try:
            if sh.shape_type == MSO_SHAPE_TYPE.LINE and (sh.top or 0) < limit:
                return True
        except (TypeError, ValueError):
            continue
    return False


def _remove_shape(sh) -> None:
    sh._element.getparent().remove(sh._element)


def _empty_pic_placeholders(slide) -> list:
    """Пустые picture-плейсхолдеры шаблона (рамки без картинки) — мусор."""
    out = []
    for sh in slide.shapes:
        if not sh.is_placeholder:
            continue
        try:
            if sh.placeholder_format.type != PP_PLACEHOLDER.PICTURE:
                continue
        except (TypeError, ValueError):
            continue
        try:
            sh.image
            continue          # картинка есть — не трогаем
        except Exception:
            out.append(sh)
    return out


def _tpl_image_left(slide, prs):
    """Левая граница (доли ширины) самой правой картинки шаблона на слайде,
    чтобы текст её не перекрывал. None — картинок нет."""
    left = None
    for sh in slide.shapes:
        is_pic = sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE)
        if not is_pic and sh.is_placeholder:
            try:
                is_pic = sh.placeholder_format.type == PP_PLACEHOLDER.PICTURE
            except (TypeError, ValueError):
                is_pic = False
        if not is_pic:
            continue
        try:
            sh.image
        except Exception:
            continue          # пустой плейсхолдер
        l = (sh.left or 0) / max(prs.slide_width, 1)
        if left is None or l > left:
            left = l
    return left


def _shrink_body_to(prs, body, obstacle_left) -> None:
    """Сузить текстовый бокс так, чтобы его правый край не доходил до
    obstacle_left (доли ширины) на 0.03. Не сжимает ниже 30% слайда."""
    if body is None or body.left is None or obstacle_left is None:
        return
    limit = int(prs.slide_width * (obstacle_left - 0.03)) - body.left
    if limit < body.width and limit >= int(prs.slide_width * 0.3):
        body.width = limit


def _fit_ph_text(prs, ph, start_pt, min_pt):
    """Подогнать текст титула в исходном плейсхолдере дизайнера."""
    if ph.left is None or ph.width is None:
        return
    ph.text_frame.word_wrap = True
    text = ph.text_frame.text
    if not text:
        return
    pt = start_pt
    box_pt = max((ph.height or 0) / 12700.0, 30)
    while pt > min_pt and _est_visual_lines(text, pt, ph.width) * pt * 1.18 > box_pt:
        pt -= 1
    if pt != start_pt:
        for para in ph.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = Pt(int(pt))


def _body_below_title(slide, prs, title_shape):
    """На исходном слайде есть заголовок, но нет готового поля под новый текст."""
    w, h = prs.slide_width, prs.slide_height
    left = title_shape.left if title_shape is not None else int(w * .06)
    top = int(h * .30)
    if title_shape is not None:
        top = title_shape.top + title_shape.height + int(h * .025)
        if top > h * .70:
            top = int(h * .14)
            bottom = max(top + int(h * .15), title_shape.top - int(h * .025))
            return slide.shapes.add_textbox(left, top, w - left - int(w * .05), bottom - top)
    return slide.shapes.add_textbox(left, top, w - left - int(w * .05), int(h * .94) - top)


def uses_body(slide):
    return slide.get("kind", "content") == "content" or any(
        slide.get(key) for key in ("bullets", "stats", "cards", "image", "table"))


def _retext_slide(prs, slide, sd: dict, t: dict, logo=None) -> None:
    kind = "content" if uses_body(sd) else sd.get("kind", "content")
    title = None
    bodies = []
    for ph in slide.placeholders:
        ptype = ph.placeholder_format.type
        if not ph.has_text_frame:
            continue
        if ptype in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE):
            _set_shape_lines(ph, [sd.get("title", "")])
            title = ph
        elif ptype == PP_PLACEHOLDER.SUBTITLE:
            _set_shape_lines(ph, [sd.get("subtitle", "")])
        elif ptype in (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY):
            bodies.append(ph)
    for shape in slide.shapes:
        if not shape.is_placeholder and shape.has_text_frame and shape.text:
            _set_shape_lines(shape, [])
    for sh in _empty_pic_placeholders(slide):
        _remove_shape(sh)
    for ph in bodies[1:]:
        _set_shape_lines(ph, [])
    if kind == "content":
        body = bodies[0] if bodies else _body_below_title(slide, prs, title)
        _native_content(slide, prs, body, sd, t, logo)
        if title is not None and sd.get("underline") and not _has_title_decor(slide, prs):
            _add_title_underline(slide, title, t["accent"])
    elif bodies:
        _set_shape_lines(bodies[0], [sd.get("subtitle", "")])
    if title is not None and kind in ("title", "closing", "section"):
        _fit_ph_text(prs, title, _ph_font_pt(slide, title), min_pt=24)
    if kind != "content" and _slide_image(sd):
        _add_template_image(slide, prs, _slide_image(sd), avoid=logo)
    if sd.get("frame"):
        _add_frame(slide, prs, int(prs.slide_width*.025), int(prs.slide_height*.025),
                   int(prs.slide_width*.95), int(prs.slide_height*.95), t["accent"])


def _retext_freeform(prs, slide, sd: dict, t: dict, logo=None) -> None:
    kind = "content" if uses_body(sd) else sd.get("kind", "content")
    for sh in _empty_pic_placeholders(slide):
        _remove_shape(sh)
    slots = _text_shapes(slide)
    if not slots:
        title = slide.shapes.add_textbox(int(prs.slide_width*.06), int(prs.slide_height*.05),
                                         int(prs.slide_width*.88), int(prs.slide_height*.16))
        _native_heading(title, sd.get("title", ""), t.get("heading_font", t["font"]), t["text"], 24, 36)
        slots = [title]
        if kind != "content" and sd.get("subtitle"):
            subtitle_box = _body_below_title(slide, prs, title)
            _native_heading(subtitle_box, sd["subtitle"], t["font"], t["text"], 18, 28, bold=False)
            slots.append(subtitle_box)
    subtitle = sd.get("subtitle", "")
    title_lines = [sd.get("title", "")]
    if kind != "content" and subtitle and len(slots) == 1:
        title_lines.append(subtitle)
    _set_shape_lines(slots[0], title_lines)
    for shape in slots[2:]:
        _set_shape_lines(shape, [])
    if kind == "content":
        body = slots[1] if len(slots) > 1 else _body_below_title(slide, prs, slots[0])
        _native_content(slide, prs, body, sd, t, logo)
        if sd.get("underline") and not _has_title_decor(slide, prs):
            _add_title_underline(slide, slots[0], t["accent"])
    else:
        if len(slots) > 1:
            _set_shape_lines(slots[1], [subtitle])
        for box in slots[:2]:
            box.text_frame.word_wrap = True
            box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        if _slide_image(sd):
            _add_template_image(slide, prs, _slide_image(sd), avoid=logo)
    if sd.get("frame"):
        _add_frame(slide, prs, int(prs.slide_width*.025), int(prs.slide_height*.025),
                   int(prs.slide_width*.95), int(prs.slide_height*.95), t["accent"])


def _bullet_lines(bullets) -> list:
    """Пункты -> [(текст, уровень)]: уровень 0 — пункт, 1 — подпункт."""
    lines = []
    for b in _plain_bullets(bullets):
        lines.append((b["text"], 0))
        for sub in b["sub"]:
            lines.append((sub, 1))
    return lines


def _est_visual_lines(text, font_pt, width_emu, char_ratio=0.54):
    """Оценка числа строк после переноса: средний символ ~0.54 кегля."""
    import math
    text = str(text or "")
    if not text:
        return 1
    usable_pt = max(width_emu / 12700.0 - 16, 60)
    cpl = max(int(usable_pt / (font_pt * char_ratio)), 8)
    return max(math.ceil(len(text) / cpl), 1)


def _text_size(shape, rows, low, high, font, indent=0):
    from slide_layout import fit_body
    frame = shape.text_frame
    w = (shape.width - frame.margin_left - frame.margin_right) / 12700 - indent
    h = (shape.height - frame.margin_top - frame.margin_bottom) / 12700
    return fit_body(rows, w, h, low=low, high=high, font=font or "Arial")


def _check_source_text(slide, shape, replacement, source_index):
    """Protect designer-sized slots when a replacement adds substantial content."""
    from slide_layout import SlideFitError
    old = shape.text
    if (len(replacement) <= max(len(old) + 12, len(old) * 1.4)
            and replacement.count("\n") <= old.count("\n")):
        return
    paragraphs = shape.text_frame.paragraphs
    runs = [run for para in paragraphs for run in para.runs]
    font = next((run.font.name for run in runs if run.font.name), "Arial")
    base = _proto_font_pt(shape)
    if not base:
        base = _ph_font_pt(slide, shape) if shape.is_placeholder else 18
    rows = []
    for index, line in enumerate(replacement.splitlines() or [""]):
        proto = paragraphs[min(index, len(paragraphs) - 1)]
        size = next((r.font.size.pt for r in proto.runs if r.font.size), base)
        rows.append((line, size / base))
    try:
        _text_size(shape, rows, max(1, round(base)), max(1, round(base)), font)
    except SlideFitError as exc:
        raise ValueError(f"Исходный слайд {source_index}, поле {shape.shape_id}: "
                         "новый текст не помещается с исходным шрифтом. "
                         "Сократите текст поля или выберите больший текстовый блок.") from exc


def _native_heading(shape, text, font, color, low, high, centered=False, bold=True):
    frame = shape.text_frame
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.text = text
    size = _text_size(shape, [text], low, high, font)
    for para in frame.paragraphs:
        para.alignment = PP_ALIGN.CENTER if centered else PP_ALIGN.LEFT
        for run in para.runs:
            run.font.name = font
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = _rgb(color)


def _native_content(slide, prs, body, sd, theme, logo=None, markers=False):
    """Разделить реальную область контента между текстом, фото, карточками и цифрами."""
    sw, sh = prs.slide_width, prs.slide_height
    image = _slide_image(sd)
    stats, cards = sd.get("stats") or [], sd.get("cards") or []
    if "user_font" in theme:
        theme = slide_theme(slide, theme, body)
    lines = _bullet_lines(sd.get("bullets"))
    x, y = body.left / sw, body.top / sh
    right = min(.97, x + body.width / sw)
    bottom = min(.72 if stats else .96, y + body.height / sh)
    if image:
        if x > .46:
            image_zone = (.05, .25, x - .075, (.71 if stats else .93) - .25)
        else:
            image_zone = (.56, .25, .39, (.71 if stats else .93) - .25)
            right = min(right, .535)
        _add_template_image(slide, prs, image, image_zone, logo)
    x, y, w, h = _avoid_rect(x, y, right - x, bottom - y, logo)
    if sd.get("table"):
        from slide_table import add_table
        _set_shape_lines(body, [])
        surface = tint(theme["bg"], theme["accent"])
        add_table(slide, sd["table"], (int(x*sw), int(y*sh), int(w*sw), int(h*sh)),
                  theme["font"], theme["text"], surface, theme["bg"], readable(theme["text"], surface))
        return
    body_h = h
    if cards:
        if lines:
            share = min(.6, max(.32, len(lines) / (len(lines) + len(cards) * 1.5)))
            body_h = h * share
            card_zone = (x, y + body_h + .018, w, h - body_h - .018)
        else:
            card_zone = (x, y, w, h)
        _add_cards(slide, prs, cards, theme["accent"], font=theme["font"], zone=card_zone,
                   bg=theme["bg"], ink=theme["text"], rounded=theme.get("rounded_cards", True))
    if lines:
        body.left, body.top = int(sw*x), int(sh*y)
        body.width, body.height = int(sw*w), int(sh*body_h)
        _fill_native_body(body, lines, theme, sd, markers)
    else:
        _set_shape_lines(body, [])
    if stats:
        _add_stats_row(slide, prs, stats, readable(theme["accent"], theme["bg"], (theme["text"],)),
                       font=theme["font"], avoid=logo, ink=theme["text"])


def _fill_native_body(body, lines, theme, sd, markers):
    frame = body.text_frame
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.word_wrap = True
    texts = [(("– " if level else "• ") + text) if markers else text for text, level in lines]
    _set_shape_lines(body, texts)
    font = next((r.font.name for p in frame.paragraphs for r in p.runs if r.font.name),
                theme["font"])
    indent = 20 if body.is_placeholder or any(level for _, level in lines) else 0
    size = _text_size(body, texts, 18, sd.get("font_size") or 32, font, indent)
    for i, (para, (_, level)) in enumerate(zip(frame.paragraphs, lines)):
        para.level = level
        para.space_before = Pt(0)
        para.space_after = Pt(size * .18 if i + 1 < len(lines) else 0)
        para.line_spacing = 1.
        for run in para.runs:
            run.font.size = Pt(max(18, round(size * .9)) if level else size)
            if markers or ("user_font" in theme and not body.is_placeholder and not run.font.name):
                run.font.name = font
                run.font.color.rgb = _rgb(theme["text"])


def _add_frame(slide, prs, x, y, w, h, accent):
    """Тонкая рамка акцентного цвета без заливки."""
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 int(x), int(y), int(w), int(h))
    box.fill.background()
    box.line.color.rgb = _rgb(accent)
    box.line.width = Pt(1.5)
    box.shadow.inherit = False
    return box


def _add_title_underline(slide, title_shape, accent):
    """Короткая акцентная линейка под заголовком контентного слайда."""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 int(title_shape.left),
                                 int(title_shape.top + title_shape.height),
                                 Inches(1.3), Inches(0.045))
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb(accent)
    bar.line.fill.background()
    bar.shadow.inherit = False


def _add_stats_row(slide, prs, stats, accent, font=None, avoid=None, ink="2B2B2B"):
    """Ряд крупных цифр с подписями и акцентной чертой слева.
    avoid — прямоугольник логотипа в долях страницы: ряд его не пересекает."""
    sw, sh = prs.slide_width, prs.slide_height
    n = max(len(stats), 1)
    gap = int(sw * 0.02)
    fx, fy, fw, fh = _avoid_rect(0.055, 0.74, 0.89, 0.22, avoid)
    x, y = int(sw * fx), int(sh * fy)
    w, h = int(sw * fw), int(sh * fh)
    w = int((w - gap * (n - 1)) / n)
    for st in stats:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y,
                                     int(sw * 0.006), h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(accent)
        bar.line.fill.background()
        bar.shadow.inherit = False
        value_h = int(h * .43) if st.get("caption") else h
        value = slide.shapes.add_textbox(x + int(sw*.014), y, w - int(sw*.02), value_h)
        _native_heading(value, str(st.get("value", "")), font or "Arial", accent, 18, 40)
        if st.get("caption"):
            cap = slide.shapes.add_textbox(value.left, y + value_h, value.width, h - value_h)
            _native_heading(cap, str(st["caption"]), font or "Arial", ink, 18, 26, bold=False)
        x += w + gap
def _slide_image(sd: dict) -> Path | None:
    """Путь к картинке слайда из активов пользователя (или None)."""
    name = sd.get("image")
    if not name:
        return None
    path = ASSETS_DIR / Path(str(name)).name
    return path if path.exists() else None


def _add_template_image(slide, prs, image_path, zone=(.56, .25, .39, .68), avoid=None):
    from slide_layout import add_fitted_picture
    if image_path:
        x, y, w, h = _avoid_rect(*zone, avoid)
        return add_fitted_picture(slide, image_path, int(prs.slide_width*x), int(prs.slide_height*y),
                                  int(prs.slide_width*w), int(prs.slide_height*h))


def _set_shape_lines(shape, lines: list) -> None:
    """Replace text in place, keeping paragraph and surviving character styles."""
    frame = shape.text_frame
    if frame.text == "\n".join(lines):
        return
    txBody = frame._txBody
    lines = lines or [""]
    originals = list(frame.paragraphs)
    for _ in range(len(originals), len(lines)):
        txBody.append(copy.deepcopy(originals[-1]._p))
    for para in originals[len(lines):]:
        txBody.remove(para._p)
    for para, text in zip(frame.paragraphs, lines):
        if para.text == text:
            continue
        runs = para.runs
        if not runs:
            para.text = text
            continue
        # Keep the original font/color/hyperlink wherever matching characters
        # survive a replacement, including changes within a multi-style line.
        if len(runs) == 1:
            runs[0].text = text
        else:
            old = "".join(r.text for r in runs)
            owners = [i for i, r in enumerate(runs) for _ in r.text]
            parts = [""] * len(runs)
            for op, a, b, c, d in SequenceMatcher(None, old, text, autojunk=False).get_opcodes():
                if op == "equal":
                    for pos in range(a, b):
                        parts[owners[pos]] += old[pos]
                elif op in ("replace", "insert"):
                    idx = owners[min(a, len(owners) - 1)] if owners else 0
                    parts[idx] += text[c:d]
            for run, part in zip(runs, parts):
                run.text = part
        for child in list(para._p):
            if child.tag in (qn("a:br"), qn("a:fld")):
                para._p.remove(child)


def _duplicate_slide(prs, source):
    """Копия слайда с его фигурами, фоном, переходами и связями."""
    new = prs.slides.add_slide(source.slide_layout)
    root = new._element
    root.clear()
    root.attrib.update(source._element.attrib)
    root.extend(copy.deepcopy(child) for child in source._element)
    # add_slide has already accessed these lazy wrappers for the former tree.
    for prop in ("shapes", "placeholders", "background"):
        new.__dict__.pop(prop, None)
    # Images, chart data, hyperlinks and picture backgrounds refer to this
    # slide's relationship table. Relocate those references on the clone.
    rel_ns = _REL_ID.split("}")[0] + "}"
    for element in root.iter():
        for attr, rid in list(element.attrib.items()):
            if attr.startswith(rel_ns) and rid in source.part.rels:
                rel = source.part.rels[rid]
                target = rel.target_ref if rel.is_external else rel.target_part
                element.set(attr, new.part.relate_to(target, rel.reltype, rel.is_external))
    return new


def _delete_slide(prs, slide) -> None:
    xml_slides = prs.slides._sldIdLst
    for slide_id in list(xml_slides):
        rel_id = slide_id.get(_REL_ID)
        if rel_id and prs.part.related_part(rel_id) is slide.part:
            prs.part.drop_rel(rel_id)
            xml_slides.remove(slide_id)
            return


def _reorder(prs, ordered: list) -> None:
    """Расставить слайды в порядке нашей структуры (append переносит элемент)."""
    sldIdLst = prs.slides._sldIdLst
    part_to_el = {}
    for el in list(sldIdLst):
        rel_id = el.get(_REL_ID)
        if rel_id:
            part_to_el[prs.part.related_part(rel_id)] = el
    for s in ordered:
        el = part_to_el.get(s.part)
        if el is not None:
            sldIdLst.append(el)


def _wipe_template_slides(prs) -> None:
    """Удалить все слайды шаблона (для фолбэка на лейауты)."""
    for s in list(prs.slides):
        _delete_slide(prs, s)


def _layout_with(prs, *wanted_types):
    """Первый лейаут, содержащий все указанные типы плейсхолдеров."""
    for layout in prs.slide_layouts:
        types = {ph.placeholder_format.type for ph in layout.placeholders}
        if all(wt in types for wt in wanted_types):
            return layout
    return None


def _fill_placeholder(slide, types, text, bold=False, size=None):
    """Вписать текст в плейсхолдер по типу. Шрифт и цвет не трогаем —
    они наследуются от шаблона; задаём только жирность/кегль."""
    for ph in slide.placeholders:
        if ph.placeholder_format.type in types:
            frame = ph.text_frame
            frame.text = str(text)
            for para in frame.paragraphs:
                for run in para.runs:
                    if bold:
                        run.font.bold = True
                    if size:
                        run.font.size = Pt(size)
            return ph
    return None


# ---------- сборка по PDF-шаблону ----------

def _corner_color(page) -> tuple:
    """Цвет фона страницы PDF по её углам (для заливки вырезанного текста)."""
    import fitz
    try:
        r = page.rect
        samples = []
        for x, y in ((3, 3), (r.width - 3, 3), (3, r.height - 3), (r.width - 3, r.height - 3)):
            pix = page.get_pixmap(clip=fitz.Rect(x - 1, y - 1, x + 1, y + 1))
            samples.append(pix.pixel(0, 0))
        avg = tuple(sum(c[i] for c in samples) // len(samples) for i in range(3))
        return tuple(c / 255 for c in avg)
    except Exception:
        return (1, 1, 1)

def _shrink_runs(shape, factor):
    """Уменьшить кегль всех ранов текстбокса (при плотном тексте)."""
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size:
                run.font.size = Pt(max(10, int(run.font.size.pt * factor)))


def _proto_font_pt(shape):
    """Кегль первого содержательного рана текстбокса (или None)."""
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size:
                return run.font.size.pt
    return None


def _ph_font_pt(slide, ph, default=20):
    """Эффективный кегль плейсхолдера: layout-раны → lstStyle layout/мастера
    → txStyles мастера (bodyStyle/titleStyle) → default."""
    from pptx.oxml.ns import qn
    idx = ph.placeholder_format.idx

    def _from_txbody(holder_ph):
        for para in holder_ph.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    return run.font.size.pt
        ls = holder_ph.text_frame._txBody.find(qn("a:lstStyle"))
        if ls is not None:
            l1 = ls.find(qn("a:lvl1pPr"))
            d = l1.find(qn("a:defRPr")) if l1 is not None else None
            if d is not None and d.get("sz"):
                return int(d.get("sz")) / 100.0
        return None

    for lph in slide.slide_layout.placeholders:
        if lph.placeholder_format.idx == idx:
            v = _from_txbody(lph)
            if v:
                return v
    layout = slide.slide_layout
    for holder in (layout, layout.slide_master):
        for mph in holder.placeholders:
            if mph.placeholder_format.idx == idx:
                v = _from_txbody(mph)
                if v:
                    return v
    tx = layout.slide_master._element.find(qn("p:txStyles"))
    if tx is not None:
        kind = "titleStyle" if ph.placeholder_format.type in (1, 13) else "bodyStyle"
        st = tx.find(qn("p:" + kind))
        if st is not None:
            l1 = st.find(qn("a:lvl1pPr"))
            d = l1.find(qn("a:defRPr")) if l1 is not None else None
            if d is not None and d.get("sz"):
                return int(d.get("sz")) / 100.0
    return default


def _avoid_rect(x, y, w, h, logo):
    """Наибольший прямоугольник внутри исходной области вне логотипа."""
    from slide_layout import SlideFitError
    if w <= 0 or h <= 0:
        raise SlideFitError("Для контента не осталось места в выбранной области шаблона.")
    if not logo:
        return x, y, w, h
    left, top, right, bottom = logo
    if x + w <= left or x >= right or y + h <= top or y >= bottom:
        return x, y, w, h
    candidates = [(x, y, left - .015 - x, h), (right + .015, y, x + w - right - .015, h),
                  (x, y, w, top - .015 - y), (x, bottom + .015, w, y + h - bottom - .015)]
    available = [r for r in candidates if r[2] > 0 and r[3] > 0]
    if not available:
        raise SlideFitError("Логотип занимает всю область контента.")
    return max(available, key=lambda rect: rect[2] * rect[3])


def _add_cards(slide, prs, cards, accent, font, zone, bg, ink, rounded):
    """«Окошки»: 2-4 карточки с заголовком и текстом (поле cards)."""
    sw, sh = prs.slide_width, prs.slide_height
    zx, zy, zw, zh = zone
    zone_x, zone_y = int(sw * zx), int(sh * zy)
    zone_w, zone_h = int(sw * zw), int(sh * zh)
    n = min(max(len(cards), 1), 4)
    cols = 1 if n < 3 and zone_w < zone_h * 1.7 else min(n, 2)
    rows = (n + cols - 1) // cols
    gap = int(sw * 0.015)
    cw = int((zone_w - gap * (cols - 1)) / cols)
    ch = int((zone_h - gap * (rows - 1)) / rows)
    if min(cw, ch) <= 0:
        from slide_layout import SlideFitError
        raise SlideFitError("Для карточек не осталось места.")
    for i, card in enumerate(cards[:4]):
        cx = zone_x + (i % cols) * (cw + gap)
        cy = zone_y + (i // cols) * (ch + gap)
        form = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        box = slide.shapes.add_shape(form, cx, cy, cw, ch)
        box.fill.solid()
        box.fill.fore_color.rgb = _rgb(bg)
        srgb = box.fill.fore_color._xFill.find(qn("a:srgbClr"))
        alpha = srgb.makeelement(qn("a:alpha"), {"val": "92000"})
        srgb.append(alpha)
        box.line.color.rgb = _rgb(accent)
        box.line.width = Pt(1.25)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = Emu(int(Inches(0.14)))
        tf.margin_right = Emu(int(Inches(0.14)))
        tf.margin_top = Emu(int(Inches(0.08)))
        tf.margin_bottom = tf.margin_top
        title, body_text = str(card.get("title", "")), str(card.get("text", ""))
        rows_to_fit = [(title, 1.12)] if title else []
        if body_text:
            rows_to_fit.append((body_text, 1.))
        size = _text_size(box, rows_to_fit, 18, 28, font)
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = title or body_text
        run.font.size = Pt(size * 1.12 if title else size)
        run.font.bold = True
        run.font.color.rgb = _rgb(readable(accent, bg, (ink,)))
        if font:
            run.font.name = font
        if body_text and title:
            para.space_after = Pt(size * .18)
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.LEFT
            r2 = p2.add_run()
            r2.text = body_text
            r2.font.size = Pt(size)
            r2.font.color.rgb = _rgb(ink)
            if font:
                r2.font.name = font


def _template_font_pptx(prs) -> "str | None":
    """Самый частый шрифт в текстах шаблона .pptx."""
    counts: dict = {}
    for s in prs.slides:
        for sp in s.shapes:
            if not sp.has_text_frame:
                continue
            for para in sp.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.name:
                        counts[run.font.name] = counts.get(run.font.name, 0) + 1
    return max(counts, key=counts.get) if counts else None


def _template_font_pdf(doc) -> "str | None":
    """Самый частый шрифт по спанам PDF (семейство без начертания)."""
    counts: dict = {}
    for page in doc:
        for b in page.get_text("dict").get("blocks", []):
            for l in b.get("lines", []):
                for sp in l.get("spans", []):
                    name = sp.get("font")
                    if name:
                        counts[name] = counts.get(name, 0) + len(sp.get("text", ""))
    if not counts:
        return None
    name = max(counts, key=counts.get)
    for suf in ("-Bold", "-Regular", "-Light", "-Semibold", "-Medium",
                " Bold", " Regular"):
        if name.endswith(suf):
            return name[: -len(suf)]
    return name


def _logo_rect(page) -> "tuple | None":
    """Ищет логотип: компактный кластер пикселей, отличных от фона угла.
    Возвращает доли страницы (x0, y0, x1, y1) или None."""
    import fitz
    r = page.rect
    bg255 = tuple(c * 255 for c in (_corner_color(page) or (1, 1, 1)))
    zones = ((0.03, 0.02, 0.32, 0.20), (0.68, 0.02, 0.97, 0.20),
             (0.03, 0.74, 0.32, 0.97), (0.68, 0.74, 0.97, 0.97))
    best = None
    for fx0, fy0, fx1, fy1 in zones:
        clip = fitz.Rect(r.width * fx0, r.height * fy0,
                         r.width * fx1, r.height * fy1)
        pix = page.get_pixmap(clip=clip, matrix=fitz.Matrix(0.3, 0.3))
        w, h, n, data = pix.width, pix.height, pix.n, pix.samples
        pts = []
        for yy in range(h):
            row = yy * w * n
            for xx in range(w):
                off = row + xx * n
                dist = (abs(data[off] - bg255[0]) + abs(data[off + 1] - bg255[1])
                        + abs(data[off + 2] - bg255[2]))
                if dist > 130:
                    pts.append((xx, yy))
        frac = len(pts) / max(w * h, 1)
        if len(pts) < 15 or frac > 0.45:
            continue  # зона пустая или это крупный цветной элемент, не лого
        xs = [q[0] for q in pts]
        ys = [q[1] for q in pts]
        rect = ((clip.x0 + min(xs) / w * clip.width) / r.width,
                (clip.y0 + min(ys) / h * clip.height) / r.height,
                (clip.x0 + max(xs) / w * clip.width) / r.width,
                (clip.y0 + max(ys) / h * clip.height) / r.height)
        bw = max(max(xs) - min(xs), 1)
        bh = max(max(ys) - min(ys), 1)
        if min(bw, bh) <= 6:
            continue   # линии и прочий плоский декор — не лого
        density = len(pts) / (bw * bh)   # лого — компактный плотный кластер
        if density < 0.12:
            continue
        if best is None or density > best[1]:
            best = (rect, density)
    return best[0] if best else None


def _pptx_to_pdf(template: Path, out_pdf: Path) -> None:
    import subprocess
    script = Path(__file__).resolve().parents[1] / "scripts" / "pptx2pdf.ps1"
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File",
                    str(script), "-in", str(template.resolve()),
                    "-out", str(out_pdf.resolve())],
                   capture_output=True, timeout=180, check=True)


def _template_analysis(template: Path) -> dict:
    """Кэш анализа привязан к содержимому загруженного файла."""
    import json as _json
    import hashlib
    import tempfile
    import fitz
    cache = template.parent / "pdf_cache" / f"{template.name}-analysis"
    cache.mkdir(parents=True, exist_ok=True)
    marker = cache / "data.json"
    digest = hashlib.sha256(template.read_bytes()).hexdigest()
    if marker.exists():
        try:
            d = _json.loads(marker.read_text(encoding="utf-8"))
            if d.get("sha256") == digest:
                return d
        except Exception:
            pass
    with tempfile.TemporaryDirectory(dir=cache) as tmp:
        pdf_path = template
        if template.suffix.lower() != ".pdf":
            pdf_path = Path(tmp) / "template.pdf"
            _pptx_to_pdf(template, pdf_path)
        with fitz.open(str(pdf_path)) as doc:
            pages = list(doc)
            out = {"sha256": digest, "font": _template_font_pdf(doc),
                   "logo_first": _logo_rect(pages[0]) if pages else None,
                   "logo_content": _logo_rect(pages[1]) if len(pages) > 2 else None,
                   "logo_last": _logo_rect(pages[-1]) if len(pages) > 1 else None}
        staged = Path(tmp) / "analysis.json"
        staged.write_text(_json.dumps(out, ensure_ascii=False), encoding="utf-8")
        staged.replace(marker)
    return out

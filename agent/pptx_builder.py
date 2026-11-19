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
import re
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
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
    strip_re = re.compile(r"^\s*(?:\d+[.)]\s*|[-•*–]\s+)")
    cleaned = []
    for b in bullets or []:
        for ln in str(b).splitlines():
            ln = _no_emoji(strip_re.sub("", ln.strip()).strip())
            if ln:
                cleaned.append(ln)
    s["bullets"] = cleaned
    raw = bullets or []
    # если модель прислала словари с подпунктами ({"text": ..., "sub": [...]}) —
    # сохраняем структуру; иначе работаем со строками
    if any(isinstance(b, dict) for b in raw):
        items = []
        for b in raw:
            if isinstance(b, dict):
                text = _no_emoji(strip_re.sub("", str(b.get("text") or b.get("title") or "").strip()).strip())
                if not text:
                    continue
                subs = [_no_emoji(strip_re.sub("", str(x).strip()))
                        for x in (b.get("sub") or b.get("items") or [])][:MAX_SUB]
                subs = [x[:MAX_SUB_LEN] for x in subs if x]
                items.append({"text": text[:MAX_BULLET_LEN], "sub": subs})
            else:
                for ln in str(b).splitlines():
                    ln = _no_emoji(strip_re.sub("", ln.strip()).strip())
                    if ln:
                        items.append({"text": ln[:MAX_BULLET_LEN], "sub": []})
        s["bullets"] = items
    else:
        cleaned = []
        for b in raw:
            for ln in str(b).splitlines():
                ln = strip_re.sub("", ln.strip()).strip()
                if ln:
                    cleaned.append(ln)
        s["bullets"] = cleaned
    # блоки с крупными цифрами: {"value": "100%", "caption": "..."}
    stats = []
    for st in (s.get("stats") or [])[:MAX_STATS]:
        if isinstance(st, dict):
            v = str(st.get("value") or st.get("number") or "").strip()[:14]
            c = str(st.get("caption") or st.get("label") or "").strip()[:80]
            if v:
                stats.append({"value": v, "caption": c})
    s["stats"] = stats
    s["frame"] = bool(s.get("frame"))
    stats = []
    for st in (s.get("stats") or [])[:MAX_STATS]:
        if isinstance(st, dict):
            v = _no_emoji(str(st.get("value") or st.get("number") or "")).strip()[:14]
            c = _no_emoji(str(st.get("caption") or st.get("label") or "")).strip()[:80]
            if v:
                stats.append({"value": v, "caption": c})
    s["stats"] = stats
    cards = []
    for c in (s.get("cards") or [])[:4]:
        if isinstance(c, dict):
            ttl = _no_emoji(str(c.get("title") or "")).strip()[:60]
            txt = _no_emoji(str(c.get("text") or c.get("body") or "")).strip()[:200]
            if ttl or txt:
                cards.append({"title": ttl, "text": txt})
    s["cards"] = cards
    s["frame"] = bool(s.get("frame"))
    fs = s.get("font_size") or s.get("body_size") or s.get("font")
    try:
        s["font_size"] = min(40, max(12, int(fs)))
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
        kind = slide.get("kind", "content")
        if kind not in kinds:
            issues.append(f"слайд {i}: неизвестный kind «{kind}» (есть: {', '.join(sorted(kinds))})")
        if not slide.get("title"):
            issues.append(f"слайд {i}: нет title")
        if kind in ("title", "closing"):
            if len(slide.get("title") or "") > 70:
                issues.append(f"слайд {i}: заголовок титула длиннее 70 символов — сократи")
            if len(slide.get("subtitle") or "") > 90:
                issues.append(f"слайд {i}: подзаголовок длиннее 90 символов — сократи")
        bullets = _plain_bullets(slide.get("bullets"))
        cards = slide.get("cards") or []
        if kind == "content" and not bullets and not cards:
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
        vis = sum(_est_visual_lines(b["text"], 18, int(SLIDE_W * 0.66)) for b in bullets)
        vis += sum(_est_visual_lines(s, 14, int(SLIDE_W * 0.58))
                   for b in bullets for s in b["sub"])
        if vis > 11:
            issues.append(
                f"слайд {i}: после переносов получится ~{vis} строк — текст не влезет; "
                "сократи формулировки (пункты до ~60 символов) или раздели на два слайда")
        if slide.get("font_size") is not None and \
                not isinstance(slide.get("font_size"), int):
            issues.append(f"слайд {i}: font_size — целое число 12-40")
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
    if slides[0].get("kind", "content") != "title":
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
    # верхнеуровневый подзаголовок (компания) — на титульный слайд, если у него нет своего
    for sd in data.get("slides", []):
        if sd.get("kind", sd.get("type", "")) in ("title", "cover", "intro") \
                and not sd.get("subtitle") and data.get("subtitle"):
            sd["subtitle"] = data["subtitle"]
    if template and template.exists() and template.suffix.lower() == ".pdf":
        # PDF-шаблон: страницы становятся фонами слайдов, текст — на карточках
        prs = Presentation()
        _build_on_pdf(prs, data, t, template, remove_images=template_images)
        prs.save(path)
        return
    if template and template.exists():
        prs = Presentation(str(template))
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
        if analysis.get("font") and not t.get("font"):
            t["font"] = analysis["font"]
        if _build_on_skeleton(prs, data, t,
                              logo_title=analysis.get("logo_first"),
                              logo_content=analysis.get("logo_content"),
                              logo_last=analysis.get("logo_last")):
            prs.save(path)
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
    prs.save(path)


def _emit_slides(prs, data, t, title_layout, content_layout, section_layout,
                 use_template_layouts: bool) -> None:
    """Старая логика генерации слайдов, но с выбором лейаута.

    use_template_layouts=True  — текст в плейсхолдерах лейаутов шаблона
                                 (фон/шрифты/цвета наследуются от шаблона);
    use_template_layouts=False — собственный дизайн по умолчанию."""
    add_slide = None
    add_text = None
    if not use_template_layouts:
        def add_slide(bg_color: str):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = _rgb(bg_color)
            return slide

        def add_text(slide, left, top, width, height, text, size, color,
                     bold=False, align=PP_ALIGN.LEFT):
            box = slide.shapes.add_textbox(left, top, width, height)
            frame = box.text_frame
            frame.word_wrap = True
            para = frame.paragraphs[0]
            para.alignment = align
            run = para.add_run()
            run.text = text
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.name = t["font"]
            run.font.color.rgb = _rgb(color)
            return box

    for slide_data in data["slides"]:
        kind = slide_data.get("kind", "content")
        title = slide_data.get("title", "")
        subtitle = slide_data.get("subtitle", "")

        if kind in ("title", "closing"):
            slide = prs.slides.add_slide(title_layout)
            if use_template_layouts:
                # дизайн — из лейаута шаблона, меняем только текст
                _fill_placeholder(slide, (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE),
                                  title, bold=True)
                if subtitle:
                    _fill_placeholder(slide, (PP_PLACEHOLDER.SUBTITLE,), subtitle)
            else:
                add_slide(t["dark"])
                add_text(slide, Inches(1), Inches(2.4), SLIDE_W - Inches(2), Inches(1.4),
                         title, 44, t["light"], bold=True, align=PP_ALIGN.CENTER)
                if subtitle:
                    add_text(slide, Inches(1), Inches(3.9), SLIDE_W - Inches(2), Inches(1),
                             subtitle, 20, t["accent"] if kind == "title" else t["light"],
                             align=PP_ALIGN.CENTER)

        elif kind == "section":
            slide = prs.slides.add_slide(section_layout)
            if use_template_layouts:
                _fill_placeholder(slide, (PP_PLACEHOLDER.TITLE,), title, bold=True)
            else:
                add_slide(t["accent"])
                add_text(slide, Inches(1), Inches(3), SLIDE_W - Inches(2), Inches(1.4),
                         title, 40, t["light"], bold=True, align=PP_ALIGN.CENTER)

        else:  # content
            bullets = slide_data.get("bullets") or []
            stats = slide_data.get("stats") or []
            image_name = slide_data.get("image")
            image_path = ASSETS_DIR / Path(str(image_name)).name if image_name else None
            has_image = bool(image_path and image_path.exists())
            slide = prs.slides.add_slide(content_layout)
            if use_template_layouts:
                _fill_placeholder(slide, (PP_PLACEHOLDER.TITLE,), title, bold=True, size=30)
                body = _fill_placeholder(slide, (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY), "")
                if body is not None:
                    body.text_frame.word_wrap = True
                    if has_image:
                        body.width = int(prs.slide_width * 0.52)
                    for i, bullet in enumerate(bullets):
                        para = body.text_frame.paragraphs[0] if i == 0 else body.text_frame.add_paragraph()
                        para.space_after = Pt(12)
                        run = para.add_run()
                        run.text = str(bullet)
                        run.font.size = Pt(20 if len(bullets) <= 4 else 16)
            else:
                add_slide(t["bg"])
                # акцентная полоса слева
                bar = slide.shapes.add_shape(1, 0, 0, Inches(0.18), SLIDE_H)  # 1 = прямоугольник
                bar.fill.solid()
                bar.fill.fore_color.rgb = _rgb(t["accent"])
                bar.line.fill.background()
                title_box = add_text(slide, Inches(0.7), Inches(0.5), SLIDE_W - Inches(1.4), Inches(1),
                         title, 32, t["accent"], bold=True)
                if sd.get("underline", True):
                    _add_title_underline(slide, title_box, t["accent"])
                text_width = Inches(6.4) if has_image else SLIDE_W - Inches(1.8) if has_image else SLIDE_W - Inches(1.8)
                body = slide.shapes.add_textbox(Inches(0.9), Inches(1.7),
                                                text_width, SLIDE_H - Inches(2.2))
                frame = body.text_frame
                frame.word_wrap = True
                if stats:
                    body.height = int(SLIDE_H * 0.5)
                _fill_bullets_box(body, bullets, sd.get("font_size") or (22 if len(bullets) <= 4 else 18),
                                  height=body.height)
            if sd.get("frame"):
                _add_frame(slide, prs,
                           Inches(0.3), Inches(0.35), SLIDE_W - Inches(0.6), SLIDE_H - Inches(0.7),
                           t["accent"])
            if stats:
                _add_stats_row(slide, prs, stats, t["accent"])
            if has_image:
                slide.shapes.add_picture(str(image_path), Inches(7.5), Inches(1.7),
                                         width=Inches(5.2))


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
    Возвращает False, если в шаблоне меньше двух слайдов."""
    slides = list(prs.slides)
    if len(slides) < 2:
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

    def take(role: str):
        pool = pools[role]
        for s in pool:
            if id(s) not in used:
                used.add(id(s))
                return s
        clone = _duplicate_slide(prs, pool[-1])
        modes[id(clone)] = modes[id(pool[-1])]
        used.add(id(clone))  # клон — тоже «наш», иначе цикл удаления его снесёт
        return clone

    for sd in data["slides"]:
        kind = sd.get("kind", "content")
        if kind == "title":
            base = take("title")
        elif kind == "closing":
            base = take("closing")
        else:  # content и section — средние листы
            base = take("content")
        logo = (logo_title if kind == "title"
                else logo_last if kind == "closing" else logo_content)
        img_left = _tpl_image_left(base, prs)
        if modes[id(base)] == "ph":
            _retext_slide(prs, base, sd, t, logo=logo, img_left=img_left)
        else:
            _retext_freeform(prs, base, sd, t, logo=logo, img_left=img_left)
        ordered.append(base)

    # чужие слайды, не понадобившиеся ни под один наш, — удалить
    for s in list(prs.slides):
        if id(s) not in used:
            _delete_slide(prs, s)
    _reorder(prs, ordered)
    return True


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
    """Титульные плейсхолдеры: правый край не дальше 95% слайда, перенос
    включён, кегль подобран под реальную ширину/высоту бокса."""
    if ph.left is None or ph.width is None:
        return
    ph.text_frame.word_wrap = True
    right_edge = int(prs.slide_width * 0.95)
    if ph.left + ph.width > right_edge:
        ph.width = max(int(prs.slide_width * 0.4), right_edge - ph.left)
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


def _retext_slide(prs, slide, sd: dict, t: dict, logo=None, img_left=None) -> None:
    """Заменить текст в ПЛЕЙСХОЛДЕРАХ слайда шаблона (дизайн остаётся).
    Посторонний текст шаблона (не в плейсхолдерах) очищается."""
    kind = sd.get("kind", "content")
    title = str(sd.get("title") or "")
    subtitle = str(sd.get("subtitle") or "")
    bullets = (sd.get("bullets") or []) if kind == "content" else []
    if sd.get("cards"):
        bullets = []  # карточки заменяют пункты — иначе текст торчит из-под окошек
    stats = sd.get("stats") or []
    accent = t.get("accent", DEFAULT_THEME["accent"])
    title_ph = None
    bodies = []  # object/body плейсхолдеры слайда (первый заполняем, остальные чистим)
    for ph in slide.placeholders:
        ptype = ph.placeholder_format.type
        if ptype in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE):
            frame = ph.text_frame
            frame.clear()
            frame.paragraphs[0].text = title
            title_ph = ph
        elif ptype == PP_PLACEHOLDER.SUBTITLE:
            frame = ph.text_frame
            frame.clear()
            frame.paragraphs[0].text = subtitle or ("Контакты" if kind == "closing" else "")
        elif ptype in (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY):
            bodies.append(ph)
    body = bodies[0] if bodies else None
    extra_bodies = bodies[1:]  # вторые object-плейсхолдеры — чужой текст
    image_path = _slide_image(sd)
    if body is not None:
        frame = body.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # PowerPoint сам подожмёт при необходимости
        if image_path:
            _shrink_body_to(prs, body, 0.55)      # наша картинка встанет на 0.55
        if img_left is not None:
            _shrink_body_to(prs, body, img_left)  # картинка из шаблона
        if stats:
            body.height = int(prs.slide_height * 0.34)
        _fill_bullets_ph(frame, bullets, body.width, body.height,
                         start_pt=sd.get("font_size") or _ph_font_pt(slide, body))
    for ph in extra_bodies:
        ph.text_frame.clear()
    for sh in _empty_pic_placeholders(slide):
        _remove_shape(sh)          # пустые рамки-плейсхолдеры шаблона
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.is_placeholder \
                and shape.text_frame.text.strip():
            shape.text_frame.clear()
    if kind in ("title", "closing"):
        # титульные листы: геометрия и кегль подгониаются под шаблон
        for ph in slide.placeholders:
            ptype = ph.placeholder_format.type
            if ptype in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE,
                         PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY,
                         PP_PLACEHOLDER.OBJECT):
                _fit_ph_text(prs, ph, _ph_font_pt(slide, ph),
                             min_pt=14 if ptype == PP_PLACEHOLDER.SUBTITLE else 24)
    if title_ph is not None and kind == "content" and sd.get("underline", True) \
            and not _has_title_decor(slide, prs) \
            and _est_visual_lines(title, _ph_font_pt(slide, title_ph), title_ph.width) <= 1:
        _add_title_underline(slide, title_ph, accent)
    if image_path:
        x = 0.55
        w = 0.40
        if img_left is not None and img_left + 0.02 > x:
            x = min(img_left + 0.02, 0.60)
            w = max(0.2, min(w, 0.95 - x))
        _add_template_image(slide, prs, image_path, x, 0.25, w)
    if sd.get("frame"):
        _add_frame(slide, prs,
                   int(prs.slide_width * 0.025), int(prs.slide_height * 0.025),
                   int(prs.slide_width * 0.95), int(prs.slide_height * 0.95), accent)
    if stats:
        _add_stats_row(slide, prs, stats, accent, font=t.get("font"), avoid=logo)
    if sd.get("cards"):
        _add_cards(slide, prs, sd["cards"], accent, font=t.get("font"), avoid=logo)


def _retext_freeform(prs, slide, sd: dict, t: dict, logo=None, img_left=None) -> None:
    """То же для «дизайнерских» слайдов: текст живёт в обычных текстбоксах.
    Крупнейший текстбокс считаем заголовком, следующий — телом/подзаголовком.
    Форматирование ранов-прототипов сохраняем."""
    kind = sd.get("kind", "content")
    title = str(sd.get("title") or "")
    subtitle = str(sd.get("subtitle") or "")
    bullets = (sd.get("bullets") or []) if kind == "content" else []
    if sd.get("cards"):
        bullets = []  # карточки заменяют пункты
    stats = sd.get("stats") or []
    accent = t.get("accent", DEFAULT_THEME["accent"])
    image_path = _slide_image(sd)
    for sh in _empty_pic_placeholders(slide):
        _remove_shape(sh)          # пустые рамки-плейсхолдеры шаблона
    texts = _text_shapes(slide)
    if not texts:
        if image_path:
            _add_template_image(slide, prs, image_path)
        return

    if kind == "content":
        _set_shape_lines(texts[0], [title])
        title_shape = texts[0]
        body = texts[1] if len(texts) > 1 else None
        if body is not None:
            if bullets:
                _set_shape_lines(body, bullets)
                proto_pt = sd.get("font_size") or _proto_font_pt(body)
                pt, _f = _fit_bullets_pt(bullets, body.width, body.height,
                                         start_pt=proto_pt or 20, min_pt=12)
                if proto_pt and pt < proto_pt:
                    _shrink_runs(body, pt / proto_pt)
                body.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            if image_path:
                _shrink_body_to(prs, body, 0.55)
            if img_left is not None:
                _shrink_body_to(prs, body, img_left)
            if stats:
                body.height = int(prs.slide_height * 0.30)
            if sd.get("frame"):
                _add_frame(slide, prs,
                           int(prs.slide_width * 0.03), int(prs.slide_height * 0.20),
                           int(prs.slide_width * 0.94), int(prs.slide_height * 0.72), accent)
            if sd.get("underline", True) \
                    and not _has_title_decor(slide, prs) \
                    and _est_visual_lines(title, _proto_font_pt(title_shape) or 36,
                                          title_shape.width) <= 1:
                _add_title_underline(slide, title_shape, accent)
            for sp in texts[2:]:
                _set_shape_lines(sp, [])
        if stats:
            _add_stats_row(slide, prs, stats, accent, font=t.get("font"), avoid=logo)
        if sd.get("cards"):
            _add_cards(slide, prs, sd["cards"], accent, font=t.get("font"), avoid=logo)
        if image_path:
            _add_template_image(slide, prs, image_path)
        return

    # титул / финал / разделитель: крупный шейп <- заголовок (+ подзаголовок
    # вторым абзацем, как «Должность докладчика» в шаблоне Политеха)
    note = subtitle or ("Контакты" if kind == "closing" else "")
    lines = [title] + ([note] if note else [])
    _set_shape_lines(texts[0], lines)
    if kind in ("title", "closing"):
        box = texts[0]
        box.text_frame.word_wrap = True          # иначе текст уезжает за край
        if box.left is not None and box.width is not None:
            # нормальная ширина титульной зоны (узкие колонки шаблона не читаются)
            target_w = int(prs.slide_width * 0.85) - box.left
            if box.width < target_w:
                box.width = target_w
            right_edge = int(prs.slide_width * 0.9)
            if box.left + box.width > right_edge:
                box.width = right_edge - box.left
        box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for sp in texts[1:]:
        _set_shape_lines(sp, [])
    if image_path:
        _add_template_image(slide, prs, image_path, 0.62, 0.40, 0.30)


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


def _fit_bullets_pt(bullets, width_emu, height_emu, start_pt=20, min_pt=12):
    """Подобрать базовый кегль, чтобы оценённый объём (с переносами)
    влез в бокс по высоте. Возвращает (кегль, влезает ли)."""
    lines = _bullet_lines(bullets)
    if not lines or not width_emu or not height_emu:
        return start_pt, True
    box_pt = height_emu / 12700.0
    pt = start_pt
    while True:
        used = 0.0
        for text, level in lines:
            size = pt * (0.8 if level else 1.0)
            used += _est_visual_lines(text, size, width_emu) * size * 1.22
        used += 6 * (len(lines) - 1)          # интервалы между абзацами
        if used <= box_pt or pt <= min_pt:
            return pt, used <= box_pt
        pt -= 1


def _est_visual_lines(text, font_pt, width_emu, char_ratio=0.54):
    """Оценка числа строк после переноса: средний символ ~0.54 кегля."""
    import math
    text = str(text or "")
    if not text:
        return 1
    usable_pt = max(width_emu / 12700.0 - 16, 60)
    cpl = max(int(usable_pt / (font_pt * char_ratio)), 8)
    return max(math.ceil(len(text) / cpl), 1)


def _fit_bullets_pt(bullets, width_emu, height_emu, start_pt=20, min_pt=12):
    """Подобрать базовый кегль, чтобы оценённый объём (с переносами)
    влез в бокс по высоте. Возвращает (кегль, влезает ли)."""
    lines = _bullet_lines(bullets)
    if not lines or not width_emu or not height_emu:
        return start_pt, True
    box_pt = height_emu / 12700.0
    pt = start_pt
    while True:
        used = 0.0
        for text, level in lines:
            size = pt * (0.8 if level else 1.0)
            used += _est_visual_lines(text, size, width_emu) * size * 1.22
        used += 6 * (len(lines) - 1)          # интервалы между абзацами
        if used <= box_pt or pt <= min_pt:
            return pt, used <= box_pt
        pt -= 1


def _fill_bullets_ph(frame, bullets, width=None, height=None, start_pt=20):
    """Буллеты в плейсхолдере: подпункты через level=1; кегль подобран
    под оценённый объём текста (переносы учитываются)."""
    lines = _bullet_lines(bullets)
    if not lines:
        return
    pt, _fits = _fit_bullets_pt(bullets, width, height, start_pt=start_pt, min_pt=12)
    for i, (text, level) in enumerate(lines):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = text
        if level:
            para.level = 1
        for run in para.runs:
            run.font.size = Pt(int(pt * (0.8 if level else 1.0)))


def _fill_bullets_box(body, bullets, base_size, height=None):
    """Буллеты в обычном текстбоксе: подпункты — тире с отступом и мельче;
    кегль ужимается под оценённый объём."""
    lines = _bullet_lines(bullets)
    if not lines:
        return
    body.text_frame.word_wrap = True
    if height:
        base_size, _fits = _fit_bullets_pt(bullets, body.width, height,
                                           start_pt=base_size, min_pt=11)
    for i, (text, level) in enumerate(lines):
        para = body.text_frame.paragraphs[0] if i == 0 else body.text_frame.add_paragraph()
        para.text = ("      -  " if level else "*  ") + text
        for run in para.runs:
            run.font.size = Pt(int(base_size * (0.8 if level else 1)))


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


def _add_stats_row(slide, prs, stats, accent, font=None, avoid=None):
    """Ряд крупных цифр с подписями и акцентной чертой слева.
    avoid — прямоугольник логотипа в долях страницы: ряд его не пересекает."""
    sw, sh = prs.slide_width, prs.slide_height
    n = max(len(stats), 1)
    gap = int(sw * 0.02)
    fx, fy, fw, fh = _avoid_rect(0.05, 0.74, 0.88, 0.18, avoid)
    x, y = int(sw * fx), int(sh * fy)
    w, h = int(sw * fw), int(sh * fh)
    w = int((w - gap * (n - 1)) / n)
    font_name = font
    for st in stats:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y,
                                     int(sw * 0.006), h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(accent)
        bar.line.fill.background()
        bar.shadow.inherit = False
        value = slide.shapes.add_textbox(x + int(sw * 0.014), y,
                                         w - int(sw * 0.02), int(h * 0.55))
        para = value.text_frame.paragraphs[0]
        run = para.add_run()
        run.text = str(st.get("value", ""))
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = _rgb(accent)
        if font_name:
            run.font.name = font_name
        cap = slide.shapes.add_textbox(x + int(sw * 0.014), y + int(h * 0.55),
                                       w - int(sw * 0.02), int(h * 0.45))
        cap.text_frame.word_wrap = True
        cpara = cap.text_frame.paragraphs[0]
        crun = cpara.add_run()
        crun.text = str(st.get("caption", ""))
        crun.font.size = Pt(11)
        crun.font.color.rgb = _rgb("2B2B2B")
        if font_name:
            crun.font.name = font_name
        x += w + gap
def _slide_image(sd: dict) -> Path | None:
    """Путь к картинке слайда из активов пользователя (или None)."""
    name = sd.get("image")
    if not name:
        return None
    path = ASSETS_DIR / Path(str(name)).name
    return path if path.exists() else None


def _add_template_image(slide, prs, image_path: Path | None,
                        x_frac: float = 0.55, y_frac: float = 0.25,
                        w_frac: float = 0.40) -> None:
    """Картинка справа, геометрия — в долях от РЕАЛЬНОГО размера слайда
    (шаблон может быть 4:3, а не 16:9 — фиксированные дюймы наедут на текст).
    Для «дизайнерских» слайдов (freeform) передаются x/y/w ниже и правее,
    чтобы не накрывать рамку заголовка и текст."""
    if not image_path:
        return
    sw, sh = prs.slide_width, prs.slide_height
    width = int(sw * w_frac)
    x = int(sw * x_frac)
    y = int(sh * y_frac)
    slide.shapes.add_picture(str(image_path), x, y, width=width)


def _set_shape_lines(shape, lines: list) -> None:
    """Заменить текст шейпа списком строк (абзацев), сохранив форматирование:
    первый ран существующего абзаца — прототип, остальные абзацы — его клоны."""
    frame = shape.text_frame
    txBody = frame._txBody
    paras = txBody.findall(qn("a:p"))
    proto = next((p for p in paras if p.findall(qn("a:r"))), None)
    if proto is None:
        frame.clear()
        frame.paragraphs[0].text = "\n".join(lines)
        return

    def set_para_text(p_el, text: str) -> None:
        runs = p_el.findall(qn("a:r"))
        keep = runs[0]
        for child in list(p_el):
            if child is keep or child.tag == qn("a:pPr"):
                continue
            p_el.remove(child)
        t_el = keep.find(qn("a:t"))
        if t_el is None:
            from lxml import etree
            t_el = etree.SubElement(keep, qn("a:t"))
        t_el.text = text

    set_para_text(proto, lines[0] if lines else "")
    for p_el in paras:
        if p_el is not proto:
            txBody.remove(p_el)
    for line in lines[1:]:
        clone = copy.deepcopy(proto)
        set_para_text(clone, line)
        txBody.append(clone)


def _duplicate_slide(prs, source):
    """Клон слайда шаблона: тот же лейаут, тот же дизайн (включая фон слайда)."""
    new = prs.slides.add_slide(source.slide_layout)
    for shape in list(new.shapes):  # пустые клоны плейсхолдеров лейаута не нужны
        shape._element.getparent().remove(shape._element)
    bg = source._element.cSld.find(qn("p:bg"))
    if bg is not None:
        dst = new._element.cSld
        old = dst.find(qn("p:bg"))
        if old is not None:
            dst.remove(old)
        dst.insert(0, copy.deepcopy(bg))
    for shape in source.shapes:
        new.shapes._spTree.append(copy.deepcopy(shape._element))
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

def _rasterize_pdf(template: Path, remove_images: str = "auto") -> list[Path]:
    """Страницы PDF -> PNG-картинки (кэш в data/templates/pdf_cache/<stem>-<режим>/).

    remove_images: "auto" — чужие фото распознаются и удаляются, гладкие фоны
    (градиенты/паттерны) остаются; "remove" — убрать все картинки; "keep" —
    убрать только мелкие (<60% страницы).
    Логика на страницу: (1) вырезаем чужой ТЕКСТ по спанам цветом окружения
    и счищаем векторный декор, задетый вырезкой; (2) уже на очищенной странице
    решаем судьбу КАРТИНОК и вырезаем лишние."""
    import fitz
    cache = template.parent / "pdf_cache" / f"{template.stem}-{remove_images}"
    pdf_mtime = template.stat().st_mtime
    marker = cache / "source.mtime"
    pages = sorted(cache.glob("p*.png"), key=lambda f: int(f.stem[1:]))
    if pages and marker.exists():
        try:
            if abs(float(marker.read_text()) - pdf_mtime) < 1:
                return pages
        except (ValueError, OSError):
            pass
    doc = fitz.open(str(template))
    cache.mkdir(parents=True, exist_ok=True)
    for old_png in cache.glob("p*.png"):
        old_png.unlink()
    zoom = 1600 / max(doc[0].rect.width, 1)
    for i, page in enumerate(doc):
        corner = _corner_color(page)
        page_area = page.rect.width * page.rect.height

        # --- проход 1: ЧУЖОЙ ТЕКСТ (по словам — накрывает и маркеры списков) ---
        for word_rect in _word_rects(page):
            page.add_redact_annot(word_rect, fill=_block_fill_color(page, word_rect))
        if i == 0:
            # титульная страница: сносим центральную полосу и правую часть шапки
            # (там их заголовок/меню и декор), оставляем логотип слева сверху
            r = page.rect
            page.add_redact_annot(
                fitz.Rect(0, r.height * 0.26, r.width, r.height * 0.90), fill=corner)
            page.add_redact_annot(
                fitz.Rect(r.width * 0.30, 0, r.width, r.height * 0.25), fill=corner)
        graphics_mode = (fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED if i == 0
                         else fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED)
        if page.get_text():
            try:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                      graphics=graphics_mode)
            except TypeError:
                page.apply_redactions()

        # --- проход 2: ЧУЖИЕ КАРТИНКИ (решаем на уже очищенной странице) ---
        for img in page.get_images(full=True):
            for rect in page.get_image_rects(img[0]):
                area_frac = (rect.width * rect.height) / page_area
                drop = area_frac < 0.6
                if not drop and remove_images == "all":
                    drop = True
                elif not drop and remove_images == "auto":
                    drop = _is_photo(page, rect)
                if drop:
                    page.add_redact_annot(fitz.Rect(rect) + (-2, -2, 2, 2),
                                          fill=_block_fill_color(page, rect))
        if page.get_images():
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE)

        page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(
            str(cache / f"p{i + 1}.png"))
    marker.write_text(str(pdf_mtime))
    # НАТУРАЛЬНАЯ сортировка: иначе p10, p11... встают между p1 и p2,
    # и «последняя страница» оказывается девятой
    return sorted(cache.glob("p*.png"), key=lambda f: int(f.stem[1:]))


def _block_fill_color(page, rect) -> tuple:
    """Цвет заливки вырезаемого блока: сэмплируем точки-кольцо вокруг блока
    и берём поканальную медиану — так цветные плашки дизайна сохраняются."""
    r = page.rect
    pad = 6
    cx, cy = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
    pts = [(max(rect.x0 - pad, 0), cy), (min(rect.x1 + pad, r.width), cy),
           (cx, max(rect.y0 - pad, 0)), (cx, min(rect.y1 + pad, r.height)),
           (max(rect.x0 - pad, 0), max(rect.y0 - pad, 0)),
           (min(rect.x1 + pad, r.width), min(rect.y1 + pad, r.height))]
    samples = []
    for x, y in pts:
        try:
            pix = page.get_pixmap(clip=fitz.Rect(x - 2, y - 2, x + 2, y + 2))
            samples.append(pix.pixel(0, 0))
        except Exception:
            continue
    if not samples:
        return _corner_color(page)
    n = len(samples)
    med = tuple(sorted(s[i] for s in samples)[n // 2] for i in range(3))
    return tuple(c / 255 for c in med)


def _corner_color(page) -> tuple:
    """Цвет фона страницы PDF по её углам (для заливки вырезанного текста)."""
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

def _word_rects(page) -> list:
    """Прямоугольники всех слов страницы (для вырезки чужого текста)."""
    import fitz
    return [fitz.Rect(w[:4]) + (-2, -2, 2, 2) for w in page.get_text("words")
            if w[4].strip()]


def _dominant_accent(page) -> tuple | None:
    """Самый частый НАСЫЩЕННЫЙ цвет страницы (фирменный акцент, не фон/текст)."""
    import fitz
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.12, 0.12))
        counts: dict = {}
        data = pix.samples
        step = pix.n
        w = pix.width
        for y in range(0, pix.height, 2):
            row = y * w * step
            for x in range(0, w, 2):
                off = row + x * step
                r, g, b = data[off], data[off + 1], data[off + 2]
                if max(r, g, b) - min(r, g, b) < 45:  # серые/белые/чёрные мимо
                    continue
                key = (r >> 4, g >> 4, b >> 4)
                counts[key] = counts.get(key, 0) + 1
        if not counts:
            return None
        key = max(counts, key=counts.get)
        return ((key[0] << 4) / 255, (key[1] << 4) / 255, (key[2] << 4) / 255)
    except Exception:
        return None

def _is_photo(page, rect) -> bool:
    """Похоже ли изображение на фотографию (высокая динамика яркости),
    а не на гладкий фон/градиент/паттерн."""
    try:
        pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(0.035, 0.035))
        w, h, n, data = pix.width, pix.height, pix.n, pix.samples
        diffs, prev = [], None
        for y in range(h):
            row = y * w * n
            for x in range(w):
                off = row + x * n
                lum = (data[off] + data[off + 1] + data[off + 2]) / 3
                if prev is not None:
                    diffs.append(abs(lum - prev))
                prev = lum
        if not diffs:
            return False
        return sum(diffs) / len(diffs) > 7
    except Exception:
        return False


def _card(slide, prs, x_frac, y_frac, w_frac, h_frac, alpha_pct=88):
    """Полупрозрачная белая карточка — читаемый фон для текста поверх дизайна PDF."""
    from pptx.enum.shapes import MSO_SHAPE
    sw, sh = prs.slide_width, prs.slide_height
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        int(sw * x_frac), int(sh * y_frac),
        int(sw * w_frac), int(sh * h_frac))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    srgb = box.fill.fore_color._xFill.find(qn("a:srgbClr"))
    alpha = srgb.makeelement(qn("a:alpha"), {"val": str(int(alpha_pct * 1000))})
    srgb.append(alpha)
    box.line.fill.background()
    box.shadow.inherit = False
    return box


def _pdf_text(slide, prs, text, x_frac, y_frac, w_frac, size, color,
              bold=False, align=PP_ALIGN.LEFT, h_frac=0.1):
    sw, sh = prs.slide_width, prs.slide_height
    box = slide.shapes.add_textbox(int(sw * x_frac), int(sh * y_frac),
                                   int(sw * w_frac), int(sh * h_frac))
    frame = box.text_frame
    frame.word_wrap = True
    para = frame.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    return box


def _build_on_pdf(prs, data: dict, t: dict, template: Path,
                  remove_images: str = "auto") -> None:
    """Слайды поверх страниц PDF: фон — картинка страницы, текст — на карточке.

    Титул — первая страница шаблона, ФИНАЛ — последняя (у брендовых PDF там
    контакты/призыв к действию), контентные слайды — средние страницы
    (циклически). Пропорции слайда = пропорции страницы PDF.
    """
    pages = _rasterize_pdf(template, remove_images=remove_images)
    if not pages:
        raise RuntimeError(f"не удалось растрировать PDF-шаблон {template.name}")
    import fitz
    first = fitz.open(str(template))[0]
    aspect = first.rect.height / max(first.rect.width, 1)
    prs.slide_width = SLIDE_W
    prs.slide_height = Emu(int(SLIDE_W * aspect))
    sw, sh = prs.slide_width, prs.slide_height
    accent = _dominant_accent(fitz.open(str(template))[0])
    if accent:
        # антиалиасинг размывает цвет — затемняем, пока не станет достаточно тёмным
        while sum(accent) / 3 > 0.62:
            accent = tuple(c * 0.78 for c in accent)
        title_color = "%02X%02X%02X" % tuple(int(c * 255) for c in accent)
    else:
        title_color = t.get("accent", DEFAULT_THEME["accent"])
    middle_pages = pages[1:-1] or pages  # средние страницы — под контент
    content_idx = 0

    for sd in data["slides"]:
        kind = sd.get("kind", "content")
        title = str(sd.get("title") or "")
        subtitle = str(sd.get("subtitle") or ("Контакты" if kind == "closing" else ""))
        bullets = [str(b) for b in (sd.get("bullets") or [])] if kind == "content" else []
        if kind == "title":
            bg = pages[0]              # первая страница — дизайн титула
        elif kind == "closing":
            bg = pages[-1]             # последняя — дизайн финала (контакты)
        else:
            bg = middle_pages[min(content_idx, len(middle_pages) - 1)]
            content_idx += 1
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(str(bg), 0, 0, width=sw, height=sh)

        if kind in ("title", "closing"):
            _card(slide, prs, 0.14, 0.38, 0.72, 0.24, alpha_pct=94)
            _pdf_text(slide, prs, title, 0.18, 0.43, 0.64, 32, title_color,
                      bold=True, align=PP_ALIGN.CENTER, h_frac=0.09)
            if subtitle:
                _pdf_text(slide, prs, subtitle, 0.18, 0.53, 0.64, 14, "2B2B2B",
                          align=PP_ALIGN.CENTER, h_frac=0.06)
        elif kind == "section":
            _card(slide, prs, 0.14, 0.40, 0.72, 0.18, alpha_pct=90)
            _pdf_text(slide, prs, title, 0.18, 0.45, 0.64, 28, title_color,
                      bold=True, align=PP_ALIGN.CENTER, h_frac=0.08)
        else:
            _card(slide, prs, 0.05, 0.18, 0.90, 0.74, alpha_pct=95)
            _pdf_text(slide, prs, title, 0.10, 0.24, 0.80, 26, title_color,
                      bold=True, h_frac=0.08)
            body = slide.shapes.add_textbox(int(sw * 0.10), int(sh * 0.36),
                                            int(sw * 0.80), int(sh * 0.52))
            frame = body.text_frame
            frame.word_wrap = True
            size = 16 if len(bullets) <= 5 else 13
            for i, bullet in enumerate(bullets):
                para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
                para.space_after = Pt(10)
                run = para.add_run()
                run.text = "•  " + bullet
                run.font.size = Pt(size)
                run.font.color.rgb = _rgb("2B2B2B")


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
    """Не пересекаться с логотипом: сдвиг правее или выше (доли страницы)."""
    if not logo:
        return x, y, w, h
    lx0, ly0, lx1, ly1 = logo
    if x + w <= lx0 or x >= lx1 or y + h <= ly0 or y >= ly1:
        return x, y, w, h
    if lx1 + 0.02 < 0.98:                       # вариант 1: правее лого
        nx = lx1 + 0.02
        return nx, y, max(0.2, 0.98 - nx), h
    ny = max(0.05, ly0 - h - 0.02)              # вариант 2: выше лого
    if ny + h <= ly0:
        return x, ny, w, h
    return x, y, w, h


def _add_cards(slide, prs, cards, accent, font=None, avoid=None):
    """«Окошки»: 2-4 карточки с заголовком и текстом (поле cards)."""
    sw, sh = prs.slide_width, prs.slide_height
    zx, zy, zw, zh = _avoid_rect(0.06, 0.30, 0.88, 0.48, avoid)
    zone_x, zone_y = int(sw * zx), int(sh * zy)
    zone_w, zone_h = int(sw * zw), int(sh * zh)
    n = min(max(len(cards), 1), 4)
    cols = 2 if n > 2 else n
    rows = (n + cols - 1) // cols
    gap = int(sw * 0.015)
    cw = int((zone_w - gap * (cols - 1)) / cols)
    ch = int((zone_h - gap * (rows - 1)) / rows)
    for i, card in enumerate(cards[:4]):
        cx = zone_x + (i % cols) * (cw + gap)
        cy = zone_y + (i // cols) * (ch + gap)
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch)
        box.fill.solid()
        box.fill.fore_color.rgb = _rgb("FFFFFF")
        srgb = box.fill.fore_color._xFill.find(qn("a:srgbClr"))
        alpha = srgb.makeelement(qn("a:alpha"), {"val": "92000"})
        srgb.append(alpha)
        box.line.color.rgb = _rgb(accent)
        box.line.width = Pt(1.25)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(int(Inches(0.14)))
        tf.margin_right = Emu(int(Inches(0.14)))
        tf.margin_top = Emu(int(Inches(0.08)))
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = str(card.get("title", ""))
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = _rgb(accent)
        if font:
            run.font.name = font
        body_text = str(card.get("text", ""))
        if body_text:
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.LEFT
            r2 = p2.add_run()
            r2.text = body_text[:200]
            r2.font.size = Pt(11)
            r2.font.color.rgb = _rgb("2B2B2B")
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
                   capture_output=True, timeout=180)


def _template_analysis(template: Path) -> dict:
    """Фирменный шрифт + положение логотипа на 1/контентном/последнем листе.
    Кэш по mtime шаблона: data/templates/pdf_cache/<stem>-analysis/data.json."""
    import json as _json
    cache = template.parent / "pdf_cache" / f"{template.stem}-analysis"
    cache.mkdir(parents=True, exist_ok=True)
    marker = cache / "data.json"
    mtime = template.stat().st_mtime
    if marker.exists():
        try:
            d = _json.loads(marker.read_text(encoding="utf-8"))
            if abs(float(d.get("mtime", 0)) - mtime) < 1:
                return d
        except Exception:
            pass
    if template.suffix.lower() == ".pdf":
        pdf_path = template
    else:
        pdf_path = cache / "template.pdf"
        if not pdf_path.exists():
            _pptx_to_pdf(template, pdf_path)
    import fitz
    doc = fitz.open(str(pdf_path))
    pages = list(doc)
    out = {"mtime": mtime, "font": _template_font_pdf(doc),
           "logo_first": _logo_rect(pages[0]) if pages else None,
           "logo_content": _logo_rect(pages[1]) if len(pages) > 2 else None,
           "logo_last": _logo_rect(pages[-1]) if len(pages) > 1 else None}
    marker.write_text(_json.dumps(out, ensure_ascii=False), encoding="utf-8")
    return out

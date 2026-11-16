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
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
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

# Модели часто используют свои имена полей — принимаем синонимы.
_KIND_ALIASES = {
    "text": "content", "bullets": "content", "list": "content",
    "cover": "title", "intro": "title",
    "end": "closing", "final": "closing", "contacts": "closing",
}
_BULLET_ALIASES = ("bullets", "content", "points", "items", "text")


def normalize_slide(slide: dict) -> dict:
    """Приводит слайд от модели к каноническому виду.
    kind: type/cover/bullets/... -> title|content|section|closing
    буллеты: ищем в bullets/content/points/items; строку режем на предложения."""
    s = dict(slide)
    kind = str(s.get("kind") or s.get("type") or "content").lower()
    s["kind"] = _KIND_ALIASES.get(kind, kind)
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
    cleaned = []
    for b in bullets or []:
        for ln in str(b).splitlines():
            ln = re.sub(r"^\s*(?:\d+[.)]\s*|[-•*–]\s+)", "", ln.strip()).strip()
            if ln:
                cleaned.append(ln)
    s["bullets"] = cleaned
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
        bullets = slide.get("bullets") or []
        if kind == "content" and not bullets:
            got = ", ".join(sorted(slide.keys()))
            issues.append(
                f"слайд {i}: content-слайд без буллетов "
                f"(пришли поля: {got}; нужно поле \"bullets\": [\"строка\", ...])"
            )
        if len(bullets) > MAX_BULLETS:
            issues.append(f"слайд {i}: буллетов {len(bullets)} > {MAX_BULLETS} — сократи")
        for b in bullets:
            if len(str(b)) > MAX_BULLET_LEN:
                issues.append(f"слайд {i}: буллет длиннее {MAX_BULLET_LEN} символов — сократи")
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
        if _build_on_skeleton(prs, data):
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
                add_text(slide, Inches(0.7), Inches(0.5), SLIDE_W - Inches(1.4), Inches(1),
                         title, 32, t["accent"], bold=True)
                text_width = Inches(6.4) if has_image else SLIDE_W - Inches(1.8)
                body = slide.shapes.add_textbox(Inches(0.9), Inches(1.7),
                                                text_width, SLIDE_H - Inches(2.2))
                frame = body.text_frame
                frame.word_wrap = True
                for i, bullet in enumerate(bullets):
                    para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
                    para.space_after = Pt(14)
                    run = para.add_run()
                    run.text = f"•  {bullet}"
                    run.font.size = Pt(22 if len(bullets) <= 4 else 18)
                    run.font.name = t["font"]
                    run.font.color.rgb = _rgb(t["text"])
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


def _build_on_skeleton(prs, data: dict) -> bool:
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
        if modes[id(base)] == "ph":
            _retext_slide(prs, base, sd)
        else:
            _retext_freeform(prs, base, sd)
        ordered.append(base)

    # чужие слайды, не понадобившиеся ни под один наш, — удалить
    for s in list(prs.slides):
        if id(s) not in used:
            _delete_slide(prs, s)
    _reorder(prs, ordered)
    return True


def _retext_slide(prs, slide, sd: dict) -> None:
    """Заменить текст в ПЛЕЙСХОЛДЕРАХ слайда шаблона (дизайн остаётся).
    Посторонний текст шаблона (не в плейсхолдерах) очищается."""
    kind = sd.get("kind", "content")
    title = str(sd.get("title") or "")
    subtitle = str(sd.get("subtitle") or "")
    bullets = [str(b) for b in (sd.get("bullets") or [])] if kind == "content" else []
    body = None
    for ph in slide.placeholders:
        ptype = ph.placeholder_format.type
        if ptype in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE):
            frame = ph.text_frame
            frame.clear()
            frame.paragraphs[0].text = title
        elif ptype == PP_PLACEHOLDER.SUBTITLE:
            frame = ph.text_frame
            frame.clear()
            frame.paragraphs[0].text = subtitle or ("Контакты" if kind == "closing" else "")
        elif ptype in (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY):
            body = ph
    image_path = _slide_image(sd)
    if body is not None:
        frame = body.text_frame
        frame.clear()
        frame.word_wrap = True
        if image_path:
            body.width = int(prs.slide_width * 0.52)
        for i, bullet in enumerate(bullets):
            para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            para.text = bullet
    _add_template_image(slide, prs, image_path, 0.62, 0.40, 0.30)
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.is_placeholder \
                and shape.text_frame.text.strip():
            shape.text_frame.clear()


def _retext_freeform(prs, slide, sd: dict) -> None:
    """То же для «дизайнерских» слайдов: текст живёт в обычных текстбоксах.
    Крупнейший текстбокс считаем заголовком, следующий — телом/подзаголовком.
    Форматирование ранов-прототипов сохраняем."""
    kind = sd.get("kind", "content")
    title = str(sd.get("title") or "")
    subtitle = str(sd.get("subtitle") or "")
    bullets = [str(b) for b in (sd.get("bullets") or [])] if kind == "content" else []
    image_path = _slide_image(sd)
    texts = _text_shapes(slide)
    if not texts:
        if image_path:
            _add_template_image(slide, prs, image_path)
        return

    if kind == "content":
        _set_shape_lines(texts[0], [title])
        if len(texts) > 1:
            body = texts[1]
            _set_shape_lines(body, bullets)
            if image_path and body.width > int(prs.slide_width * 0.55):
                body.width = int(prs.slide_width * 0.52)
            for sp in texts[2:]:
                _set_shape_lines(sp, [])
        if image_path:
            _add_template_image(slide, prs, image_path, 0.62, 0.40, 0.30)
        return

    # титул / финал / разделитель: крупный шейп <- заголовок (+ подзаголовок
    # вторым абзацем, как «Должность докладчика» в шаблоне Политеха)
    note = subtitle or ("Контакты" if kind == "closing" else "")
    lines = [title] + ([note] if note else [])
    _set_shape_lines(texts[0], lines)
    for sp in texts[1:]:
        _set_shape_lines(sp, [])
    if image_path:
        _add_template_image(slide, prs, image_path, 0.62, 0.40, 0.30)


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

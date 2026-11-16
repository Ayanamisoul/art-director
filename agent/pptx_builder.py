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
                       template: Path | None = None) -> None:
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
                        body.width = Inches(6.4)
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
    Возвращает False, если в шаблоне нет ни титульного, ни контентного слайда."""
    skeletons: dict[str, list] = {"title": [], "content": [], "section": []}
    modes: dict[int, str] = {}
    for s in list(prs.slides):
        role = _slide_role(s)
        mode = "ph" if role else None
        if not role:
            role = _freeform_role(s)
            mode = "ff" if role else None
        if role:
            skeletons[role].append(s)
            modes[id(s)] = mode
    if not skeletons["title"] or not skeletons["content"]:
        return False

    used: set[int] = set()
    ordered: list = []

    def take(role: str):
        pool = skeletons.get(role) or skeletons["content"] or skeletons["title"]
        for s in pool:
            if id(s) not in used:
                used.add(id(s))
                return s
        clone = _duplicate_slide(prs, pool[0])
        modes[id(clone)] = modes[id(pool[0])]
        used.add(id(clone))  # клон — тоже «наш», иначе цикл удаления его снесёт
        return clone

    for sd in data["slides"]:
        kind = sd.get("kind", "content")
        if kind in ("title", "closing"):
            base = take("title")
        elif kind == "section":
            base = take("section") if skeletons["section"] else take("content")
        else:
            base = take("content")
        if modes[id(base)] == "ph":
            _retext_slide(base, sd)
        else:
            _retext_freeform(base, sd)
        ordered.append(base)

    # чужие слайды, не понадобившиеся ни под один наш, — удалить
    for s in list(prs.slides):
        if id(s) not in used:
            _delete_slide(prs, s)
    _reorder(prs, ordered)
    return True


def _retext_slide(slide, sd: dict) -> None:
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
    if body is not None:
        frame = body.text_frame
        frame.clear()
        frame.word_wrap = True
        for i, bullet in enumerate(bullets):
            para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            para.text = bullet
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.is_placeholder \
                and shape.text_frame.text.strip():
            shape.text_frame.clear()


def _retext_freeform(slide, sd: dict) -> None:
    """То же для «дизайнерских» слайдов: текст живёт в обычных текстбоксах.
    Крупнейший текстбокс считаем заголовком, следующий — телом/подзаголовком.
    Форматирование ранов сохраняем."""
    kind = sd.get("kind", "content")
    title = str(sd.get("title") or "")
    subtitle = str(sd.get("subtitle") or "")
    bullets = [str(b) for b in (sd.get("bullets") or [])] if kind == "content" else []
    texts = _text_shapes(slide)
    if not texts:
        return
    _set_shape_lines(texts[0], [title])
    rest = texts[1:]
    if kind == "content" and rest:
        _set_shape_lines(rest[0], bullets)
        for sp in rest[1:]:
            _set_shape_lines(sp, [])
    else:
        note = subtitle or ("Контакты" if kind == "closing" else "")
        if rest:
            _set_shape_lines(rest[0], [note] if note else [])
        for sp in rest[1:]:
            _set_shape_lines(sp, [])


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

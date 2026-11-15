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
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

SLIDE_W = Inches(13.333)  # 16:9
SLIDE_H = Inches(7.5)

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
        parts = [p.strip() for p in re.split(r"(?<=[.!?])\s+", bullets) if p.strip()]
        bullets = parts if len(bullets) > MAX_BULLET_LEN else [bullets]
    s["bullets"] = bullets or []
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
    if slides[0].get("kind", "content") != "title":
        issues.append("первым должен быть титульный слайд (kind=title)")
    return issues


def build_presentation(path: Path, data: dict, theme: dict | None = None) -> None:
    """Собрать .pptx по структуре. Путь должен заканчиваться на .pptx."""
    t = {**DEFAULT_THEME, **(theme or {})}
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]  # пустой макет — всё рисуем сами

    def add_slide(bg_color: str):
        slide = prs.slides.add_slide(blank)
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
            slide = add_slide(t["dark"])
            add_text(slide, Inches(1), Inches(2.4), SLIDE_W - Inches(2), Inches(1.4),
                     title, 44, t["light"], bold=True, align=PP_ALIGN.CENTER)
            if subtitle:
                add_text(slide, Inches(1), Inches(3.9), SLIDE_W - Inches(2), Inches(1),
                         subtitle, 20, t["accent"] if kind == "title" else t["light"],
                         align=PP_ALIGN.CENTER)

        elif kind == "section":
            slide = add_slide(t["accent"])
            add_text(slide, Inches(1), Inches(3), SLIDE_W - Inches(2), Inches(1.4),
                     title, 40, t["light"], bold=True, align=PP_ALIGN.CENTER)

        else:  # content
            slide = add_slide(t["bg"])
            # акцентная полоса слева
            bar = slide.shapes.add_shape(1, 0, 0, Inches(0.18), SLIDE_H)  # 1 = прямоугольник
            bar.fill.solid()
            bar.fill.fore_color.rgb = _rgb(t["accent"])
            bar.line.fill.background()
            add_text(slide, Inches(0.7), Inches(0.5), SLIDE_W - Inches(1.4), Inches(1),
                     title, 32, t["accent"], bold=True)
            body = slide.shapes.add_textbox(Inches(0.9), Inches(1.7),
                                            SLIDE_W - Inches(1.8), SLIDE_H - Inches(2.2))
            frame = body.text_frame
            frame.word_wrap = True
            for i, bullet in enumerate(slide_data.get("bullets") or []):
                para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
                para.space_after = Pt(14)
                run = para.add_run()
                run.text = f"•  {bullet}"
                run.font.size = Pt(22)
                run.font.name = t["font"]
                run.font.color.rgb = _rgb(t["text"])

    prs.save(path)

def _bullet_lines(bullets) -> list:
    """Пункты -> [(текст, уровень)]: уровень 0 — пункт, 1 — подпункт."""
    lines = []
    for b in _plain_bullets(bullets):
        lines.append((b["text"], 0))
        for sub in b["sub"]:
            lines.append((sub, 1))
    return lines


def _fill_bullets_ph(frame, bullets, size):
    """Буллеты в плейсхолдере: подпункты через level=1 (стиль шаблона)."""
    lines = _bullet_lines(bullets)
    if not lines:
        return
    for i, (text, level) in enumerate(lines):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = text
        if level:
            para.level = 1


def _fill_bullets_box(body, bullets, size):
    """Буллеты в обычном текстбоксе: подпункты — тире с отступом, мельче."""
    lines = _bullet_lines(bullets)
    if not lines:
        return
    body.text_frame.word_wrap = True
    for i, (text, level) in enumerate(lines):
        para = body.text_frame.paragraphs[0] if i == 0 else body.text_frame.add_paragraph()
        para.text = ("      –  " if level else "•  ") + text
        for run in para.runs:
            run.font.size = Pt(int(size * (0.8 if level else 1)))


def _add_frame(slide, prs, x, y, w, h, accent):
    """Тонкая рамка акцентного цвета (без заливки) — «рамка» по просьбе."""
    from pptx.enum.shapes import MSO_SHAPE
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 int(x), int(y), int(w), int(h))
    box.fill.background()
    box.line.color.rgb = _rgb(accent)
    box.line.width = Pt(1.5)
    box.shadow.inherit = False
    return box


def _add_title_underline(slide, title_shape, accent):
    """Короткая акцентная линейка под заголовком контентного слайда."""
    from pptx.enum.shapes import MSO_SHAPE
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 int(title_shape.left),
                                 int(title_shape.top + title_shape.height),
                                 Inches(1.3), Inches(0.045))
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb(accent)
    bar.line.fill.background()
    bar.shadow.inherit = False


def _add_stats_row(slide, prs, stats, accent):
    """Ряд крупных цифр с подписями — «уместная графика» для данных."""
    sw, sh = prs.slide_width, prs.slide_height
    n = max(len(stats), 1)
    gap = int(sw * 0.02)
    w = int((sw * 0.88 - gap * (n - 1)) / n)
    x = int(sw * 0.05)
    y = int(sh * 0.74)
    h = int(sh * 0.18)
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
        cap = slide.shapes.add_textbox(x + int(sw * 0.014), y + int(h * 0.55),
                                       w - int(sw * 0.02), int(h * 0.45))
        cap.text_frame.word_wrap = True
        cpara = cap.text_frame.paragraphs[0]
        crun = cpara.add_run()
        crun.text = str(st.get("caption", ""))
        crun.font.size = Pt(11)
        crun.font.color.rgb = _rgb("2B2B2B")
        x += w + gap


def _content_extras(slide, prs, sd, accent, body_shape, base_size):
    """Рамка, линейка под заголовком, блоки цифр + уменьшение тела под них."""
    if sd.get("frame"):
        _add_frame(slide, prs,
                   int(prs.slide_width * 0.025), int(prs.slide_height * 0.025),
                   int(prs.slide_width * 0.95), int(prs.slide_height * 0.95),
                   accent)
    if body_shape is not None and body_shape.is_placeholder is False or True:
        pass  # линейку ставим по заголовку ниже
    title_shape = None
    for ph in slide.placeholders:
        if ph.placeholder_format.type in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE):
            title_shape = ph
    if title_shape is None:
        # в freeform заголовок — первый текстбокс
        for sp in slide.shapes:
            if sp.has_text_frame and sp.text_frame.text.strip():
                title_shape = sp
                break
    if title_shape is not None and sd.get("underline", True):
        _add_title_underline(slide, title_shape, accent)
    stats = sd.get("stats") or []
    if stats and body_shape is not None:
        body_shape.height = max(int(prs.slide_height * 0.10),
                                int(prs.slide_height * 0.72 - body_shape.top))
        _add_stats_row(slide, prs, stats, accent)


def _retext_slide(prs, slide, sd: dict, t: dict | None = None) -> None:

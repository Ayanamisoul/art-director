"""Build editable slides from the type and layout of a PDF reference."""
from collections import Counter
from dataclasses import dataclass, replace
from functools import lru_cache
from io import BytesIO
from math import ceil
from pathlib import Path
import os
import re

import fitz
from PIL import Image, ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Pt
from slide_layout import add_fitted_picture, fit_body, fitting_slide, SlideFitError, text_height
from template_palette import readable, tint


@dataclass(frozen=True)
class TextStyle:
    name: str
    size: float  # PDF page units
    color: int
    bold: bool = False
    italic: bool = False


_FONTS = [
    ("Arial", "arial.ttf", False), ("Arial", "arialbd.ttf", True),
    ("Segoe UI", "segoeui.ttf", False), ("Segoe UI Semibold", "seguisb.ttf", False),
    ("Calibri", "calibri.ttf", False), ("Calibri", "calibrib.ttf", True),
    ("Verdana", "verdana.ttf", False), ("Verdana", "verdanab.ttf", True),
]


@lru_cache(maxsize=None)
def _local_font(filename):
    try:
        return ImageFont.truetype(str(Path(os.environ.get("WINDIR", "")) / "Fonts" / filename), 100)
    except OSError:
        return None


def _style(span, preview, zoom):
    name = re.sub(r"^[A-Z]{6}\+", "", span["font"])
    bold = bool(span["flags"] & 16) or bool(re.search("Bold|Heavy|Black", name, re.I))
    italic = bool(span["flags"] & 2) or bool(re.search("Italic|Oblique", name, re.I))
    if not name or name.startswith("Type3"):
        # Outlined Figma fonts carry no family metadata. Match the measured run
        # width against installed Cyrillic-capable faces, preserving editability.
        text = span["text"].replace("\u2028", "").replace("\n", "")
        rect = fitz.Rect(span["bbox"])
        size = max(span["size"], 1)
        target = rect.width / size
        density = None
        if preview and span["color"] == 0:
            sample = preview.crop(tuple(round(v*zoom) for v in rect))
            dark = sum(sample.histogram()[:100])
            if 0 < dark < sample.width*sample.height*.68:
                density = dark/((size*zoom)**2)
        candidates = []
        for family, file, weight in _FONTS:
            font = _local_font(file)
            if font is None:
                continue
            width = font.getlength(text)/100
            error = abs(width-target)/max(target, 1)
            if density is not None:
                mask = font.getmask(text)
                ink = (sum(bytes(mask))/255)/(100**2)
                error += .4*abs(ink-density)/max(density, .1)
            candidates.append((error, family, weight))
        _, name, bold = min(candidates, default=(0, "Arial", False))
    else:
        name = re.sub(r"(?:[-, ]?(?:BoldItalic|BoldOblique|Italic|Oblique|Bold|Regular|Roman))?(?:MT|PSMT)?$", "", name)
        name = {"Times": "Times New Roman", "TimesNewRomanPS": "Times New Roman",
                "Helvetica": "Arial", "Courier": "Courier New"}.get(name, name) or "Arial"
    return TextStyle(name, span["size"], span["color"], bold, italic)


@dataclass
class TextBlock:
    box: fitz.Rect
    style: TextStyle
    align: PP_ALIGN


def _blocks(page):
    blocks = []
    preview = None
    zoom = 1000/page.rect.width
    flags = fitz.TEXTFLAGS_DICT & ~fitz.TEXT_PRESERVE_IMAGES
    for block in page.get_text("dict", flags=flags)["blocks"]:
        lines = [line for line in block.get("lines", []) if abs(line["dir"][1]) < .1]
        spans = [s for line in lines for s in line["spans"] if s["text"].strip()]
        if not spans:
            continue
        box = fitz.Rect(block["bbox"])
        span = max(spans, key=lambda s: (s["size"], len(s["text"])))
        if preview is None and (span["font"].startswith("Type3") or not span["font"]):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csGRAY, alpha=False)
            preview = Image.frombytes("L", (pix.width, pix.height), pix.samples)
        center = (box.x0 + box.x1) / 2
        align = PP_ALIGN.CENTER if abs(center - page.rect.width / 2) < page.rect.width * .025 else PP_ALIGN.LEFT
        if len(lines) > 1:
            ends = [line["bbox"][2] for line in lines]
            starts = [line["bbox"][0] for line in lines]
            if max(ends) - min(ends) < span["size"] * .3 < max(starts) - min(starts):
                align = PP_ALIGN.RIGHT
        blocks.append(TextBlock(box, _style(span, preview, zoom), align))
    return blocks


def _rect_path(path):
    items = path["items"]
    if len(items) == 1 and items[0][0] == "re":
        return True
    return 3 <= len(items) <= 4 and all(i[0] == "l" and
                                        (abs(i[1].x - i[2].x) < .1 or abs(i[1].y - i[2].y) < .1)
                                        for i in items)


class Layout:
    def __init__(self, page):
        self.page = page
        self.width, self.height = w, h = page.rect.width, page.rect.height
        self.paths = page.get_drawings()
        self.images = page.get_image_info(xrefs=True)
        blocks = _blocks(page)
        if blocks:
            largest = max(b.style.size for b in blocks)
            self.title = min((b for b in blocks if b.style.size >= largest * .87),
                             key=lambda b: (b.box.y0, b.box.x0))
        else:
            # A flattened page may already contain baked-in text. Only sample its
            # color and use a clean layout so donor copy cannot survive as pixels.
            pix = page.get_pixmap(matrix=fitz.Matrix(30 / w, 30 / w), alpha=False)
            color = pix.pixel(0, 0)
            ink = 0xffffff if sum(color[:3]) < 330 else 0x141414
            self.title = TextBlock(fitz.Rect(.065*w, .14*h, .91*w, .32*h),
                                   TextStyle("Arial", .052*w, ink, True), PP_ALIGN.LEFT)
        self.others = [b for b in blocks if b is not self.title]
        self.margin = min(max(self.title.box.x0, w*.035), w*.10)
        self.vertical = self.horizontal = None
        self.panels = []
        for path in self.paths:
            r = path["rect"]
            rectangular = _rect_path(path)
            if path.get("fill") and rectangular:
                if .22*w < r.x0 < .76*w and r.width > .23*w and r.height >= .96*h:
                    self.vertical = r.x0
                elif .20*h < r.y0 < .70*h and r.width >= .96*w and r.height > .22*h:
                    self.horizontal = r.y0
            rounded = len(path["items"]) <= 12 and sum(i[0] == "c" for i in path["items"]) == 4
            if (rounded or rectangular) and .17*w < r.width < .94*w and .12*h < r.height < .82*h:
                if not any(abs(r.x0 - other.x0) < 3 and abs(r.y0 - other.y0) < 3
                           for other in self.panels):
                    self.panels.append(r)
        self.has_text = bool(blocks)

    def body_box(self):
        w, h, m = self.width, self.height, self.margin
        if self.vertical:
            return fitz.Rect(self.vertical + m*.8, .15*h, w-m, .91*h)
        next_blocks = [b for b in self.others if b.box.y0 > self.title.box.y1 - h*.015]
        y = min((b.box.y0 for b in next_blocks), default=self.title.box.y1 + .06*h)
        if self.horizontal:
            y = max(self.horizontal + .045*h, min(y, self.horizontal+.13*h))
        else:
            y = min(y, .65*h)
        return fitz.Rect(m, y, w-m, .93*h)

    def body_style(self, area):
        # Use the type of actual body copy within this layout, omitting tiny axes,
        # footnotes and very large display figures.
        candidates = [b for b in self.others if area.intersects(b.box) and
                      .30*self.title.style.size <= b.style.size <= .85*self.title.style.size]
        if not candidates:
            return replace(self.title.style, size=self.title.style.size*.46, bold=False)
        sizes = Counter(round(b.style.size, 1) for b in candidates)
        size = sizes.most_common(1)[0][0]
        return next(b.style for b in candidates if round(b.style.size, 1) == size)

    def card_row(self, count):
        if count < 2:
            return []
        for sample in self.panels:
            row = sorted((r for r in self.panels if abs(r.y0-sample.y0) < self.height*.025
                          and abs(r.height-sample.height) < self.height*.06), key=lambda r: r.x0)
            if len(row) == count and all(a.x1 <= b.x0 for a, b in zip(row, row[1:])):
                return row
        return []


def _path_key(path):
    rect = path["rect"]
    return (round(rect.width, 1), round(rect.height, 1), len(path["items"]),
            path.get("fill"), path.get("color"))


def _replay(page, path):
    pen = page.new_shape()
    for op, *args in path["items"]:
        if op == "l":
            pen.draw_line(*args)
        elif op == "c":
            pen.draw_bezier(*args)
        elif op == "re":
            pen.draw_rect(args[0])
        elif op == "qu":
            pen.draw_quad(args[0])
    pen.finish(color=path.get("color"), fill=path.get("fill"),
               width=path.get("width") or 0, closePath=path.get("closePath", False),
               even_odd=path.get("even_odd", False),
               fill_opacity=path.get("fill_opacity") if path.get("fill_opacity") is not None else 1,
               stroke_opacity=path.get("stroke_opacity") if path.get("stroke_opacity") is not None else 1)
    pen.commit()


def _background(layout, common_art, common_images, mode):
    """Replay only decorative source layers. Text outlined as vector paths in
    Figma and charts containing old values must never be new-slide backgrounds."""
    w, h = layout.width, layout.height
    source = layout.page
    with fitz.open() as clean:
        page = clean.new_page(width=w, height=h)
        if not layout.has_text:
            pix = source.get_pixmap(matrix=fitz.Matrix(48/w, 48/w), alpha=False)
            color = pix.pixel(0, 0)
            page.draw_rect(page.rect, fill=tuple(v/255 for v in color[:3]), color=None)
        # PDF drawing and image sequence numbers refer to the same display list.
        layers = []
        for path in layout.paths:
            r = path["rect"]
            edge = r.x0 < w*.012 or r.x1 > w*.988 or r.y0 < h*.012 or r.y1 > h*.988
            surface = edge and path.get("fill") and _rect_path(path)
            peripheral = r.y1 < .22*h or r.y0 > .85*h
            small_art = (r.width < .38*w and r.height < .16*h and peripheral and
                         (_path_key(path) in common_art or len(layout.paths) < 8))
            if small_art:
                small_art = not any(r.intersects(b.box) for b in [layout.title, *layout.others])
            if surface or small_art:
                layers.append((path["seqno"], "path", path))
        for img in layout.images:
            r = fitz.Rect(img["bbox"])
            logo = r.width < .35*w and r.height < .18*h and (r.y1 < .24*h or r.y0 > .77*h)
            # Flat page screenshots can contain old copy even in the keep mode.
            recurring = img["digest"] in common_images
            keep = logo and recurring if mode in ("remove", "all") else mode == "keep" or recurring or logo
            if img["xref"] and (layout.has_text or logo) and keep:
                layers.append((img["number"], "image", img))
        for _, kind, layer in sorted(layers, key=lambda item: item[0]):
            if kind == "path":
                _replay(page, layer)
            else:
                # The page and its image resources share a document only in the
                # source. Transfer the image, including any transparency mask.
                xref = layer["xref"]
                raw = source.parent.extract_image(xref)
                if raw.get("smask"):
                    base = fitz.Pixmap(source.parent, xref)
                    mask = fitz.Pixmap(source.parent, raw["smask"])
                    base = fitz.Pixmap(base, 0) if base.alpha else base
                    pixels = fitz.Pixmap(base, mask).tobytes("png")
                else:
                    pixels = raw["image"]
                page.insert_image(fitz.Rect(layer["bbox"]), stream=pixels, keep_proportion=False)
        return page.get_pixmap(matrix=fitz.Matrix(1600/w, 1600/w), alpha=False).tobytes("png")


class Canvas:
    def __init__(self, slide, prs, layout, background):
        self.slide = slide
        self.background = background
        self.shapes = slide.shapes
        self.sx = prs.slide_width / layout.width
        self.sy = prs.slide_height / layout.height

    def bounds(self, r):
        return (round(r.x0*self.sx), round(r.y0*self.sy), max(1, round(r.width*self.sx)),
                max(1, round(r.height*self.sy)))

    def text(self, area, rows, align=PP_ALIGN.LEFT, spacing=.19, *, body_style=None, max_pt=32):
        """Keep display type in the source style; fit body and cards in physical points."""
        if not rows:
            return
        if area.is_empty:
            raise SlideFitError("Для текста не осталось места в выбранной композиции.")
        if body_style:
            ratios = [style.size / body_style.size for _, style, _ in rows]
            minimum = ceil(18 / min(ratios) - .00001)
            indent = max(level * max_pt for _, _, level in rows)
            size = fit_body([(row[0], ratio) for row, ratio in zip(rows, ratios)],
                            area.width * self.sx / 12700 - indent, area.height * self.sy / 12700,
                            low=minimum, high=max_pt, font=body_style.name, spacing=spacing)
            rows = [(text, replace(style, size=size*ratio*12700/self.sx), level)
                    for (text, style, level), ratio in zip(rows, ratios)]
        box = self.shapes.add_textbox(*self.bounds(area))
        frame = box.text_frame
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.NONE
        scale = 1.
        if body_style is None:
            source_pt = rows[0][1].size * self.sx / 12700
            try:
                fitted = fit_body([text for text, _, _ in rows], area.width * self.sx / 12700,
                                  area.height * self.sy / 12700, low=max(1, min(18, int(source_pt))),
                                  high=ceil(source_pt), font=rows[0][1].name, spacing=spacing)
            except SlideFitError as exc:
                raise ValueError(f"Заголовок «{rows[0][0]}» слишком длинный для поля шаблона.") from exc
            scale = min(1., fitted / source_pt)
        for i, (text, style, indent) in enumerate(rows):
            p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            p.alignment = align
            size = max(1, style.size * scale * self.sx / 12700)
            p.line_spacing = 1.
            p.space_after = Pt(size * spacing) if i < len(rows)-1 else Pt(0)
            if indent:
                p._p.get_or_add_pPr().set("marL", str(round(Pt(size)*indent)))
            run = p.add_run()
            run.text = text
            run.font.name, run.font.size = style.name, Pt(size)
            run.font.color.rgb = RGBColor.from_string(f"{style.color:06X}")
            run.font.bold, run.font.italic = style.bold, style.italic

    def image(self, path, area):
        return add_fitted_picture(self.slide, path, *self.bounds(area))


def _title_area(layout, bottom):
    t, w = layout.title, layout.width
    if layout.vertical:
        right = layout.vertical - min(layout.margin*.7, .045*w)
        left = t.box.x0
        if t.align != PP_ALIGN.LEFT or left >= right - .15*w:
            left = min(layout.margin, .055*w)
        return fitz.Rect(left, t.box.y0, right, bottom)
    left, right = t.box.x0, w - layout.margin*.7
    if t.align == PP_ALIGN.CENTER:
        left, right = layout.margin, w-layout.margin
    if t.align == PP_ALIGN.RIGHT:
        left, right = layout.margin, t.box.x1
    return fitz.Rect(left, t.box.y0, max(right, left+.12*w), bottom)


def _rows(items, style):
    result = []
    for item in items:
        result.append((item["text"], style, 0))
        for sub in item["sub"]:
            result.append((sub, replace(style, size=style.size*.9), 1))
    return result


def _columns(area, n, gap):
    width = (area.width - (n-1)*gap) / n
    return [fitz.Rect(area.x0+i*(width+gap), area.y0, area.x0+i*(width+gap)+width, area.y1)
            for i in range(n)]


def _content(canvas, layout, sd, plain_bullets, image):
    w, h = layout.width, layout.height
    body = layout.body_box()
    bullets = plain_bullets(sd.get("bullets"))
    cards, stats = sd.get("cards") or [], sd.get("stats") or []
    style = layout.body_style(body)
    title_pt = layout.title.style.size * canvas.sx / 12700
    max_body = min(32, max(26, round(title_pt * .93)))
    if sd.get("font_size"):
        max_body = min(40, max(24, sd["font_size"]))
    title_bottom = .89*h if layout.vertical else (layout.horizontal or body.y0)-.012*h
    title = _title_area(layout, title_bottom)
    side_note = bool(layout.vertical and stats and bullets and not cards and not image)
    header_note = None
    if (layout.horizontal and not layout.vertical and cards and not image and len(bullets) == 1
            and not bullets[0]["sub"] and len(bullets[0]["text"]) <= 140):
        note = fitz.Rect(.57*w, title.y0, w - layout.margin, title.y1)
        available = title.height * canvas.sy / 12700 * .98
        if (text_height(sd["title"], (.53*w - title.x0) * canvas.sx / 12700, min(title_pt, 24),
                        layout.title.style.name) <= available
                and text_height(bullets[0]["text"], note.width * canvas.sx / 12700, 18, style.name) <= available):
            title.x1 = .53*w
            header_note = note
    if layout.vertical and (image or side_note):
        height_pt = text_height(sd["title"], title.width * canvas.sx / 12700,
                                title_pt, layout.title.style.name)
        title.y1 = min(.64*h, title.y0 + height_pt * 12700 / canvas.sy + .012*h)
    canvas.text(title, [(sd["title"], layout.title.style, 0)], layout.title.align, spacing=0)
    if header_note is not None:
        canvas.text(header_note, _rows(bullets, style), body_style=style, max_pt=min(24, max_body))
        bullets = []
    if sd.get("table"):
        from slide_table import add_table
        # Sample the clean decorative layer, after old text and charts have been removed.
        with Image.open(BytesIO(canvas.background)) as img:
            def surface_of(area):
                box = (area.x0*img.width/w, area.y0*img.height/h, area.x1*img.width/w, area.y1*img.height/h)
                counts = img.crop(box).resize((24, 24)).convert("RGB").getcolors(576)
                return bytes(max(counts)[1]).hex().upper()
            bg, surface = surface_of(body), surface_of(title)
        ink = readable(f"{style.color:06X}", bg)
        if surface == bg:
            surface = tint(bg, ink)
        add_table(canvas.slide, sd["table"], canvas.bounds(body), style.name, ink, surface, bg,
                  readable(f"{layout.title.style.color:06X}", surface, (ink,)))
        return
    if side_note:
        note = fitz.Rect(title.x0, title.y1+.025*h, title.x1, .93*h)
        canvas.text(note, _rows(bullets, style), body_style=style, max_pt=min(24, max_body))
        bullets = []
    if image:
        if layout.vertical:
            canvas.image(image, body)
            body = fitz.Rect(title.x0, title.y1+.025*h, title.x1, .93*h)
        else:
            text, picture = _columns(body, 2, .035*w)
            canvas.image(image, picture)
            body = text
    if sd.get("subtitle"):
        bullets = [{"text": sd["subtitle"], "sub": []}, *bullets]
    parts = []
    if bullets:
        parts.append(("bullets", max(1.35, sum(1 + len(b["sub"])*.7 for b in bullets))))
    if cards:
        parts.append(("cards", max(1.7, max(len(c["text"]) for c in cards)/55)))
    if stats:
        parts.append(("stats", 1.6))
    offset = body.y0
    gap = .025*h
    available = body.height - max(0, len(parts)-1)*gap
    for kind, weight in parts:
        height = available*weight/sum(v for _, v in parts)
        area = fitz.Rect(body.x0, offset, body.x1, offset+height)
        offset += height+gap
        if kind == "bullets":
            canvas.text(area, _rows(bullets, style), spacing=.28, body_style=style, max_pt=max_body)
            continue
        groups = cards if kind == "cards" else stats
        source_panels = layout.card_row(len(groups))
        exact = (len(parts) == 1 and not image and source_panels
                 and all(body.intersects(r) for r in source_panels))
        cells = source_panels if exact else _columns(area, len(groups), .014*w)
        if (kind == "stats" and layout.vertical and not image) or (area.width/len(groups) < .14*w and not exact):
            cells = [fitz.Rect(area.x0, area.y0+i*height/len(groups), area.x1,
                               area.y0+(i+1)*height/len(groups)) for i in range(len(groups))]
        for cell, group in zip(cells, groups):
            if exact and kind == "cards":
                surfaces = [p for p in layout.paths if
                            max(abs(a-b) for a, b in zip(p["rect"], cell)) < 2.5]
                rounded = any(any(item[0] == "c" for item in p["items"]) for p in surfaces)
                form = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
                outline = canvas.shapes.add_shape(form, *canvas.bounds(cell))
                fill = next((p["fill"] for p in surfaces if p.get("fill")), None)
                if fill:
                    outline.fill.solid()
                    outline.fill.fore_color.rgb = RGBColor(*(round(255*c) for c in fill))
                else:
                    outline.fill.background()
                edge = next((p.get("color") or p["fill"] for p in surfaces
                             if p.get("color") or (p.get("fill") and p["fill"] != fill)), None)
                if edge:
                    outline.line.color.rgb = RGBColor(*(round(255*c) for c in edge))
                    outline.line.width = Pt(.5)
                else:
                    outline.line.fill.background()
                outline.shadow.inherit = False
                if rounded:
                    outline.adjustments[0] = .09
                pad = min(.016*w, cell.width*.05)
                cell = cell + (pad, pad, -pad, -pad)
            if kind == "cards":
                label, caption, size = group["title"], group["text"], style.size*1.30
            else:
                label, caption, size = group["value"], group["caption"], style.size*2.1
            rows = [(value, face, 0) for value, face in
                    [(label, replace(style, size=size, bold=True)), (caption, style)] if value]
            canvas.text(cell, rows, body_style=style, max_pt=min(max_body, 25))


def render_pdf_template(prs, slides, template, mode, slide_width, plain_bullets, slide_image):
    from pptx_builder import uses_body
    with fitz.open(template) as doc:
        if not len(doc):
            raise ValueError("Пустой PDF-шаблон")
        pages = list(doc)
        for page in pages:
            if page.rotation:
                page.remove_rotation()
        layouts = [Layout(page) for page in pages]
        first = layouts[0]
        prs.slide_width = slide_width
        prs.slide_height = round(slide_width * first.height / first.width)
        seen_paths, seen_images = Counter(), Counter()
        for layout in layouts:
            seen_paths.update({_path_key(p) for p in layout.paths})
            seen_images.update({i["digest"] for i in layout.images})
        common_art = {key for key, count in seen_paths.items() if count > 1}
        common_images = {key for key, count in seen_images.items() if count > 1}
        middle = layouts[1:-1] or layouts[1:] or layouts
        backgrounds, uses = {}, Counter()
        for index, sd in enumerate(slides):
            kind, image = sd.get("kind", "content"), slide_image(sd)
            if kind == "title":
                layout = first
            elif kind == "closing":
                layout = layouts[-1]
            else:
                pool = middle
                if image:
                    pool = [p for p in middle if p.vertical] or pool
                elif sd.get("cards"):
                    pool = [p for p in middle if p.card_row(len(sd["cards"]))] or pool
                elif sd.get("stats"):
                    pool = [p for p in middle if p.vertical] or pool
                elif sd.get("table"):
                    pool = [p for p in middle if p.horizontal and not p.vertical] or pool
                layout = min(pool, key=lambda p: uses[p])
                uses[layout] += 1
            if layout not in backgrounds:
                backgrounds[layout] = _background(layout, common_art, common_images, mode)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(BytesIO(backgrounds[layout]), 0, 0,
                                     width=prs.slide_width, height=prs.slide_height)
            canvas = Canvas(slide, prs, layout, backgrounds[layout])
            if uses_body(sd):
                with fitting_slide(index):
                    _content(canvas, layout, sd, plain_bullets, image)
                continue
            h, w = layout.height, layout.width
            title = layout.title
            candidates = [b for b in layout.others if b.box.y0 >= title.box.y1-.015*h
                          and abs(b.box.x0-title.box.x0) < .12*w]
            subtitle = min(candidates, key=lambda b: b.box.y0, default=None)
            text = sd.get("subtitle")
            y = min(subtitle.box.y0 if subtitle else title.box.y1+.07*h, .85*h)
            title_area = _title_area(layout, y-.015*h if text else .93*h)
            canvas.text(title_area, [(sd["title"], title.style, 0)], title.align, spacing=0)
            if text:
                style = subtitle.style if subtitle else replace(title.style, size=title.style.size*.45, bold=False)
                area = fitz.Rect(title_area.x0, y, title_area.x1, .96*h)
                max_pt = max(24, min(32, round(title.style.size * canvas.sx / 12700 * .9)))
                with fitting_slide(index):
                    canvas.text(area, [(text, style, 0)], title.align, body_style=style, max_pt=max_pt)

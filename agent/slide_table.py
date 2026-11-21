"""Editable PowerPoint tables for slide-sized datasets."""
from math import sqrt

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Pt

from slide_layout import fit_body


def checked_table(value):
    if not isinstance(value, dict):
        raise ValueError("Для таблицы нужны headers и rows.")
    headers, rows = value.get("headers"), value.get("rows")
    if not isinstance(headers, list) or not 2 <= len(headers) <= 5 or not all(
            isinstance(h, str) and h.strip() for h in headers):
        raise ValueError("Для таблицы задайте от 2 до 5 названий колонок.")
    if not isinstance(rows, list) or not rows or any(
            not isinstance(r, list) or len(r) != len(headers) or
            any(type(c) not in (str, int, float) for c in r) for r in rows):
        raise ValueError("В каждой строке таблицы нужно значение для каждой колонки.")
    rows = [[str(c) for c in row] for row in rows]
    if sum(len(c) for r in [headers, *rows] for c in r) > 8000:
        raise ValueError("Разделите большую таблицу по темам.")
    return {"headers": headers, "rows": rows}


def add_table(slide, data, bounds, font, ink, surface, background="FFFFFF", header_ink=None):
    x, y, w, h = bounds
    values = [data["headers"], *data["rows"]]
    weights = [sqrt(max(4, min(64, max(map(len, column))))) for column in zip(*values)]
    widths = [round(w * weight / sum(weights)) for weight in weights]
    widths[-1] += w - sum(widths)
    row_h = min(h // len(values), Pt(72))
    size = min(fit_body([value], width / 12700 - 16, row_h / 12700 - 8, high=24, font=font)
               for row in values for width, value in zip(widths, row))
    table = slide.shapes.add_table(len(values), len(widths), x, y, w, row_h * len(values)).table
    table.first_row = table.horz_banding = False
    for col, width in zip(table.columns, widths):
        col.width = width
    for i, row in enumerate(values):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(surface if i == 0 else background)
            cell.margin_left = cell.margin_right = Pt(8)
            cell.margin_top = cell.margin_bottom = Pt(4)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.word_wrap = True
            run = cell.text_frame.paragraphs[0].add_run()
            run.text = value
            run.font.name, run.font.size, run.font.bold = font, Pt(size), i == 0
            run.font.color.rgb = RGBColor.from_string((header_ink or ink) if i == 0 else ink)

"""Измерение текста и размещение изображений в физических координатах слайда."""
from contextlib import contextmanager
from contextvars import ContextVar
from functools import lru_cache
from io import BytesIO
from math import floor
from pathlib import Path

from PIL import Image, ImageFont, ImageOps


class SlideFitError(ValueError):
    index = None


_custom_face = ContextVar("custom_face", default=None)


@contextmanager
def measure_with(path):
    token = _custom_face.set(path)
    try:
        yield
    finally:
        _custom_face.reset(token)


@contextmanager
def fitting_slide(index):
    try:
        yield
    except SlideFitError as exc:
        exc.index = index
        raise


@lru_cache(maxsize=64)
def _face(name):
    if isinstance(name, Path):
        return ImageFont.truetype(str(name), 100)
    # Pillow разбирает имена файлов системных шрифтов в Windows.
    family = name.casefold().replace(" ", "")
    filenames = {"segoeui": "segoeui", "segoeuisemibold": "seguisb",
                 "timesnewroman": "times", "couriernew": "cour"}
    for face in (filenames.get(family, family), "arial", "DejaVuSans"):
        try:
            return ImageFont.truetype(face + ".ttf", 100)
        except OSError:
            continue
    return ImageFont.load_default(size=100)


@lru_cache(maxsize=8192)
def _cached_advance(word, font):
    return _face(font).getlength(word) / 100


def _advance(word, font):
    return _cached_advance(word, _custom_face.get() or font)


def _measure(text, font):
    return [[_advance(word, font) for word in line.split()] for line in text.split("\n")]


def _wrapped_lines(paragraph, available, space):
    lines = len(paragraph)
    for words in paragraph:
        used = 0
        for word in words:
            if word > available:
                return float("inf")
            if used and used + space + word > available:
                lines += 1
                used = word
            else:
                used += (space if used else 0) + word
    return lines


def text_height(text, width_pt, size, font):
    lines = _wrapped_lines(_measure(text, font), .96 * width_pt / size, _advance(" ", font))
    return lines * 1.18 * size


def fit_body(rows, width_pt, height_pt, *, low=18, high=32, font="Arial", spacing=.18):
    """Largest whole-point body size that fits measured words and paragraph spacing.

    A (text, ratio) row shares the baseline size with the other rows. Callers
    subtract shape margins and indent before passing the available dimensions.
    """
    rows = [(row, 1.) if isinstance(row, str) else row for row in rows]
    measured = [(_measure(text, font), ratio) for text, ratio in rows]
    space = _advance(" ", font)
    upper = max(low, floor(high))
    for size in range(upper, low - 1, -1):
        height = 0
        for index, (paragraph, ratio) in enumerate(measured):
            pt = size * ratio
            available = .96 * width_pt / pt
            lines = _wrapped_lines(paragraph, available, space)
            height += pt * (1.18 * lines + (spacing if index < len(rows)-1 else 0))
            if height > height_pt * .98:
                break
        if height <= height_pt * .98:
            return size
    raise SlideFitError(f"Текст не помещается в поле {width_pt:.0f} × {height_pt:.0f} пт "
                        f"при кегле {low} пт. Разделите этот блок или сократите текст.")


def add_fitted_picture(slide, path, x, y, w, h):
    """Вписать изображение целиком; EXIF и пропорции определяются по пикселям."""
    import fitz
    path = Path(path)
    if w <= 0 or h <= 0:
        raise SlideFitError("Для изображения не осталось места на слайде.")
    try:
        if path.suffix.lower() == ".svg":
            with fitz.open(stream=path.read_bytes(), filetype="svg") as doc:
                page = doc[0]
                zoom = 1600 / max(page.rect.width, page.rect.height)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=True)
                stream = BytesIO(pix.tobytes("png"))
                iw, ih = pix.width, pix.height
        else:
            with Image.open(path) as original:
                image = ImageOps.exif_transpose(original)
                iw, ih = image.size
                if original.format in ("JPEG", "PNG") and original.getexif().get(274, 1) == 1:
                    stream = str(path)
                else:
                    stream = BytesIO()
                    image.convert("RGBA").save(stream, format="PNG")
                    stream.seek(0)
    except (OSError, ValueError, RuntimeError) as exc:
        raise ValueError(f"Не удалось прочитать изображение «{path.name}».") from exc
    scale = min(w / iw, h / ih)
    pw, ph = max(1, round(iw * scale)), max(1, round(ih * scale))
    picture = slide.shapes.add_picture(stream, int(x + (w-pw)/2), int(y + (h-ph)/2),
                                       width=pw, height=ph)
    picture.name = path.name
    return picture

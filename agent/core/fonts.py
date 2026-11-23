"""Пользовательские TrueType-шрифты для редактируемых презентаций."""
import hashlib
import re
import struct
import tempfile
from io import BytesIO
from pathlib import Path

from fontTools.ttLib import TTFont
from PIL import ImageFont
from pptx.opc.package import Part
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

from core.artifacts import ROOT

FONT_DIR = ROOT / "data" / "fonts"
MAX_BYTES = 16 * 1024 * 1024


def _family(font):
    return font["name"].getDebugName(1)


def accept_font(data):
    if not data or len(data) > MAX_BYTES:
        raise ValueError("Выберите файл шрифта размером до 16 МБ.")
    try:
        with TTFont(BytesIO(data)) as font:
            if font.flavor or "glyf" not in font or "fvar" in font:
                raise ValueError("Выберите статический шрифт TrueType (.ttf).")
            flags = font["OS/2"].fsType
            if flags & 0x200 or (flags & 0xE and not flags & 8):
                raise ValueError("Лицензия файла ограничивает встраивание шрифта в редактируемую презентацию.")
            family = _family(font)
            if not family or not font.getBestCmap():
                raise ValueError("В шрифте отсутствует название семейства или таблица символов.")
            ImageFont.truetype(BytesIO(data), 24)
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Не удалось прочитать файл. Выберите шрифт TrueType (.ttf).") from exc
    key = hashlib.sha256(data).hexdigest()
    FONT_DIR.mkdir(parents=True, exist_ok=True)
    path = FONT_DIR / (key + ".ttf")
    if not path.exists():
        with tempfile.TemporaryDirectory(dir=FONT_DIR) as folder:
            tmp = Path(folder) / "font.ttf"
            tmp.write_bytes(data)
            tmp.replace(path)
    return key, family


def font_path(key):
    if isinstance(key, str) and re.fullmatch("[0-9a-f]{64}", key):
        path = FONT_DIR / (key + ".ttf")
        if path.is_file():
            return path
    raise ValueError("Загрузите файл пользовательского шрифта ещё раз.")


def _eot(font, data):
    """Полный несжатый EOT 2.1 для части ppt/fonts/*.fntdata в PowerPoint."""
    os2 = font["OS/2"]
    buf = bytearray(struct.pack(
        "<4I10sBBIHH4I2II4I", 0, len(data), 0x20001, 0,
        font.getTableData("OS/2")[32:42], 1, os2.fsSelection & 1, os2.usWeightClass,
        os2.fsType, 0x504C, *(getattr(os2, f"ulUnicodeRange{i}") for i in range(1, 5)),
        getattr(os2, "ulCodePageRange1", 0), getattr(os2, "ulCodePageRange2", 0),
        font["head"].checkSumAdjustment, 0, 0, 0, 0))
    for index in (1, 2, 5, 4):
        name = (font["name"].getDebugName(index) or "").encode("utf-16le")
        buf.extend(struct.pack("<HH", 0, len(name)) + name)
    buf.extend(bytes(4))
    struct.pack_into("<I", buf, 0, len(buf) + len(data))
    return bytes(buf) + data


def embed(prs, path):
    data = path.read_bytes()
    with TTFont(BytesIO(data)) as font:
        family = _family(font)
        bold = bool(font["head"].macStyle & 1)
        italic = bool(font["head"].macStyle & 2)
        package = prs.part.package
        part = Part(package.next_partname("/ppt/fonts/font%d.fntdata"),
                    "application/x-fontdata", package, _eot(font, data))
    rid = prs.part.relate_to(part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font")
    root = prs.part._element
    root.set("embedTrueTypeFonts", "1")
    group = root.find(qn("p:embeddedFontLst"))
    if group is None:
        group = OxmlElement("p:embeddedFontLst")
        root.insert_element_before(group, "p:custShowLst", "p:photoAlbum", "p:custDataLst",
                                   "p:kinsoku", "p:defaultTextStyle", "p:modifyVerifier", "p:extLst")
    entry = next((e for e in group if e.find(qn("p:font")).get("typeface") == family), None)
    if entry is None:
        entry = OxmlElement("p:embeddedFont")
        group.append(entry)
        face = OxmlElement("p:font")
        face.set("typeface", family)
        entry.append(face)
    variants = ("regular", "bold", "italic", "boldItalic")
    style = int(bold) + 2 * int(italic)
    tag = "p:" + variants[style]
    ref = entry.find(qn(tag))
    if ref is None:
        ref = OxmlElement(tag)
        before = next((entry.find(qn("p:" + variant)) for variant in variants[style+1:]
                       if entry.find(qn("p:" + variant)) is not None), None)
        entry.insert(entry.index(before) if before is not None else len(entry), ref)
    ref.set(qn("r:id"), rid)
    for slide in prs.slides:
        for run in slide._element.xpath(".//a:r | .//a:fld"):
            props = run.find(qn("a:rPr"))
            if props is None:
                props = OxmlElement("a:rPr")
                run.insert(0, props)
            if bold:
                props.set("b", "1")
            if italic:
                props.set("i", "1")
            props.get_or_add_latin().set("typeface", family)
            for script in ("ea", "cs"):
                face = props.find(qn("a:" + script))
                if face is None:
                    face = OxmlElement("a:" + script)
                    after = ("a:cs",) if script == "ea" else ()
                    props.insert_element_before(face, *after, "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:extLst")
                face.set("typeface", family)

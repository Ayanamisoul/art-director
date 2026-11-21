"""Resolve the active PowerPoint master and build readable colors for new objects."""
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT


def _channels(color):
    return tuple(bytes.fromhex(color))


def tint(base, color, fraction=.14):
    return bytes(round(a + (b-a)*fraction) for a, b in zip(_channels(base), _channels(color))).hex().upper()


def _luminance(color):
    values = [v / 255 for v in _channels(color)]
    return sum(w * (v / 12.92 if v <= .04045 else ((v+.055)/1.055)**2.4)
               for w, v in zip((.2126, .7152, .0722), values))


def readable(ink, background, alternatives=()):
    def contrast(color):
        lo, hi = sorted((_luminance(color), _luminance(background)))
        return (hi + .05) / (lo + .05)
    if contrast(ink) >= 4.5:
        return ink
    return next((c for c in alternatives if contrast(c) >= 4.5),
                max(("FFFFFF", "000000"), key=contrast))


def slide_theme(slide, defaults, body=None):
    layout = slide.slide_layout
    master = layout.slide_master
    try:
        theme = parse_xml(master.part.part_related_by(RT.THEME).blob)
    except KeyError:
        return defaults
    palette = {}
    for scheme in theme.xpath(".//a:clrScheme/*"):
        if len(scheme):
            color = scheme[0]
            value = color.get("lastClr") or color.get("val", "")
            if len(value) == 6:
                palette[scheme.tag.split("}")[-1]] = value.upper()
    mapping = {}
    for node in (master, layout, slide):
        for cmap in node._element.xpath("./p:clrMap | ./p:clrMapOvr/a:overrideClrMapping"):
            mapping.update(cmap.attrib)

    def color(node):
        if node.tag == qn("a:srgbClr"):
            return node.get("val")
        key = node.get("val")
        return palette.get(mapping.get(key, key))

    bg = palette.get(mapping.get("bg1", "lt1"), defaults["bg"])
    for node in (master, layout, slide):
        fills = node._element.xpath("./p:cSld/p:bg/p:bgPr/a:solidFill/* | ./p:cSld/p:bg/p:bgRef/*")
        if fills:
            bg = color(fills[0]) or bg
    ink = palette.get(mapping.get("tx1", "dk1"), defaults["text"])
    ink = readable(ink, bg, (palette.get("lt1", "FFFFFF"), palette.get("dk1", "000000")))
    result = {**defaults, "bg": bg, "text": ink,
              "accent": defaults.get("user_accent") or palette.get("accent1", defaults["accent"])}
    result["rounded_cards"] = any(node._element.xpath('.//a:prstGeom[@prst="roundRect"]')
                                  for node in (slide, layout, master))
    for key, scheme in (("font", "minorFont"), ("heading_font", "majorFont")):
        faces = theme.xpath(f".//a:fontScheme/a:{scheme}/a:latin")
        result[key] = defaults.get("user_font") or (faces[0].get("typeface") if faces else None) or defaults["font"]
    fonts = {"+mj-lt": result["heading_font"], "+mn-lt": result["font"]}
    for role, key in (("titleStyle", "heading_font"), ("bodyStyle", "font")):
        properties = master._element.xpath(f"./p:txStyles/p:{role}/a:lvl1pPr/a:defRPr")
        if not properties:
            continue
        latin = properties[0].find(qn("a:latin"))
        if latin is not None and latin.get("typeface"):
            face = latin.get("typeface")
            result[key] = defaults.get("user_font") or fonts.get(face, face)
        if role == "bodyStyle":
            fill = properties[0].xpath("./a:solidFill/*")
            if fill and color(fill[0]):
                result["text"] = readable(color(fill[0]), bg, (ink,))
    if body is not None:
        prototypes = [body]
        if body.is_placeholder:
            parent = layout.placeholders.get(idx=body.placeholder_format.idx)
            if parent is not None:
                prototypes.insert(0, parent)
        for prototype in prototypes:
            props = prototype._element.xpath(
                "./p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr | ./p:txBody/a:p[1]/a:pPr/a:defRPr | "
                "./p:txBody/a:p[1]/a:r[1]/a:rPr")
            for prop in props:
                faces = prop.xpath("./a:latin")
                if faces and faces[0].get("typeface"):
                    face = faces[0].get("typeface")
                    result["font"] = defaults.get("user_font") or fonts.get(face, face)
                fills = prop.xpath("./a:solidFill/*")
                if fills and color(fills[0]):
                    result["text"] = readable(color(fills[0]), bg, (ink,))
        if not defaults.get("user_accent"):
            title = slide.shapes.title
            if title is not None:
                colors = title._element.xpath("./p:txBody/a:p//a:solidFill/*")
                local_ink = color(colors[0]) if colors else None
                if local_ink and max(_channels(local_ink)) - min(_channels(local_ink)) > 32:
                    result["accent"] = local_ink
    return result

"""Разбиение слайдов по смысловым границам до подбора кегля."""
import re
import textwrap

from pptx_builder import MAX_BULLETS, MAX_BULLET_LEN, MAX_SUB, MAX_SUB_LEN, MAX_TOTAL_LINES, uses_body


def _phrases(text, limit):
    if len(text) <= limit:
        return [text] if text else []
    pieces = []
    for sentence in re.split(r"(?<=[.!?…])\s+|\n+", text):
        pieces.extend(textwrap.wrap(sentence, limit, break_long_words=False, break_on_hyphens=False))
    return pieces


def _items(bullets):
    for b in bullets:
        text = b["text"] if isinstance(b, dict) else b
        parts = _phrases(text, MAX_BULLET_LEN)
        subs = [s for sub in b.get("sub", []) for s in _phrases(sub, MAX_SUB_LEN)] if isinstance(b, dict) else []
        if subs:
            yield from parts[:-1]
            for i in range(0, len(subs), MAX_SUB):
                yield {"text": parts[-1], "sub": subs[i:i+MAX_SUB]}
        else:
            yield from parts


def _lines(item):
    return [item["text"], *item.get("sub", [])] if isinstance(item, dict) else [item]


def _continued(slide):
    title = slide["title"]
    return {**slide, "title": title if title.endswith(" (продолжение)") else title + " (продолжение)",
            "subtitle": "", "image": ""}


def paginate_slide(slide):
    if "source_slide" in slide or not uses_body(slide):
        return [slide]
    role = slide.get("kind", "content")
    if role in ("title", "closing", "section"):
        content = {**slide, "kind": "content"}
        if role == "title":
            cover = {**slide, "bullets": [], "stats": [], "cards": [], "image": ""}
            cover.pop("table", None)
            return [cover, *paginate_slide({**content, "subtitle": ""})]
        pages = paginate_slide(content)
        pages[-1 if role == "closing" else 0]["kind"] = role
        return pages
    if slide.get("subtitle"):
        slide = {**slide, "bullets": [slide["subtitle"], *slide.get("bullets", [])], "subtitle": ""}
    bullets, cards, stats = (slide.get(key) or [] for key in ("bullets", "cards", "stats"))
    if slide.get("table"):
        table = slide["table"]
        other = dict(slide)
        del other["table"]
        result = paginate_slide(other) if any((bullets, cards, stats, slide.get("image"))) else []
        for start in range(0, len(table["rows"]), 6):
            page = _continued(slide) if result else slide
            result.append({**page, "bullets": [], "cards": [], "stats": [], "image": "",
                           "table": {**table, "rows": table["rows"][start:start+6]}})
        return result
    length = sum(len(line) for b in bullets for line in _lines(b))
    length += sum(len(value) for c in [*cards, *stats] for value in c.values())
    if length > 8000:
        raise ValueError(f"«{slide['title']}»: в одном блоке {length} символов. "
                         "Выделите основные мысли и разделите исходник по темам.")
    # Карточки с длинной аргументацией получают свои страницы.
    if cards and (len(cards) not in (2, 3, 4) or any(len(c["text"]) > 200 or len(c["title"]) > 60 for c in cards)):
        result = []
        if bullets or stats:
            result.extend(paginate_slide({**slide, "cards": []}))
        for card in cards:
            page = {**slide, "title": card["title"] or slide["title"], "bullets": [card["text"] or card["title"]],
                    "subtitle": "", "cards": [], "stats": []}
            if result:
                page.pop("image", None)
            result.extend(paginate_slide(page))
        return result
    if len(stats) > 3:
        result = []
        for i in range(0, len(stats), 3):
            page = slide if i == 0 else {**_continued(slide), "bullets": [], "cards": []}
            result.extend(paginate_slide({**page, "stats": stats[i:i+3]}))
        return result

    pages, group = [], []
    chars = lines = 0
    for item in _items(bullets):
        values = _lines(item)
        size = sum(map(len, values))
        # Редакционный предел, дополнительно проверяемый по физическим размерам поля.
        budget = 340 if not pages and slide.get("image") else 540
        repeated_heading = (group and isinstance(group[-1], dict) and isinstance(item, dict)
                            and item["text"] == group[-1]["text"])
        if group and (repeated_heading or len(group) >= MAX_BULLETS or lines + len(values) > MAX_TOTAL_LINES
                      or chars + size > budget):
            pages.append(group)
            group, chars, lines = [], 0, 0
        group.append(item)
        chars += size
        lines += len(values)
    pages.append(group)

    result = []
    for i, group in enumerate(pages):
        page = {**(slide if i == 0 else _continued(slide)), "bullets": group}
        if i < len(pages)-1:
            page.update(cards=[], stats=[])
        result.append(page)
    return result


def split_for_space(slide):
    """Дополнительный разрез при переполнении конкретного поля выбранного шаблона."""
    if "source_slide" in slide or not uses_body(slide):
        return []
    if slide.get("kind") in ("section", "closing"):
        pages = split_for_space({**slide, "kind": "content"})
        if pages:
            pages[-1 if slide["kind"] == "closing" else 0]["kind"] = slide["kind"]
        return pages
    first, more = dict(slide), _continued(slide)
    items, cards, stats = (slide.get(key) or [] for key in ("bullets", "cards", "stats"))
    if slide.get("table"):
        table = slide["table"]
        if len(table["rows"]) < 2:
            return []
        middle = (len(table["rows"]) + 1) // 2
        first["table"] = {**table, "rows": table["rows"][:middle]}
        more["table"] = {**table, "rows": table["rows"][middle:]}
    elif stats and (items or cards):
        first["stats"] = []
        more.update(bullets=[], cards=[])
    elif cards and items:
        first["cards"] = []
        more["bullets"] = []
    elif len(items) > 1:
        middle = (len(items) + 1) // 2
        first["bullets"], more["bullets"] = items[:middle], items[middle:]
    elif len(items) == 1 and isinstance(items[0], dict) and items[0].get("sub"):
        heading, details = items[0]["text"], items[0]["sub"]
        if len(details) > 1:
            middle = (len(details) + 1) // 2
            first["bullets"] = [{"text": heading, "sub": details[:middle]}]
            more["bullets"] = [{"text": heading, "sub": details[middle:]}]
        elif slide.get("image"):
            first["bullets"] = []
        elif len(title := slide["title"].removesuffix(" (продолжение)") + " — " + heading) <= 110:
            return [{**slide, "title": title, "bullets": details}]
        else:
            return []
    elif cards:
        if len(cards) == 4:
            first["cards"], more["cards"] = cards[:2], cards[2:]
        else:
            return [{**(first if i == 0 else more), "cards": [], "title": c["title"] or slide["title"],
                     "bullets": _phrases(c["text"] or c["title"], MAX_BULLET_LEN)}
                    for i, c in enumerate(cards)]
    elif len(stats) > 1:
        first["stats"], more["stats"] = stats[:1], stats[1:]
    elif slide.get("image") and (items or stats):
        first.update(bullets=[], stats=[])
    else:
        return []
    return [first, more]

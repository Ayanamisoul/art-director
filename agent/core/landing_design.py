"""Чтение правил оформления из HTML-примера: очистка, токены, CSS.

Спека: docs/superpowers/specs/2026-11-23-landing-examples-design.md.
Только stdlib — модуль используется и веб-приложением, и MCP-сервером.
"""
import re
from collections import Counter
from html.parser import HTMLParser

DEFAULTS = {
    "ground": "#ffffff",
    "ink": "#1a1a1a",
    "font": "Arial, sans-serif",
    "heading_font": "",
    "h1_size": "48px",
    "h1_weight": "600",
    "accent": "#2563a8",
    "button_bg": "",
    "button_color": "",
    "button_radius": "12px",
    "header_bg": "",
    "footer_bg": "",
    "width": "1220px",
    "radius": "12px",
}

_HTML_MARKERS = ("<html", "<body", "<style")
_SCRIPT_RE = re.compile(r"<(script|template)\b[^>]*>[\s\S]*?</\1\s*>", re.IGNORECASE)
_SCRIPT_OPEN_RE = re.compile(r"<(script|template)\b[^>]*/?>", re.IGNORECASE)
_HANDLER_RE = re.compile(r"""\son[a-z]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)""", re.IGNORECASE)
_DATA_URL_RE = re.compile(r"""(data:[-a-z0-9.+/]+;base64,)[a-z0-9+/=]{16,}""", re.IGNORECASE)
_DATA_PLAIN_RE = re.compile(r"""(data:[-a-z0-9.+/]+),[^)"'\s>]{16,}""", re.IGNORECASE)


def sanitize_example(data: bytes) -> str:
    """Очистить HTML-пример: скрипты, обработчики, payloads data:-URL.

    Пример — только источник правил оформления, но храним и показываем
    модели уже очищенную копию.
    """
    text = data.decode("utf-8", errors="replace")
    if not any(marker in text.lower() for marker in _HTML_MARKERS):
        raise ValueError("Файл не похож на HTML-страницу: нет тегов html, body или style.")
    text = _SCRIPT_RE.sub("", text)
    text = _SCRIPT_OPEN_RE.sub("", text)
    text = _HANDLER_RE.sub("", text)
    text = _DATA_URL_RE.sub(r"\1…", text)
    text = _DATA_PLAIN_RE.sub(r"\1…", text)
    return text


# --- цвета: разбор и контраст (по образцу template_palette, без pptx) ---

_COLOR_RE = re.compile(
    r"""(?:#[0-9a-f]{3}(?:[0-9a-f]{3})?\b|rgba?\(\s*\d+[\s,]+\d+[\s,]+\d+[^)]*\))""",
    re.IGNORECASE)


def _channels(color: str) -> tuple[int, int, int]:
    value = color.strip()
    if value.startswith("#"):
        hexa = value[1:]
        if len(hexa) == 3:
            hexa = "".join(ch * 2 for ch in hexa)
        return tuple(int(hexa[i:i + 2], 16) for i in (0, 2, 4))
    numbers = re.findall(r"\d+", value)
    return tuple(int(n) for n in numbers[:3])


def _hex(color: str) -> str:
    return "#{:02X}{:02X}{:02X}".format(*_channels(color))


def _luminance(color: str) -> float:
    def channel(value: int) -> float:
        ratio = value / 255
        return ratio / 12.92 if ratio <= .04045 else ((ratio + .055) / 1.055) ** 2.4
    r, g, b = (_channels(_hex(color)) if not color.startswith("#") else _channels(color))
    return .2126 * channel(r) + .7152 * channel(g) + .0722 * channel(b)


def contrast_ratio(first: str, second: str) -> float:
    lo, hi = sorted((_luminance(first), _luminance(second)))
    return (hi + .05) / (lo + .05)


def readable(ink: str, background: str) -> str:
    """Чернила, читаемые на фоне: иначе чёрный или белый — что контрастнее."""
    if contrast_ratio(ink, background) >= 4.5:
        return _hex(ink)
    return max(("#000000", "#FFFFFF"), key=lambda c: contrast_ratio(c, background))


def _saturation(color: str) -> float:
    r, g, b = (v / 255 for v in _channels(color))
    return (max(r, g, b) - min(r, g, b)) / max(r, g, b, 1e-6)


def _neutral(color: str) -> bool:
    """Почти чёрный/белый/серый — не годится как акцент."""
    return _saturation(color) < .15 or _luminance(color) < .08 or _luminance(color) > .93


# --- разбор CSS примера ---

_RULE_RE = re.compile(r"([^{}]+)\{([^{}]*)\}")
_STYLE_BLOCK_RE = re.compile(r"<style\b[^>]*>([\s\S]*?)</style>", re.IGNORECASE)
_TAG_STYLE_RE = re.compile(r"<(body|header|footer|h1)\b([^>]*)>", re.IGNORECASE)
_DECL_RE = re.compile(r"([-a-z]+)\s*:\s*([^;]+)", re.IGNORECASE)


def _rules(html: str) -> list[tuple[str, dict]]:
    """Правила CSS из <style> и style= ключевых тегов. @media раскрывается
    естественным образом: внутренние правила матчатся тем же regex."""
    found: list[tuple[str, dict]] = []
    css = "\n".join(_STYLE_BLOCK_RE.findall(html))
    for selector, body in _RULE_RE.findall(css):
        found.append((selector.strip().casefold(),
                      {name.strip().casefold(): value.strip()
                       for name, value in _DECL_RE.findall(body)}))
    for tag, attrs in _TAG_STYLE_RE.findall(html):
        match = re.search(r"""style\s*=\s*["']([^"']*)["']""", attrs, re.IGNORECASE)
        if match:
            found.append((tag.strip().casefold(),
                          {name.strip().casefold(): value.strip()
                           for name, value in _DECL_RE.findall(match.group(1))}))
    return found


def _first_px(value: str, default: str) -> str:
    match = re.search(r"(-?\d+(?:\.\d+)?)px", value)
    if not match:
        return default
    number = float(match.group(1))
    return f"{max(2, round(number))}px"


def _find(rules, *needles, prop=()) -> str:
    """Значение первого подходящего свойства: селектор содержит needle."""
    for selector, decls in rules:
        if not any(needle in selector for needle in needles):
            continue
        for name, value in decls.items():
            if prop and name in prop or not prop and name in (
                    "background", "background-color", "color"):
                return value
    return ""


def extract_tokens(html: str) -> dict:
    """Токены оформления примера. Каждый ключ всегда заполнен."""
    rules = _rules(html)
    tokens = dict(DEFAULTS)

    def background_of(*selectors) -> str:
        for selector, decls in rules:
            if any(needle in selector for needle in selectors):
                value = decls.get("background") or decls.get("background-color") or ""
                color = next(iter(_COLOR_RE.findall(value)), "")
                if color and not color.startswith("url"):
                    return _hex(color)
        return ""

    body = background_of("body", "html") or DEFAULTS["ground"]
    tokens["ground"] = body
    ink = ""
    for selector, decls in rules:
        if "body" in selector or "html" in selector:
            if decls.get("color"):
                ink = next(iter(_COLOR_RE.findall(decls["color"])), "")
                break
    tokens["ink"] = readable(_hex(ink), body) if ink else readable(DEFAULTS["ink"], body)

    fonts = _find(rules, "body", "html", prop=("font-family",))
    if fonts:
        tokens["font"] = ", ".join(part.strip() for part in fonts.split(",")[:2])
    tokens["heading_font"] = _find(rules, "h1", "h2", prop=("font-family",)) or tokens["font"]

    for selector, decls in rules:
        if selector in ("h1",) or selector.endswith(" h1"):
            if decls.get("font-size"):
                tokens["h1_size"] = _first_px(decls["font-size"], DEFAULTS["h1_size"])
            if decls.get("font-weight"):
                tokens["h1_weight"] = decls["font-weight"].strip()
            break

    button_bg = background_of("btn", "button") or ""
    button_radius = _find(rules, "btn", "button", prop=("border-radius",))
    if button_radius:
        tokens["button_radius"] = _first_px(button_radius, DEFAULTS["button_radius"])
    tokens["button_bg"] = button_bg or tokens["accent"]
    button_color = ""
    for selector, decls in rules:
        if "btn" in selector or "button" in selector:
            if decls.get("color"):
                button_color = next(iter(_COLOR_RE.findall(decls["color"])), "")
                break
    tokens["button_color"] = (readable(_hex(button_color), tokens["button_bg"])
                              if button_color else readable("#FFFFFF", tokens["button_bg"]))

    tokens["header_bg"] = background_of("header") or DEFAULTS["header_bg"] or body
    tokens["footer_bg"] = background_of("footer") or DEFAULTS["footer_bg"] or body

    width = ""
    for selector, decls in rules:
        if any(part in selector for part in ("container", "wrap", "content", "main", "body")):
            if decls.get("max-width"):
                width = _first_px(decls["max-width"], "")
                if width:
                    break
    tokens["width"] = width or DEFAULTS["width"]

    radii = [value for _, decls in rules
             for value in [decls.get("border-radius")] if value]
    if radii:
        tokens["radius"] = _first_px(Counter(radii).most_common(1)[0][0], DEFAULTS["radius"])

    weights: Counter[str] = Counter()
    for selector, decls in rules:
        boost = 3 if ("btn" in selector or "button" in selector or
                      re.search(r"(^|\s|,)a[\s,.:{[]", selector)) else 1
        for name, value in decls.items():
            if name in ("background", "background-color", "color", "border",
                        "border-color", "border-top-color"):
                for color in _COLOR_RE.findall(value):
                    normalized = _hex(color)
                    if not _neutral(normalized):
                        weights[normalized] += boost
    if button_bg:
        weights[button_bg] += 3
    tokens["accent"] = weights.most_common(1)[0][0] if weights else DEFAULTS["accent"]
    return tokens


def generate_css(tokens: dict) -> str:
    """Параметрический CSS под классы лендинга (нейтральный макет)."""
    accent = tokens["accent"]
    hero_tint = _mix(tokens["ground"], accent, .10)
    section_alt = _mix(tokens["ground"], accent, .05)
    border = _mix(tokens["ground"], tokens["ink"], .16)
    return f""":root {{ --ink: {tokens['ink']}; --ground: {tokens['ground']};
  --font: {tokens['font']}; --heading-font: {tokens['heading_font']}; --width: {tokens['width']}; }}
body {{ font-family: var(--font); }}
header {{ background: {tokens['header_bg']};
  color: {readable(tokens['ink'], tokens['header_bg'])}; }}
header a {{ color: {readable(tokens['ink'], tokens['header_bg'])}; }}
footer {{ background: {tokens['footer_bg']};
  color: {readable(tokens['ink'], tokens['footer_bg'])}; }}
footer a {{ color: {readable(tokens['ink'], tokens['footer_bg'])}; }}
.hero {{ background: linear-gradient(160deg, {hero_tint} 0%, {tokens['ground']} 72%);
  color: {tokens['ink']}; padding-block: clamp(56px, 8vw, 110px); }}
.hero-inner {{ display: grid; grid-template-columns: minmax(0, 1.2fr) minmax(0, .8fr);
  gap: 6%; align-items: center; }}
.hero-media img {{ border-radius: {tokens['radius']}; border: 1px solid {border};
  box-shadow: 0 18px 44px rgba(0,0,0,.16); }}
h1, h2, h3 {{ font-family: var(--heading-font); }}
h1 {{ font-size: clamp(30px, 4.4vw, {tokens['h1_size']}); font-weight: {tokens['h1_weight']}; }}
h1, h2 {{ letter-spacing: -.02em; }}
.button {{ background: {tokens['button_bg']}; color: {tokens['button_color']};
  border-radius: {tokens['button_radius']}; }}
.button:hover {{ filter: brightness(1.06); }}
.chapter-inner {{ display: grid; grid-template-columns: minmax(0, .9fr) minmax(0, 1.1fr);
  gap: 8%; align-items: start; }}
.chapter + .chapter {{ border-top: 1px solid {border}; }}
.chapter:nth-of-type(3n) {{ background: {section_alt}; }}
.items li {{ border-top: 1px solid {border}; padding: 16px 0 16px 22px; position: relative; }}
.items li::before {{ content: ""; width: 7px; height: 7px; background: {accent};
  border-radius: 50%; position: absolute; top: 1.5em; left: 0; }}
a {{ color: {accent}; }}
details div {{ border: 1px solid {border}; border-radius: {tokens['radius']}; }}
@media (max-width: 740px) {{
  .hero-inner {{ grid-template-columns: 1fr; gap: 22px; }}
  .hero-copy {{ text-align: center; }}
  .chapter-inner {{ grid-template-columns: 1fr; gap: 24px; }}
}}"""


def _mix(first: str, second: str, fraction: float) -> str:
    a, b = _channels(first), _channels(second)
    return "#{:02X}{:02X}{:02X}".format(*(round(x + (y - x) * fraction) for x, y in zip(a, b)))


# --- валидация CSS-профиля (белые списки под структуру render_landing) ---

KNOWN_CLASSES = {"wrap", "bar", "footer-inner", "desktop-links", "wordmark", "hero",
                 "hero-inner", "hero-copy", "hero-art", "hero-media", "lead", "button",
                 "chapter", "chapter-inner", "chapter-content", "subsection",
                 "section-text", "items", "mark", "law-title", "sculpture"}
KNOWN_ELEMENTS = {"html", "body", "header", "footer", "main", "nav", "section", "article",
                  "aside", "div", "span", "p", "a", "ul", "ol", "li", "h1", "h2", "h3",
                  "img", "details", "summary", "button", "i", "em", "strong", "b"}
KNOWN_PSEUDOS = {"root", "hover", "active", "focus", "focus-visible", "visited",
                 "first-child", "last-child", "first-of-type", "last-of-type",
                 "nth-child", "nth-of-type", "before", "after", "selection", "marker",
                 "placeholder"}
KNOWN_AT_RULES = {"@media", "@font-face"}
MAX_CSS_SIZE = 48_000
ROOT_ISSUE = "нет правила :root с переменными --ink и --ground"
DEFAULT_ROOT = (":root { --ink: #1A1A1A; --ground: #FFFFFF; --font: Arial, sans-serif;"
                " --width: 1220px; }\n")

_PART_RE = re.compile(r"^(?=[.a-zA-Z*:])[a-zA-Z*0-9-]*(?:\.[a-zA-Z][a-zA-Z0-9_-]*)*"
                      r"(?:::?[a-zA-Z-]+(?:\([^()]*\))?)*$")
_COMBINATOR_SPLIT = re.compile(r"\s*[\s>+~]\s*")
_URL_RE = re.compile(r"url\(([^)]*)\)", re.IGNORECASE)


def _selector_issues(selector: str) -> list[str]:
    issues = []
    for part in selector.split(","):
        part = part.strip()
        if not part:
            continue
        for segment in _COMBINATOR_SPLIT.split(part):
            segment = segment.strip()
            if not segment or not _PART_RE.match(segment):
                issues.append(f"непонятный селектор «{segment}» в «{part.strip()}»")
                continue
            element = re.match(r"[a-zA-Z][a-zA-Z0-9-]*", segment)
            if element and element.group(0) not in KNOWN_ELEMENTS:
                issues.append(f"тег «{element.group(0)}» не входит в структуру лендинга")
            for name in re.findall(r"\.([a-zA-Z][a-zA-Z0-9_-]*)", segment):
                if name not in KNOWN_CLASSES:
                    issues.append(f"класс «.{name}» не входит в структуру лендинга")
            for pseudo in re.findall(r"(?<!:):(?!:)([a-zA-Z-]+)", segment):
                if pseudo not in KNOWN_PSEUDOS:
                    issues.append(f"псевдокласс «:{pseudo}» не разрешён")
    return issues


def _walk_rules(css: str):
    """Итератор (селектор, тело) верхнего уровня; @media — внутрь рекурсивно."""
    index = 0
    while True:
        brace = css.find("{", index)
        if brace == -1:
            break
        selector = css[index:brace].strip()
        depth, cursor = 1, brace + 1
        while cursor < len(css) and depth:
            if css[cursor] == "{":
                depth += 1
            elif css[cursor] == "}":
                depth -= 1
            cursor += 1
        yield selector, css[brace + 1:cursor - 1]
        index = cursor


def validate_profile_css(css: str) -> list[str]:
    """Замечания к CSS-профилю. Пустой список — годен для встраивания."""
    issues: list[str] = []
    css = css.strip()
    if len(css) > MAX_CSS_SIZE:
        return [f"CSS больше {MAX_CSS_SIZE // 1000} КБ"]
    if css.count("{") != css.count("}"):
        issues.append("не сбалансированы фигурные скобки")
    if "<" in css:
        issues.append("в CSS не должно быть HTML-разметки")
    for at_rule in set(re.findall(r"@[a-zA-Z-]+", css)):
        if at_rule not in KNOWN_AT_RULES:
            issues.append(f"at-правило {at_rule} не разрешено (допустимы @media и @font-face)")
    for argument in (match.group(1).strip().strip("'\"") for match in _URL_RE.finditer(css)):
        if not argument.casefold().startswith("data:"):
            issues.append(f"внешний ресурс запрещён: url({argument[:60]})")
    for selector, body in _walk_rules(css):
        lowered = selector.casefold()
        if lowered.startswith("@font-face"):
            continue
        if lowered.startswith("@media"):
            issues.extend(_selector_issues(body.split("{", 1)[0]) if "{" not in body else [])
            continue
        issues.extend(_selector_issues(selector))
    has_root = False
    for selector, body in _walk_rules(css):
        if selector.casefold().strip() == ":root":
            has_root = "--ink" in body and "--ground" in body
    if not has_root:
        issues.append(ROOT_ISSUE)
    return issues


# --- дизайнерский вызов модели ---

class DesignError(ValueError):
    """Модель дважды не смогла дать годный CSS-профиль."""


DESIGNER_SYSTEM = """Ты — веб-дизайнер. Тебе дают правила оформления HTML-примера\
(стили и скелет структуры) и список классов целевого лендинга.
Напиши CSS, переносящий оформление примера: фон страницы, шапку и подвал,
первый экран (hero), типографику и кегли, отступы, кнопки, рамки фотографий,
адаптив под узкий экран.
Структура целевой страницы (менять нельзя, только стилизовать):
header > .wrap.bar > .wordmark + nav(.desktop-links, details); main > section.hero
> .wrap.hero-inner > .hero-copy(h1, p.lead, a.button) + .hero-media(img);
section.chapter > .wrap.chapter-inner > h2 + .chapter-content(p.section-text,
ul.items > li, section.subsection > h3); footer > .wrap.footer-inner.
Обязательные правила: :root с переменными --ink, --ground, --font, --width.
Запрещено: url() с внешними адресами (только data:), @import, @font-face
с внешним src, селекторы вне перечисленных классов и тегов.
Верни только CSS, без пояснений и markdown-рамок."""


def prepare_model_input(html: str, *, style_cap: int = 24_000, skeleton_cap: int = 12_000) -> str:
    """Компактный пакет «куда смотреть»: стили + скелет структуры."""
    styles = "\n".join(part.strip() for part in _STYLE_BLOCK_RE.findall(html) if part.strip())
    if len(styles) > style_cap:
        styles = styles[:style_cap] + "\n/* …стили обрезаны… */"
    outline: list[str] = []

    class Skeleton(HTMLParser):
        depth = 0

        def handle_starttag(self, tag, attrs):
            nonlocal outline
            if tag not in KNOWN_ELEMENTS and tag not in ("details", "summary"):
                return
            if len(outline) >= 120:
                return
            names = dict(attrs)
            line = "  " * min(self.depth, 8) + tag
            if names.get("class"):
                line += "." + ".".join(names["class"].split()[:3])
            if names.get("style"):
                line += " style=" + names["style"][:120]
            outline.append(line)
            if tag not in ("img", "br", "hr", "input", "li", "span", "a", "i", "em",
                           "strong", "b", "summary", "p", "h1", "h2", "h3"):
                self.depth += 1

        def handle_endtag(self, tag):
            if tag not in ("img", "br", "hr", "input", "li", "span", "a", "i", "em",
                           "strong", "b", "summary", "p", "h1", "h2", "h3"):
                self.depth = max(0, self.depth - 1)

    Skeleton().feed(html)
    skeleton = "\n".join(outline)
    if len(skeleton) > skeleton_cap:
        skeleton = skeleton[:skeleton_cap] + "\n…"
    return f"=== СТИЛИ ПРИМЕРА ===\n{styles}\n\n=== СКЕЛЕТ СТРУКТУРЫ ПРИМЕРА ===\n{skeleton}"


def _strip_fences(answer: str) -> str:
    text = answer.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text.strip())
    return text.strip()


def read_design(html_text: str, chat) -> str:
    """Спросить модель о правилах оформления примера.

    chat(system, user) -> str. Одна повторная попытка с перечнем ошибок;
    после второй неудачи — DesignError. Отсутствующий :root дописываем сами.
    """
    payload = prepare_model_input(html_text)
    feedback = ""
    for _ in range(2):
        answer = chat(DESIGNER_SYSTEM, payload + feedback)
        css = _strip_fences(answer)
        issues = validate_profile_css(css)
        if issues == [ROOT_ISSUE]:
            css = DEFAULT_ROOT + css
            issues = validate_profile_css(css)
        if not issues:
            return css
        feedback = ("\n\nПредыдущий CSS не прошёл проверку:\n- "
                    + "\n- ".join(issues[:10]) + "\nИсправь и пришли CSS заново.")
    raise DesignError("Модель не смогла дать годный CSS: " + "; ".join(issues[:5]))


def model_designer():
    """Вызвать настроенную модель проекта без инструментов (один текстовый ход)."""
    import llm
    from config import BACKENDS, DEFAULT_BACKEND

    backend = BACKENDS[DEFAULT_BACKEND]

    def chat(system: str, user: str) -> str:
        return llm.chat(backend["base_url"], backend["api_key"], backend["model"],
                        [{"role": "system", "content": system},
                         {"role": "user", "content": user}])

    return chat

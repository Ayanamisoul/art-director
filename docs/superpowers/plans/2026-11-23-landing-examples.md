# Лендинги из примера — план реализации

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Лендинги берут оформление из загруженного HTML-примера: модель один раз читает его правила и пишет CSS-профиль, код валидирует, кэширует и рендерит страницы в этом стиле.

**Architecture:** Новый модуль `core/landing_design.py` (очистка примера, фолбек-токены, валидация CSS, дизайнерский вызов модели) и `core/landing_examples.py` (реестр `data/landing_examples/`, кэш CSS). `render_landing` разрешает `profile` = имя примера; `artifacts.save_landing` пишет копию CSS рядом с артефактом. Веб: роутер загрузки, селектор «Образец лендинга», «Авто» = свежий пример.

**Tech Stack:** Python 3.14 stdlib (html.parser, re), FastAPI, существующий `llm.chat` (Z.ai/OpenAI-совместимый). Новых зависимостей нет.

**Spec:** `docs/superpowers/specs/2026-11-23-landing-examples-design.md`

## Global Constraints

- Новые core-модули — только stdlib (без PIL/pptx/fastapi импортов).
- Запуск: `py` = системный Python 3.14 с зависимостями; тесты `py -m unittest discover -s tests -v`.
- Тексты UI и сообщения — по-русски, без эмодзи.
- В репозитории есть посторонние незакоммиченные изменения: коммитим только файлы задачи (`git add <пути>`), никогда `git add -A`.
- Тесты изолируют каталоги через `patch.object(module, "DIR", tmp)` — как в `tests/test_template_flow.py`.
- Модель в тестах мокается; реальный вызов — только в ручной проверке (задача 7).

## Review Focus

1. Огромный пример с data:-URL (мегабайты встроенных картинок) — очистка до передачи модели и хранения. Тест в задаче 1 (sanitize).
2. Модель возвращает CSS в markdown-рамке ``` — снимаем рамку до валидации. Тест в задаче 2 (read_design).
3. Пример с именем встроенного профиля (`glavbuh.html`) — встроенный профиль побеждает. Тест в задаче 3.
4. Повторное сохранение лендинга после удаления примера — работает по боковой копии CSS. Тест в задаче 4.
5. Мусорный/не-UTF8 файл вместо HTML — внятная ошибка, не трейсбек. Тест в задаче 3.
6. CSS модели без `:root`-переменных — дописываем дефолты, а не отвергаем. Тест в задаче 2.

---

### Task 1: Очистка примера и фолбек-токены

**Files:**
- Create: `agent/core/landing_design.py`
- Test: `tests/test_landing_design.py`

**Interfaces:**
- Produces: `sanitize_example(data: bytes) -> str`; `extract_tokens(html: str) -> dict` (ключи: `ground, ink, font, heading_font, h1_size, h1_weight, accent, button_bg, button_color, button_radius, header_bg, footer_bg, width, radius` — все str, всегда заполнены дефолтами); `generate_css(tokens: dict) -> str`.

- [ ] **Step 1: Пишем падающие тесты**

```python
EXAMPLE = """<!DOCTYPE html><html><head><style>
body { background: #101214; color: #f5f2ee; font-family: "Lab Grotesque", Arial, sans-serif; max-width: 1200px; }
h1 { font-family: "Lab Grotesque"; font-size: 54px; font-weight: 600; }
.btn { background: #e3b7b2; color: #211b1b; border-radius: 200px; max-width: 354px; }
header { background: #121317; } footer { background: #131418; }
</style></head><body><script>alert(1)</script>
<img src="data:image/png;base64,AAAA...несколько-КБ...==" onerror="hack()">
<p onmouseover="x()">Текст</p></body></html>"""

def test_sanitize_removes_scripts_handlers_and_data_urls(self):
    cleaned = landing_design.sanitize_example(EXAMPLE.encode("utf-8"))
    self.assertNotIn("<script", cleaned); self.assertNotIn("alert(1)", cleaned)
    self.assertNotIn("onerror", cleaned); self.assertNotIn("onmouseover", cleaned)
    self.assertLess(len(cleaned), len(EXAMPLE))       # data:-URL урезан до заглушки
    self.assertIn("Текст", cleaned)                    # контент сохранён

def test_sanitize_rejects_non_html(self):
    with self.assertRaises(ValueError):
        landing_design.sanitize_example("это не html".encode("utf-8"))
    with self.assertRaises(ValueError):
        landing_design.sanitize_example(b"\xff\xd8\xff\xe0 бинарный мусор")

def test_tokens_from_example(self):
    tokens = landing_design.extract_tokens(EXAMPLE)
    self.assertEqual(tokens["ground"].upper(), "#101214")
    self.assertEqual(tokens["ink"].upper(), "#F5F2EE")
    self.assertIn("Lab Grotesque", tokens["font"])
    self.assertEqual(tokens["h1_size"], "54px"); self.assertEqual(tokens["h1_weight"], "600")
    self.assertEqual(tokens["button_bg"].upper(), "#E3B7B2"); self.assertEqual(tokens["button_radius"], "200px")
    self.assertEqual(tokens["header_bg"].upper(), "#121317"); self.assertEqual(tokens["width"], "1200px")
    self.assertEqual(tokens["accent"].upper(), "#E3B7B2")

def test_tokens_defaults_without_styles(self):
    tokens = landing_design.extract_tokens("<html><body><p>Привет</p></body></html>")
    for value in tokens.values():
        self.assertTrue(value)                          # все ключи заполнены

def test_fallback_css_carries_tokens_and_is_valid(self):
    css = landing_design.generate_css(landing_design.extract_tokens(EXAMPLE))
    self.assertIn("#101214", css); self.assertIn("Lab Grotesque", css)
    self.assertEqual(landing_design.validate_profile_css(css), [])   # задача 2 добавит валидатор
```

- [ ] **Step 2: Прогон — убеждаемся, что падают (нет модуля)**

Run: `py -m unittest tests.test_landing_design -v` → ImportError/FAIL.

- [ ] **Step 3: Реализуем `landing_design.py`**

`sanitize_example`: decode utf-8 (errors="replace" не для проверки: если после замены нет `<html`/`<body`/`<style` (casefold) — `ValueError("Файл не похож на HTML-страницу…")`). Затем `html.parser.HTMLParser`: drop `<script>/<template>` целиком, drop атрибуты `on*`, значения `src`/`href` с `data:` урезать до `data:<mime>` (заглушка), остальное пересобрать. `extract_tokens`: regex-разбор правил `selector { … }` из `<style>` и `style=` у `body|header|footer|h1`; селектор-подсказки: кнопка = селектор содержит `btn|button`; контейнер = `container|wrap|content|main`; `accent` = самый частый цвет из кнопок/ссылок/границ, иначе самый частый насыщенный. Цвета: `#rgb`, `#rrggbb`, `rgb()/rgba()` → `#RRGGBB`. Контраст чернил на фоне: локальные `_luminance/_contrast` (по образцу `template_palette.py`, без импорта pptx) — если < 4.5, чернила → контрастная альтернатива. `generate_css`: параметрический шаблон под классы `:root, header, .bar, .hero, h1, .lead, .button, .hero-media img, .chapter, .chapter-inner, .items li, footer` + `@media (max-width: 740px)`; один нейтральный макет (центрированный hero, двухколоночные главы), значения из токенов.

- [ ] **Step 4: Прогон — зелёный, кроме теста валидации (задача 2)**

Run: `py -m unittest tests.test_landing_design -v` → 4 PASS, 1 FAIL (validate_profile_css не существует).

- [ ] **Step 5: Коммит**

```bash
git add agent/core/landing_design.py tests/test_landing_design.py
git commit -m "Лендинги из примера: очистка HTML и фолбек-токены оформления"
```

---

### Task 2: Валидация CSS и дизайнерский вызов модели

**Files:**
- Modify: `agent/core/landing_design.py`
- Test: `tests/test_landing_design.py`

**Interfaces:**
- Consumes: токены/очистку из задачи 1.
- Produces: `validate_profile_css(css: str) -> list[str]` (пустой список = годен); `prepare_model_input(html: str) -> str`; `read_design(html_text: str, chat: Callable[[str, str], str]) -> str` — поднимает `DesignError(ValueError)` после двух неудач; `model_designer() -> Callable[[str, str], str]` — привязка `llm.chat` к `config.BACKENDS[config.DEFAULT_BACKEND]`.

- [ ] **Step 1: Падающие тесты**

```python
GOOD_CSS = (":root { --ink: #111; --ground: #fff; }\n"
            ".hero { background: #f0f0f0; }\n@media (max-width: 740px) { .hero { padding: 20px; } }\n")

def test_validator_accepts_good_css(self):
    self.assertEqual(landing_design.validate_profile_css(GOOD_CSS), [])

def test_validator_rejects_danger(self):
    for bad, why in (("a { color: red }", "нет :root с --ink/--ground"),
                     ("@import url('x.css');", "импорт"),
                     (".a { background: url(https://evil/x.png); }", "внешний url"),
                     (".a { background: url(data:image/png;base64,AAAA); } @font-face { font-family: X; src: url(https://f.woff2); }", "внешний шрифт"),
                     (".a { color: red;", "скобки"),
                     (".weird-unknown { color: red; } :root { --ink: #111; --ground: #fff; }", "чужой селектор"),
                     ("#" * 50000, "лимит 48 КБ")):
        self.assertTrue(landing_design.validate_profile_css(bad), why)

def test_validator_allows_data_font_face_and_known_pseudos(self):
    css = (":root { --ink: #111; --ground: #fff; } @font-face { font-family: F; src: url(data:font/woff2;base64,AA); }"
           " .button:hover, .items li::before { color: red; } .chapter:nth-of-type(3n) { background: #eee; }")
    self.assertEqual(landing_design.validate_profile_css(css), [])

def test_read_design_retries_then_raises_and_strips_fences(self):
    bad, fenced = ".a {", "```css\n" + GOOD_CSS + "\n```"
    chat = iter_calls([bad, fenced])                    # мок: 1-й вызов — мусор, 2-й — CSS в рамке
    css = landing_design.read_design(GOOD_HTML, chat)
    self.assertIn(":root", css); self.assertNotIn("```", css)
    with self.assertRaises(landing_design.DesignError):
        landing_design.read_design(GOOD_HTML, iter_calls([bad, bad]))

def test_read_design_prepends_missing_root(self):
    chat = iter_calls([".hero { background: #f0f0f0; }"])   # без :root
    css = landing_design.read_design(GOOD_HTML, chat)
    self.assertIn("--ink:", css); self.assertIn("--ground:", css)

def test_prepare_model_input_is_bounded_and_structured(self):
    payload = landing_design.prepare_model_input(BIG_HTML)   # BIG_HTML: 30KB стилей + скелет
    self.assertLess(len(payload), 60000)
    self.assertIn("<style>", payload); self.assertIn("СЕКЦИИ/СКЕЛЕТ", payload.upper())
```

- [ ] **Step 2: Прогон — новые тесты падают**

Run: `py -m unittest tests.test_landing_design -v`.

- [ ] **Step 3: Реализация**

`validate_profile_css`: баланс `{}`; отклонять `@import`, `url(` с аргументом не `data:`, `@font-face` без `data:`-src, `<`, `javascript:`, `expression(`; объём ≤ 48_000 символов; at-rules — только `@media` и `@font-face`; селекторы — только из белого списка: теги (`html body header footer main nav section article aside div span p a ul ol li h1 h2 h3 img details summary button i em strong b`), `*`, `:root`, классы из набора `KNOWN_CLASSES = {wrap bar footer-inner desktop-links wordmark hero hero-inner hero-copy hero-art hero-media lead button chapter chapter-inner chapter-content subsection section-text items mark law-title}` (значения из `common.css`/`render_landing`), псевдо `:hover :active :focus :focus-visible :visited :first-child :last-child ::before ::after :nth-child(…) :nth-of-type(…)`, комбинаторы потомка/`>`/`+`; `@media` — рекурсивно к внутренним правилам. Разбор — разбиение по скобкам, не полный парсер. Ошибки — по-русски, с именем селектора/причины. `prepare_model_input`: все `<style>` целиком (до 24К суммарно) + скелет тела (HTMLParser: структурные теги с классами и их `style=`, до 120 элементов, до 12К) + подписи разделов. `read_design`: системный текст — «ты дизайнер: вот правила оформления примера, вот скелет классов (список KNOWN_CLASSES и краткая структура render_landing), перенеси фон, шапку/подвал, hero, типографику, отступы, кнопки, рамки фото; пиши только CSS; без url()/@import/внешних @font-face; обязателен :root с --ink/--ground»; ответ: снять ```-рамки → валидация → при провале один повтор с перечнем ошибок → `DesignError`. Если `:root` без `--ink/--ground` — дописать дефолты и принять. `model_designer()`: замыкание над `config.BACKENDS[config.DEFAULT_BACKEND]`, вызывающее `llm.chat(base_url, api_key, model, [{"role":"system",...},{"role":"user",...}])`.

- [ ] **Step 4: Прогон — всё зелёное**

Run: `py -m unittest tests.test_landing_design -v` → PASS.

- [ ] **Step 5: Коммит**

```bash
git add agent/core/landing_design.py tests/test_landing_design.py
git commit -m "Лендинги из примера: валидация CSS-профиля и вызов модели-дизайнера"
```

---

### Task 3: Реестр примеров

**Files:**
- Create: `agent/core/landing_examples.py`
- Test: `tests/test_landing_examples.py`

**Interfaces:**
- Consumes: `landing_design.sanitize_example`, `read_design`, `generate_css`, `extract_tokens`, `DesignError`, `model_designer`.
- Produces: `EXAMPLES_DIR: Path`; `MAX_SIZE = 8 * 1024 * 1024`; `list_names() -> list[str]` (имена с `.html`); `newest() -> str` (`""` если пусто); `resolve(profile: str) -> str | None` — каноническое имя примера (`"имя"` или `"имя.html"` → `"имя.html"`; `None` если не пример); `accept_upload_bytes(name: str, data: bytes, designer=None) -> dict` = `{"name", "replaced", "css_source": "model"|"fallback", "notes": [str]}`; `profile_css_for(profile: str, fallback: Path | None = None) -> str | None`; `is_available_profile(profile: str) -> bool` (встроенные `PROFILES` из `core.landings` + примеры).

- [ ] **Step 1: Падающие тесты** (изоляция: `patch.object(landing_examples, "EXAMPLES_DIR", tmp)`)

```python
def test_upload_saves_cleaned_html_and_generated_css(self):
    result = landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML_BYTES, designer=fixed_designer(GOOD_CSS))
    self.assertEqual(result["css_source"], "model")
    self.assertTrue((tmp / "vp.html").read_text(encoding="utf-8").count("<script") == 0)
    self.assertIn(":root", (tmp / "vp.css").read_text(encoding="utf-8"))
    self.assertEqual(landing_examples.list_names(), ["vp.html"])
    self.assertEqual(landing_examples.resolve("vp"), "vp.html")
    self.assertEqual(landing_examples.resolve("vp.html"), "vp.html")

def test_upload_falls_back_without_model(self):
    result = landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML_BYTES)  # designer=None
    self.assertEqual(result["css_source"], "fallback")
    self.assertTrue((tmp / "vp.css").exists()); self.assertTrue(result["notes"])

def test_upload_rejects_garbage_and_oversize(self):
    with self.assertRaises(ValueError): landing_examples.accept_upload_bytes("x.html", b"not html")
    with patch.object(landing_examples, "MAX_SIZE", 100), self.assertRaises(ValueError):
        landing_examples.accept_upload_bytes("big.html", b"<html>" + b" " * 200)

def test_builtin_profile_wins_over_same_name_example(self):
    landing_examples.accept_upload_bytes("glavbuh.html", EXAMPLE_HTML_BYTES, designer=fixed_designer(GOOD_CSS))
    self.assertIsNone(landing_examples.resolve("glavbuh"))
    self.assertTrue(landing_examples.is_available_profile("glavbuh"))
    self.assertTrue(landing_examples.is_available_profile("glavbuh.html") is False)  # канонизация: это встроенное имя без .html
```

Примечание к последнему: `resolve` возвращает не-None только для примеров, чьё каноническое имя не совпадает с ключом `PROFILES` (`glavbuh.html` → каноническое `glavbuh.html`, но профиль без расширения `glavbuh` — встроенный; правило: пример `X.html` недоступен как профиль, если `X` ∈ PROFILES).

```python
def test_reupload_regenerates_css(self):
    calls = []
    landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML_BYTES, designer=recording_designer(calls, "a {} :root{--ink:#111;--ground:#fff}"))
    landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML_BYTES, designer=recording_designer(calls, "b {} :root{--ink:#111;--ground:#fff}"))
    self.assertEqual(len(calls), 2)
    self.assertIn("b {", (tmp / "vp.css").read_text(encoding="utf-8"))

def test_profile_css_for_cache_and_fallback_copy(self):
    landing_examples.accept_upload_bytes("vp.html", EXAMPLE_HTML_BYTES, designer=fixed_designer(GOOD_CSS))
    css = landing_examples.profile_css_for("vp.html")
    self.assertIn(":root", css)
    sidecar = tmp / "saved.profile.css"; sidecar.write_text(GOOD_CSS, encoding="utf-8")
    (tmp / "vp.html").unlink(); (tmp / "vp.css").unlink()
    self.assertIn(":root", landing_examples.profile_css_for("vp.html", fallback=sidecar))
    self.assertIsNone(landing_examples.profile_css_for("vp.html"))
```

- [ ] **Step 2: Прогон — падают (нет модуля)**

- [ ] **Step 3: Реализация**

По образцу `core/templates.py`. `accept_upload_bytes`: `sanitize_example` (ValueError наверх), лимит `MAX_SIZE`, атомарная запись через `tempfile` в `EXAMPLES_DIR`; затем генерация CSS: `designer or model_designer()` → `read_design(cleaned, designer)`; при `DesignError`/исключении сети — `generate_css(extract_tokens(cleaned))`, `css_source="fallback"`, заметка «Стиль упрощён: модель недоступна/ошибка»; провал валидации и фолбека — CSS гарантированно валиден (шаблон). `<имя>.css` пишется рядом. Кэш `profile_css_for`: `{(имя, mtime_html, mtime_css): text}`; порядок: `EXAMPLES_DIR/<имя>.css` → `fallback`. `is_available_profile`: `profile in PROFILES or resolve(profile) is not None`.

- [ ] **Step 4: Прогон — зелёный**

Run: `py -m unittest tests.test_landing_examples -v`.

- [ ] **Step 5: Коммит**

```bash
git add agent/core/landing_examples.py tests/test_landing_examples.py
git commit -m "Лендинги из примера: реестр примеров и кэш CSS-профиля"
```

---

### Task 4: Рендер лендинга в стиле примера

**Files:**
- Modify: `agent/core/landings.py` (render_landing, `_visual`)
- Modify: `agent/core/artifacts.py` (`save_landing`)
- Test: `tests/test_landing_renderer.py` (дополнить)

**Interfaces:**
- Consumes: `landing_examples.profile_css_for`, `resolve`, `list_names`; `PROFILES`, `render_landing`, `read_landing`.
- Produces: `render_landing(..., example_css: str = "")` — новый kwarg; в `save_landing` боковая копия `data/artifacts/<имя>.profile.css`.

- [ ] **Step 1: Падающие тесты**

```python
def test_render_landing_with_example_profile(self):
    with tmp_examples():                                   # patch EXAMPLES_DIR + загрузка vp.html c фиксированным CSS
        html = render_landing("Тема", "Лид", SECTIONS, "vp.html")
        self.assertIn("--ink: #111", html); self.assertEqual(html, render_landing("Тема", "Лид", SECTIONS, "vp.html"))
        doc = Document(html); self.assertEqual(doc.attributes("html", "data-profile"), ["vp.html"])
        self.assertFalse(doc.attributes("script", "src")); self.assertFalse(doc.attributes("link", "href"))

def test_render_landing_error_lists_options(self):
    with self.assertRaises(ValueError) as ctx:
        render_landing("Т", "", SECTIONS, "нет такого")
    self.assertIn("glavbuh", str(ctx.exception))

def test_example_hero_without_image_is_clean(self):       # никакой «скульптуры» у примеров
    with tmp_examples():
        html = render_landing("Т", "", SECTIONS, "vp.html")
        self.assertNotIn("sculpture", html)

def test_save_landing_roundtrip_after_example_deleted(self):
    with tmp_examples(), tmp_artifacts():                 # patch artifacts.DATA_DIR
        save_landing("page", "Тема", "Лид", SECTIONS, "vp.html")
        self.assertTrue((artifacts_dir / "page.profile.css").exists())
        (examples_dir / "vp.html").unlink(); (examples_dir / "vp.css").unlink()
        save_landing("page", "Тема-2", "Лид", SECTIONS, "")   # profile="" → прежний из брифа
        brief = read_landing("page"); self.assertEqual(brief["profile"], "vp.html")
        html = (artifacts_dir / "page.html").read_text(encoding="utf-8"); self.assertIn("Тема-2", html)
```

- [ ] **Step 2: Прогон — падают**

- [ ] **Step 3: Реализация**

`render_landing`: если `profile in PROFILES` — как раньше; иначе `resolve(profile)`; для примера: `css = example_css or landing_examples.profile_css_for(каноническое имя)`; нет CSS → `ValueError("Неизвестный стиль лендинга. Доступны glavbuh, school, law и примеры: …")`. `_visual`: для примера — только `hero-media` при наличии `image`, иначе пусто. `artifacts.save_landing`: перед рендером `css = landing_examples.profile_css_for(каноническое, fallback=DATA_DIR / f"{safe_name(name)}.profile.css")`; если источником был каталог примеров — скопировать в боковую копию; `render_landing(..., example_css=css or "")`.

- [ ] **Step 4: Прогон всей suites**

Run: `py -m unittest discover -s tests -v` → всё зелёное (встроенные профили не изменились).

- [ ] **Step 5: Коммит**

```bash
git add agent/core/landings.py agent/core/artifacts.py tests/test_landing_renderer.py
git commit -m "Лендинги: рендер в стиле загруженного примера и боковая копия CSS"
```

---

### Task 5: Веб-интеграция и MCP

**Files:**
- Create: `agent/routers/landing_examples.py`
- Modify: `agent/web.py`, `agent/prompts.py`, `agent/mcp_servers/artifacts.py`
- Test: ядро покрыто задачами 1–4; здесь — ручные проверки curl.

**Interfaces:**
- Consumes: `landing_examples.accept_upload_bytes/list_names/newest/is_available_profile`, `landing_design.model_designer`.
- Produces: `POST /api/landing/examples` (multipart `file`) → `{"ok", "name", "replaced", "css_source", "notes"}`; `GET /api/landing/profiles` → `{"profiles", "examples"}`.

- [ ] **Step 1: Роутер** — копия структуры `routers/templates.py`: POST читает `UploadFile`, зовёт `accept_upload_bytes` через `asyncio.to_thread` с `model_designer()`; `ValueError` → 400. Подключить в `web.py` (`app.include_router`).

- [ ] **Step 2: web.py** — `/api/landing/profiles` добавить `"examples": landing_examples.list_names()`. В `_chat_turn` проверку `profile not in PROFILES` заменить на `not landing_examples.is_available_profile(profile)`. «Авто»: после бренд-автодетекта, если `profile` пуст и `fmt != "presentation"` и `landing_examples.newest()` → `profile = newest()`. В `_system_prompt`: при выбранном примере строка «Образец лендинга: <имя>. Оформление примера применяется при сборке.»

- [ ] **Step 3: prompts.py** — в LANDING_DESIGNER: «Стиль — glavbuh/school/law либо имя загруженного примера лендинга (список виден в контексте чата). Для правки оставляй прежний profile.»

- [ ] **Step 4: MCP** — docstring `save_landing` в `mcp_servers/artifacts.py`: упомянуть примеры; сообщение `render_landing` об ошибке уже перечисляет имена (задача 4).

- [ ] **Step 5: Ручная проверка API**

```powershell
# перезапустить сервер, затем:
curl.exe -F "file=@data/artifacts/lab22_school.html" http://127.0.0.1:8000/api/landing/examples
curl.exe http://127.0.0.1:8000/api/landing/profiles   # пример в списке
# чат с output_format=landing и landing_profile=lab22_school.html → в saved html есть data-profile="lab22_school.html"
```

- [ ] **Step 6: Коммит**

```bash
git add agent/routers/landing_examples.py agent/web.py agent/prompts.py agent/mcp_servers/artifacts.py
git commit -m "Лендинги из примера: загрузка через API, выбор в чате и MCP-описание"
```

---

### Task 6: Интерфейс «Чаты и лендинги»

**Files:**
- Modify: `agent/web/index.html`, `agent/web/static/app.js`

**Interfaces:**
- Consumes: `GET /api/landing/profiles` (`examples`), `POST /api/landing/examples`.

- [ ] **Step 1: index.html** — кнопка загрузки: подпись «Загрузить шаблон PPTX, PDF или пример лендинга (HTML)», input `accept=".pptx,.pdf,.html"`.

- [ ] **Step 2: app.js** — в обработчике загрузки файлы `.html` отправлять на `/api/landing/examples`; после ответа обновлять селектор `#landing-profile` (имена примеров группой после встроенных, подпись — имя файла) и показать статус («стиль прочитан моделью» / «стиль упрощён» по `css_source`); заполнение селектора — из поля `examples` ответа `/api/landing/profiles`.

- [ ] **Step 3: Ручная проверка в браузере** — загрузить пример, выбрать его в селекторе, создать лендинг, открыть HTML.

- [ ] **Step 4: Коммит**

```bash
git add agent/web/index.html agent/web/static/app.js
git commit -m "Лендинги из примера: загрузка и выбор образца в интерфейсе"
```

---

### Task 7: Документация и полная проверка

**Files:**
- Modify: `README.md`, `test_mcp.py`

- [ ] **Step 1: README** — раздел «Лендинги»: пример как образец (загрузка, что переносится, ограничения: только стили внутри файла; фолбек без модели); таблица «Где лежат файлы»: `agent/core/landing_design.py`, `agent/core/landing_examples.py`, `data/landing_examples/`; «Проверка установки»: ручной сценарий.

- [ ] **Step 2: test_mcp.py** — после существующих шагов: записать мини-пример в `data/landing_examples/` (с `css`, чтобы без модели), вызвать `save_landing` с `profile` примера, проверить файл.

- [ ] **Step 3: Полная проверка**

```powershell
py -m unittest discover -s tests -v    # все тесты
py test_mcp.py                          # MCP-путь с примером
```

- [ ] **Step 4: Ручной сценарий с реальной моделью** — `curl -F "file=@data/artifacts/lab22_school.html" …/api/landing/examples` → `css_source: "model"`; создать лендинг в чате на другую тему; убедиться: палитра/шрифт School-примера в CSS, узкий экран, `POST /api/landing/{имя}/pdf`.

- [ ] **Step 5: Коммит**

```bash
git add README.md test_mcp.py
git commit -m "Лендинги из примера: документация и MCP-проверка"
```

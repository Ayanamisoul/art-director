"""Реестр примеров лендингов: data/landing_examples/, генерация CSS-профиля.

По образцу core/templates.py. Пример — HTML-страница пользователя; из неё
модель читает правила оформления (core.landing_design), код валидирует и
кэширует результат как CSS-профиль примера.
"""
import tempfile
from pathlib import Path

from . import landing_design

ROOT = Path(__file__).resolve().parents[2]
EXAMPLES_DIR = ROOT / "data" / "landing_examples"
EXAMPLES_DIR.mkdir(parents=True, exist_ok=True)

MAX_SIZE = 8 * 1024 * 1024
_CSS_CACHE: dict[tuple, str] = {}


def _builtins() -> set[str]:
    """Встроенные профили; ленивый импорт — core.landings тянет PIL и сам
    импортирует этот модуль (задача 4), прямой импорт дал бы цикл."""
    from core.landings import PROFILES
    return set(PROFILES)


def list_names() -> list[str]:
    return sorted(path.name for path in EXAMPLES_DIR.iterdir()
                  if path.is_file() and path.suffix.lower() == ".html")


def newest() -> str:
    """Самый свежий пример по времени изменения (для «Авто»)."""
    files = [path for path in EXAMPLES_DIR.iterdir()
             if path.is_file() and path.suffix.lower() == ".html"]
    files.sort(key=lambda path: path.stat().st_mtime)
    return files[-1].name if files else ""


def resolve(profile: str) -> str | None:
    """Каноническое имя примера для значения profile или None.

    Встроенный профиль всегда побеждает: пример glavbuh.html недоступен
    как стиль, потому что glavbuh — встроенный профиль.
    """
    profile = (profile or "").strip()
    if not profile or profile in _builtins():
        return None
    base = profile[:-5] if profile.casefold().endswith(".html") else profile
    if base in _builtins():
        return None
    for candidate in (f"{base}.html", profile):
        path = EXAMPLES_DIR / candidate
        if path.is_file() and path.suffix.lower() == ".html":
            return path.name
    return None


def is_available_profile(profile: str) -> bool:
    """Профиль известен: встроенный или загруженный пример."""
    return (profile or "").strip() in _builtins() or resolve(profile) is not None


def accept_upload_bytes(name: str, data: bytes, designer=None) -> dict:
    """Сохранить пример и сгенерировать его CSS-профиль.

    designer — callable(system, user) -> str; по умолчанию настроенная модель.
    Недоступная модель не ломает загрузку: стиль собирается из токенов.
    """
    name = Path(name or "example.html").name
    if Path(name).suffix.lower() != ".html":
        raise ValueError("Для примера лендинга нужен файл .html")
    if len(data) > MAX_SIZE:
        raise ValueError("Файл больше 8 МБ. Сохраните страницу как «только HTML».")
    cleaned = landing_design.sanitize_example(data)
    path = EXAMPLES_DIR / name
    replaced = path.exists()
    with tempfile.TemporaryDirectory(dir=EXAMPLES_DIR) as folder:
        staged = Path(folder) / name
        staged.write_text(cleaned, encoding="utf-8")
        staged.replace(path)

    notes: list[str] = []
    css_source = "model"
    if designer is None:
        designer = landing_design.model_designer()
    try:
        css = landing_design.read_design(cleaned, designer)
    except Exception as exc:  # модель недоступна или дважды дала негодный CSS
        css = landing_design.generate_css(landing_design.extract_tokens(cleaned))
        css_source = "fallback"
        notes.append("Стиль упрощён: модель недоступна или ответ не прошёл проверку "
                     f"({type(exc).__name__}). Загрузите файл заново, чтобы повторить.")
    with tempfile.TemporaryDirectory(dir=EXAMPLES_DIR) as folder:
        staged = Path(folder) / (path.stem + ".css")
        staged.write_text(css, encoding="utf-8")
        staged.replace(EXAMPLES_DIR / f"{path.stem}.css")
    return {"name": path.name, "replaced": replaced, "css_source": css_source,
            "notes": notes}


def profile_css_for(profile: str, fallback: Path | None = None) -> str | None:
    """CSS-профиль примера: каталог примеров, иначе боковая копия артефакта."""
    name = resolve(profile)
    if name:
        html_path = EXAMPLES_DIR / name
        css_path = EXAMPLES_DIR / f"{Path(name).stem}.css"
        key = (str(EXAMPLES_DIR.resolve()), name,
               html_path.stat().st_mtime_ns if html_path.exists() else 0,
               css_path.stat().st_mtime_ns if css_path.exists() else 0)
        if key in _CSS_CACHE:
            return _CSS_CACHE[key]
        if css_path.is_file():
            css = css_path.read_text(encoding="utf-8")
            _CSS_CACHE.clear()  # в кэше держим один свежий профиль
            _CSS_CACHE[key] = css
            return css
    if fallback is not None and fallback.is_file():
        return fallback.read_text(encoding="utf-8") or None
    return None

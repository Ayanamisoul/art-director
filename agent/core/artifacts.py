"""Сохранённые артефакты (data/artifacts/): HTML-лендинги, PDF, списки."""
import json
import subprocess
import tempfile
from pathlib import Path

from validation import validate_html

ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "data" / "artifacts"
DATA_DIR.mkdir(parents=True, exist_ok=True)

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


def safe_name(name: str) -> str:
    """Только буквы/цифры/дефис/подчёркивание; срезаем .html/.pdf, если модель
    передала расширение вопреки docstring'у."""
    import re
    name = re.sub(r"\.(html?|pdf)$", "", name.strip(), flags=re.IGNORECASE)
    slug = re.sub(r"[^\w\-]", "_", name)
    return slug or "artifact"


def list_files() -> list[str]:
    return sorted(p.name for p in DATA_DIR.iterdir()
                  if p.is_file() and p.suffix != ".json")


def list_text() -> str:
    files = list_files()
    return "Сохранённые файлы:\n" + "\n".join(files) if files else "Пока ничего не сохранено."


def saved_results(messages: list[dict]) -> dict:
    """Возвращаем только результаты подтверждённых сохранений этого набора сообщений."""
    saved = {"presentation": None, "html": None}
    tools = {call["id"]: call["function"]["name"] for message in messages
             for call in message.get("tool_calls") or []}
    for message in messages:
        if message.get("role") != "tool":
            continue
        name = tools.get(message.get("tool_call_id"))
        content = message.get("content") or ""
        if name == "save_presentation" and content.startswith("Презентация сохранена и прошла проверку: "):
            path = content.split(": ", 1)[1].partition(".pptx (")[0]
            saved["presentation"] = Path(path).name
        if name in ("save_landing", "save_html") and content.startswith("Сохранено и прошло проверку: "):
            saved["html"] = Path(content.split(": ", 1)[1]).stem
    return saved


def save_html(name: str, html: str) -> str:
    """Сохранить HTML-лендинг с проверкой качества. Строка-ответ для модели."""
    path = DATA_DIR / f"{safe_name(name)}.html"
    path.write_text(html, encoding="utf-8")
    issues = validate_html(html)
    if issues:
        return (f"Сохранено: {path}. НО есть замечания валидации:\n- "
                + "\n- ".join(issues)
                + "\nИсправь их и сохрани заново тем же именем.")
    return f"Сохранено и прошло проверку: {path}"


_BRIEF_TAG = '<script type="application/json" id="art-director-brief">'


def read_landing(name: str) -> dict | None:
    """Прочитать модель контента из самого HTML, в том числе после переноса файла."""
    path = DATA_DIR / f"{safe_name(name)}.html"
    if not path.is_file():
        return None
    html = path.read_text(encoding="utf-8")
    _, found, tail = html.rpartition(_BRIEF_TAG)
    return json.loads(tail.split("</script>", 1)[0]) if found else None


def save_landing(name: str, title: str, lead: str, sections: list[dict],
                 profile: str = "", brand: str = "", cta: str = "", image: str = "",
                 font_id: str = "") -> str:
    from core.landings import DEFAULT_PROFILE, render_landing

    previous = read_landing(name) or {}
    brief = dict(title=title, lead=lead, sections=sections, brand=brand, cta=cta,
                 image=image, profile=profile or previous.get("profile") or DEFAULT_PROFILE)
    if font_id or previous.get("font_id"):
        brief["font_id"] = font_id or previous["font_id"]
    html = render_landing(**brief)
    # Escape '<' even inside a JSON string so user text cannot close the script element.
    encoded = json.dumps(brief, ensure_ascii=False).replace("<", r"\u003c")
    html = html.replace("</body>", f"{_BRIEF_TAG}{encoded}</script></body>")
    path = DATA_DIR / f"{safe_name(name)}.html"
    with tempfile.TemporaryDirectory(dir=DATA_DIR) as folder:
        staged = Path(folder) / "page.html"
        staged.write_text(html, encoding="utf-8")
        staged.replace(path)
    return f"Сохранено и прошло проверку: {path}"


def html_to_pdf(name: str) -> str:
    """Конвертировать сохранённый HTML в PDF (headless Chrome). Строка для модели."""
    html_path = DATA_DIR / f"{safe_name(name)}.html"
    if not html_path.exists():
        return f"Ошибка: файл {html_path} не найден. Сначала сохрани лендинг."
    pdf_path = html_path.with_suffix(".pdf")
    with tempfile.TemporaryDirectory() as profile, tempfile.TemporaryDirectory(dir=DATA_DIR) as output:
        staged_pdf = Path(output) / "output.pdf"
        result = subprocess.run(
            [CHROME, "--headless=new", "--disable-gpu",
             f"--user-data-dir={profile}", f"--print-to-pdf={staged_pdf}",
             "--no-pdf-header-footer", html_path.as_uri()],
            capture_output=True, timeout=120)
        if result.returncode or not staged_pdf.is_file():
            return f"Ошибка конвертации: {result.stderr.decode('utf-8', errors='replace')[:300]}"
        staged_pdf.replace(pdf_path)
    return f"PDF готов: {pdf_path}"

"""Сохранённые артефакты (data/artifacts/): HTML-лендинги, PDF, списки."""
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


def html_to_pdf(name: str) -> str:
    """Конвертировать сохранённый HTML в PDF (headless Chrome). Строка для модели."""
    html_path = DATA_DIR / f"{safe_name(name)}.html"
    if not html_path.exists():
        return f"Ошибка: файл {html_path} не найден. Сначала вызови save_html."
    pdf_path = html_path.with_suffix(".pdf")
    with tempfile.TemporaryDirectory() as profile:
        result = subprocess.run(
            [CHROME, "--headless=new", "--disable-gpu",
             f"--user-data-dir={profile}", f"--print-to-pdf={pdf_path}",
             "--no-pdf-header-footer", html_path.as_uri()],
            capture_output=True, timeout=120)
    if not pdf_path.exists():
        return f"Ошибка конвертации: {result.stderr.decode('utf-8', errors='replace')[:300]}"
    return f"PDF готов: {pdf_path}"

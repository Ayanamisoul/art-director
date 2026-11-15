"""MCP-сервер «artifacts» — первые инструменты агента.

Инструменты (их увидит LLM через MCP-клиент и сможет вызывать):
  save_html(name, html)   — сохранить сгенерированный HTML на диск пользователя
  html_to_pdf(name)       — сконвертировать сохранённый HTML в PDF (headless Chrome)
  list_artifacts()        — показать, что уже сохранено

Local-first: всё пишется в <репо>/data/artifacts/ на машине пользователя.
Сервер ничего не отправляет наружу и ничего не хранит «у себя».

Запуск (stdio-транспорт, сервер поднимается клиентом на каждую сессию):
  py agent/mcp_servers/artifacts.py
"""
import re
import subprocess
import tempfile
from pathlib import Path

from mcp.server.mcpserver import MCPServer

# <репо>/data/artifacts/ — вне git (см. .gitignore)
DATA_DIR = Path(__file__).resolve().parents[2] / "data" / "artifacts"
DATA_DIR.mkdir(parents=True, exist_ok=True)

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

mcp = MCPServer("artifacts")


def _safe_name(name: str) -> str:
    """Только буквы/цифры/дефис/подчёркивание — защита от путей вида ../.."""
    slug = re.sub(r"[^\w\-]", "_", name.strip())
    return slug or "artifact"


@mcp.tool()
def save_html(name: str, html: str) -> str:
    """Сохранить HTML-страницу на диск. name — короткое имя без расширения,
    html — полный код страницы. Возвращает путь к файлу."""
    path = DATA_DIR / f"{_safe_name(name)}.html"
    path.write_text(html, encoding="utf-8")
    return f"Сохранено: {path}"


@mcp.tool()
def html_to_pdf(name: str) -> str:
    """Конвертировать ранее сохранённый HTML (save_html) в PDF.
    name — то же имя, что передавали в save_html. Возвращает путь к PDF."""
    html_path = DATA_DIR / f"{_safe_name(name)}.html"
    if not html_path.exists():
        return f"Ошибка: файл {html_path} не найден. Сначала вызови save_html."
    pdf_path = html_path.with_suffix(".pdf")
    # Отдельный user-data-dir, чтобы не конфликтовать с открытым Chrome пользователя
    with tempfile.TemporaryDirectory() as profile:
        result = subprocess.run(
            [
                CHROME,
                "--headless=new",
                "--disable-gpu",
                f"--user-data-dir={profile}",
                f"--print-to-pdf={pdf_path}",
                "--no-pdf-header-footer",
                html_path.as_uri(),
            ],
            capture_output=True,
            timeout=120,
        )
    if not pdf_path.exists():
        return f"Ошибка конвертации: {result.stderr.decode('utf-8', errors='replace')[:300]}"
    return f"PDF готов: {pdf_path}"


@mcp.tool()
def list_artifacts() -> str:
    """Показать список сохранённых файлов (HTML и PDF)."""
    files = sorted(p.name for p in DATA_DIR.iterdir() if p.is_file())
    return "Сохранённые файлы:\n" + "\n".join(files) if files else "Пока ничего не сохранено."


if __name__ == "__main__":
    mcp.run()  # stdio: общается с MCP-клиентом через stdin/stdout

# -*- coding: utf-8 -*-
"""Дымовой тест MCP-сервера artifacts: поднимаем сервер по stdio,
спрашиваем список инструментов и прогоняем save_html -> html_to_pdf,
затем save_landing со стилем-примером."""
import asyncio
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

REPO = Path(__file__).resolve().parent
HTML = """<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">
<style>body{font-family:Arial;background:#3b2f1a;color:#fff;text-align:center;padding:80px}</style>
</head><body><h1>Тест MCP</h1><p>Если ты видишь это в PDF — конвейер работает.</p></body></html>"""
EXAMPLE = """<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><style>
body { background: #f0f7ff; color: #102a43; font-family: Verdana, sans-serif; }
.btn { background: #2563eb; color: #ffffff; border-radius: 8px; }
</style></head><body><h1>Пример</h1><p>Стиль-образец для smoke-теста</p></body></html>"""


async def main() -> None:
    params = StdioServerParameters(
        command=sys.executable,  # тот же Python, что запустил тест
        args=[str(REPO / "agent" / "mcp_servers" / "artifacts.py")],
    )
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            tools = await session.list_tools()
            print("ИНСТРУМЕНТЫ:", ", ".join(t.name for t in tools.tools))

            r1 = await session.call_tool("save_html", {"name": "mcp-test", "html": HTML})
            print("save_html ->", r1.content[0].text)

            r2 = await session.call_tool("html_to_pdf", {"name": "mcp-test"})
            print("html_to_pdf ->", r2.content[0].text)

            r3 = await session.call_tool("list_artifacts", {})
            print("list_artifacts ->", r3.content[0].text)

            # лендинг в стиле загруженного примера: кладём файл напрямую
            # и собираем страницу через профиль-пример (без модели)
            examples = REPO / "data" / "landing_examples"
            examples.mkdir(parents=True, exist_ok=True)
            (examples / "mcp_smoke_example.html").write_text(EXAMPLE, encoding="utf-8")
            sys.path.insert(0, str(REPO / "agent"))
            from core import landing_design, landing_examples
            css = landing_design.generate_css(landing_design.extract_tokens(EXAMPLE))
            (examples / "mcp_smoke_example.css").write_text(css, encoding="utf-8")
            r4 = await session.call_tool("save_landing", {
                "name": "mcp-test-landing", "title": "Лендинг по примеру",
                "lead": "Проверка стиля-примера",
                "sections": [{"title": "Раздел", "text": "Текст", "items": ["Пункт"]}],
                "profile": "mcp_smoke_example.html"})
            print("save_landing ->", r4.content[0].text)

    pdf = REPO / "data" / "artifacts" / "mcp-test.pdf"
    print("PDF на диске:", pdf.exists(), pdf.stat().st_size if pdf.exists() else 0, "байт")
    landing = REPO / "data" / "artifacts" / "mcp-test-landing.html"
    print("Лендинг на диске:", landing.exists(), landing.stat().st_size if landing.exists() else 0, "байт")


asyncio.run(main())

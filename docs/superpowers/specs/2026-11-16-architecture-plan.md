# План: реструктуризация + фичи

1. core/: previews.py (перенос), sessions.py, templates.py, assets.py, artifacts.py, decks.py
2. mcp_servers/artifacts.py -> тонкий адаптер над core (те же имена инструментов и докстринги)
3. routers/: sessions.py, templates.py, assets.py, files.py (включая NEW: DELETE asset, POST pdf)
4. web.py -> сборка app + chat-эндпоинт + static mount
5. Фронт: style.css + app.js из index.html; index.html -> shell; новые: PDF-кнопка, удаление картинок, скроллы
6. prompts.py: правило «фото только по-пользовательски»; editor.html: PDF-кнопка
7. Проверки: py_compile, node --check, рестарт, смоук API (шаблоны/сессии/PDF/удаление), чат через MCP-адаптер (zai), скрин UI
8. Коммит + push

"""Семантическая валидация: второй LLM-вызов проверяет, соответствует ли
артефакт исходному запросу пользователя по СМЫСЛУ (тема, состав).

Код-валидация (validation.py) ловит форму, этот модуль — содержание.
Пример, который он должен поймать: просили «презентацию по Counter-Strike»,
а модель сделала слайды про Computer Science.

Принцип fail-open: если судья ответил мусором — не блокируем выдачу,
помечаем ok=True и пишем raw в лог. Ложный «стоп» хуже пропущенной ошибки.
"""
import json
import re

from llm import chat

JUDGE_SYSTEM = """Ты — строгий проверяющий качества. Тебе дают запрос пользователя и результат,
который подготовил другой ИИ. Проверь ТОЛЬКО соответствие смысла:
1) тема результата совпадает с запрошенной;
2) все явно запрошенные элементы присутствуют (количество слайдов, цвета, разделы);
3) нет выдуманных фактов, противоречащих запросу.
Не придирайся к формулировкам и стилю. Ответь СТРОГО одним JSON-объектом без пояснений:
{"ok": true} или {"ok": false, "issues": ["короткое замечание 1", ...]}"""


def _extract_json(raw: str) -> dict | None:
    match = re.search(r"\{[\s\S]*\}", raw)
    if not match:
        return None
    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError:
        return None


def semantic_check(backend: dict, user_request: str, artifact_summary: str) -> dict:
    """Возвращает {"ok": bool, "issues": [str], "raw": str | None}."""
    messages = [
        {"role": "system", "content": JUDGE_SYSTEM},
        {"role": "user", "content":
            f"Запрос пользователя:\n{user_request}\n\nРезультат (структура/содержимое):\n{artifact_summary}"},
    ]
    try:
        raw = chat(backend["base_url"], backend["api_key"], backend["model"], messages)
    except Exception as exc:
        return {"ok": True, "issues": [], "raw": f"судья недоступен: {exc}"}
    verdict = _extract_json(raw)
    if verdict is None or "ok" not in verdict:
        return {"ok": True, "issues": [], "raw": raw}  # fail-open
    return {"ok": bool(verdict["ok"]),
            "issues": [str(i) for i in verdict.get("issues", [])][:5],
            "raw": None}

"""Клиент для API генерации текста с поддержкой вызова инструментов."""
import json
import urllib.request


def complete(base_url: str, api_key: str, model: str, messages: list[dict],
             tools: list[dict] | None = None, timeout: int = 600) -> dict:
    """Один шаг диалога. Возвращает сырое сообщение модели:
    {"role": "assistant", "content": ...} и, если модель решила вызвать
    инструмент, ещё и "tool_calls": [{id, function: {name, arguments}}].
    """
    payload: dict = {"model": model, "messages": messages}
    if tools:
        payload["tools"] = tools
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        base_url.rstrip("/") + "/chat/completions",
        data=body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "art-director-agent/0.1",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = json.loads(response.read().decode("utf-8"))
    return data["choices"][0]["message"]


def chat(base_url: str, api_key: str, model: str, messages: list[dict], timeout: int = 600) -> str:
    """Упрощённый вариант без инструментов: вернуть только текст ответа."""
    return complete(base_url, api_key, model, messages, timeout=timeout)["content"]

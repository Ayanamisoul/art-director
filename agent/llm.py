"""Единый клиент к LLM. Работает с любым OpenAI-совместимым API:
локальная Ollama, шлюз xkiro, а в перспективе — Qwen на сервере Актиона.
"""
import json
import urllib.request


def chat(base_url: str, api_key: str, model: str, messages: list[dict], timeout: int = 600) -> str:
    """Отправляет историю сообщений, возвращает текст ответа модели."""
    body = json.dumps({"model": model, "messages": messages}).encode("utf-8")
    request = urllib.request.Request(
        base_url.rstrip("/") + "/chat/completions",
        data=body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            # без нормального User-Agent xkiro отвечает 403
            "User-Agent": "art-director-agent/0.1",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = json.loads(response.read().decode("utf-8"))
    return data["choices"][0]["message"]["content"]

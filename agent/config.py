"""Конфигурация агента: читает .env и выбирает бэкенд (ollama или xkiro).

Оба бэкенда говорят по OpenAI-совместимому протоколу (POST /v1/chat/completions),
поэтому переключение между ними = смена base_url + api_key + model.
"""
import os
from pathlib import Path

# .env лежит в корне репозитория (уровнем выше папки agent/)
ENV_PATH = Path(__file__).resolve().parent.parent / ".env"


def load_env(path: Path = ENV_PATH) -> None:
    """Простейший парсер .env без сторонних библиотек (формат KEY=VALUE)."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip())


load_env()

BACKENDS = {
    "ollama": {
        "base_url": os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434/v1"),
        "api_key": "ollama",  # Ollama ключ не проверяет, но заголовок нужен по протоколу
        "model": os.environ.get("OLLAMA_MODEL", "qwen2.5:7b-instruct"),
    },
    "xkiro": {
        "base_url": os.environ.get("XKIRO_BASE_URL", "https://xkiro.com/v1"),
        "api_key": os.environ.get("XKIRO_API_KEY", ""),
        "model": os.environ.get("XKIRO_MODEL", "mistralai/mistral-large-2512"),
    },
    # Z.ai (Zhipu) — тот же провайдер, на котором работает ассистент разработчика:
    # OpenAI-совместимый API, у базового URL нет /v1, поэтому /chat/completions
    # дописывается в llm.py к base_url как есть
    "zai": {
        "base_url": os.environ.get("ZAI_BASE_URL", "https://api.z.ai/api/paas/v4"),
        "api_key": os.environ.get("ZAI_API_KEY", ""),
        "model": os.environ.get("ZAI_MODEL", "glm-5.3-flash"),
    },
}

DEFAULT_BACKEND = "ollama"

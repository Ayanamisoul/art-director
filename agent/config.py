"""Конфигурация агента: читает .env и выбирает бэкенд LLM.

Бэкенд один — Z.ai (GLM), тот же OpenAI-совместимый протокол, поэтому если
понадобится другой провайдер — достаточно дополнить словарь BACKENDS.
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

# Z.ai (Zhipu) — OpenAI-совместимый API, у базового URL нет /v1, поэтому
# /chat/completions дописывается в llm.py к base_url как есть
BACKENDS = {
    "zai": {
        "base_url": os.environ.get("ZAI_BASE_URL", "https://api.z.ai/api/paas/v4"),
        "api_key": os.environ.get("ZAI_API_KEY", ""),
        "model": os.environ.get("ZAI_MODEL", "glm-5.3-flash"),
    },
}

DEFAULT_BACKEND = "zai"

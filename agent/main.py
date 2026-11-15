"""CLI-оболочка агента — теперь с настоящим агентным циклом.

Цикл: сообщение -> LLM -> если модель попросила инструмент, выполняем его
через MCP и возвращаем результат модели -> повторяем, пока не получим
обычный текстовый ответ. Максимум MAX_TOOL_ROUNDS вызовов подряд.

Команды в чате:
  /backend ollama|xkiro  — переключить бэкенд
  /skill chat|landing    — переключить системный промпт (скилл)
  /new                   — начать диалог заново
  /exit                  — выйти
"""
import asyncio
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(errors="replace")  # Windows-консоль cp1251 не любит эмодзи в логах
except Exception:
    pass

from config import BACKENDS, DEFAULT_BACKEND
from llm import complete
from mcp_client import MCPToolbox
from prompts import SKILLS

ARTIFACTS_SERVER = Path(__file__).parent / "mcp_servers" / "artifacts.py"
MAX_TOOL_ROUNDS = 5  # защита от бесконечного цикла вызовов


async def agent_reply(backend: dict, messages: list[dict], toolbox: MCPToolbox,
                      log: list[str] | None = None) -> str:
    """Один полный ход агента: модель + при необходимости инструменты.
    В log (если передан) пишутся строки вида «вызван save_html(...)» —
    их веб-клиент показывает пользователю."""
    tools = toolbox.openai_tools()
    for _ in range(MAX_TOOL_ROUNDS):
        message = complete(backend["base_url"], backend["api_key"], backend["model"],
                           messages, tools=tools)
        messages.append(message)  # важно: с tool_calls, чтобы модель видела свои вызовы

        if not message.get("tool_calls"):
            return message.get("content") or "(пустой ответ)"

        for call in message["tool_calls"]:
            name = call["function"]["name"]
            try:
                args = json.loads(call["function"]["arguments"] or "{}")
            except json.JSONDecodeError:
                args = {}
            short_args = {k: (v[:60] + "…" if isinstance(v, str) and len(v) > 60 else v)
                          for k, v in args.items()}
            line = f"{name}({short_args})"
            print(f"  [tool] {line}")
            if log is not None:
                log.append(line)
            try:
                result = await toolbox.call(name, args)
            except Exception as exc:
                result = f"ошибка инструмента: {exc}"
            print(f"  [tool] -> {result[:400]}")
            messages.append({
                "role": "tool",
                "tool_call_id": call["id"],
                "content": result,
            })

    return "слишком много вызовов инструментов подряд — остановился, попробуй переформулировать"


async def main() -> None:
    backend_name = DEFAULT_BACKEND
    skill_name = "chat"
    messages: list[dict] = []

    toolbox = MCPToolbox(ARTIFACTS_SERVER)
    await toolbox.connect()
    tool_names = ", ".join(t.name for t in toolbox.tools)

    print(f"Арт-директор v0.2 | бэкенд: {backend_name} | скилл: {skill_name}")
    print(f"MCP-сервер artifacts подключён, инструменты: {tool_names}")
    print("Команды: /backend, /skill, /new, /exit\n")

    try:
        while True:
            try:
                user_input = (await asyncio.to_thread(input, "вы > ")).strip()
            except (EOFError, KeyboardInterrupt):
                print()
                break
            if not user_input:
                continue

            if user_input.startswith("/"):
                parts = user_input.split()
                command = parts[0]
                if command == "/exit":
                    break
                elif command == "/new":
                    messages = []
                    print("(диалог очищен)\n")
                elif command == "/backend" and len(parts) == 2 and parts[1] in BACKENDS:
                    backend_name = parts[1]
                    print(f"(бэкенд: {backend_name})\n")
                elif command == "/skill" and len(parts) == 2 and parts[1] in SKILLS:
                    skill_name = parts[1]
                    print(f"(скилл: {skill_name})\n")
                else:
                    print("Команды: /backend ollama|xkiro, /skill chat|landing, /new, /exit\n")
                continue

            backend = BACKENDS[backend_name]
            if not messages:
                messages.append({"role": "system", "content": SKILLS[skill_name]})
            messages.append({"role": "user", "content": user_input})

            print("агент думает...")
            try:
                answer = await agent_reply(backend, messages, toolbox)
            except Exception as exc:
                print(f"[ошибка бэкенда: {exc}]\n")
                messages.pop()
                continue
            print("агент >", answer[:800] + ("..." if len(answer) > 800 else ""), "\n")
    finally:
        await toolbox.close()


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))

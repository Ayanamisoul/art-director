"""CLI-оболочка агента — минимальный цикл для проверки связки.

Веб-клиент появится позже; смысл этого файла — доказать, что ядро работает:
сообщения -> LLM -> ответ, с переключаемым бэкендом и скиллами.

Команды в чате:
  /backend ollama|xkiro  — переключить бэкенд
  /skill chat|landing    — переключить системный промпт (скилл)
  /new                   — начать диалог заново
  /exit                  — выйти
"""
import sys

from config import BACKENDS, DEFAULT_BACKEND
from llm import chat
from prompts import SKILLS


def main() -> None:
    backend_name = DEFAULT_BACKEND
    skill_name = "chat"
    messages: list[dict] = []

    print(f"Арт-директор v0.1 | бэкенд: {backend_name} | скилл: {skill_name}")
    print("Команды: /backend, /skill, /new, /exit\n")

    while True:
        try:
            user_input = input("вы > ").strip()
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

        print("агент > ", end="", flush=True)
        try:
            answer = chat(backend["base_url"], backend["api_key"], backend["model"], messages)
        except Exception as exc:  # на этапе скелета показываем ошибку как есть
            print(f"[ошибка бэкенда: {exc}]\n")
            messages.pop()  # не оставляем в истории вопрос без ответа
            continue
        print(answer[:500] + ("..." if len(answer) > 500 else "") + "\n")
        messages.append({"role": "assistant", "content": answer})


if __name__ == "__main__":
    sys.exit(main())

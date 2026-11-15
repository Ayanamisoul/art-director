"""MCP-клиент: поднимает MCP-сервер как дочерний процесс и даёт агенту
доступ к его инструментам в формате, который понимает LLM.

Жизненный цикл: connect() -> ... работаем ... -> close().
Сервер живёт, пока жив toolbox.
"""
import sys
from contextlib import AsyncExitStack
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


class MCPToolbox:
    def __init__(self, server_script: Path):
        self._params = StdioServerParameters(
            command=sys.executable,  # тем же Python, что и агент
            args=[str(server_script)],
        )
        self._stack = AsyncExitStack()
        self._session: ClientSession | None = None
        self.tools: list = []

    async def connect(self) -> None:
        read, write = await self._stack.enter_async_context(stdio_client(self._params))
        self._session = await self._stack.enter_async_context(ClientSession(read, write))
        await self._session.initialize()
        self.tools = (await self._session.list_tools()).tools

    def openai_tools(self) -> list[dict]:
        """Переводит MCP-описания инструментов в формат OpenAI tools,
        который понимает LLM (Qwen, Mistral и др.)."""
        return [
            {
                "type": "function",
                "function": {
                    "name": t.name,
                    "description": t.description or "",
                    "parameters": t.input_schema,
                },
            }
            for t in self.tools
        ]

    async def call(self, name: str, arguments: dict) -> str:
        """Выполнить инструмент на MCP-сервере, вернуть текстовый результат."""
        assert self._session is not None, "сначала connect()"
        result = await self._session.call_tool(name, arguments)
        texts = [c.text for c in result.content if getattr(c, "text", None)]
        return "\n".join(texts) or "(инструмент вернул пустой ответ)"

    async def close(self) -> None:
        await self._stack.aclose()

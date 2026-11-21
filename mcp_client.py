"""Клиент официального SDK для готового локального Excel MCP-сервера по stdio."""

from __future__ import annotations

import asyncio
import os
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any, AsyncIterator

from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


class MCPError(RuntimeError):
    """Ошибка соединения с MCP или вызова его инструмента."""


def jsonable(value: Any) -> Any:
    if hasattr(value, "model_dump"):
        return value.model_dump(mode="json", exclude_none=True)
    if isinstance(value, dict):
        return {str(k): jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [jsonable(v) for v in value]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


class ExcelMCPClient:
    """Запускает внешний сервер с доступом только в проектный каталог data."""

    def __init__(self, base_dir: Path, timeout_seconds: float = 45) -> None:
        self.base_dir = base_dir.resolve()
        self.data_dir = (self.base_dir / "data").resolve()
        self.timeout_seconds = timeout_seconds

    def server_params(self) -> StdioServerParameters:
        local_uvx = self.base_dir / ".venv" / "Scripts" / "uvx.exe"
        command = os.getenv("EXCEL_MCP_COMMAND", "").strip() or (str(local_uvx) if local_uvx.exists() else "uvx")
        safe_env = {
            key: os.environ[key]
            for key in ("PATH", "SYSTEMROOT", "WINDIR", "TEMP", "TMP", "USERPROFILE", "APPDATA", "LOCALAPPDATA")
            if os.environ.get(key)
        }
        return StdioServerParameters(
            command=command,
            args=["--from", os.getenv("EXCEL_MCP_PACKAGE", "excel-mcp-server==2.3.3"),
                  "excel-mcp-server", "stdio", "--allow-dir", str(self.data_dir)],
            # OpenAI secrets and unrelated environment variables are never inherited by the child MCP process.
            env={**safe_env, "PYTHONUTF8": "1"},
        )

    @asynccontextmanager
    async def connected(self) -> AsyncIterator[ClientSession]:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        try:
            async with asyncio.timeout(self.timeout_seconds):
                async with stdio_client(self.server_params(), errlog=None) as (read_stream, write_stream):
                    async with ClientSession(read_stream, write_stream) as session:
                        await session.initialize()
                        yield session
        except TimeoutError as exc:
            raise MCPError(f"Истекло время ожидания Excel MCP ({self.timeout_seconds:g} с).") from exc
        except MCPError:
            raise
        except Exception as exc:
            raise MCPError(f"Excel MCP недоступен: {type(exc).__name__}: {exc}") from exc

    async def list_tools(self) -> list[dict[str, Any]]:
        async with self.connected() as session:
            try:
                result = await session.list_tools()
            except Exception as exc:
                raise MCPError(f"Ошибка tools/list: {type(exc).__name__}: {exc}") from exc
            return [{"name": tool.name, "description": tool.description or "", "inputSchema": jsonable(tool.inputSchema)}
                    for tool in result.tools]

    async def call_tool(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        async with self.connected() as session:
            try:
                result = await session.call_tool(name, arguments=arguments)
            except Exception as exc:
                raise MCPError(f"Ошибка tools/call ({name}): {type(exc).__name__}: {exc}") from exc
            return {"isError": bool(result.isError),
                    "content": [jsonable(item) for item in result.content],
                    "structuredContent": jsonable(getattr(result, "structuredContent", None))}

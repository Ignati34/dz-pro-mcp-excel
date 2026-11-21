from __future__ import annotations

import os
from contextlib import asynccontextmanager
from pathlib import Path
from uuid import UUID

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openai import AsyncOpenAI
from pydantic import BaseModel, Field

from excel_agent import ExcelAgent
from mcp_client import ExcelMCPClient, MCPError

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

api_key = os.getenv("OPENAI_API_KEY", "").strip()
base_url = os.getenv("OPENAI_BASE_URL", "").strip() or "https://api.openai.com/v1"
llm_client = AsyncOpenAI(
    api_key=api_key or "missing-key",
    base_url=base_url,
    timeout=float(os.getenv("OPENAI_TIMEOUT_SECONDS", "45")),
    max_retries=0,
)
excel_mcp = ExcelMCPClient(BASE_DIR, float(os.getenv("EXCEL_MCP_TIMEOUT_SECONDS", "45")))
agent = ExcelAgent(
    mcp=excel_mcp,
    client=llm_client,
    model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
    max_tool_calls=int(os.getenv("MAX_TOOL_CALLS", "5")),
)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    yield
    await llm_client.close()


app = FastAPI(title="AI-помощник по заявкам · Pro", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")


class ChatRequest(BaseModel):
    session_id: UUID
    message: str = Field(min_length=1, max_length=4000)


class SessionRequest(BaseModel):
    session_id: UUID


class ActionRequest(BaseModel):
    session_id: UUID
    action_id: UUID


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.get("/api/health")
async def health() -> dict:
    try:
        tools = await excel_mcp.list_tools()
        names = {tool["name"] for tool in tools}
        required = {"create_workbook", "read_range", "write_range"}
        return {"app": "ok", "mcp": "available", "transport": "stdio", "tool_count": len(tools),
                "compatible": required.issubset(names), "missing": sorted(required - names),
                "tools": tools}
    except MCPError as exc:
        return {"app": "ok", "mcp": "unavailable", "transport": "stdio", "error": str(exc), "tools": []}


@app.get("/api/requests")
async def requests() -> dict:
    traces: list[dict] = []
    try:
        rows = await agent.read_requests(traces)
        return {"requests": rows, "tool_calls": traces, "error": None}
    except Exception as exc:
        return {"requests": [], "tool_calls": traces, "error": str(exc)}


@app.post("/api/chat")
async def chat(request: ChatRequest) -> dict:
    text = request.message.strip()
    if not text:
        raise HTTPException(status_code=422, detail="Введите сообщение.")
    if not api_key:
        return {"reply": "Добавьте OPENAI_API_KEY в локальный .env и перезапустите приложение.",
                "tool_calls": [], "pending": None}
    try:
        return await agent.reply(str(request.session_id), text)
    except Exception:
        return {"reply": "Не удалось обработать запрос. Проверьте доступность OpenAI API и Excel MCP.",
                "tool_calls": [], "pending": None}


@app.post("/api/new-chat")
async def new_chat(request: SessionRequest) -> dict[str, bool]:
    agent.reset(str(request.session_id))
    return {"ok": True}


@app.post("/api/requests/confirm")
async def confirm(request: ActionRequest) -> dict:
    try:
        return await agent.confirm(str(request.session_id), str(request.action_id))
    except Exception as exc:
        return {"ok": False, "message": f"Не удалось подтвердить действие: {exc}", "tool_calls": []}


@app.post("/api/requests/cancel")
async def cancel(request: ActionRequest) -> dict:
    return await agent.cancel(str(request.session_id), str(request.action_id))

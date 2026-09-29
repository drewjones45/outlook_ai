import os
import sys

import anyio

from outlook_ai.mcp_server import build_server

EXPECTED_TOOLS = {
    "inbox_overview", "emails_needing_reply", "read_email", "search_emails", "calendar",
    "meeting_prep", "create_reply_draft", "sync_outlook", "outlook_status",
}


def _text(result) -> str:
    content = getattr(result, "content", result)
    if isinstance(content, tuple):  # SDK 1.x returns (content, structured)
        content = content[0]
    return "\n".join(getattr(block, "text", "") for block in content)


def test_tools_prompts_and_annotations(service):
    server = build_server(service)

    async def go():
        tools = {t.name: t for t in await server.list_tools()}
        prompts = {p.name for p in await server.list_prompts()}
        return tools, prompts

    tools, prompts = anyio.run(go)
    assert set(tools) == EXPECTED_TOOLS
    assert prompts == {"inbox_summary", "draft_replies", "daily_briefing"}
    hints = {name: t.annotations.model_dump(by_alias=True) for name, t in tools.items()}
    writers = {name for name, h in hints.items() if not h.get("readOnlyHint")}
    assert writers == {"create_reply_draft"}
    assert all(h.get("destructiveHint") is False for h in hints.values())


def test_tool_calls_round_trip(service, fake):
    server = build_server(service)

    async def go():
        overview = _text(await server.call_tool("inbox_overview", {"days": 7}))
        draft = _text(await server.call_tool("create_reply_draft", {"message_id": 101, "body": "On it."}))
        prompt = await server.get_prompt("draft_replies", {"days": "3", "max_drafts": "2"})
        return overview, draft, prompt

    overview, draft, prompt = anyio.run(go)
    assert "[101]" in overview
    assert "not sent" in draft and fake.drafts
    assert "create_reply_draft" in prompt.messages[0].content.text


def test_outlook_errors_become_readable_tool_output(service, fake):
    from outlook_ai.osa import OutlookError

    fake.fail["create_reply_draft"] = OutlookError("Not authorized to send Apple events", -1743)
    server = build_server(service)

    async def go():
        await server.call_tool("sync_outlook", {})
        return _text(await server.call_tool("create_reply_draft", {"message_id": 101, "body": "Hi"}))

    text = anyio.run(go)
    assert text.startswith("Outlook error:") and "Automation" in text


def test_stdio_server_starts_and_answers(tmp_path):
    """Spawn `python -m outlook_ai serve` and talk MCP to it over stdio."""
    from mcp import ClientSession
    from mcp.client.stdio import StdioServerParameters, stdio_client

    env = dict(os.environ, OUTLOOK_AI_DATA_DIR=str(tmp_path / "d"), OUTLOOK_AI_CONFIG=str(tmp_path / "none.toml"))
    params = StdioServerParameters(command=sys.executable, args=["-m", "outlook_ai", "serve"], env=env)

    async def go():
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                tools = await session.list_tools()
                status = await session.call_tool("outlook_status", {})
                return {t.name for t in tools.tools}, _text(status)

    names, status = anyio.run(go)
    assert names == EXPECTED_TOOLS
    assert "Config:" in status

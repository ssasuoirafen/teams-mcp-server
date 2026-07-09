"""Registration tests for the Teams MCP server.

The server is a module-level FastMCP singleton (`mcp`) with `@mcp.tool()`
decorators applied at import time. Importing the module registers the tool
closures without calling them; env vars are only read inside `_init()` (via
`main()`), not at import. We set dummy required env vars defensively before the
import so the test stays robust if that ever changes.

This is the executable spec for the tool surface: assert the exact count and
the exact set of tool names.
"""

import os

os.environ.setdefault("TEAMS_MCP_TENANT_ID", "test-tenant")
os.environ.setdefault("TEAMS_MCP_CLIENT_ID", "test-client")

from teams_mcp.server import _format_message, mcp  # noqa: E402

EXPECTED_TOOLS = {
    # auth
    "login",
    "complete_login",
    # read
    "list_teams",
    "list_channels",
    "list_chats",
    "list_channel_messages",
    "list_thread_replies",
    "list_chat_messages",
    "list_team_members",
    "list_channel_members",
    "list_chat_members",
    "list_pinned_messages",
    "get_user_presence",
    "get_user",
    "search_messages",
    "download_attachment",
    # write
    "send_channel_message",
    "send_chat_message",
    "reply_to_channel_message",
    "create_chat",
    "create_group_chat",
    "set_reaction",
    "unset_reaction",
    "pin_message",
    "unpin_message",
    "delete_message",
    "update_message",
    "mark_chat_read",
    "mark_chat_unread",
}


def test_format_message_null_body_content():
    """Graph can send body.content = null (e.g. deleted reply in a thread).

    Present-but-null defeats .get(k, default) defaults; formatting must not
    crash (regression: re.finditer(None) in _format_hosted_contents).
    """
    msg = {
        "id": "1783130166074",
        "createdDateTime": "2026-07-04T10:00:00Z",
        "from": {"user": {"displayName": "Someone"}},
        "body": {"content": None},
    }
    result = _format_message(msg)
    assert result["content"] == ""
    assert "hostedContents" not in result


def test_tool_count():
    names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert len(names) == 29


def test_tool_names():
    names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert names == EXPECTED_TOOLS

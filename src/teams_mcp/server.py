import json
import os
import re
import sys

from mcp.server.fastmcp import FastMCP

from teams_mcp.auth import AuthManager
from teams_mcp.graph import GraphClient

mcp = FastMCP(
    "teams-mcp",
    instructions=(
        "Microsoft Teams MCP server. Call login first to authenticate. "
        "For channels: list_teams -> list_channels -> list_channel_messages or send_channel_message. "
        "For chats: list_chats -> list_chat_messages or send_chat_message."
    ),
)

auth: AuthManager | None = None
graph: GraphClient | None = None


def _init():
    global auth, graph
    tenant_id = os.environ["TEAMS_MCP_TENANT_ID"]
    client_id = os.environ["TEAMS_MCP_CLIENT_ID"]
    auth = AuthManager(tenant_id=tenant_id, client_id=client_id)
    graph = GraphClient(token_provider=auth.get_token)


def _init_if_needed():
    if auth is None:
        _init()


def _require_auth() -> GraphClient:
    if graph is None or not auth.is_authenticated():
        raise RuntimeError("Not authenticated. Call the login tool first.")
    return graph


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "")


def _format_message(msg: dict) -> dict:
    return {
        "id": msg.get("id"),
        "sender": (msg.get("from") or {}).get("user", {}).get("displayName"),
        "createdDateTime": msg.get("createdDateTime"),
        "content": _strip_html((msg.get("body") or {}).get("content", "")),
    }


# Global to hold pending device flow between login and complete_login calls
_pending_flow: dict | None = None


# Tool: login
# Annotations: openWorldHint=True
@mcp.tool()
def login() -> str:
    """Start authentication with Microsoft Teams via device code flow.

    If already authenticated, returns current account info.
    Otherwise, returns a device code and URL. The user must open the URL in a browser
    and enter the code. Then call complete_login to finish authentication."""
    global _pending_flow
    _init_if_needed()
    accounts = auth._app.get_accounts()
    if accounts:
        return json.dumps(
            {"status": "already_authenticated", "account": accounts[0].get("username")},
            ensure_ascii=False,
            indent=2,
        )
    _pending_flow = auth.login()
    return json.dumps(
        {
            "status": "action_required",
            "message": _pending_flow.get("message", ""),
            "user_code": _pending_flow.get("user_code", ""),
            "verification_uri": _pending_flow.get("verification_uri", ""),
            "instructions": "Open the URL, enter the code, then call complete_login.",
        },
        ensure_ascii=False,
        indent=2,
    )


# Tool: complete_login
# Annotations: openWorldHint=True
@mcp.tool()
def complete_login() -> str:
    """Complete the device code authentication after the user has entered the code in the browser.

    Call this AFTER the user has opened the URL from login and entered the device code.
    Blocks until authentication completes (up to 15 minutes)."""
    global _pending_flow
    _init_if_needed()
    if _pending_flow is None:
        return json.dumps({"status": "error", "message": "No pending login. Call login first."})
    flow = _pending_flow
    _pending_flow = None
    result = auth.complete_login(flow)
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: list_teams
# Annotations: readOnlyHint=True, openWorldHint=True
@mcp.tool()
async def list_teams() -> str:
    """List all Microsoft Teams you are a member of.

    Returns team id, name, and description for each team.
    Use a team_id with list_channels to see its channels.
    """
    _init_if_needed()
    client = _require_auth()
    teams = await client.list_teams()
    result = [
        {
            "id": t.get("id"),
            "name": t.get("displayName"),
            "description": t.get("description"),
        }
        for t in teams
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: list_channels
# Annotations: readOnlyHint=True, openWorldHint=True
@mcp.tool()
async def list_channels(team_id: str) -> str:
    """List channels in a Microsoft Teams team.

    Use list_teams first to get the team_id.
    Returns channel id, name, description, and membership type.
    """
    _init_if_needed()
    client = _require_auth()
    channels = await client.list_channels(team_id)
    result = [
        {
            "id": c.get("id"),
            "name": c.get("displayName"),
            "description": c.get("description"),
            "membershipType": c.get("membershipType"),
        }
        for c in channels
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: list_chats
# Annotations: readOnlyHint=True, openWorldHint=True
@mcp.tool()
async def list_chats(limit: int = 20) -> str:
    """List recent chats with participant names.

    Does NOT include channel conversations - use list_teams + list_channels for those.
    Returns chat id, topic, type, and member names.
    """
    _init_if_needed()
    client = _require_auth()
    chats = await client.list_chats(limit=limit)
    result = [
        {
            "id": c.get("id"),
            "topic": c.get("topic"),
            "chatType": c.get("chatType"),
            "lastUpdatedDateTime": c.get("lastUpdatedDateTime"),
        }
        for c in chats
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: list_channel_messages
# Annotations: readOnlyHint=True, openWorldHint=True
@mcp.tool()
async def list_channel_messages(team_id: str, channel_id: str, limit: int = 20) -> str:
    """List recent messages in a Teams channel.

    Use list_teams -> list_channels to get team_id and channel_id.
    Returns message id, sender, timestamp, and plain text content.
    System messages are excluded.
    """
    _init_if_needed()
    client = _require_auth()
    messages = await client.list_channel_messages(team_id, channel_id, limit=limit)
    result = [
        _format_message(m)
        for m in messages
        if m.get("messageType") == "message"
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: list_chat_messages
# Annotations: readOnlyHint=True, openWorldHint=True
@mcp.tool()
async def list_chat_messages(chat_id: str, limit: int = 20) -> str:
    """List recent messages in a chat.

    Use list_chats to get the chat_id.
    Returns message id, sender, timestamp, and plain text content.
    System messages are excluded.
    """
    _init_if_needed()
    client = _require_auth()
    messages = await client.list_chat_messages(chat_id, limit=limit)
    result = [
        _format_message(m)
        for m in messages
        if m.get("messageType") == "message"
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: send_channel_message
# Annotations: openWorldHint=True
@mcp.tool()
async def send_channel_message(team_id: str, channel_id: str, content: str) -> str:
    """Send a message to a Teams channel.

    Use list_teams -> list_channels to get team_id and channel_id.
    For replies to existing messages, use reply_to_channel_message instead.
    """
    _init_if_needed()
    client = _require_auth()
    result = await client.send_channel_message(team_id, channel_id, content)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: send_chat_message
# Annotations: openWorldHint=True
@mcp.tool()
async def send_chat_message(chat_id: str, content: str) -> str:
    """Send a message to a Teams chat.

    Use list_chats to get the chat_id.
    Chat messages don't support threaded replies - just send a new message.
    """
    _init_if_needed()
    client = _require_auth()
    result = await client.send_chat_message(chat_id, content)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: reply_to_channel_message
# Annotations: openWorldHint=True
@mcp.tool()
async def reply_to_channel_message(
    team_id: str, channel_id: str, message_id: str, content: str
) -> str:
    """Reply to a message in a Teams channel thread.

    Use list_channel_messages to get the message_id to reply to.
    For new top-level messages, use send_channel_message instead.
    """
    _init_if_needed()
    client = _require_auth()
    result = await client.reply_to_channel_message(team_id, channel_id, message_id, content)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


def main():
    _init()
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

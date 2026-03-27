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
        "Microsoft Teams MCP server. Call login first to authenticate.\n"
        "Channels: list_teams -> list_channels -> list_channel_messages/send_channel_message.\n"
        "Chats: list_chats -> list_chat_messages/send_chat_message.\n"
        "Members: list_team_members, list_channel_members, list_chat_members.\n"
        "Search: search_messages for full-text search across all chats/channels.\n"
        "Users: get_user to find users, get_user_presence for online status.\n"
        "Reactions: set_reaction/unset_reaction. Pins: pin_message/unpin_message.\n"
        "Message ops: update_message, delete_message.\n"
        "Chat management: create_chat, create_group_chat, mark_chat_read/mark_chat_unread.\n"
        "Send tools support @mentions via the mentions parameter."
    ),
)

auth: AuthManager | None = None
graph: GraphClient | None = None


def _init():
    global auth, graph
    tenant_id = os.environ["TEAMS_MCP_TENANT_ID"]
    client_id = os.environ["TEAMS_MCP_CLIENT_ID"]
    scopes_env = os.environ.get("TEAMS_MCP_SCOPES")
    scopes = [s.strip() for s in scopes_env.split(",") if s.strip()] if scopes_env else None
    auth = AuthManager(tenant_id=tenant_id, client_id=client_id, scopes=scopes)
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


def _extract_element_text(element: dict) -> list[str]:
    """Extract text lines from a single Adaptive Card element."""
    t = element.get("type", "")
    lines: list[str] = []

    if t == "TextBlock":
        text = element.get("text", "")
        if text:
            lines.append(text)

    elif t == "FactSet":
        for fact in element.get("facts", []):
            title = fact.get("title", "")
            value = fact.get("value", "")
            if title or value:
                lines.append(f"{title}: {value}" if title and value else title or value)

    elif t == "RichTextBlock":
        parts = []
        for inline in element.get("inlines", []):
            text = inline.get("text", "")
            if text:
                parts.append(text)
        if parts:
            lines.append("".join(parts))

    elif t in ("Container", "Column", "TableCell"):
        for item in element.get("items", []):
            lines.extend(_extract_element_text(item))

    elif t == "ColumnSet":
        for col in element.get("columns", []):
            lines.extend(_extract_element_text(col))

    elif t == "Table":
        for row in element.get("rows", []):
            for cell in row.get("cells", []):
                lines.extend(_extract_element_text(cell))

    elif t == "ImageSet":
        for img in element.get("images", []):
            alt = img.get("altText", "")
            if alt:
                lines.append(alt)

    elif t == "Image":
        alt = element.get("altText", "")
        if alt:
            lines.append(alt)

    elif t == "ActionSet":
        for action in element.get("actions", []):
            lines.extend(_extract_element_text(action))

    elif t == "Action.OpenUrl":
        title = element.get("title", "")
        url = element.get("url", "")
        if title and url:
            lines.append(f"{title} ({url})")
        elif title:
            lines.append(title)

    elif t == "Action.Submit":
        title = element.get("title", "")
        if title:
            lines.append(title)

    return lines


def _extract_adaptive_card_text(card: dict) -> str:
    """Extract all readable text from an Adaptive Card as plain text."""
    lines: list[str] = []
    for element in card.get("body", []):
        lines.extend(_extract_element_text(element))
    for action in card.get("actions", []):
        lines.extend(_extract_element_text(action))
    return "\n".join(lines)


def _extract_attachments_text(attachments: list) -> str:
    """Extract text from Adaptive Card attachments."""
    lines: list[str] = []
    for att in attachments:
        if att.get("contentType") != "application/vnd.microsoft.card.adaptive":
            continue
        try:
            card = json.loads(att.get("content", "{}"))
        except (json.JSONDecodeError, TypeError):
            continue
        text = _extract_adaptive_card_text(card)
        if text:
            lines.append(text)
    return "\n".join(lines)


def _format_member(member: dict) -> dict:
    return {
        "id": member.get("userId") or member.get("id"),
        "displayName": member.get("displayName"),
        "email": member.get("email"),
        "roles": member.get("roles", []),
    }


def _format_message(msg: dict) -> dict:
    body_text = _strip_html((msg.get("body") or {}).get("content", ""))
    card_text = _extract_attachments_text(msg.get("attachments") or [])
    content = "\n".join(filter(None, [body_text, card_text]))
    return {
        "id": msg.get("id"),
        "sender": (msg.get("from") or {}).get("user", {}).get("displayName"),
        "createdDateTime": msg.get("createdDateTime"),
        "content": content,
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
    if auth.is_authenticated():
        accounts = auth._app.get_accounts()
        username = accounts[0].get("username") if accounts else "unknown"
        return json.dumps(
            {"status": "already_authenticated", "account": username},
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
    result = []
    for c in chats:
        members = [
            m.get("displayName", "")
            for m in (c.get("members") or [])
            if m.get("displayName")
        ]
        result.append({
            "id": c.get("id"),
            "topic": c.get("topic") or ", ".join(members),
            "chatType": c.get("chatType"),
            "lastUpdatedDateTime": c.get("lastUpdatedDateTime"),
            "members": members,
        })
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


# Tool: list_thread_replies
# Annotations: readOnlyHint=True, openWorldHint=True
@mcp.tool()
async def list_thread_replies(
    team_id: str, channel_id: str, message_id: str, limit: int = 20
) -> str:
    """List replies in a channel message thread.

    Use list_channel_messages to get the parent message_id.
    Returns the parent message followed by all replies in the thread."""
    _init_if_needed()
    client = _require_auth()
    parent = await client.get_channel_message(team_id, channel_id, message_id)
    replies = await client.list_thread_replies(team_id, channel_id, message_id, limit=limit)
    result = [_format_message(parent)] + [
        _format_message(m)
        for m in replies
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
async def send_channel_message(
    team_id: str, channel_id: str, content: str, mentions: str | None = None,
) -> str:
    """Send a message to a Teams channel.

    Use list_teams -> list_channels to get team_id and channel_id.
    For replies to existing messages, use reply_to_channel_message instead.

    mentions: optional JSON array of users to @mention.
    Format: [{"user_id": "...", "name": "Display Name"}]
    Use @DisplayName in content where the mention should appear.
    Get user_id from list_team_members, list_channel_members, or get_user.
    """
    _init_if_needed()
    client = _require_auth()
    parsed_mentions = None
    if mentions:
        try:
            parsed_mentions = json.loads(mentions)
        except (json.JSONDecodeError, TypeError):
            return json.dumps({"error": "Invalid mentions format. Expected JSON array: [{\"user_id\": \"...\", \"name\": \"...\"}]"})
    result = await client.send_channel_message(team_id, channel_id, content, mentions=parsed_mentions)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: send_chat_message
# Annotations: openWorldHint=True
@mcp.tool()
async def send_chat_message(chat_id: str, content: str, mentions: str | None = None) -> str:
    """Send a message to a Teams chat.

    Use list_chats to get the chat_id.
    Chat messages don't support threaded replies - just send a new message.

    mentions: optional JSON array of users to @mention.
    Format: [{"user_id": "...", "name": "Display Name"}]
    Use @DisplayName in content where the mention should appear.
    """
    _init_if_needed()
    client = _require_auth()
    parsed_mentions = None
    if mentions:
        try:
            parsed_mentions = json.loads(mentions)
        except (json.JSONDecodeError, TypeError):
            return json.dumps({"error": "Invalid mentions format. Expected JSON array: [{\"user_id\": \"...\", \"name\": \"...\"}]"})
    result = await client.send_chat_message(chat_id, content, mentions=parsed_mentions)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: reply_to_channel_message
# Annotations: openWorldHint=True
@mcp.tool()
async def reply_to_channel_message(
    team_id: str, channel_id: str, message_id: str, content: str, mentions: str | None = None,
) -> str:
    """Reply to a message in a Teams channel thread.

    Use list_channel_messages to get the message_id to reply to.
    For new top-level messages, use send_channel_message instead.

    mentions: optional JSON array of users to @mention.
    Format: [{"user_id": "...", "name": "Display Name"}]
    Use @DisplayName in content where the mention should appear.
    """
    _init_if_needed()
    client = _require_auth()
    parsed_mentions = None
    if mentions:
        try:
            parsed_mentions = json.loads(mentions)
        except (json.JSONDecodeError, TypeError):
            return json.dumps({"error": "Invalid mentions format. Expected JSON array: [{\"user_id\": \"...\", \"name\": \"...\"}]"})
    result = await client.reply_to_channel_message(
        team_id, channel_id, message_id, content, mentions=parsed_mentions,
    )
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: create_chat
# Annotations: openWorldHint=True
@mcp.tool()
async def create_chat(user_email: str, message: str) -> str:
    """Create a new 1:1 chat with a user and send the first message.

    Use this when no existing chat is found via list_chats.
    Requires the user's email address (e.g. amaksudov@avo.uz)."""
    _init_if_needed()
    client = _require_auth()
    me = await client.get_me()
    chat = await client.create_chat(me["id"], user_email)
    chat_id = chat["id"]
    msg = await client.send_chat_message(chat_id, message)
    return json.dumps({
        "status": "sent",
        "chat_id": chat_id,
        "message": _format_message(msg),
    }, ensure_ascii=False, indent=2)


@mcp.tool()
async def list_team_members(team_id: str) -> str:
    """List members of a Microsoft Teams team.

    Returns member id, display name, email, and roles (owner/member).
    Use list_teams to get the team_id.
    """
    _init_if_needed()
    client = _require_auth()
    members = await client.list_team_members(team_id)
    return json.dumps([_format_member(m) for m in members], ensure_ascii=False, indent=2)


@mcp.tool()
async def list_channel_members(team_id: str, channel_id: str) -> str:
    """List members of a specific channel.

    Returns member id, display name, email, and roles.
    Use list_channels to get the channel_id.
    """
    _init_if_needed()
    client = _require_auth()
    members = await client.list_channel_members(team_id, channel_id)
    return json.dumps([_format_member(m) for m in members], ensure_ascii=False, indent=2)


@mcp.tool()
async def list_chat_members(chat_id: str) -> str:
    """List members of a chat.

    Returns member id, display name, email, and roles.
    Use list_chats to get the chat_id.
    """
    _init_if_needed()
    client = _require_auth()
    members = await client.list_chat_members(chat_id)
    return json.dumps([_format_member(m) for m in members], ensure_ascii=False, indent=2)


@mcp.tool()
async def delete_message(
    message_id: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
) -> str:
    """Soft-delete a message you sent.

    For channel messages: provide team_id + channel_id + message_id.
    For chat messages: provide chat_id + message_id.
    The message can be recovered by an admin within 7 days.
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.soft_delete_chat_message(chat_id, message_id)
    elif team_id and channel_id:
        await client.soft_delete_channel_message(team_id, channel_id, message_id)
    else:
        return json.dumps({"error": "Provide chat_id OR (team_id + channel_id)"})
    return json.dumps({"status": "ok", "deleted": message_id})


@mcp.tool()
async def update_message(
    message_id: str,
    content: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
) -> str:
    """Edit a message you sent.

    For channel messages: provide team_id + channel_id + message_id.
    For chat messages: provide chat_id + message_id.
    Only available in Global cloud (not GCC/DOD).
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.update_chat_message(chat_id, message_id, content)
    elif team_id and channel_id:
        await client.update_channel_message(team_id, channel_id, message_id, content)
    else:
        return json.dumps({"error": "Provide chat_id OR (team_id + channel_id)"})
    return json.dumps({"status": "ok", "updated": message_id})


@mcp.tool()
async def set_reaction(
    message_id: str,
    reaction: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
) -> str:
    """React to a message with an emoji.

    For channel messages: provide team_id + channel_id + message_id.
    For chat messages: provide chat_id + message_id.
    Common reactions: like, angry, sad, laugh, heart, surprised.
    Custom reactions: any unicode emoji.
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.set_reaction_chat(chat_id, message_id, reaction)
    elif team_id and channel_id:
        await client.set_reaction_channel(team_id, channel_id, message_id, reaction)
    else:
        return json.dumps({"error": "Provide chat_id OR (team_id + channel_id)"})
    return json.dumps({"status": "ok", "reaction": reaction})


@mcp.tool()
async def unset_reaction(
    message_id: str,
    reaction: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
) -> str:
    """Remove a reaction from a message.

    For channel messages: provide team_id + channel_id + message_id.
    For chat messages: provide chat_id + message_id.
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.unset_reaction_chat(chat_id, message_id, reaction)
    elif team_id and channel_id:
        await client.unset_reaction_channel(team_id, channel_id, message_id, reaction)
    else:
        return json.dumps({"error": "Provide chat_id OR (team_id + channel_id)"})
    return json.dumps({"status": "ok", "reaction_removed": reaction})


@mcp.tool()
async def create_group_chat(member_emails: str, topic: str | None = None, message: str | None = None) -> str:
    """Create a new group chat with multiple users.

    member_emails: comma-separated email addresses (e.g. "a@org.com, b@org.com").
    topic: optional chat topic/name.
    message: optional first message to send.
    """
    _init_if_needed()
    client = _require_auth()
    me = await client.get_me()
    emails = [e.strip() for e in member_emails.split(",") if e.strip()]
    if len(emails) < 2:
        return json.dumps({"error": "Group chat requires at least 2 other members"})
    chat = await client.create_group_chat(me["id"], emails, topic=topic)
    chat_id = chat["id"]
    result: dict = {"status": "created", "chat_id": chat_id, "topic": topic}
    if message:
        msg = await client.send_chat_message(chat_id, message)
        result["message"] = _format_message(msg)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
async def pin_message(chat_id: str, message_id: str) -> str:
    """Pin a message in a chat.

    Only works in chats, not channels. Use list_chat_messages to get the message_id.
    """
    _init_if_needed()
    client = _require_auth()
    result = await client.pin_message(chat_id, message_id)
    return json.dumps({"status": "ok", "pinned_message_info_id": result.get("id")}, ensure_ascii=False, indent=2)


@mcp.tool()
async def unpin_message(chat_id: str, pinned_message_info_id: str) -> str:
    """Unpin a message from a chat.

    Use list_pinned_messages to get the pinned_message_info_id (NOT the message_id).
    """
    _init_if_needed()
    client = _require_auth()
    await client.unpin_message(chat_id, pinned_message_info_id)
    return json.dumps({"status": "ok", "unpinned": pinned_message_info_id})


@mcp.tool()
async def list_pinned_messages(chat_id: str) -> str:
    """List pinned messages in a chat.

    Returns pinned message info including the message content.
    Use list_chats to get the chat_id.
    """
    _init_if_needed()
    client = _require_auth()
    pinned = await client.list_pinned_messages(chat_id)
    result = []
    for p in pinned:
        msg = p.get("message", {})
        result.append({
            "pinned_message_info_id": p.get("id"),
            "message": _format_message(msg) if msg else None,
        })
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
async def mark_chat_read(chat_id: str) -> str:
    """Mark a chat as read for the current user.

    Use list_chats to get the chat_id.
    """
    _init_if_needed()
    client = _require_auth()
    me = await client.get_me()
    await client.mark_chat_read(chat_id, me["id"])
    return json.dumps({"status": "ok", "chat_id": chat_id, "marked": "read"})


@mcp.tool()
async def mark_chat_unread(chat_id: str, last_message_read_date_time: str) -> str:
    """Mark a chat as unread for the current user.

    last_message_read_date_time: ISO 8601 timestamp of the last message
    that should be considered as read (e.g. "2026-03-26T10:00:00Z").
    Messages after this timestamp will appear as unread.
    """
    _init_if_needed()
    client = _require_auth()
    me = await client.get_me()
    await client.mark_chat_unread(chat_id, me["id"], last_message_read_date_time)
    return json.dumps({"status": "ok", "chat_id": chat_id, "marked": "unread"})


@mcp.tool()
async def get_user_presence(user_id: str) -> str:
    """Get the presence/availability status of a user.

    Returns availability (Available, Busy, DoNotDisturb, Away, Offline, etc.)
    and activity (InACall, InAMeeting, Presenting, etc.).
    Get the user_id from list_team_members, list_chat_members, or get_user.
    """
    _init_if_needed()
    client = _require_auth()
    presence = await client.get_user_presence(user_id)
    return json.dumps({
        "availability": presence.get("availability"),
        "activity": presence.get("activity"),
        "statusMessage": (presence.get("statusMessage") or {}).get("message", {}).get("content"),
    }, ensure_ascii=False, indent=2)


@mcp.tool()
async def search_messages(query: str, size: int = 25) -> str:
    """Search for messages across all chats and channels.

    Full-text search on message body and attachments.
    Returns matching messages ranked by relevance with sender and context.
    Uses beta API - results may vary.
    """
    _init_if_needed()
    client = _require_auth()
    hits = await client.search_messages(query, size=size)
    result = []
    for hit in hits:
        resource = hit.get("resource", {})
        sender = resource.get("from", {}).get("emailAddress", {})
        result.append({
            "summary": hit.get("summary"),
            "sender": sender.get("name"),
            "senderEmail": sender.get("address"),
            "createdDateTime": resource.get("createdDateTime"),
            "chatId": resource.get("chatId"),
            "channelIdentity": resource.get("channelIdentity"),
            "webLink": resource.get("webLink"),
        })
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
async def get_user(query: str, limit: int = 10) -> str:
    """Search for users by name or email.

    Returns user id, display name, email, and job title.
    Useful for finding user_id needed by other tools (e.g. get_user_presence, create_chat).
    """
    _init_if_needed()
    client = _require_auth()
    users = await client.search_users(query, limit=limit)
    result = [
        {
            "id": u.get("id"),
            "displayName": u.get("displayName"),
            "email": u.get("mail") or u.get("userPrincipalName"),
            "jobTitle": u.get("jobTitle"),
        }
        for u in users
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


def main():
    _init()
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

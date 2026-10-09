import argparse
import functools
import html
import inspect
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import webbrowser
from datetime import UTC, datetime
from importlib.metadata import PackageNotFoundError, version
from urllib.parse import parse_qs, unquote, urlsplit

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from teams_mcp.auth import NOT_AUTHENTICATED, AuthError, AuthManager
from teams_mcp.graph import GraphApiError, GraphClient

try:
    _VERSION = version("teams-mcp-server")
except PackageNotFoundError:  # running from a source tree without an install
    _VERSION = "0.0.0+dev"

# mcp 2.x defaults version to "" (1.x reported the SDK's own version), so set it
# explicitly or clients see a blank version in serverInfo.
mcp = MCPServer(
    "teams-mcp",
    version=_VERSION,
    instructions=(
        "Microsoft Teams via Microsoft Graph, acting as the signed-in user: anything "
        "sent, edited, deleted or reacted to appears under their name. Use these tools "
        "for Teams chats and channels - reading, searching, sending, editing and deleting "
        "messages (with @mentions), thread replies, reactions, pins, read state, creating "
        "1:1 and group chats, listing teams, channels, members and tags, finding users "
        "and their presence, and downloading inline images.\n"
        "The sign-in is cached between sessions. Only when a tool reports \"Not "
        "authenticated\", ask the user to run `teams-mcp login` in a terminal (the "
        "command that starts this server, with `login` appended) and retry once they have "
        "signed in; the server picks up the sign-in without a restart."
    ),
)

auth: AuthManager | None = None
graph: GraphClient | None = None

# Largest `limit` a list tool accepts. Graph returns at most 50 messages per page and the
# pages are fetched one after another; a longer result also risks exceeding what an MCP
# client accepts as one tool output, so older history is paged with a cursor instead.
MAX_LIST_LIMIT = 200


def _init():
    global auth, graph
    try:
        tenant_id = os.environ["TEAMS_MCP_TENANT_ID"]
        client_id = os.environ["TEAMS_MCP_CLIENT_ID"]
    except KeyError:
        print("teams-mcp: missing TEAMS_MCP_TENANT_ID/CLIENT_ID", file=sys.stderr)
        raise SystemExit(1) from None
    scopes_env = os.environ.get("TEAMS_MCP_SCOPES")
    scopes = [s.strip() for s in scopes_env.split(",") if s.strip()] if scopes_env else None
    auth = AuthManager(tenant_id=tenant_id, client_id=client_id, scopes=scopes)
    graph = GraphClient(token_provider=auth.get_token)


def _init_if_needed():
    if auth is None:
        _init()


def _require_auth() -> GraphClient:
    if graph is None or not auth.is_authenticated():
        raise AuthError(NOT_AUTHENTICATED)
    return graph


# Failures the caller can act on. mcp 2.x passes only ToolError text to the client; any
# other exception reaches it as a bare "Error executing tool <name>" and is logged with
# its traceback as a crash, which is what a real bug should stay. Bad arguments are
# checked in the tools and raised as ToolError directly, so ValueError is not listed:
# it would also swallow bugs such as a non-JSON Graph body.
_ANTICIPATED_ERRORS = (AuthError, GraphApiError)


def _tool(fn):
    """Register async fn as an MCP tool whose anticipated failures reach the client as text."""
    if not inspect.iscoroutinefunction(fn):
        # the wrapper awaits fn, so a sync tool would fail on every call
        raise TypeError(f"tool {fn.__name__} must be async")

    @functools.wraps(fn)
    async def wrapper(*args, **kwargs):
        try:
            return await fn(*args, **kwargs)
        except _ANTICIPATED_ERRORS as exc:
            raise ToolError(str(exc)) from exc

    return mcp.tool()(wrapper)


def _check_limit(limit: int) -> None:
    if not 1 <= limit <= MAX_LIST_LIMIT:
        raise ToolError(f"limit must be between 1 and {MAX_LIST_LIMIT}, got {limit}")


def _parse_timestamp(name: str, value: str | None) -> datetime | None:
    """Read an ISO 8601 argument as an aware UTC datetime; no offset means UTC.

    A blank value means no bound: some clients send "" for an omitted argument.
    """
    if value is None or not value.strip():
        return None
    try:
        parsed = datetime.fromisoformat(value.strip())
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=UTC)
        return parsed.astimezone(UTC)
    except (ValueError, OverflowError):  # OverflowError: shifting e.g. year 1 to UTC
        raise ToolError(
            f"{name} must be an ISO 8601 timestamp such as 2026-10-07T12:00:00Z, got {value!r}"
        ) from None


_GUID = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")


def _parse_message_link(link: str) -> tuple[dict[str, str], str]:
    """Split a Teams message link into where the message lives and its id.

    The path is /l/message/<chat or channel id>/<message id>. A channel link carries the
    team as groupId and the thread root as parentMessageId (the message's own id for a
    root message); a chat link has neither. Links come from message content, which
    anyone in a chat can write, so every id is checked before it reaches a Graph path.
    """
    # a link copied out of message content keeps its HTML escaping (&amp;)
    url = urlsplit(html.unescape(link.strip()))
    match = re.fullmatch(r"/l/message/([^/]+)/([0-9]+)/?", url.path)
    if not match:
        raise ToolError(
            "link must be a Teams message link such as https://teams.microsoft.com/l/message/"
            f"<chat or channel id>/<message id>, got {link!r}"
        )
    conversation_id, message_id = unquote(match.group(1)), match.group(2)
    if "/" in conversation_id:
        raise ToolError(f"link has an invalid chat or channel id: {link!r}")
    query = parse_qs(url.query)
    if "groupId" not in query:
        if conversation_id.endswith(("@thread.tacv2", "@thread.skype")):
            raise ToolError(
                "This channel link has no groupId (team). Pass team_id, channel_id and "
                "message_id instead."
            )
        return {"chat_id": conversation_id}, message_id
    team_id = query["groupId"][0]
    parent = query.get("parentMessageId", [message_id])[0]
    if not _GUID.fullmatch(team_id) or not re.fullmatch(r"[0-9]+", parent):
        raise ToolError(f"link has an invalid groupId or parentMessageId: {link!r}")
    location = {"team_id": team_id, "channel_id": conversation_id}
    if parent != message_id:
        location["parent_message_id"] = parent
    return location, message_id


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

    elif t == "Action.ShowCard":
        title = element.get("title", "")
        if title:
            lines.append(title)
        nested = element.get("card", {})
        if nested:
            lines.append(_extract_adaptive_card_text(nested))

    return lines


def _extract_adaptive_card_text(card: dict) -> str:
    """Extract all readable text from an Adaptive Card as plain text."""
    lines: list[str] = []
    for element in card.get("body", []):
        lines.extend(_extract_element_text(element))
    for action in card.get("actions", []):
        lines.extend(_extract_element_text(action))
    return "\n".join(lines)


def _extract_forwarded_text(att: dict) -> str:
    """Extract text from a forwarded/quoted message reference attachment.

    forwardedMessageReference uses: originalMessageSender, originalMessageContent
    messageReference uses: messageSender, messagePreview, body
    """
    raw = att.get("content", "")
    if not raw:
        return ""
    if isinstance(raw, str):
        try:
            data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            return _strip_html(raw)
    else:
        data = raw

    parts = []

    # Sender: try forwarded fields first, then quoted-reply fields
    # (.get("user") may be present-but-null for bot/application senders)
    sender = (
        ((data.get("originalMessageSender") or {}).get("user") or {}).get("displayName")
        or ((data.get("messageSender") or {}).get("user") or {}).get("displayName")
    )
    if sender:
        parts.append(f"[Forwarded from {sender}]")

    # Content: try forwarded field first, then quoted-reply fields
    content = data.get("originalMessageContent", "")
    if content:
        parts.append(_strip_html(content))
    else:
        preview = data.get("messagePreview", "")
        if preview:
            parts.append(preview)
        else:
            body = data.get("body")
            if isinstance(body, dict):
                body_content = body.get("content", "")
                if body_content:
                    parts.append(_strip_html(body_content))
            elif isinstance(body, str):
                parts.append(body)

    return "\n".join(parts)


def _extract_attachments_text(attachments: list) -> str:
    """Extract text from Adaptive Card and forwarded message attachments."""
    lines: list[str] = []
    for att in attachments:
        ct = att.get("contentType", "")

        if ct == "application/vnd.microsoft.card.adaptive":
            raw = att.get("content", "{}")
            if isinstance(raw, dict):
                card = raw
            else:
                try:
                    card = json.loads(raw)
                except (json.JSONDecodeError, TypeError):
                    continue
            text = _extract_adaptive_card_text(card)
            if text:
                lines.append(text)

        elif "messageReference" in ct or "forwardedMessage" in ct:
            text = _extract_forwarded_text(att)
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


_MENTION_SHAPE = '{"user_id": "...", "name": "..."} or {"tag_id": "...", "name": "..."}'


def _parse_mentions(mentions: list | str | None) -> list[dict] | None:
    """Accept mentions as list (deserialized by the SDK) or JSON string.

    Anything malformed is an error before sending: a message cannot gain mentions later
    (update_message turns them into plain text). A blank string or JSON null means no
    mentions, since some clients send those for an omitted argument.
    """
    if isinstance(mentions, str):
        if not mentions.strip():
            return None
        try:
            mentions = json.loads(mentions)
        except json.JSONDecodeError as exc:
            raise ToolError(f"mentions is not valid JSON ({exc}); expected a list of "
                            f"{_MENTION_SHAPE}") from None
    if mentions is None:
        return None
    if not isinstance(mentions, list):
        raise ToolError(f"mentions must be a list of {_MENTION_SHAPE}")
    for i, mention in enumerate(mentions):
        valid = (
            isinstance(mention, dict)
            and isinstance(mention.get("name"), str)
            and mention["name"] != ""
            and bool(mention.get("user_id")) != bool(mention.get("tag_id"))
        )
        if not valid:
            raise ToolError(f"mentions[{i}] must be {_MENTION_SHAPE}, got {json.dumps(mention)}")
    return mentions


def _format_attachments(attachments: list) -> list[dict]:
    """Extract attachment metadata for file/image attachments."""
    result = []
    for att in attachments:
        ct = att.get("contentType", "")
        if ct == "application/vnd.microsoft.card.adaptive":
            continue
        if "messageReference" in ct or "forwardedMessage" in ct:
            continue  # handled by _extract_attachments_text
        info: dict = {"id": att.get("id"), "name": att.get("name"), "contentType": ct}
        url = att.get("contentUrl")
        if url:
            info["contentUrl"] = url
        result.append(info)
    return result


def _format_hosted_contents(body_html: str) -> list[dict]:
    """Extract hosted content IDs from inline <img> tags."""
    result = []
    for match in re.finditer(r'src="[^"]*?/hostedContents/([^/"]+)/\$value"', body_html):
        result.append({"hostedContentId": match.group(1)})
    return result


def _format_message(msg: dict) -> dict:
    # body.content may be present-but-null (e.g. deleted reply) - .get() default won't cover it
    body_html = (msg.get("body") or {}).get("content") or ""
    body_text = _strip_html(body_html)
    card_text = _extract_attachments_text(msg.get("attachments") or [])
    content = "\n".join(filter(None, [body_text, card_text]))
    # bot/workflow messages carry from.user = null and from.application instead
    frm = msg.get("from") or {}
    sender = (
        (frm.get("user") or {}).get("displayName")
        or (frm.get("application") or {}).get("displayName")
    )
    result: dict = {
        "id": msg.get("id"),
        "sender": sender,
        "createdDateTime": msg.get("createdDateTime"),
        "content": content,
    }
    file_attachments = _format_attachments(msg.get("attachments") or [])
    hosted = _format_hosted_contents(body_html)
    if file_attachments:
        result["attachments"] = file_attachments
    if hosted:
        result["hostedContents"] = hosted
    mentions = msg.get("mentions") or []
    if mentions:
        result["mentions"] = mentions  # raw entities - ids are needed for @mention replies
    return result


# Tool: list_teams
# Annotations: readOnlyHint=True, openWorldHint=True
@_tool
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
@_tool
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


# Tool: list_team_tags
# Annotations: readOnlyHint=True, openWorldHint=True
@_tool
async def list_team_tags(team_id: str) -> str:
    """List TEAM-level tags of a team (id, name, member count).

    Requires the TeamworkTag.Read delegated permission. These ids work as
    mention tag_ids in STANDARD channels. Shared channels have their own
    channel-scoped tags that Graph does not expose via any API - a team-level
    id posted there renders a phantom tag (no name, 0 members). For shared
    channels, reuse the tag id from the "mentions" entities of an existing
    message where a human @mentioned the tag."""
    _init_if_needed()
    client = _require_auth()
    tags = await client.list_team_tags(team_id)
    result = [
        {
            "id": t.get("id"),
            "name": t.get("displayName"),
            "memberCount": t.get("memberCount"),
        }
        for t in tags
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: list_chats
# Annotations: readOnlyHint=True, openWorldHint=True
@_tool
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
@_tool
async def list_channel_messages(team_id: str, channel_id: str, limit: int = 20) -> str:
    """List recent top-level messages in a Teams channel.

    Use list_teams -> list_channels to get team_id and channel_id. Replies are not
    included - read a thread with list_thread_replies. Returns up to `limit` (1-200)
    messages, the threads with the latest activity first (a new reply moves its thread
    up); there is no cursor for older ones. Each message has id, sender, timestamp and
    plain-text content, plus attachments, hostedContents (inline image ids for
    download_attachment) and mention entities when present. System messages are
    excluded.
    """
    _check_limit(limit)
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
@_tool
async def list_thread_replies(
    team_id: str, channel_id: str, message_id: str, limit: int = 20
) -> str:
    """List replies in a channel message thread.

    Use list_channel_messages to get the parent message_id (a top-level message).
    Returns the parent message followed by up to `limit` (1-200) replies, newest first.
    In a longer thread the oldest replies, the ones right after the parent, are left
    out and nothing marks the cut. System messages are excluded.
    """
    _check_limit(limit)
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
@_tool
async def list_chat_messages(
    chat_id: str, limit: int = 20, before: str | None = None, after: str | None = None,
) -> str:
    """List messages in a chat, newest first, a page at a time.

    Use list_chats to get the chat_id. Returns {"messages": [...], "next_before": ...}
    with up to `limit` (1-200) messages created after `after` and before `before`: ISO
    8601 timestamps, both exclusive and optional; one without an offset is read as UTC.
    For the next older page call again with before=next_before and the same `after`;
    next_before is null when no older messages are left in that range. Pages come
    newest first, so `after` alone returns the newest messages of the chat, not the
    ones right after that time.
    Each message has id, sender, createdDateTime and plain-text content, plus
    attachments, hostedContents (inline image ids for download_attachment) and mention
    entities when present. System messages are left out, so a page can hold fewer
    than `limit` messages.
    """
    _check_limit(limit)
    before_ts = _parse_timestamp("before", before)
    after_ts = _parse_timestamp("after", after)
    _init_if_needed()
    client = _require_auth()
    raw = await client.list_chat_messages(chat_id, limit=limit, before=before_ts, after=after_ts)
    # A full page may have older messages behind it. The cursor is the oldest item Graph
    # returned, system messages included, so the next page starts where this one ended.
    next_before = raw[-1]["createdDateTime"] if len(raw) == limit else None
    result = {
        "messages": [_format_message(m) for m in raw if m.get("messageType") == "message"],
        "next_before": next_before,
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


# Tool: get_message
# Annotations: readOnlyHint=True, openWorldHint=True
@_tool
async def get_message(
    link: str | None = None,
    message_id: str | None = None,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
    parent_message_id: str | None = None,
) -> str:
    """Get one chat or channel message by its Teams link or by ids.

    link: a message link copied from Teams, shaped like
    https://teams.microsoft.com/l/message/<chat or channel id>/<message id>?...
    Without a link, pass message_id with chat_id, or with team_id + channel_id (plus
    parent_message_id, the thread root id, for a reply in a channel thread).
    Returns where the message lives (chat_id, or team_id + channel_id and, for a reply,
    parent_message_id) and the message with id, sender, createdDateTime and content.
    To see what led up to a chat message, call list_chat_messages with
    before=<its createdDateTime>. Pages come newest first, so for what followed pass
    after=<its createdDateTime> together with a before a little later (an hour on, for
    example); after alone returns the newest messages of the chat.
    """
    if link:
        location, message_id = _parse_message_link(link)
    elif message_id and chat_id:
        location = {"chat_id": chat_id}
    elif message_id and team_id and channel_id:
        location = {"team_id": team_id, "channel_id": channel_id}
        if parent_message_id:
            location["parent_message_id"] = parent_message_id
    else:
        raise ToolError("Provide link, or message_id with chat_id or team_id + channel_id")
    _init_if_needed()
    client = _require_auth()
    if "chat_id" in location:
        msg = await client.get_chat_message(location["chat_id"], message_id)
    elif "parent_message_id" in location:
        msg = await client.get_channel_reply(
            location["team_id"], location["channel_id"], location["parent_message_id"], message_id,
        )
    else:
        msg = await client.get_channel_message(
            location["team_id"], location["channel_id"], message_id,
        )
    return json.dumps({**location, "message": _format_message(msg)}, ensure_ascii=False, indent=2)


# Tool: send_channel_message
# Annotations: openWorldHint=True
@_tool
async def send_channel_message(
    team_id: str, channel_id: str, content: str, mentions: list | str | None = None,
) -> str:
    """Send a message to a Teams channel.

    Use list_teams -> list_channels to get team_id and channel_id.
    For replies to existing messages, use reply_to_channel_message instead.

    content is plain text: newlines are kept and URLs become links; Markdown and HTML
    show literally. Each mention's "@<name>" must appear in content exactly as given in
    name, or that mention is dropped without an error.

    mentions: optional JSON array of users or team tags to @mention.
    Format: [{"user_id": "...", "name": "Display Name"} | {"tag_id": "...", "name": "TagName"}]
    Get tag_id from list_team_tags (tags work only in channel messages).
    Use @DisplayName in content where the mention should appear.
    Get user_id from list_team_members, list_channel_members, or get_user.
    """
    _init_if_needed()
    client = _require_auth()
    parsed_mentions = _parse_mentions(mentions)
    result = await client.send_channel_message(team_id, channel_id, content, mentions=parsed_mentions)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: send_chat_message
# Annotations: openWorldHint=True
@_tool
async def send_chat_message(
    chat_id: str, content: str, mentions: list | str | None = None, reply_to: str | None = None,
) -> str:
    """Send a message to a Teams chat.

    Use list_chats to get the chat_id.
    reply_to: optional message ID to reply to (shows as a quoted reply).
    Use list_chat_messages to get the message_id.

    content is plain text: newlines are kept and URLs become links; Markdown and HTML
    show literally. Each mention's "@<name>" must appear in content exactly as given in
    name, or that mention is dropped without an error.

    mentions: optional JSON array of users to @mention:
    [{"user_id": "...", "name": "Display Name"}]. Team tags cannot be mentioned in
    chats. Get user_id from list_chat_members or get_user.
    Use @DisplayName in content where the mention should appear.
    """
    _init_if_needed()
    client = _require_auth()
    parsed_mentions = _parse_mentions(mentions)
    result = await client.send_chat_message(chat_id, content, mentions=parsed_mentions, reply_to_id=reply_to)
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: reply_to_channel_message
# Annotations: openWorldHint=True
@_tool
async def reply_to_channel_message(
    team_id: str, channel_id: str, message_id: str, content: str, mentions: list | str | None = None,
) -> str:
    """Reply to a message in a Teams channel thread.

    Use list_channel_messages to get the message_id to reply to.
    For new top-level messages, use send_channel_message instead.

    content is plain text: newlines are kept and URLs become links; Markdown and HTML
    show literally. Each mention's "@<name>" must appear in content exactly as given in
    name, or that mention is dropped without an error.

    mentions: optional JSON array of users or team tags to @mention.
    Format: [{"user_id": "...", "name": "Display Name"} | {"tag_id": "...", "name": "TagName"}]
    Get tag_id from list_team_tags (tags work only in channel messages).
    Use @DisplayName in content where the mention should appear.
    """
    _init_if_needed()
    client = _require_auth()
    parsed_mentions = _parse_mentions(mentions)
    result = await client.reply_to_channel_message(
        team_id, channel_id, message_id, content, mentions=parsed_mentions,
    )
    return json.dumps(_format_message(result), ensure_ascii=False, indent=2)


# Tool: create_chat
# Annotations: openWorldHint=True
@_tool
async def create_chat(user_email: str, message: str) -> str:
    """Send a message in the 1:1 chat with a user, creating the chat if needed.

    If a 1:1 chat with this user already exists, Graph returns it, so there is no need
    to look it up with list_chats first. Returns chat_id and the sent message.
    user_email: the user's sign-in address (userPrincipalName) or user id from
    get_user; a mail alias that differs from the sign-in address may not resolve.
    """
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


@_tool
async def list_team_members(team_id: str) -> str:
    """List members of a Microsoft Teams team.

    Returns member id, display name, email, and roles (owner/member).
    Use list_teams to get the team_id.
    """
    _init_if_needed()
    client = _require_auth()
    members = await client.list_team_members(team_id)
    return json.dumps([_format_member(m) for m in members], ensure_ascii=False, indent=2)


@_tool
async def list_channel_members(team_id: str, channel_id: str) -> str:
    """List members of a specific channel.

    Returns member id, display name, email, and roles.
    Use list_channels to get the channel_id.
    """
    _init_if_needed()
    client = _require_auth()
    members = await client.list_channel_members(team_id, channel_id)
    return json.dumps([_format_member(m) for m in members], ensure_ascii=False, indent=2)


@_tool
async def list_chat_members(chat_id: str) -> str:
    """List members of a chat.

    Returns member id, display name, email, and roles.
    Use list_chats to get the chat_id.
    """
    _init_if_needed()
    client = _require_auth()
    members = await client.list_chat_members(chat_id)
    return json.dumps([_format_member(m) for m in members], ensure_ascii=False, indent=2)


@_tool
async def delete_message(
    message_id: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
) -> str:
    """Soft-delete a message you sent.

    For channel messages: provide team_id + channel_id + message_id.
    For chat messages: provide chat_id + message_id.
    Works on top-level channel messages and on chat messages; channel thread replies
    are not supported (Graph addresses them under their parent message).
    Other members then see "This message has been deleted"; this server has no undo tool.
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.soft_delete_chat_message(chat_id, message_id)
    elif team_id and channel_id:
        await client.soft_delete_channel_message(team_id, channel_id, message_id)
    else:
        raise ToolError("Provide chat_id OR (team_id + channel_id)")
    return json.dumps({"status": "ok", "deleted": message_id})


@_tool
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
    Works on top-level channel messages and on chat messages; channel thread replies
    are not supported (Graph addresses them under their parent message).
    Replaces the whole body with plain-text content; @mentions cannot be added on edit
    and existing ones become plain text. Only available in Global cloud (not GCC/DOD).
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.update_chat_message(chat_id, message_id, content)
    elif team_id and channel_id:
        await client.update_channel_message(team_id, channel_id, message_id, content)
    else:
        raise ToolError("Provide chat_id OR (team_id + channel_id)")
    return json.dumps({"status": "ok", "updated": message_id})


@_tool
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
    Works on top-level channel messages and on chat messages; channel thread replies
    are not supported (Graph addresses them under their parent message).
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
        raise ToolError("Provide chat_id OR (team_id + channel_id)")
    return json.dumps({"status": "ok", "reaction": reaction})


@_tool
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
    Works on top-level channel messages and on chat messages; channel thread replies
    are not supported (Graph addresses them under their parent message).
    """
    _init_if_needed()
    client = _require_auth()
    if chat_id:
        await client.unset_reaction_chat(chat_id, message_id, reaction)
    elif team_id and channel_id:
        await client.unset_reaction_channel(team_id, channel_id, message_id, reaction)
    else:
        raise ToolError("Provide chat_id OR (team_id + channel_id)")
    return json.dumps({"status": "ok", "reaction_removed": reaction})


@_tool
async def create_group_chat(member_emails: str, topic: str | None = None, message: str | None = None) -> str:
    """Create a new group chat with you and at least two other users.

    For one other person use create_chat. Every call creates a new chat; Graph does not
    reuse an existing group chat with the same members.
    member_emails: comma-separated sign-in addresses of the other members
    (e.g. "a@org.com, b@org.com").
    topic: optional chat topic/name.
    message: optional first message to send.
    """
    _init_if_needed()
    client = _require_auth()
    emails = [e.strip() for e in member_emails.split(",") if e.strip()]
    if len(emails) < 2:
        raise ToolError("Group chat requires at least 2 other members")
    me = await client.get_me()
    chat = await client.create_group_chat(me["id"], emails, topic=topic)
    chat_id = chat["id"]
    result: dict = {"status": "created", "chat_id": chat_id, "topic": topic}
    if message:
        msg = await client.send_chat_message(chat_id, message)
        result["message"] = _format_message(msg)
    return json.dumps(result, ensure_ascii=False, indent=2)


@_tool
async def pin_message(chat_id: str, message_id: str) -> str:
    """Pin a message in a chat.

    Only works in chats, not channels. Use list_chat_messages to get the message_id.
    """
    _init_if_needed()
    client = _require_auth()
    result = await client.pin_message(chat_id, message_id)
    return json.dumps({"status": "ok", "pinned_message_info_id": result.get("id")}, ensure_ascii=False, indent=2)


@_tool
async def unpin_message(chat_id: str, pinned_message_info_id: str) -> str:
    """Unpin a message from a chat.

    Use list_pinned_messages to get the pinned_message_info_id (NOT the message_id).
    """
    _init_if_needed()
    client = _require_auth()
    await client.unpin_message(chat_id, pinned_message_info_id)
    return json.dumps({"status": "ok", "unpinned": pinned_message_info_id})


@_tool
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


@_tool
async def mark_chat_read(chat_id: str) -> str:
    """Mark a chat as read for the current user.

    Use list_chats to get the chat_id.
    """
    _init_if_needed()
    client = _require_auth()
    me = await client.get_me()
    await client.mark_chat_read(chat_id, me["id"])
    return json.dumps({"status": "ok", "chat_id": chat_id, "marked": "read"})


@_tool
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


@_tool
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


@_tool
async def search_messages(query: str, size: int = 25) -> str:
    """Search Teams messages in all chats and channels the signed-in user can see.

    Matches message content, including attachments (Microsoft Search). Returns up to
    `size` hits, newest first, with no further pages. Each hit has a snippet
    (`summary`, not the full body), sender name and email, timestamp, chatId or
    channelIdentity, and a webLink. Message ids are not returned - to reply to, react
    to or delete a hit, find it with list_chat_messages or list_channel_messages.
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


@_tool
async def get_user(query: str, limit: int = 10) -> str:
    """Find users whose display name or email address starts with `query`.

    Prefix match only (Graph startsWith on displayName or mail): a surname or a
    fragment from the middle of a name finds nothing. Returns up to `limit` users with
    id, display name, email (mail, else userPrincipalName) and job title. The id is
    what get_user_presence and mention user_id take; create_chat takes the email.
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


@_tool
async def download_attachment(
    message_id: str,
    hosted_content_id: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
    parent_message_id: str | None = None,
) -> str:
    """Download an inline image (hosted content) from a message.

    For channel messages: provide team_id + channel_id + message_id.
    For content inside a channel REPLY: additionally pass parent_message_id
    (the thread root id) - Graph serves reply content only under
    /messages/{parent}/replies/{reply}.
    For chat messages: provide chat_id + message_id.
    hosted_content_id: from the hostedContents array in message data.
    Returns the local file path to the downloaded image. File attachments listed under
    `attachments` (SharePoint/OneDrive files) cannot be downloaded with this tool. The
    image is written to the system temp directory on the machine running this server
    and is not deleted afterwards.
    """
    if not chat_id and not (team_id and channel_id):
        raise ToolError("Provide chat_id OR (team_id + channel_id)")
    _init_if_needed()
    client = _require_auth()
    data = await client.download_hosted_content(
        chat_id=chat_id,
        team_id=team_id,
        channel_id=channel_id,
        message_id=message_id,
        hosted_content_id=hosted_content_id,
        parent_message_id=parent_message_id,
    )
    suffix = ".png"
    if data[:3] == b"\xff\xd8\xff":
        suffix = ".jpg"
    elif data[:4] == b"GIF8":
        suffix = ".gif"
    elif data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        suffix = ".webp"
    fd, path = tempfile.mkstemp(suffix=suffix, prefix="teams_attachment_")
    os.write(fd, data)
    os.close(fd)
    return json.dumps({"path": path, "size": len(data)}, ensure_ascii=False, indent=2)


def _clipboard_command() -> list[str] | None:
    """The command that puts its stdin on the system clipboard here, if there is one."""
    if sys.platform == "darwin":
        return ["pbcopy"]
    if sys.platform == "win32":
        return ["clip"]
    for command in (
        ["wl-copy"],
        # without -selection clipboard, xclip fills the PRIMARY selection instead
        ["xclip", "-selection", "clipboard"],
        ["xsel", "--clipboard", "--input"],
    ):
        if shutil.which(command[0]):
            return command
    return None


def _copy_to_clipboard(text: str) -> bool:
    """Put text on the system clipboard, as `gh auth login --clipboard` does.

    Returns False when there is no clipboard here (no tool, or no display over SSH); the
    caller has printed the text anyway, so that is not an error.
    """
    command = _clipboard_command()
    if command is None:
        return False
    try:
        # xclip and wl-copy fork to serve the clipboard and keep inherited pipes open, so
        # capturing their output would block until the timeout
        subprocess.run(
            command, input=text, text=True, check=True, timeout=5,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return True


def _can_open_browser() -> bool:
    """True on a local desktop session.

    Over SSH a browser would open on the remote machine's desktop, out of sight, and on
    Linux without a display server webbrowser falls back to a console browser that takes
    over the terminal.
    """
    if os.environ.get("SSH_CONNECTION") or os.environ.get("SSH_TTY"):
        return False
    if sys.platform.startswith("linux"):
        return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))
    return True


def _open_in_browser(url: str) -> bool:
    """Open url in the default browser, as `gh auth login --web` does; False if it did not."""
    if not _can_open_browser():
        return False
    try:
        return webbrowser.open(url)
    except webbrowser.Error:  # no runnable browser; the URL is printed anyway
        return False


def _login_in_terminal() -> int:
    """Device code sign-in on the terminal; it writes the cache the server reads."""
    try:
        if auth.is_authenticated():
            print(f"Already signed in as {auth.username()}.")
            return 0
        flow = auth.login()
        print(flow["message"], flush=True)
        if _copy_to_clipboard(flow["user_code"]):
            print("The code is copied to the clipboard.", flush=True)
        if _open_in_browser(flow["verification_uri"]):
            print(f"Opened {flow['verification_uri']} in the browser.", flush=True)
        result = auth.complete_login(flow)
    except AuthError as exc:
        print(f"teams-mcp login: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("teams-mcp login: cancelled", file=sys.stderr)
        return 130
    print(f"Signed in as {result['account']}.")
    return 0


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="teams-mcp",
        description=(
            "MCP server for Microsoft Teams over stdio. Needs TEAMS_MCP_TENANT_ID and "
            "TEAMS_MCP_CLIENT_ID (TEAMS_MCP_SCOPES is optional)."
        ),
    )
    commands = parser.add_subparsers(dest="command")
    commands.add_parser(
        "login",
        help="sign in with a device code in this terminal; a running server picks it up",
    )
    args = parser.parse_args(argv)
    _init()
    if args.command == "login":
        raise SystemExit(_login_in_terminal())
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

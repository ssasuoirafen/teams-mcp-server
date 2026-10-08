<!-- last reviewed: 2026-06-26 -->

# Teams MCP Server

MCP server for Microsoft Teams via Microsoft Graph API. Delegated auth (device code flow), no client secret.

## Stack

Python >=3.12 (dev pin: 3.14 via `.python-version`), `mcp[cli]` 2.x (`MCPServer`; the major is bounded `<3`, so the next SDK major is a deliberate port rather than a broken fresh install), httpx, msal. Tests in `tests/` (`uv run pytest`): `test_server.py` pins the tool surface and formatting, `test_tools.py` calls tools through an in-process `mcp.Client` over a fake Graph transport (`httpx.MockTransport`) and fake sign-in, with no network; ruff + pytest in the dev group (`uv sync`). CI runs on push/PR via `.github/workflows/ci.yml` (uv sync + pytest + advisory ruff).

## Commands

```bash
uv run teams-mcp              # run server (needs TEAMS_MCP_TENANT_ID, TEAMS_MCP_CLIENT_ID)
```

## Environment Variables

- `TEAMS_MCP_TENANT_ID` (required) - Azure AD tenant ID
- `TEAMS_MCP_CLIENT_ID` (required) - App registration client ID
- `TEAMS_MCP_SCOPES` (optional) - comma-separated scope override (default: `.default`; scope semantics in the Scopes section)

## Architecture

Three-file pattern, all logic flows through the same pipeline:

```
server.py (MCP tools) -> graph.py (Graph API client) -> Microsoft Graph REST API
                          auth.py (MSAL device code flow, token cache)
```

- `src/teams_mcp/auth.py` - AuthManager: device code flow, token cache at `~/.teams-mcp/token_cache.json`. Default scope `https://graph.microsoft.com/.default` (semantics in the Scopes section)
- `src/teams_mcp/graph.py` - GraphClient: async httpx wrapper with `_get`, `_post`, `_post_no_content`, `_patch`, `_delete` helpers. All raise `GraphApiError` on failure (parses Graph API JSON error body). `GRAPH_BASE` constant points at v1.0 — all endpoints currently used are GA.
- `src/teams_mcp/server.py` - MCP tools + response post-processing (Adaptive Card text extraction, forwarded/quoted message parsing, hosted-content inlining, mention parsing). Global `auth`/`graph` initialized lazily. `_require_auth()` guard on every tool.

## Tools

### Auth
- `login` / `complete_login` - device code flow (two-phase)

### Read
- `list_teams`, `list_channels`, `list_chats`
- `list_channel_messages`, `list_chat_messages`, `list_thread_replies`
- `list_chat_messages` pages through history: `before`/`after` (ISO 8601, exclusive), `limit` up to `MAX_LIST_LIMIT` across several Graph pages, and it returns `{"messages": [...], "next_before": ...}` - pass `next_before` as `before` for the next older page
- `list_channel_messages` and `list_thread_replies` also follow `@odata.nextLink` up to `MAX_LIST_LIMIT` but have no cursor: Graph takes only `$top` there, and channel messages are sorted by the latest activity in the whole thread, so a timestamp cursor cannot work
- `list_team_members`, `list_channel_members`, `list_chat_members`
- `list_team_tags` (tag ids for @tag mentions; needs `TeamworkTag.Read`)
- `list_pinned_messages`
- `get_message` - one message by Teams link (`/l/message/<chat or channel id>/<message id>`; a channel link carries `groupId` = team id and `parentMessageId` = thread root) or by ids; returns where it lives plus the formatted message
- `get_user_presence`, `get_user` (search by name/email)
- `search_messages` (Microsoft Search API, full-text)
- `download_attachment` (inline images via hostedContents, magic-byte format detection — png/jpg/gif/webp, returns temp file path)

### Write
- `send_channel_message`, `send_chat_message`, `reply_to_channel_message` - all support `mentions` (list | JSON string)
- `send_chat_message` also accepts `reply_to` (message_id) for quoted replies - uses `replyWithQuote` endpoint
- `create_chat` (1:1), `create_group_chat`
- `set_reaction`, `unset_reaction`
- `pin_message`, `unpin_message`
- `delete_message` (soft delete), `update_message` (edit)
- `mark_chat_read`, `mark_chat_unread`

### Tool patterns
- Channel tools need `team_id` + `channel_id`
- Chat tools need `chat_id`
- Dual-context tools (reactions, delete, update, download_attachment, get_message) accept either `chat_id` OR `team_id + channel_id`
- Send tools accept `mentions` as list or JSON string: `[{"user_id": "...", "name": "..."}]` for users, `[{"tag_id": "...", "name": "..."}]` for team tags (channel messages only; ids from `list_team_tags`); use `@Name` in content (longest-name-first replacement to avoid partial matches). Malformed `mentions` (invalid JSON, not a list, an entry without `name` or without exactly one of `user_id`/`tag_id`) fail with a `ToolError` before anything is sent; an entry whose `@Name` is missing from content is still dropped silently

## Error handling

- mcp 2.x passes only `ToolError` text to the client (an `is_error` result prefixed `Error executing tool <name>: `). Any other exception is a crash: the client sees only `Error executing tool <name>` and the traceback goes to the server log.
- Tools are registered with `@_tool`, not `@mcp.tool()`: it re-raises the anticipated failures (`AuthError`, `GraphApiError`, `ValueError`) as `ToolError` and leaves everything else a crash, so real bugs keep their traceback
- Argument checks inside a tool raise `ToolError` directly; `graph.py` and `auth.py` raise their own exceptions and stay unaware of MCP
- `GraphApiError(status_code, code, message)` - raised by all graph helpers, contains parsed Graph API error JSON
- 403 errors surface the Graph API message directly (e.g. "Insufficient privileges to complete the operation") - tools work with whatever scopes the user has, missing scopes produce clear errors
- `AuthError` (`auth.py`) - sign-in missing ("Not authenticated...", from `_require_auth()` or a token provider that returns nothing) or failed (device flow errors from `AuthManager`)

## Scopes (delegated)

Default scope literal: `https://graph.microsoft.com/.default` — Azure issues a token covering all permissions configured and consented on the app registration. Override via `TEAMS_MCP_SCOPES` env var (comma-separated) to restrict. Tools that request a missing scope return 403 with Graph API's error message.

Typical app registration permissions (any subset works — `.default` picks them up):

```
User.Read, User.ReadBasic.All, Team.ReadBasic.All, TeamMember.Read.All,
Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All,
ChannelMessage.Send, ChannelMessage.ReadWrite, Chat.Read, Chat.ReadWrite,
Presence.Read.All, TeamworkTag.Read
```

Scopes requiring admin consent: `TeamMember.Read.All`, `ChannelMember.Read.All`, `ChannelMessage.Read.All`, `ChannelMessage.ReadWrite`.

## Adding a new tool

1. Add graph method in `graph.py` (use existing `_get`/`_post`/`_post_no_content`/`_patch`/`_delete`)
2. Add MCP tool in `server.py`: register it with `@_tool` and follow the `_init_if_needed() -> _require_auth() -> call graph -> json.dumps` pattern
3. If a new scope is needed: no code change required - `.default` picks it up once consented (see the Scopes section).

## Known Quirks

- Chat replies (`send_chat_message` with `reply_to`) use `/chats/{id}/messages/replyWithQuote` - the `/replies` sub-collection only exists on channel messages
- Some scopes require admin consent (see the Scopes section) - tools return 403 if not consented
- `delete_message` is soft delete - message shows "This message has been deleted" to other users
- Token cache at `~/.teams-mcp/token_cache.json` - delete file to force re-auth
- Hosted content inside a channel REPLY is served only under `/messages/{parent}/replies/{reply}/hostedContents` - pass `parent_message_id` to `download_attachment`, the parent-form URL 404s for reply ids
- Chat messages (`GET /chats/{id}/messages`): `$top` max 50; the default order is `lastModifiedDateTime desc`, so a reaction or edit moves an old message to the top; `$filter` on `createdDateTime` supports only `lt` and is silently ignored unless `$orderby` uses the same property. `list_chat_messages` therefore always sends `$orderby=createdDateTime desc` and applies `after` itself, stopping the paging at the first older message
- Graph transport flakes intermittently (ReadTimeout etc.); httpx exceptions often have empty `str()` - the client wraps them in `GraphApiError` with the exception type so tool errors are never blank
- Tag mentions: `GET /teams/{id}/tags` returns TEAM-level tags - valid as mention ids in standard channels only. Shared channels (`membershipType` reads as `unknownFutureValue` on v1.0) use channel-scoped tags with no Graph API at all; a team-level id posted there renders a phantom tag (0 members). Channel-tag ids appear only inside `mentions[]` of stored messages - harvest from a manual mention. Note: sending `mentioned.tag` is formally documented only on Graph beta; v1.0 accepts it (open type) and round-trips it on reads

## Conventions

- Tool docstrings in English (they become user-facing MCP tool descriptions)
- `_format_message(msg)` / `_format_member(member)` for consistent output shapes
- HTML for outgoing messages (`_to_html` / `_build_message_body`), stripped for display (`_strip_html`)
- `_to_html` auto-links `http(s)://` URLs and escapes `& < >` - don't pre-HTML-encode content
- `_format_message` extracts text from: body HTML, Adaptive Card attachments (`_extract_adaptive_card_text`), forwarded messages (`_extract_forwarded_text` handles both `forwardedMessageReference` and `messageReference`)
- Inline images: `_format_hosted_contents` parses `/hostedContents/<id>/$value` from `<img src>` - use `download_attachment` tool to fetch bytes

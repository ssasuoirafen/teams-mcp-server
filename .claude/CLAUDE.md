# Teams MCP Server

MCP server for Microsoft Teams via Microsoft Graph API. Delegated auth (device code flow), no client secret.

## Stack

Python 3.12+, FastMCP, httpx, msal. Tests: pytest + pytest-asyncio.

## Commands

```bash
uv run pytest -v              # run tests
uv run teams-mcp              # run server (needs TEAMS_MCP_TENANT_ID, TEAMS_MCP_CLIENT_ID)
```

## Architecture

Three-file pattern, all logic flows through the same pipeline:

```
server.py (MCP tools) -> graph.py (Graph API client) -> Microsoft Graph REST API
                          auth.py (MSAL device code flow, token cache)
```

- `src/teams_mcp/auth.py` - AuthManager: device code flow, token cache at `~/.teams-mcp/token_cache.json`, DEFAULT_SCOPES list
- `src/teams_mcp/graph.py` - GraphClient: async httpx wrapper with `_get`, `_post`, `_post_no_content`, `_patch`, `_delete` helpers. All raise `GraphApiError` on failure (parses Graph API JSON error body). `GRAPH_BASE` (v1.0) and `GRAPH_BETA` constants.
- `src/teams_mcp/server.py` - 28 MCP tools. Global `auth`/`graph` initialized lazily. `_require_auth()` guard on every tool.
- `tests/test_graph.py` - Graph client tests with `mock_transport` / `make_client` helpers. Pattern: mock httpx transport -> call graph method -> assert URL/body/response.
- `tests/test_auth.py` - Auth tests with mocked MSAL.

## Tools (28)

### Auth
- `login` / `complete_login` - device code flow (two-phase)

### Read
- `list_teams`, `list_channels`, `list_chats`
- `list_channel_messages`, `list_chat_messages`, `list_thread_replies`
- `list_team_members`, `list_channel_members`, `list_chat_members`
- `list_pinned_messages`
- `get_user_presence`, `get_user` (search by name/email)
- `search_messages` (beta API, full-text)

### Write
- `send_channel_message`, `send_chat_message`, `reply_to_channel_message` - all support `mentions` param (JSON string)
- `create_chat` (1:1), `create_group_chat`
- `set_reaction`, `unset_reaction`
- `pin_message`, `unpin_message`
- `delete_message` (soft delete), `update_message` (edit)
- `mark_chat_read`, `mark_chat_unread`

### Tool patterns
- Channel tools need `team_id` + `channel_id`
- Chat tools need `chat_id`
- Dual-context tools (reactions, delete, update) accept either `chat_id` OR `team_id + channel_id`
- Send tools accept optional `mentions` as JSON string: `[{"user_id": "...", "name": "..."}]`, use `@Name` in content

## Error handling

- `GraphApiError(status_code, code, message)` - raised by all graph helpers, contains parsed Graph API error JSON
- 403 errors surface the Graph API message directly (e.g. "Insufficient privileges to complete the operation") - tools work with whatever scopes the user has, missing scopes produce clear errors
- `RuntimeError("Not authenticated...")` - when no token available

## Scopes (delegated)

All 12 scopes requested at login. Tools that need a missing scope return 403 with Graph API's error message.

```
User.Read, User.ReadBasic.All, Team.ReadBasic.All, TeamMember.Read.All,
Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All,
ChannelMessage.Send, ChannelMessage.ReadWrite, Chat.Read, Chat.ReadWrite,
Presence.Read.All
```

## Adding a new tool

1. Add graph method in `graph.py` (use existing `_get`/`_post`/`_post_no_content`/`_patch`/`_delete`)
2. Add MCP tool in `server.py` (follow `_init_if_needed() -> _require_auth() -> call graph -> json.dumps` pattern)
3. Add test in `tests/test_graph.py` (use `mock_transport` or custom handler with `captured` list)
4. If new scope needed: add to `DEFAULT_SCOPES` in `auth.py`, update `.env.example`

## Conventions

- Code in English, tool docstrings in English
- `_format_message(msg)` / `_format_member(member)` for consistent output shapes
- HTML for outgoing messages (`_to_html` / `_build_message_body`), stripped for display (`_strip_html`)
- `asyncio_mode = "strict"` in pytest config

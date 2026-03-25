# Teams MCP Server — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local stdio MCP server that exposes Microsoft Teams (channels + chats) to Claude Code via Microsoft Graph API with delegated permissions (device code flow).

**Architecture:** Three modules — `auth.py` (MSAL device code flow + persistent token cache), `graph.py` (async httpx wrapper for Graph API), `server.py` (FastMCP with ~9 tools). Server runs as stdio process, configured in Claude Code `settings.local.json`.

**Tech Stack:** Python 3.12+, `mcp` SDK (FastMCP), `msal`, `httpx`, `uv`

---

## File Structure

```
teams-mcp-server/
├── pyproject.toml              # Project config, dependencies, entry point
├── .python-version             # Python version for uv
├── .gitignore                  # Python + token cache
├── .env.example                # Required env vars template
├── src/
│   └── teams_mcp/
│       ├── __init__.py         # Package init, version
│       ├── auth.py             # MSAL device code flow, token cache, silent refresh
│       ├── graph.py            # Microsoft Graph API client (httpx)
│       └── server.py           # FastMCP server, tools, entry point
└── tests/
    ├── __init__.py
    ├── test_auth.py            # Auth module tests (mocked MSAL)
    └── test_graph.py           # Graph client tests (mocked httpx)
```

## Tools Inventory

| Tool | Endpoint | Annotation | Description |
|------|----------|------------|-------------|
| `login` | MSAL device code flow | — | Authenticate with Microsoft. Prints device code + URL to follow. |
| `list_teams` | `GET /me/joinedTeams` | readOnly, openWorld | List all Teams the user has joined. |
| `list_channels` | `GET /teams/{id}/channels` | readOnly, openWorld | List channels in a team. Use list_teams first to get team_id. |
| `list_chats` | `GET /me/chats` | readOnly, openWorld | List recent 1:1 and group chats with participant names. |
| `list_channel_messages` | `GET /teams/{id}/channels/{id}/messages` | readOnly, openWorld | List recent messages in a channel. Returns sender, timestamp, content. |
| `list_chat_messages` | `GET /chats/{id}/messages` | readOnly, openWorld | List recent messages in a chat. Use list_chats first to get chat_id. |
| `send_channel_message` | `POST /teams/{id}/channels/{id}/messages` | openWorld | Send a message to a Teams channel. |
| `send_chat_message` | `POST /chats/{id}/messages` | openWorld | Send a message to a 1:1 or group chat. |
| `reply_to_channel_message` | `POST /teams/{id}/channels/{id}/messages/{id}/replies` | openWorld | Reply to a specific message in a channel thread. |

## Graph API Scopes (Delegated)

Minimal set: `User.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.ReadWrite`, `ChatMessage.Read`

---

## Task 1: Project Scaffold

**Files:**
- Create: `pyproject.toml`
- Create: `.python-version`
- Create: `.gitignore`
- Create: `.env.example`
- Create: `src/teams_mcp/__init__.py`

- [ ] **Step 1: Create `pyproject.toml`**

```toml
[project]
name = "teams-mcp-server"
version = "0.1.0"
description = "MCP server for Microsoft Teams via Graph API"
requires-python = ">=3.12"
dependencies = [
    "mcp[cli]>=1.9.0",
    "msal>=1.31.0",
    "httpx>=0.28.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "pytest-asyncio>=0.25.0",
]

[project.scripts]
teams-mcp = "teams_mcp.server:main"

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/teams_mcp"]
```

- [ ] **Step 2: Create `.python-version`, `.gitignore`, `.env.example`, `__init__.py`**

`.python-version`: `3.12`

`.gitignore`: standard Python + `.token_cache.json` + `.env`

`.env.example`:
```
TEAMS_MCP_TENANT_ID=your-tenant-id
TEAMS_MCP_CLIENT_ID=your-client-id
```

`src/teams_mcp/__init__.py`: empty file

- [ ] **Step 3: Initialize git and install dependencies**

```bash
cd C:/Users/ssasuoirafen/Projects/teams-mcp-server
git init
uv sync
```

- [ ] **Step 4: Commit scaffold**

```bash
git add -A
git commit -m "chore: project scaffold"
```

---

## Task 2: Auth Module

**Files:**
- Create: `src/teams_mcp/auth.py`
- Create: `tests/__init__.py`
- Create: `tests/test_auth.py`

- [ ] **Step 1: Write test for auth manager initialization**

```python
# tests/test_auth.py
import pytest
from unittest.mock import patch, MagicMock
from teams_mcp.auth import AuthManager

def test_auth_manager_init():
    mgr = AuthManager(tenant_id="test-tenant", client_id="test-client")
    assert mgr.tenant_id == "test-tenant"
    assert mgr.client_id == "test-client"
    assert mgr.scopes == AuthManager.DEFAULT_SCOPES

def test_auth_manager_not_authenticated_initially():
    mgr = AuthManager(tenant_id="t", client_id="c")
    assert mgr.get_token() is None
```

- [ ] **Step 2: Run tests — expect FAIL (module doesn't exist)**

```bash
uv run pytest tests/test_auth.py -v
```

- [ ] **Step 3: Implement `auth.py`**

Core logic:
- `AuthManager` class with `tenant_id`, `client_id`, `scopes`
- `msal.PublicClientApplication` with persistent `SerializableTokenCache`
- Token cache saved to `~/.teams-mcp/token_cache.json`
- `get_token()` — tries `acquire_token_silent()` first, returns access_token or None
- `login()` — initiates device code flow, returns user_code + verification_uri, polls for completion
- `is_authenticated()` — checks if valid token exists

```python
# src/teams_mcp/auth.py
import json
import os
from pathlib import Path

import msal


class AuthManager:
    DEFAULT_SCOPES = [
        "User.Read",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
        "ChannelMessage.Send",
        "Chat.ReadWrite",
        "ChatMessage.Read",
    ]

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        scopes: list[str] | None = None,
        cache_dir: str | None = None,
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.scopes = scopes or self.DEFAULT_SCOPES
        self._cache_dir = Path(cache_dir or os.path.expanduser("~/.teams-mcp"))
        self._cache_dir.mkdir(parents=True, exist_ok=True)
        self._cache_path = self._cache_dir / "token_cache.json"
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        self._app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self._cache,
        )

    def _load_cache(self):
        if self._cache_path.exists():
            self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))

    def _save_cache(self):
        if self._cache.has_state_changed:
            self._cache_path.write_text(
                self._cache.serialize(), encoding="utf-8"
            )

    def get_token(self) -> str | None:
        accounts = self._app.get_accounts()
        if not accounts:
            return None
        result = self._app.acquire_token_silent(
            scopes=self.scopes, account=accounts[0]
        )
        self._save_cache()
        if result and "access_token" in result:
            return result["access_token"]
        return None

    def login(self) -> dict:
        flow = self._app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow failed: {flow.get('error_description', 'unknown error')}")
        return flow

    def complete_login(self, flow: dict) -> dict:
        result = self._app.acquire_token_by_device_flow(flow)
        self._save_cache()
        if "access_token" in result:
            return {"status": "ok", "account": result.get("id_token_claims", {}).get("preferred_username", "unknown")}
        raise RuntimeError(result.get("error_description", "Authentication failed"))

    def is_authenticated(self) -> bool:
        return self.get_token() is not None
```

- [ ] **Step 4: Run tests — expect PASS**

```bash
uv run pytest tests/test_auth.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/teams_mcp/auth.py tests/
git commit -m "feat: auth module with MSAL device code flow"
```

---

## Task 3: Graph API Client

**Files:**
- Create: `src/teams_mcp/graph.py`
- Create: `tests/test_graph.py`

- [ ] **Step 1: Write test for graph client**

```python
# tests/test_graph.py
import pytest
import httpx
import pytest_asyncio
from unittest.mock import AsyncMock, patch
from teams_mcp.graph import GraphClient

@pytest.fixture
def client():
    return GraphClient(token_provider=lambda: "fake-token")

@pytest.mark.asyncio
async def test_list_teams(client):
    mock_response = {
        "value": [
            {"id": "team-1", "displayName": "Engineering"},
            {"id": "team-2", "displayName": "Data"},
        ]
    }
    with patch.object(client._http, "get", new_callable=AsyncMock) as mock_get:
        mock_get.return_value = httpx.Response(200, json=mock_response)
        teams = await client.list_teams()
        assert len(teams) == 2
        assert teams[0]["displayName"] == "Engineering"
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_graph.py -v
```

- [ ] **Step 3: Implement `graph.py`**

Core: async httpx client with auth header injection, all Graph API endpoints.

```python
# src/teams_mcp/graph.py
from typing import Any, Callable

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphClient:
    def __init__(self, token_provider: Callable[[], str | None]):
        self._token_provider = token_provider
        self._http = httpx.AsyncClient(
            base_url=GRAPH_BASE,
            timeout=30.0,
        )

    def _headers(self) -> dict[str, str]:
        token = self._token_provider()
        if not token:
            raise RuntimeError("Not authenticated. Call the login tool first.")
        return {"Authorization": f"Bearer {token}"}

    async def _get(self, path: str, params: dict | None = None) -> dict[str, Any]:
        resp = await self._http.get(path, headers=self._headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    async def _post(self, path: str, json_body: dict) -> dict[str, Any]:
        resp = await self._http.post(path, headers=self._headers(), json=json_body)
        resp.raise_for_status()
        return resp.json()

    # --- Teams ---

    async def list_teams(self) -> list[dict]:
        data = await self._get("/me/joinedTeams", params={"$select": "id,displayName,description"})
        return data.get("value", [])

    async def list_channels(self, team_id: str) -> list[dict]:
        data = await self._get(
            f"/teams/{team_id}/channels",
            params={"$select": "id,displayName,description,membershipType"},
        )
        return data.get("value", [])

    # --- Chats ---

    async def list_chats(self, limit: int = 20) -> list[dict]:
        data = await self._get(
            "/me/chats",
            params={
                "$top": limit,
                "$expand": "members($select=displayName,email)",
                "$select": "id,topic,chatType,lastUpdatedDateTime",
                "$orderby": "lastUpdatedDateTime desc",
            },
        )
        return data.get("value", [])

    # --- Messages ---

    async def list_channel_messages(
        self, team_id: str, channel_id: str, limit: int = 20
    ) -> list[dict]:
        data = await self._get(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            params={"$top": limit},
        )
        return data.get("value", [])

    async def list_chat_messages(self, chat_id: str, limit: int = 20) -> list[dict]:
        data = await self._get(
            f"/chats/{chat_id}/messages",
            params={"$top": limit},
        )
        return data.get("value", [])

    # --- Send ---

    async def send_channel_message(
        self, team_id: str, channel_id: str, content: str, content_type: str = "text"
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            {"body": {"content": content, "contentType": content_type}},
        )

    async def send_chat_message(
        self, chat_id: str, content: str, content_type: str = "text"
    ) -> dict:
        return await self._post(
            f"/chats/{chat_id}/messages",
            {"body": {"content": content, "contentType": content_type}},
        )

    async def reply_to_channel_message(
        self,
        team_id: str,
        channel_id: str,
        message_id: str,
        content: str,
        content_type: str = "text",
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            {"body": {"content": content, "contentType": content_type}},
        )

    async def close(self):
        await self._http.aclose()
```

- [ ] **Step 4: Run tests — expect PASS**

```bash
uv run pytest tests/test_graph.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/teams_mcp/graph.py tests/test_graph.py
git commit -m "feat: Graph API client with Teams endpoints"
```

---

## Task 4: MCP Server + Tools

**Files:**
- Create: `src/teams_mcp/server.py`

This is the main file — FastMCP server with all 9 tools, server instructions, and entry point.

- [ ] **Step 1: Implement `server.py`**

Key design decisions:
- Server `instructions` hint about auth flow
- All read tools: `readOnlyHint=True`, `openWorldHint=True`
- All write tools: `openWorldHint=True`
- `login` tool: two-phase — first call returns device code, user authenticates in browser, tool polls for completion
- Error messages suggest which tool to call next
- Message content stripped from HTML tags for readability
- Large responses truncated with count hint

```python
# src/teams_mcp/server.py
import json
import os
import re
import sys
from typing import Any

from mcp.server.fastmcp import FastMCP

from teams_mcp.auth import AuthManager
from teams_mcp.graph import GraphClient

mcp = FastMCP(
    "teams-mcp",
    instructions=(
        "Microsoft Teams MCP server. Before using any tool, call login to authenticate. "
        "For channels: list_teams -> list_channels -> list_channel_messages or send_channel_message. "
        "For chats: list_chats -> list_chat_messages or send_chat_message."
    ),
)

auth: AuthManager | None = None
graph: GraphClient | None = None


def _init():
    global auth, graph
    tenant_id = os.environ.get("TEAMS_MCP_TENANT_ID")
    client_id = os.environ.get("TEAMS_MCP_CLIENT_ID")
    if not tenant_id or not client_id:
        raise RuntimeError(
            "TEAMS_MCP_TENANT_ID and TEAMS_MCP_CLIENT_ID env vars are required"
        )
    auth = AuthManager(tenant_id=tenant_id, client_id=client_id)
    graph = GraphClient(token_provider=auth.get_token)


def _require_auth() -> GraphClient:
    if not auth or not auth.is_authenticated():
        raise RuntimeError("Not authenticated. Call the login tool first.")
    return graph


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "")


def _format_message(msg: dict) -> dict[str, Any]:
    sender = "unknown"
    if msg.get("from") and msg["from"].get("user"):
        sender = msg["from"]["user"].get("displayName", "unknown")
    return {
        "id": msg.get("id"),
        "sender": sender,
        "created": msg.get("createdDateTime"),
        "content": _strip_html(msg.get("body", {}).get("content", "")),
    }


# --- Auth ---

@mcp.tool(
    annotations={"openWorldHint": True},
)
def login() -> str:
    """Authenticate with Microsoft Teams via device code flow.
    Opens a browser URL where you enter a code. Required before any other tool.
    If already authenticated, returns current account info."""
    _init_if_needed()
    if auth.is_authenticated():
        accounts = auth._app.get_accounts()
        name = accounts[0].get("username", "unknown") if accounts else "unknown"
        return json.dumps({"status": "already_authenticated", "account": name})
    flow = auth.login()
    # Print device code instructions to stderr so user sees them
    msg = flow.get("message", "")
    print(msg, file=sys.stderr, flush=True)
    # Block until user completes auth in browser
    result = auth.complete_login(flow)
    return json.dumps(result)


def _init_if_needed():
    if auth is None:
        _init()


# --- Teams ---

@mcp.tool(
    annotations={"readOnlyHint": True, "openWorldHint": True},
)
async def list_teams() -> str:
    """List all Microsoft Teams the user has joined.
    Returns team IDs and names. Use a team_id with list_channels to see its channels."""
    _init_if_needed()
    client = _require_auth()
    teams = await client.list_teams()
    result = [
        {"id": t["id"], "name": t.get("displayName", ""), "description": t.get("description", "")}
        for t in teams
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool(
    annotations={"readOnlyHint": True, "openWorldHint": True},
)
async def list_channels(team_id: str) -> str:
    """List channels in a Microsoft Teams team.
    Use list_teams first to get the team_id. Returns channel IDs and names.
    Use a channel_id with list_channel_messages or send_channel_message."""
    _init_if_needed()
    client = _require_auth()
    channels = await client.list_channels(team_id)
    result = [
        {
            "id": c["id"],
            "name": c.get("displayName", ""),
            "type": c.get("membershipType", ""),
            "description": c.get("description", ""),
        }
        for c in channels
    ]
    return json.dumps(result, ensure_ascii=False, indent=2)


# --- Chats ---

@mcp.tool(
    annotations={"readOnlyHint": True, "openWorldHint": True},
)
async def list_chats(limit: int = 20) -> str:
    """List recent 1:1 and group chats.
    Returns chat IDs, participants, and last activity time.
    Use a chat_id with list_chat_messages or send_chat_message.
    Does NOT include channel conversations — use list_teams + list_channels for those."""
    _init_if_needed()
    client = _require_auth()
    chats = await client.list_chats(limit=limit)
    result = []
    for c in chats:
        members = []
        for m in c.get("members", []):
            name = m.get("displayName", m.get("email", "unknown"))
            if name:
                members.append(name)
        result.append({
            "id": c["id"],
            "topic": c.get("topic") or ", ".join(members[:3]),
            "type": c.get("chatType", ""),
            "last_updated": c.get("lastUpdatedDateTime", ""),
            "members": members,
        })
    return json.dumps(result, ensure_ascii=False, indent=2)


# --- Messages ---

@mcp.tool(
    annotations={"readOnlyHint": True, "openWorldHint": True},
)
async def list_channel_messages(team_id: str, channel_id: str, limit: int = 20) -> str:
    """List recent messages in a Teams channel.
    Use list_teams + list_channels to get IDs.
    Returns sender, timestamp, and plain text content (HTML stripped).
    Use reply_to_channel_message to reply to a specific message."""
    _init_if_needed()
    client = _require_auth()
    messages = await client.list_channel_messages(team_id, channel_id, limit=limit)
    result = [_format_message(m) for m in messages if m.get("messageType") == "message"]
    if len(result) > limit:
        result = result[:limit]
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool(
    annotations={"readOnlyHint": True, "openWorldHint": True},
)
async def list_chat_messages(chat_id: str, limit: int = 20) -> str:
    """List recent messages in a 1:1 or group chat.
    Use list_chats to get the chat_id.
    Returns sender, timestamp, and plain text content (HTML stripped)."""
    _init_if_needed()
    client = _require_auth()
    messages = await client.list_chat_messages(chat_id, limit=limit)
    result = [_format_message(m) for m in messages if m.get("messageType") == "message"]
    if len(result) > limit:
        result = result[:limit]
    return json.dumps(result, ensure_ascii=False, indent=2)


# --- Send ---

@mcp.tool(
    annotations={"openWorldHint": True},
)
async def send_channel_message(team_id: str, channel_id: str, content: str) -> str:
    """Send a new message to a Teams channel.
    Use list_teams + list_channels to get IDs.
    For replies to existing messages, use reply_to_channel_message instead."""
    _init_if_needed()
    client = _require_auth()
    msg = await client.send_channel_message(team_id, channel_id, content)
    return json.dumps({"status": "sent", "message_id": msg.get("id")})


@mcp.tool(
    annotations={"openWorldHint": True},
)
async def send_chat_message(chat_id: str, content: str) -> str:
    """Send a message to a 1:1 or group chat.
    Use list_chats to get the chat_id.
    Chat messages don't support threaded replies — just send a new message."""
    _init_if_needed()
    client = _require_auth()
    msg = await client.send_chat_message(chat_id, content)
    return json.dumps({"status": "sent", "message_id": msg.get("id")})


@mcp.tool(
    annotations={"openWorldHint": True},
)
async def reply_to_channel_message(
    team_id: str, channel_id: str, message_id: str, content: str
) -> str:
    """Reply to a specific message in a Teams channel (creates a thread reply).
    Use list_channel_messages to get the message_id.
    For new top-level messages, use send_channel_message instead."""
    _init_if_needed()
    client = _require_auth()
    msg = await client.reply_to_channel_message(team_id, channel_id, message_id, content)
    return json.dumps({"status": "sent", "reply_id": msg.get("id")})


def main():
    _init()
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify it compiles (no syntax errors)**

```bash
uv run python -c "from teams_mcp.server import mcp; print('OK')"
```

- [ ] **Step 3: Commit**

```bash
git add src/teams_mcp/server.py
git commit -m "feat: MCP server with 9 Teams tools"
```

---

## Task 5: Integration + Claude Code Config

**Files:**
- No new files in project

- [ ] **Step 1: Add MCP server to Claude Code settings**

Add to `C:\Users\ssasuoirafen\.claude\settings.local.json`:

```json
{
  "mcpServers": {
    "teams": {
      "command": "uv",
      "args": ["run", "--directory", "C:/Users/ssasuoirafen/Projects/teams-mcp-server", "teams-mcp"],
      "env": {
        "TEAMS_MCP_TENANT_ID": "<from ITO-1052 App Registration>",
        "TEAMS_MCP_CLIENT_ID": "<from ITO-1052 App Registration>"
      }
    }
  }
}
```

- [ ] **Step 2: Restart Claude Code, verify server appears in `/mcp`**

- [ ] **Step 3: Test login flow**

Call `login` tool — should print device code to stderr, user opens URL, enters code, auth completes.

- [ ] **Step 4: Test read tools**

Call `list_teams` → pick a team → `list_channels` → `list_channel_messages`

- [ ] **Step 5: Test send (in a test channel)**

Call `send_channel_message` to a safe test channel.

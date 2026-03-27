# teams-mcp-server

MCP server for Microsoft Teams via Microsoft Graph API. Supports channels, chats, search, reactions, pins, and more.

## Tools

| Tool | Description |
|------|-------------|
| `login` / `complete_login` | Authenticate via device code flow |
| `list_teams` | List joined teams |
| `list_channels` | List channels in a team |
| `list_channel_messages` | List messages in a channel |
| `list_thread_replies` | List replies in a thread |
| `send_channel_message` | Send a message to a channel |
| `reply_to_channel_message` | Reply to a thread |
| `list_chats` | List chats |
| `list_chat_messages` | List messages in a chat |
| `send_chat_message` | Send a chat message |
| `search_messages` | Full-text search across chats and channels |
| `get_user` | Find a user |
| `get_user_presence` | Get user's online status |
| `create_chat` / `create_group_chat` | Create 1:1 or group chats |
| `update_message` / `delete_message` | Edit or delete messages |
| `set_reaction` / `unset_reaction` | Manage reactions |
| `pin_message` / `unpin_message` | Manage pinned messages |
| `mark_chat_read` / `mark_chat_unread` | Mark chat read status |
| `list_team_members` / `list_channel_members` / `list_chat_members` | List members |
| `list_pinned_messages` | List pinned messages |

Adaptive Card attachments (from bots) are automatically extracted as plain text.

## Configuration

Requires an Azure AD app registration with delegated permissions for Microsoft Graph.

### Claude Code (`.mcp.json`)

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/ssasuoirafen/teams-mcp-server", "teams-mcp"],
      "env": {
        "TEAMS_MCP_TENANT_ID": "your-tenant-id",
        "TEAMS_MCP_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

| Variable | Description |
|----------|-------------|
| `TEAMS_MCP_TENANT_ID` | Azure AD tenant ID |
| `TEAMS_MCP_CLIENT_ID` | App registration client ID |
| `TEAMS_MCP_SCOPES` | (Optional) Comma-separated scopes |

### Authentication

On first use, call the `login` tool. It returns a device code and URL. Open the URL in a browser, enter the code, then call `complete_login`. Tokens are cached in `~/.teams-mcp/token_cache.json`.

## Development

```bash
git clone https://github.com/ssasuoirafen/teams-mcp-server.git
cd teams-mcp-server
uv sync --extra dev
uv run pytest tests/ -v
```

## License

MIT

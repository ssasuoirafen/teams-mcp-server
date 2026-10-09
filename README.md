# teams-mcp-server

MCP server for Microsoft Teams via Microsoft Graph API.

## Tools

| Tool | Description |
|------|-------------|
| `list_teams` | List joined teams |
| `list_channels` | List channels in a team |
| `list_channel_messages` | List messages in a channel |
| `list_thread_replies` | List replies in a thread |
| `send_channel_message` | Send a message to a channel |
| `reply_to_channel_message` | Reply to a thread |
| `list_chats` | List chats |
| `list_chat_messages` | List messages in a chat, newest first; pages back through history with `before`/`after` and a `next_before` cursor |
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
| `list_team_tags` | List team tags (for @tag mentions in send/reply tools) |
| `list_pinned_messages` | List pinned messages |
| `get_message` | Open one chat or channel message by its Teams link or by ids |
| `download_attachment` | Download an inline image (hostedContents) to a temp file |

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

### Required permissions

The app registration needs delegated Microsoft Graph permissions. By default the server requests the `.default` scope, so Azure issues a token covering every permission consented on the app registration. A typical set (any subset works - `.default` picks up whatever is consented):

```
User.Read, User.ReadBasic.All, Team.ReadBasic.All, TeamMember.Read.All,
Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All,
ChannelMessage.Send, ChannelMessage.ReadWrite, Chat.Read, Chat.ReadWrite,
Presence.Read.All, TeamworkTag.Read
```

These require tenant admin consent: `TeamMember.Read.All`, `ChannelMember.Read.All`, `ChannelMessage.Read.All`, `ChannelMessage.ReadWrite`. Set `TEAMS_MCP_SCOPES` (comma-separated) to request a narrower set; a tool that hits a missing permission returns the Graph API 403 error message.

### Authentication

Sign in once from a terminal, with the same `TEAMS_MCP_TENANT_ID` and `TEAMS_MCP_CLIENT_ID` the server gets. The command is the one from `.mcp.json` with `login` appended:

```bash
uvx --from git+https://github.com/ssasuoirafen/teams-mcp-server teams-mcp login
```

It prints a URL and a device code, copies the code to the clipboard (`pbcopy`, `clip`, or `wl-copy`/`xclip`/`xsel` when available) and opens the URL in the browser on a local desktop (not over SSH). Paste the code and sign in; the command then prints the account. The token is cached in `~/.teams-mcp/token_cache.json` and renewed silently. A running server picks up the new sign-in on its next tool call, with no restart.

## Development

```bash
git clone https://github.com/ssasuoirafen/teams-mcp-server.git
cd teams-mcp-server
uv sync
uv run teams-mcp
```

## License

MIT

# MSTeams MCP Server

An MCP server for Microsoft Teams via the Microsoft Graph API. Manage presence, send messages, read channels and chats, and search users — all from Claude Code or any MCP-compatible client.

## Features

- Get and set your Teams presence / availability
- List teams, channels, and channel messages
- Send and reply to channel messages
- List and send direct chats
- List team members and search users
- Read shift schedules

## Prerequisites

- Node.js 18+
- An Azure App Registration with Microsoft Graph permissions
- An Outlook.com or Microsoft 365 account

## Setup

### 1. Register an Azure App

1. Go to [portal.azure.com](https://portal.azure.com)
2. **Azure Active Directory** → **App registrations** → **New registration**
3. Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, add:
   - `Chat.Read`, `Chat.ReadWrite`
   - `ChannelMessage.Read.All`, `ChannelMessage.Send`
   - `Presence.Read`, `Presence.Read.All`, `Presence.ReadWrite`
   - `Team.ReadBasic.All`, `TeamMember.Read.All`
   - `User.Read`, `User.ReadBasic.All`
4. Copy your **Application (client) ID** and **Directory (tenant) ID**

### 2. Install

```bash
npx -y @web4w3/msteams-mcp-server
```

No cloning or build step needed. The package is pre-compiled and available on npm as [`@web4w3/msteams-mcp-server`](https://www.npmjs.com/package/@web4w3/msteams-mcp-server).

### 3. Add to Claude config

```json
{
  "mcpServers": {
    "msteams": {
      "command": "npx",
      "args": ["-y", "@web4w3/msteams-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

On first run, a device code flow opens in your browser for OAuth consent. The token is cached locally for subsequent runs.

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | Yes | Directory (tenant) ID from your Azure App Registration |
| `AZURE_CLIENT_ID` | Yes | Application (client) ID from your Azure App Registration |
| `AZURE_CLIENT_SECRET` | No | Client secret (required for `application` auth mode) |
| `MS_AUTH_MODE` | No | `delegated` (default, device code flow) or `application` |
| `MS_GRAPH_ACCESS_TOKEN` | No | Pre-obtained access token (overrides all other auth) |

## Available Tools

| Tool | Description |
|------|-------------|
| `teams_get_my_presence` | Get your current Teams presence/availability |
| `teams_get_user_presence` | Get another user's presence by ID or email |
| `teams_get_presences_batch` | Get presence for multiple users at once |
| `teams_set_my_presence` | Set your own presence status |
| `teams_clear_my_presence` | Clear your presence override |
| `teams_set_preferred_presence` | Set preferred availability |
| `teams_list_my_teams` | List all teams you belong to |
| `teams_list_channels` | List channels in a team |
| `teams_list_channel_messages` | Read messages from a channel |
| `teams_send_channel_message` | Post a message to a channel |
| `teams_reply_to_channel_message` | Reply in a channel thread |
| `teams_list_my_chats` | List your direct chats |
| `teams_get_chat_messages` | Get messages from a chat |
| `teams_send_chat_message` | Send a direct message |
| `teams_create_chat` | Start a new 1:1 or group chat |
| `teams_list_team_members` | List members of a team |
| `teams_search_users` | Search for users by name or email |
| `teams_get_schedule` | Read shift schedules |

## License

MIT

import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

import { getConfig } from "./config.js";
import { TeamsService } from "./service.js";

const config = getConfig();
const service = new TeamsService(config);

const server = new McpServer({
  name: "msteams-mcp-server",
  version: "1.0.0",
});

// ─── Presence Tools ────────────────────────────────────────────

server.tool(
  "teams_get_my_presence",
  "Get your current Teams presence/availability status",
  {},
  async () => {
    const result = await service.getMyPresence();
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_get_user_presence",
  "Get another user's presence/availability by their user ID or email",
  { userId: z.string().min(1).describe("User ID or userPrincipalName (email)") },
  async ({ userId }) => {
    const result = await service.getUserPresence(userId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_get_presences_batch",
  "Get presence for multiple users at once",
  { userIds: z.array(z.string().min(1)).min(1).max(650) },
  async ({ userIds }) => {
    const result = await service.getPresencesByUserIds(userIds);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_set_my_presence",
  "Set your Teams presence status (Available, Busy, DoNotDisturb, Away, Offline)",
  {
    availability: z.enum(["Available", "Busy", "DoNotDisturb", "Away", "Offline"]),
    activity: z.enum([
      "Available", "InACall", "InAConferenceCall", "Busy",
      "Presenting", "DoNotDisturb", "Away", "OffWork", "Offline",
    ]),
    expirationDuration: z.string().default("PT1H").describe("ISO 8601 duration, e.g. PT1H, PT30M"),
  },
  async ({ availability, activity, expirationDuration }) => {
    const sessionId = config.clientId || "mcp-session";
    const result = await service.setMyPresence(sessionId, availability, activity, expirationDuration);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_clear_my_presence",
  "Clear your custom presence and revert to automatic status",
  {},
  async () => {
    const sessionId = config.clientId || "mcp-session";
    const result = await service.clearMyPresence(sessionId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_set_preferred_presence",
  "Set user preferred presence (like manual status in Teams app, overrides all apps)",
  {
    availability: z.enum(["Available", "Busy", "DoNotDisturb", "Away", "Offline"]),
    activity: z.enum([
      "Available", "InACall", "InAConferenceCall", "Busy",
      "Presenting", "DoNotDisturb", "Away", "OffWork", "Offline",
    ]),
    expirationDuration: z.string().default("PT8H").describe("ISO 8601 duration"),
  },
  async ({ availability, activity, expirationDuration }) => {
    const result = await service.setUserPreferredPresence(availability, activity, expirationDuration);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ─── Teams & Channels Tools ───────────────────────────────────

server.tool(
  "teams_list_my_teams",
  "List all Teams you are a member of",
  {},
  async () => {
    const result = await service.listMyTeams();
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_list_channels",
  "List channels in a specific Team",
  { teamId: z.string().min(1) },
  async ({ teamId }) => {
    const result = await service.listChannels(teamId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_list_channel_messages",
  "List recent messages from a channel",
  {
    teamId: z.string().min(1),
    channelId: z.string().min(1),
    limit: z.number().int().min(1).max(50).default(20),
  },
  async ({ teamId, channelId, limit }) => {
    const result = await service.listChannelMessages(teamId, channelId, limit);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_send_channel_message",
  "Send a message to a Teams channel",
  {
    teamId: z.string().min(1),
    channelId: z.string().min(1),
    content: z.string().min(1),
  },
  async ({ teamId, channelId, content }) => {
    const result = await service.sendChannelMessage(teamId, channelId, content);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_reply_to_channel_message",
  "Reply to a specific message in a Teams channel",
  {
    teamId: z.string().min(1),
    channelId: z.string().min(1),
    messageId: z.string().min(1),
    content: z.string().min(1),
  },
  async ({ teamId, channelId, messageId, content }) => {
    const result = await service.replyToChannelMessage(teamId, channelId, messageId, content);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ─── Chat Tools ───────────────────────────────────────────────

server.tool(
  "teams_list_my_chats",
  "List your recent chats",
  { limit: z.number().int().min(1).max(50).default(20) },
  async ({ limit }) => {
    const result = await service.listMyChats(limit);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_get_chat_messages",
  "Get messages from a specific chat",
  {
    chatId: z.string().min(1),
    limit: z.number().int().min(1).max(50).default(20),
  },
  async ({ chatId, limit }) => {
    const result = await service.getChatMessages(chatId, limit);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_send_chat_message",
  "Send a message in an existing chat",
  {
    chatId: z.string().min(1),
    content: z.string().min(1),
  },
  async ({ chatId, content }) => {
    const result = await service.sendChatMessage(chatId, content);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_create_chat",
  "Create a new 1:1 or group chat",
  {
    memberUserIds: z.array(z.string().min(1)).min(1).describe("User IDs to include in the chat"),
  },
  async ({ memberUserIds }) => {
    const members = memberUserIds.map((userId) => ({ userId }));
    const result = await service.createChat(members);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ─── Members & Users ──────────────────────────────────────────

server.tool(
  "teams_list_team_members",
  "List members of a specific Team",
  { teamId: z.string().min(1) },
  async ({ teamId }) => {
    const result = await service.listTeamMembers(teamId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "teams_search_users",
  "Search for users by name or email",
  { query: z.string().min(1) },
  async ({ query }) => {
    const result = await service.searchUsers(query);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ─── Schedule / Free-Busy ─────────────────────────────────────

server.tool(
  "teams_get_schedule",
  "Get free/busy schedule for users (calendar availability)",
  {
    schedules: z.array(z.string().min(1)).min(1).describe("Email addresses to check"),
    startTime: z.string().describe("Start datetime, e.g. 2026-05-08T09:00:00"),
    endTime: z.string().describe("End datetime, e.g. 2026-05-08T17:00:00"),
    timeZone: z.string().default("UTC").describe("IANA timezone, e.g. America/New_York"),
  },
  async ({ schedules, startTime, endTime, timeZone }) => {
    const result = await service.getSchedule(schedules, startTime, endTime, timeZone);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ─── Start Server ─────────────────────────────────────────────

async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Failed to start msteams-mcp-server", error);
  process.exit(1);
});

import { GraphClient } from "./graph-client.js";
import type { MSTeamsConfig } from "./config.js";

export class TeamsService {
  private client: GraphClient;

  constructor(config: MSTeamsConfig) {
    this.client = new GraphClient(config);
  }

  // ─── Presence ───────────────────────────────────────────────

  async getMyPresence(): Promise<unknown> {
    return this.client.request("/me/presence");
  }

  async getUserPresence(userId: string): Promise<unknown> {
    return this.client.request(`/users/${userId}/presence`);
  }

  async getPresencesByUserIds(ids: string[]): Promise<unknown> {
    return this.client.request("/communications/getPresencesByUserId", {
      method: "POST",
      body: { ids },
    });
  }

  async setMyPresence(
    sessionId: string,
    availability: string,
    activity: string,
    expirationDuration: string
  ): Promise<unknown> {
    return this.client.request("/me/presence/setPresence", {
      method: "POST",
      body: { sessionId, availability, activity, expirationDuration },
    });
  }

  async clearMyPresence(sessionId: string): Promise<unknown> {
    return this.client.request("/me/presence/clearPresence", {
      method: "POST",
      body: { sessionId },
    });
  }

  async setUserPreferredPresence(
    availability: string,
    activity: string,
    expirationDuration: string
  ): Promise<unknown> {
    return this.client.request("/me/presence/setUserPreferredPresence", {
      method: "POST",
      body: { availability, activity, expirationDuration },
    });
  }

  // ─── Teams & Channels ──────────────────────────────────────

  async listMyTeams(): Promise<unknown> {
    return this.client.request("/me/joinedTeams");
  }

  async listChannels(teamId: string): Promise<unknown> {
    return this.client.request(`/teams/${teamId}/channels`);
  }

  async listChannelMessages(teamId: string, channelId: string, limit: number): Promise<unknown> {
    return this.client.request(
      `/teams/${teamId}/channels/${channelId}/messages?$top=${limit}`
    );
  }

  async sendChannelMessage(teamId: string, channelId: string, content: string): Promise<unknown> {
    return this.client.request(
      `/teams/${teamId}/channels/${channelId}/messages`,
      {
        method: "POST",
        body: { body: { content, contentType: "text" } },
      }
    );
  }

  async replyToChannelMessage(
    teamId: string,
    channelId: string,
    messageId: string,
    content: string
  ): Promise<unknown> {
    return this.client.request(
      `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`,
      {
        method: "POST",
        body: { body: { content, contentType: "text" } },
      }
    );
  }

  // ─── Chats ─────────────────────────────────────────────────

  async listMyChats(limit: number): Promise<unknown> {
    return this.client.request(`/me/chats?$top=${limit}`);
  }

  async getChatMessages(chatId: string, limit: number): Promise<unknown> {
    return this.client.request(`/chats/${chatId}/messages?$top=${limit}`);
  }

  async sendChatMessage(chatId: string, content: string): Promise<unknown> {
    return this.client.request(`/chats/${chatId}/messages`, {
      method: "POST",
      body: { body: { content, contentType: "text" } },
    });
  }

  async createChat(members: Array<{ userId: string; roles?: string[] }>): Promise<unknown> {
    const chatType = members.length > 1 ? "group" : "oneOnOne";
    const memberPayload = members.map((m) => ({
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      roles: m.roles || ["owner"],
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${m.userId}')`,
    }));

    return this.client.request("/chats", {
      method: "POST",
      body: { chatType, members: memberPayload },
    });
  }

  // ─── Team Members ──────────────────────────────────────────

  async listTeamMembers(teamId: string): Promise<unknown> {
    return this.client.request(`/teams/${teamId}/members`);
  }

  // ─── Users (lookup) ────────────────────────────────────────

  async searchUsers(query: string): Promise<unknown> {
    return this.client.request(
      `/users?$filter=startsWith(displayName,'${query}') or startsWith(mail,'${query}')&$top=10`
    );
  }

  // ─── Calendar / Schedule ───────────────────────────────────

  async getSchedule(
    schedules: string[],
    startTime: string,
    endTime: string,
    timeZone: string
  ): Promise<unknown> {
    return this.client.request("/me/calendar/getSchedule", {
      method: "POST",
      body: {
        schedules,
        startTime: { dateTime: startTime, timeZone },
        endTime: { dateTime: endTime, timeZone },
        availabilityViewInterval: 30,
      },
    });
  }
}

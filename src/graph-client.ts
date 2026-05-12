import type { MSTeamsConfig } from "./config.js";
import { OAuth2Auth } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

interface TokenResponse {
  access_token: string;
  expires_in: number;
}

export class GraphClient {
  private config: MSTeamsConfig;
  private oauth2: OAuth2Auth | null = null;
  private token: string | null = null;
  private tokenExpiry: number = 0;

  constructor(config: MSTeamsConfig) {
    this.config = config;

    if (config.authMode === "delegated") {
      this.oauth2 = new OAuth2Auth(config.tenantId, config.clientId);
    }
  }

  private async getToken(): Promise<string> {
    if (this.config.accessToken) {
      return this.config.accessToken;
    }

    if (this.oauth2) {
      return this.oauth2.getAccessToken();
    }

    // Client credentials flow (application permissions)
    if (this.token && Date.now() < this.tokenExpiry) {
      return this.token;
    }

    const url = TOKEN_URL_TEMPLATE.replace("{tenantId}", this.config.tenantId);
    const body = new URLSearchParams({
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Token request failed (${res.status}): ${text}`);
    }

    const data = (await res.json()) as TokenResponse;
    this.token = data.access_token;
    this.tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
    return this.token;
  }

  async request(path: string, options: { method?: string; body?: unknown } = {}): Promise<unknown> {
    const token = await this.getToken();
    const { method = "GET", body } = options;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const res = await fetch(`${GRAPH_BASE}${path}`, {
      method,
      headers,
      body: body ? JSON.stringify(body) : undefined,
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`MS Graph ${method} ${path} failed (${res.status}): ${text}`);
    }

    const contentType = res.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      return res.json();
    }
    return { status: res.status, statusText: res.statusText };
  }
}

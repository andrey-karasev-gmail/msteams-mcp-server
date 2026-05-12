import fs from "node:fs";
import path from "node:path";

const TOKEN_CACHE_FILE = path.resolve(
  process.env.MS_TOKEN_CACHE_PATH || path.join(process.cwd(), "data", "token-cache.json")
);

const SCOPES = [
  "Presence.ReadWrite",
  "Presence.Read.All",
  "Chat.ReadWrite",
  "ChannelMessage.Send",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "Calendars.Read",
  "User.Read",
  "User.ReadBasic.All",
].join(" ");

interface TokenCache {
  access_token: string;
  refresh_token: string;
  expires_at: number;
}

interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  error?: string;
  error_description?: string;
}

export class OAuth2Auth {
  private tenantId: string;
  private clientId: string;
  private tokenCache: TokenCache | null = null;

  constructor(tenantId: string, clientId: string) {
    this.tenantId = tenantId;
    this.clientId = clientId;
    this.loadCache();
  }

  private get tokenUrl(): string {
    return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
  }

  private get deviceCodeUrl(): string {
    return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/devicecode`;
  }

  private loadCache(): void {
    try {
      if (fs.existsSync(TOKEN_CACHE_FILE)) {
        const data = JSON.parse(fs.readFileSync(TOKEN_CACHE_FILE, "utf-8"));
        this.tokenCache = data;
      }
    } catch {
      this.tokenCache = null;
    }
  }

  private saveCache(cache: TokenCache): void {
    this.tokenCache = cache;
    const dir = path.dirname(TOKEN_CACHE_FILE);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(TOKEN_CACHE_FILE, JSON.stringify(cache, null, 2));
  }

  async getAccessToken(): Promise<string> {
    if (this.tokenCache) {
      if (Date.now() < this.tokenCache.expires_at - 60_000) {
        return this.tokenCache.access_token;
      }
      if (this.tokenCache.refresh_token) {
        try {
          return await this.refreshToken(this.tokenCache.refresh_token);
        } catch {
          // Refresh failed, need new device code flow
        }
      }
    }

    return this.deviceCodeFlow();
  }

  private async refreshToken(refreshToken: string): Promise<string> {
    const body = new URLSearchParams({
      client_id: this.clientId,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: SCOPES,
    });

    const res = await fetch(this.tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!res.ok) {
      throw new Error(`Refresh token failed: ${res.status}`);
    }

    const data = (await res.json()) as TokenResponse;
    this.saveCache({
      access_token: data.access_token,
      refresh_token: data.refresh_token || refreshToken,
      expires_at: Date.now() + data.expires_in * 1000,
    });

    return data.access_token;
  }

  private async deviceCodeFlow(): Promise<string> {
    const body = new URLSearchParams({
      client_id: this.clientId,
      scope: `${SCOPES} offline_access`,
    });

    const res = await fetch(this.deviceCodeUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Device code request failed (${res.status}): ${text}`);
    }

    const deviceCode = (await res.json()) as DeviceCodeResponse;

    // Write instructions to stderr so the user can see them
    process.stderr.write("\n");
    process.stderr.write("╔══════════════════════════════════════════════════════════╗\n");
    process.stderr.write("║  Microsoft Teams Authentication Required                 ║\n");
    process.stderr.write("╠══════════════════════════════════════════════════════════╣\n");
    process.stderr.write(`║  Go to: ${deviceCode.verification_uri.padEnd(47)}║\n`);
    process.stderr.write(`║  Enter code: ${deviceCode.user_code.padEnd(43)}║\n`);
    process.stderr.write("╚══════════════════════════════════════════════════════════╝\n");
    process.stderr.write("\n");
    process.stderr.write("Waiting for authentication...\n");

    return this.pollForToken(deviceCode.device_code, deviceCode.interval, deviceCode.expires_in);
  }

  private async pollForToken(deviceCode: string, interval: number, expiresIn: number): Promise<string> {
    const deadline = Date.now() + expiresIn * 1000;
    const pollInterval = Math.max(interval, 5) * 1000;

    while (Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, pollInterval));

      const body = new URLSearchParams({
        client_id: this.clientId,
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        device_code: deviceCode,
      });

      const res = await fetch(this.tokenUrl, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      });

      const data = (await res.json()) as TokenResponse;

      if (data.error === "authorization_pending") {
        continue;
      }

      if (data.error === "slow_down") {
        await new Promise((resolve) => setTimeout(resolve, 5000));
        continue;
      }

      if (data.error) {
        throw new Error(`Auth failed: ${data.error_description || data.error}`);
      }

      process.stderr.write("Authentication successful!\n\n");

      this.saveCache({
        access_token: data.access_token,
        refresh_token: data.refresh_token || "",
        expires_at: Date.now() + data.expires_in * 1000,
      });

      return data.access_token;
    }

    throw new Error("Device code flow timed out — user did not authenticate in time");
  }
}

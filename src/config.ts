export type AuthMode = "delegated" | "application" | "token";

export interface MSTeamsConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  accessToken?: string;
  authMode: AuthMode;
}

function requireEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export function getConfig(): MSTeamsConfig {
  const accessToken = process.env.MS_GRAPH_ACCESS_TOKEN?.trim();

  if (accessToken) {
    return {
      tenantId: "",
      clientId: "",
      clientSecret: "",
      accessToken,
      authMode: "token",
    };
  }

  const authMode = (process.env.MS_AUTH_MODE?.trim() || "delegated") as AuthMode;

  if (authMode === "application") {
    return {
      tenantId: requireEnv("AZURE_TENANT_ID"),
      clientId: requireEnv("AZURE_CLIENT_ID"),
      clientSecret: requireEnv("AZURE_CLIENT_SECRET"),
      authMode: "application",
    };
  }

  // Delegated (device code flow) — no client secret needed
  return {
    tenantId: requireEnv("AZURE_TENANT_ID"),
    clientId: requireEnv("AZURE_CLIENT_ID"),
    clientSecret: process.env.AZURE_CLIENT_SECRET?.trim() || "",
    authMode: "delegated",
  };
}

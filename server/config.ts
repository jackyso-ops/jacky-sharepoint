import "dotenv/config";

function env(name: string): string {
  return process.env[name]?.trim() || "";
}

/** Treat .env.example placeholders as unset. */
function realSecret(value: string): string {
  const v = value.trim();
  if (!v) return "";
  if (/^your[-_]/i.test(v)) return "";
  if (v === "change-me-to-a-long-random-string") return "";
  return v;
}

export type AppConfig = {
  port: number;
  publicBaseUrl: string;
  entra: {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    redirectUri: string;
  };
  sessionSecret: string;
  hasEntra: boolean;
};

export function loadConfig(): AppConfig {
  const tenantId = realSecret(env("ENTRA_TENANT_ID"));
  const clientId = realSecret(env("ENTRA_CLIENT_ID"));
  const clientSecret = realSecret(env("ENTRA_CLIENT_SECRET"));
  const redirectUri =
    env("ENTRA_REDIRECT_URI") ||
    `${env("PUBLIC_BASE_URL") || "http://localhost:5173"}/auth/microsoft/callback`;

  return {
    port: Number(env("PORT") || 3001),
    publicBaseUrl: env("PUBLIC_BASE_URL") || "http://localhost:5173",
    entra: { tenantId, clientId, clientSecret, redirectUri },
    sessionSecret: realSecret(env("SESSION_SECRET")) || "dev-only-change-me",
    hasEntra: Boolean(tenantId && clientId && clientSecret),
  };
}

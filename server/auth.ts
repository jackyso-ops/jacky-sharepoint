/**
 * Entra (Microsoft) delegated OAuth for local-first MVP.
 * Tokens stored in local JSON file (see tokens.ts). Prod: monday-code Secure Storage.
 */

import {
  deleteMicrosoftToken,
  getMicrosoftToken,
  saveMicrosoftToken,
  type MicrosoftTokenRecord,
} from "./tokens.js";

export type { MicrosoftTokenRecord };

const DEFAULT_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Files.Read.All",
  "Sites.Read.All",
];

export function buildAuthorizeUrl(opts: {
  tenantId: string;
  clientId: string;
  redirectUri: string;
  state: string;
  scopes?: string[];
}): string {
  const scopes = (opts.scopes ?? DEFAULT_SCOPES).join(" ");
  const u = new URL(
    `https://login.microsoftonline.com/${opts.tenantId}/oauth2/v2.0/authorize`
  );
  u.searchParams.set("client_id", opts.clientId);
  u.searchParams.set("response_type", "code");
  u.searchParams.set("redirect_uri", opts.redirectUri);
  u.searchParams.set("response_mode", "query");
  u.searchParams.set("scope", scopes);
  u.searchParams.set("state", opts.state);
  return u.toString();
}

export async function exchangeCodeForTokens(opts: {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  code: string;
}): Promise<MicrosoftTokenRecord> {
  const body = new URLSearchParams({
    client_id: opts.clientId,
    client_secret: opts.clientSecret,
    grant_type: "authorization_code",
    code: opts.code,
    redirect_uri: opts.redirectUri,
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${opts.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Entra token exchange failed: ${res.status} ${text}`);
  }
  const json = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
    scope?: string;
  };
  return {
    accessToken: json.access_token,
    refreshToken: json.refresh_token,
    expiresAt: Date.now() + json.expires_in * 1000,
    scope: json.scope,
  };
}

export async function getValidAccessToken(
  accountId: string,
  userId: string,
  secrets: { tenantId: string; clientId: string; clientSecret: string }
): Promise<string | null> {
  const stored = await getMicrosoftToken(accountId, userId);
  if (!stored) return null;

  const skewMs = 60_000;
  if (stored.expiresAt > Date.now() + skewMs) {
    return stored.accessToken;
  }
  if (!stored.refreshToken) return null;

  const body = new URLSearchParams({
    client_id: secrets.clientId,
    client_secret: secrets.clientSecret,
    grant_type: "refresh_token",
    refresh_token: stored.refreshToken,
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${secrets.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    }
  );
  if (!res.ok) {
    await deleteMicrosoftToken(accountId, userId);
    return null;
  }
  const json = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
    scope?: string;
  };
  const next: MicrosoftTokenRecord = {
    accessToken: json.access_token,
    refreshToken: json.refresh_token ?? stored.refreshToken,
    expiresAt: Date.now() + json.expires_in * 1000,
    scope: json.scope ?? stored.scope,
  };
  await saveMicrosoftToken(accountId, userId, next);
  return next.accessToken;
}

export { deleteMicrosoftToken, getMicrosoftToken, saveMicrosoftToken };

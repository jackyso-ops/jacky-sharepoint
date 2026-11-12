/**
 * Local Express API — Entra OAuth callback + Microsoft Graph proxy.
 * Vite proxies /api and /auth here in development.
 * Never put ENTRA_CLIENT_SECRET in the Vite bundle.
 */

import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import cookieParser from "cookie-parser";
import cors from "cors";
import express from "express";
import {
  buildAuthorizeUrl,
  deleteMicrosoftToken,
  exchangeCodeForTokens,
  getMicrosoftToken,
  getValidAccessToken,
  saveMicrosoftToken,
} from "./auth.js";
import { loadConfig } from "./config.js";
import { getShareRootWithChildren, listChildren } from "./graph.js";
import { folderUrlToShareId } from "./sharepoint-url.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const config = loadConfig();
const app = express();

const OAUTH_STATE_COOKIE = "sp_oauth_state";
const CTX_COOKIE = "sp_ms_ctx";
const useSecureCookies = config.publicBaseUrl.startsWith("https://");

type Ctx = { accountId: string; userId: string };

function cookieBase(maxAge: number) {
  return {
    httpOnly: true as const,
    signed: true as const,
    maxAge,
    sameSite: (useSecureCookies ? "none" : "lax") as "none" | "lax",
    secure: useSecureCookies,
  };
}

app.use(
  cors({
    origin: true,
    credentials: true,
  })
);
app.use(express.json());
app.use(cookieParser(config.sessionSecret));

function readCtx(req: express.Request): Ctx | null {
  const signed = req.signedCookies?.[CTX_COOKIE] as string | undefined;
  if (signed) {
    try {
      const parsed = JSON.parse(signed) as Ctx;
      if (parsed.accountId && parsed.userId) return parsed;
    } catch {
      /* ignore */
    }
  }
  const accountId =
    (req.header("x-monday-account-id") ||
      (req.query.accountId as string) ||
      process.env.DEV_ACCOUNT_ID ||
      "").trim();
  const userId =
    (req.header("x-monday-user-id") ||
      (req.query.userId as string) ||
      process.env.DEV_USER_ID ||
      "").trim();
  if (accountId && userId) return { accountId, userId };
  return null;
}

function setCtxCookie(res: express.Response, ctx: Ctx): void {
  res.cookie(CTX_COOKIE, JSON.stringify(ctx), cookieBase(30 * 24 * 60 * 60 * 1000));
}

function entraMissing(): string {
  return [
    "Entra is not configured.",
    "Copy .env.example → .env and set ENTRA_TENANT_ID, ENTRA_CLIENT_ID, ENTRA_CLIENT_SECRET, ENTRA_REDIRECT_URI.",
    "ENTRA_REDIRECT_URI must match an Entra app redirect URI exactly (use your tunnel HTTPS URL + /auth/microsoft/callback).",
  ].join(" ");
}

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    hasEntra: config.hasEntra,
    publicBaseUrl: config.publicBaseUrl,
    redirectUri: config.entra.redirectUri,
  });
});

app.get("/api/auth/status", async (req, res) => {
  const ctx = readCtx(req);
  if (!ctx) {
    res.json({ connected: false, reason: "no_context" });
    return;
  }
  const token = await getMicrosoftToken(ctx.accountId, ctx.userId);
  res.json({
    connected: Boolean(token),
    accountId: ctx.accountId,
    userId: ctx.userId,
    expiresAt: token?.expiresAt ?? null,
  });
});

/** Persist monday account/user so OAuth callback can key the token store. */
app.post("/api/auth/context", (req, res) => {
  const accountId = String(req.body?.accountId || "").trim();
  const userId = String(req.body?.userId || "").trim();
  if (!accountId || !userId) {
    res.status(400).json({ error: "accountId and userId are required" });
    return;
  }
  const ctx = { accountId, userId };
  // In local http (no tunnel yet) SameSite=None+Secure cookies may be dropped;
  // also accept headers / query on subsequent calls.
  try {
    setCtxCookie(res, ctx);
  } catch {
    /* ignore */
  }
  res.json({ ok: true, ...ctx });
});

app.get("/auth/microsoft/start", (req, res) => {
  if (!config.hasEntra) {
    res.status(503).type("text").send(entraMissing());
    return;
  }

  const accountId = String(
    req.query.accountId || readCtx(req)?.accountId || process.env.DEV_ACCOUNT_ID || ""
  ).trim();
  const userId = String(
    req.query.userId || readCtx(req)?.userId || process.env.DEV_USER_ID || ""
  ).trim();

  if (!accountId || !userId) {
    res
      .status(400)
      .type("text")
      .send(
        "Missing monday accountId/userId. Open the item view inside monday (or set DEV_ACCOUNT_ID / DEV_USER_ID in .env)."
      );
    return;
  }

  const nonce = crypto.randomBytes(16).toString("hex");
  const statePayload = Buffer.from(
    JSON.stringify({ accountId, userId, nonce }),
    "utf8"
  ).toString("base64url");

  res.cookie(OAUTH_STATE_COOKIE, nonce, cookieBase(10 * 60 * 1000));
  setCtxCookie(res, { accountId, userId });

  const url = buildAuthorizeUrl({
    tenantId: config.entra.tenantId,
    clientId: config.entra.clientId,
    redirectUri: config.entra.redirectUri,
    state: statePayload,
  });
  res.redirect(url);
});

app.get("/auth/microsoft/callback", async (req, res) => {
  if (!config.hasEntra) {
    res.status(503).type("text").send(entraMissing());
    return;
  }

  const err = req.query.error as string | undefined;
  const errDesc = req.query.error_description as string | undefined;
  if (err) {
    res
      .status(400)
      .type("text")
      .send(
        `Microsoft sign-in failed.\n\nerror: ${err}\nerror_description: ${errDesc || "(none)"}\n\nCommon fixes:\n- ENTRA_REDIRECT_URI must exactly match an Entra redirect URI (including tunnel hostname).\n- App needs admin consent for Files.Read.All + Sites.Read.All.\n- Tenant / account mismatch.`
      );
    return;
  }

  const code = req.query.code as string | undefined;
  const stateRaw = req.query.state as string | undefined;
  if (!code || !stateRaw) {
    res.status(400).type("text").send("Missing code or state from Entra callback.");
    return;
  }

  let state: { accountId: string; userId: string; nonce: string };
  try {
    state = JSON.parse(Buffer.from(stateRaw, "base64url").toString("utf8")) as {
      accountId: string;
      userId: string;
      nonce: string;
    };
  } catch {
    res.status(400).type("text").send("Invalid OAuth state.");
    return;
  }

  const cookieNonce = req.signedCookies?.[OAUTH_STATE_COOKIE] as string | undefined;
  if (cookieNonce && cookieNonce !== state.nonce) {
    res.status(400).type("text").send("OAuth state mismatch (cookie). Retry Connect Microsoft.");
    return;
  }

  try {
    const tokens = await exchangeCodeForTokens({
      tenantId: config.entra.tenantId,
      clientId: config.entra.clientId,
      clientSecret: config.entra.clientSecret,
      redirectUri: config.entra.redirectUri,
      code,
    });
    await saveMicrosoftToken(state.accountId, state.userId, tokens);
    setCtxCookie(res, { accountId: state.accountId, userId: state.userId });
    res.clearCookie(OAUTH_STATE_COOKIE);
    res
      .type("html")
      .send(
        `<!doctype html><html><body style="font-family:system-ui;padding:1.5rem">
        <h1>Connected</h1>
        <p>Microsoft account linked. You can close this tab and return to the monday item view.</p>
        <script>try{window.close()}catch(e){}</script>
        </body></html>`
      );
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    res
      .status(500)
      .type("text")
      .send(
        `Entra token exchange failed.\n\n${message}\n\nCheck ENTRA_CLIENT_SECRET, ENTRA_REDIRECT_URI, and that the auth code was not reused.`
      );
  }
});

app.post("/auth/microsoft/disconnect", async (req, res) => {
  const ctx = readCtx(req);
  if (!ctx) {
    res.status(400).json({ error: "no_context" });
    return;
  }
  await deleteMicrosoftToken(ctx.accountId, ctx.userId);
  res.json({ ok: true });
});

async function requireAccessToken(
  req: express.Request,
  res: express.Response
): Promise<string | null> {
  const ctx = readCtx(req);
  if (!ctx) {
    res.status(401).json({
      error: "unauthorized",
      message: "Missing monday context. Open inside the item view or set DEV_ACCOUNT_ID / DEV_USER_ID.",
    });
    return null;
  }
  if (!config.hasEntra) {
    res.status(503).json({ error: "entra_not_configured", message: entraMissing() });
    return null;
  }
  const token = await getValidAccessToken(ctx.accountId, ctx.userId, {
    tenantId: config.entra.tenantId,
    clientId: config.entra.clientId,
    clientSecret: config.entra.clientSecret,
  });
  if (!token) {
    res.status(401).json({
      error: "need_auth",
      message: "Connect Microsoft to browse this folder.",
    });
    return null;
  }
  return token;
}

app.get("/api/folder", async (req, res) => {
  const accessToken = await requireAccessToken(req, res);
  if (!accessToken) return;

  let shareId = String(req.query.shareId || "").trim();
  const folderUrl = String(req.query.folderUrl || "").trim();
  if (!shareId && folderUrl) {
    shareId = folderUrlToShareId(folderUrl) || "";
  }
  if (!shareId) {
    res.status(400).json({ error: "shareId or folderUrl required" });
    return;
  }

  try {
    const item = await getShareRootWithChildren(accessToken, shareId);
    const driveId = item.parentReference?.driveId;
    res.json({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      driveId,
      children: item.children ?? [],
    });
  } catch (e) {
    const status = (e as Error & { status?: number }).status || 502;
    const message = e instanceof Error ? e.message : String(e);
    res.status(status >= 400 && status < 600 ? status : 502).json({
      error: "graph_error",
      message,
    });
  }
});

app.get("/api/folder/children", async (req, res) => {
  const accessToken = await requireAccessToken(req, res);
  if (!accessToken) return;

  const driveId = String(req.query.driveId || "").trim();
  const driveItemId = String(
    req.query.driveItemId || req.query.itemId || ""
  ).trim();
  if (!driveId || !driveItemId) {
    res.status(400).json({ error: "driveId and driveItemId are required" });
    return;
  }

  try {
    const children = await listChildren(accessToken, driveId, driveItemId);
    res.json({ children });
  } catch (e) {
    const status = (e as Error & { status?: number }).status || 502;
    const message = e instanceof Error ? e.message : String(e);
    res.status(status >= 400 && status < 600 ? status : 502).json({
      error: "graph_error",
      message,
    });
  }
});

// Production static serve (after `npm run build`)
const dist = path.resolve(__dirname, "..", "dist");
app.use(express.static(dist));
app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api") || req.path.startsWith("/auth")) {
    next();
    return;
  }
  res.sendFile(path.join(dist, "index.html"), (err) => {
    if (err) next();
  });
});

app.listen(config.port, () => {
  console.log(`[server] listening on http://127.0.0.1:${config.port}`);
  console.log(`[server] PUBLIC_BASE_URL=${config.publicBaseUrl}`);
  console.log(`[server] ENTRA_REDIRECT_URI=${config.entra.redirectUri}`);
  if (!config.hasEntra) {
    console.warn(`[server] ${entraMissing()}`);
  }
});

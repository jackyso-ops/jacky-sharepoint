# jacky-sharepoint

View-only SharePoint folder browser as a **monday.com item view** for board **HCA** (`5032301117`), URL column **`text_mm90r6dr`**.

Custom React UI via Microsoft Graph — **not** an iframe of `AllItems.aspx`.

See [docs/PLAN.md](./docs/PLAN.md) for architecture and milestones.

## Stack (local-first)

| Piece | Role |
| --- | --- |
| Vite + React | Item view UI (`src/item-view/`) |
| Express | Entra OAuth callback + Graph proxy (`server/`) |
| monday SDK | Read item context + column `text_mm90r6dr` |
| Tunnel | monday apps-cli and/or ngrok → Developer Center + Entra redirects |

Production later: same UI, confidential backend on **monday-code** + Secure Storage. Do not start monday-code until the local Vite + tunnel loop works.

## Quick start

```bash
cp .env.example .env
# fill ENTRA_* values (see below)

npm install
npm run dev
```

- Item view (Vite): http://localhost:5173  
- API (Express): http://127.0.0.1:3001  
- Vite proxies `/api` and `/auth` → Express

```bash
npm run typecheck
npm run tunnel:hint
```

## Tunnel into monday

Expose Vite (5173) so monday can load the item view over HTTPS:

**A — monday apps-cli** (preferred with monday tooling)

```bash
mapps tunnel:create -p 5173
```

**B — ngrok**

```bash
ngrok http 5173
```

Then set **all three** to the same tunnel origin:

1. Developer Center → Item view feature URL → `https://<tunnel-host>/`
2. Entra app → Redirect URI → `https://<tunnel-host>/auth/microsoft/callback`
3. Local `.env`:

```env
PUBLIC_BASE_URL=https://<tunnel-host>
ENTRA_REDIRECT_URI=https://<tunnel-host>/auth/microsoft/callback
```

Restart `npm run dev` after `.env` changes. Re-register redirect URIs whenever the tunnel hostname changes.

## Entra (Microsoft) setup

1. Entra ID → App registrations → New registration.
2. Redirect URI: tunnel callback (above). Mismatch is the #1 auth failure.
3. Client secret → `ENTRA_CLIENT_SECRET` in `.env` only (never in the Vite bundle).
4. Delegated permissions: `Files.Read.All`, `Sites.Read.All`, plus `openid`, `profile`, `offline_access`, `User.Read`.
5. Grant admin consent.

## Auth errors (exact text)

If Entra is missing from `.env`, the server returns:

```text
Entra is not configured. Copy .env.example → .env and set ENTRA_TENANT_ID, ENTRA_CLIENT_ID, ENTRA_CLIENT_SECRET, ENTRA_REDIRECT_URI. ENTRA_REDIRECT_URI must match an Entra app redirect URI exactly (use your tunnel HTTPS URL + /auth/microsoft/callback).
```

If Microsoft returns an OAuth error on callback:

```text
Microsoft sign-in failed.

error: <error>
error_description: <error_description>

Common fixes:
- ENTRA_REDIRECT_URI must exactly match an Entra redirect URI (including tunnel hostname).
- App needs admin consent for Files.Read.All + Sites.Read.All.
- Tenant / account mismatch.
```

If token exchange fails:

```text
Entra token exchange failed.

Entra token exchange failed: <status> <body>

Check ENTRA_CLIENT_SECRET, ENTRA_REDIRECT_URI, and that the auth code was not reused.
```

Graph/API 401 JSON:

```json
{ "error": "need_auth", "message": "Connect Microsoft to browse this folder." }
```

## MVP behavior

- Read column `text_mm90r6dr` → normalize `AllItems.aspx?id=…` → Graph `/shares/{u!…}`
- List children (folders first); drill-down; open files via `webUrl` in a new tab
- Connect / Disconnect Microsoft (tokens in `.tokens/` locally)
- Empty / error / reconnect states — no SharePoint iframe, no upload/edit

## Layout

```
├── docs/PLAN.md
├── .env.example
├── index.html
├── vite.config.ts
├── server/          Express Entra + Graph
├── src/
│   ├── item-view/   React item view
│   ├── lib/         URL normalize + share id
│   └── styles/
└── scripts/tunnel-hint.mjs
```

## Install on HCA

1. Create monday Developer Center app → Item view → tunnel URL.
2. Build Draft; install on jacky-demo.
3. HCA board → item → Views / Apps → add this view.
4. Connect Microsoft; open an item that already has a SharePoint folder URL.

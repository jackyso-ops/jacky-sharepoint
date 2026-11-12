# Custom monday.com SharePoint folder browser (HCA)

**Board:** HCA (`5032301117`)  
**URL column:** `text_mm90r6dr`  
**Goal:** View-only folder + file listing for SharePoint library URLs on each item — custom UI, not an iframe of `AllItems.aspx`.

**Decision locked:** Build our own monday **Developer Center** item-view app. dsapps Secure Embed is out. vibe Graph paste is out for production.

**Dev hosting (Jacky):** Local-first. Vite + React item view on his machine, exposed into monday via **monday apps-cli tunnel** and/or **ngrok**. monday-code is the **production** backend later — not the day-one host.

---

## GitHub repo

| Field | Value |
| --- | --- |
| GitHub account | Confirmed (Jacky) |
| Repo URL | https://github.com/jackyso-ops/jacky-sharepoint.git |
| Scaffold | Copied from Project store `internal/sharepoint-monday-app/` and fleshed out for local Vite + Express MVP |

---

## Development order (local-first → production)

| Phase | What | Hosting |
| --- | --- | --- |
| **0** | Repo URL pasted; scaffold copied in | GitHub (Jacky’s repo) |
| **1 — local MVP** | Vite + React item view; local Express (or similar) for Entra callback + Graph proxy; env-file secrets | `localhost` → **monday apps-cli tunnel** and/or **ngrok** as the item-view / OAuth public URL in Developer Center |
| **2 — demo on jacky-demo** | Same local stack; Entra redirect URIs include tunnel URL(s); install Draft app on HCA | Tunnel stays up while testing |
| **3 — production** | Deploy confidential backend to **monday-code**; Secure Storage for tokens; production item-view host | monday-code + published app version |

**Do not** start on monday-code hosting before the local Vite + tunnel loop works for auth + list folder.

### Local stack (phase 1)

```
Jacky’s machine
  Vite (React item view)  ──┐
  Local API (OAuth + Graph)─┼──► apps-cli tunnel / ngrok ──► monday item view iframe
  .env (Entra secrets)    ──┘
```

1. Point the Developer Center **Item view** feature URL at the tunnel (HTTPS).
2. Register the same tunnel origin(s) as Entra **redirect URIs** (and update when the tunnel URL changes).
3. Keep client secret only in local `.env` / server process — never in the Vite bundle.
4. When ready for prod: swap the local API for monday-code Secure Storage + Secrets; keep the same React UI.

---

## Architecture

```
monday item view (React + monday-sdk)
        │  reads item + column text_mm90r6dr
        ▼
Secure backend
  • Local (dev): Vite companion API on tunnel / ngrok
  • Production: monday-code (Secure Storage / Secrets)
        │  holds / refreshes Microsoft tokens
        ▼
Microsoft Entra ID — delegated OAuth (per user)
        ▼
Microsoft Graph
  • normalize AllItems.aspx?id=… → folder web URL
  • GET /shares/{u!…}/driveItem?$expand=children
  • drill-down: GET /drives/{driveId}/items/{id}/children
        ▼
Custom folder UI in the item view (list, breadcrumb, open file in new tab)
```

| Layer | Role |
| --- | --- |
| **Item view** | Hosted in the HCA item card (dev: Vite via tunnel). Reads `text_mm90r6dr`, renders folder/file list, Connect / Disconnect Microsoft. |
| **Backend** | Token exchange, refresh, Graph proxy. **Dev:** local server. **Prod:** monday-code Secure Storage. |
| **Entra (Azure AD) app** | Registration; redirect URIs for tunnel (dev) and monday-code / prod callbacks. |
| **Graph** | Delegated `Files.Read.All` + `Sites.Read.All` (or least privilege that still lists drive children). |
| **UI** | App-rendered folder browser only. **Never** set iframe `src` to SharePoint Online pages. |

**View-only.** Open files via `webUrl` in a new tab. No upload, edit, or write scopes in MVP.

---

## Why not vibe / why not Connect OAuth there

| Need | monday vibe item view | This Developer Center app |
| --- | --- | --- |
| Per-user M365 ACL (user A ≠ user B) | Poor fit — custom API tokens are **shared by everyone** using the vibe app | Delegated Entra OAuth per monday user |
| First-class Microsoft Connect OAuth | Prebuilt vibe tools cover Gmail / Outlook / Drive / etc. — **no SharePoint / Graph files tool** | Workflow Credentials or app-managed Microsoft OAuth as documented for monday apps |
| Secure token store + refresh | Not a documented vibe item-view secret store | Local encrypted/env store in dev → monday-code **Secure Storage** in prod; and/or [Manage user tokens and credentials](https://developer.monday.com/apps/docs/manage-user-tokens-and-credentials) |
| Background / confidential client | Client-only (`is_fullstack: false` vibe item views) | Local API (dev) → monday-code (prod) |

**Bottom line:** vibe can prototype a paste-token Graph UI; it cannot productize Connect OAuth for SharePoint the way a Developer Center app can. This plan is the production path Jacky chose after ruling out dsapps.

**Auth options for this app (pick during setup):**

1. **App-managed Entra OAuth** — item view → backend starts Entra auth code flow → store access/refresh (local file/db in dev; Secure Storage keyed by monday `accountId` + `userId` in prod). Backend calls Graph with the user’s token. **Recommended for MVP.**
2. **Workflow Credentials (OAuth 2.0)** — configure Microsoft as a credential provider in Developer Center ([docs](https://developer.monday.com/apps/docs/manage-user-tokens-and-credentials)). Best when also shipping automation blocks; for a pure item view, confirm token delivery at runtime (credentials are designed for workflow blocks).

---

## Setup Jacky must do

### 1. GitHub repo

https://github.com/jackyso-ops/jacky-sharepoint.git

### 2. monday Developer Center app

1. Create a new app in [monday Developer Center](https://developer.monday.com/).
2. Add an **Item view** feature; for local-first set its URL to the **tunnel HTTPS URL** (apps-cli or ngrok), not monday-code yet.
3. Note client ID / signing secrets as needed for monday SDK.
4. Build a Draft version and install it on the **jacky-demo** account.
5. Later: enable **monday-code** for production OAuth callback + Graph proxy + Secure Storage.

### 3. Local tunnel

1. Run Vite (item view) + local API on localhost.
2. Expose with **monday apps-cli tunnel** (preferred when using the monday apps toolchain) and/or **ngrok**.
3. Put that public HTTPS base URL in Developer Center (item view) and in Entra redirect URIs.
4. Re-register redirect URIs whenever the tunnel hostname changes.

### 4. Microsoft Entra app registration

1. Entra ID → App registrations → New registration (name e.g. `monday-hca-sharepoint-browser`).
2. Supported account types: org that owns `dapulse.sharepoint.com` (single tenant unless multi-tenant is required).
3. **Redirect URIs:** tunnel callback(s) for local-first; add monday-code / prod URLs when deploying. Mismatch here is the #1 auth failure.
4. Create a **client secret** (or certificate); store in local `.env` for now — later monday-code Secrets. Never in the item-view bundle.
5. **API permissions (delegated):**
   - `Files.Read.All`
   - `Sites.Read.All`
   - Plus identity basics as needed: `openid`, `profile`, `offline_access`, `User.Read`
   - Prefer least privilege if Graph Explorer proves a narrower set works for `/shares` + children on HCA folders; start with the two Read.All scopes above if unsure.
6. Grant **admin consent** in the tenant for those delegated permissions.

### 5. Install and attach on HCA

1. Install the app on **jacky-demo**.
2. Open board **HCA** (`5032301117`) → item → Views / Apps → add this item view.
3. Confirm the view reads column **`text_mm90r6dr`** (title often “URL”).
4. With tunnel running: Connect Microsoft; open an item that already has a SharePoint folder URL.

---

## URL normalization (`AllItems.aspx` → Graph `/shares`)

Sample column value:

`https://dapulse.sharepoint.com/sites/QuaffANZ/Shared%20Documents/Forms/AllItems.aspx?...&id=%2Fsites%2FQuaffANZ%2FShared%20Documents%2FHCA%20Projects%2F...`

**Steps:**

1. If the URL has an `id=` query param, build a clean folder web URL:

   `https://dapulse.sharepoint.com` + `decodeURIComponent(id)`

   Drop `Forms/AllItems.aspx` chrome; do not Graph-call the library view page URL as-is if a path `id` exists.

2. Encode for Graph shares ([shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0)):

   - UTF-8 bytes of the absolute folder URL → base64url (`+`→`-`, `/`→`_`, strip `=`) → prefix `u!`

3. Call:

   `GET https://graph.microsoft.com/v1.0/shares/{shareId}/driveItem?$expand=children`

4. On folder click: `GET /drives/{driveId}/items/{itemId}/children` (keep `driveId` / ids from the share response).

5. Fallback if `/shares` fails: parse site + library + path and use path-based drive APIs.

`AllItems.aspx` is a **UI route**, not a Graph id. Successful apps always normalize then list children.

---

## MVP milestones

| # | Milestone | Done when |
| --- | --- | --- |
| M0 | **Repo + local host** | Jacky pasted GitHub URL; scaffold in repo; Vite item view loads in monday via apps-cli tunnel / ngrok. |
| M1 | **Auth** | Connect Microsoft from the item view (local API + tunnel); tokens stored/refreshed on the local backend; Disconnect clears them. (Prod later: Secure Storage.) |
| M2 | **List root folder** | Given a filled `text_mm90r6dr`, show children of the linked folder (folders first, then files). |
| M3 | **Drill-down** | Click folder → breadcrumb updates → children of that folder load. |
| M4 | **Open file** | Click file → `webUrl` opens in a new tab. |
| M5 | **Empty / error states** | No URL → “No SharePoint folder linked”. 401/403 → reconnect prompt. Graph errors surfaced without crashing the view. |
| M6 | **Production host** | monday-code deploy; Secure Storage; production redirect URIs; published version on jacky-demo. |

**Hard constraints for MVP:** view-only; no `AllItems.aspx` iframe; live Graph data only (no mock trees).

---

## Out of scope

- Edit, upload, delete, or rename in SharePoint
- Write Graph scopes (`Files.ReadWrite.*`, etc.)
- dsapps Secure Embed / marketplace Microsoft 365 SharePoint app
- vibe Graph paste-token prototype as the product path
- File preview iframes via Graph `preview` (nice-to-have after MVP)
- Automations that create folders or sync files (later)
- Starting on monday-code before local Vite + tunnel works

---

## Recommended stack

| Piece | Choice |
| --- | --- |
| Client | **Vite + React** item view + [`monday-sdk-js`](https://developer.monday.com/apps/docs/introduction-to-the-sdk) |
| Local expose | **monday apps-cli tunnel** and/or **ngrok** |
| Backend (dev) | Local Node API (Express or equivalent) + `.env` for Entra secrets |
| Backend (prod) | [monday-code](https://developer.monday.com/apps/docs/monday-code-javascript-sdk) + [`@mondaycom/apps-sdk`](https://developer.monday.com/apps/docs/monday-code-javascript-sdk) (`SecureStorage`, Secrets) |
| Microsoft | Entra OAuth (delegated) + Microsoft Graph REST |

---

## Repo status

Remote: https://github.com/jackyso-ops/jacky-sharepoint.git

Local-first MVP scaffold is in-repo (Vite item view + Express Entra/Graph). Auth path for MVP: **app-managed Entra OAuth**.

### Jacky decisions still open

1. Entra tenancy: single-tenant (dapulse) vs multi-tenant.
2. Tunnel preference: apps-cli only, ngrok only, or both as fallback.

---

## Related research (already in Project)

- `/cursor/stores/self/docs/sharepoint-embed-research.md` — why iframes fail; Graph pattern
- `/cursor/stores/self/docs/vibe-sharepoint-path.md` — why vibe Connect OAuth / paste-token is not this product path

---

## Quick reference links

- [Manage user tokens and credentials](https://developer.monday.com/apps/docs/manage-user-tokens-and-credentials)
- [monday-code JavaScript SDK (Secure Storage)](https://developer.monday.com/apps/docs/monday-code-javascript-sdk)
- [Graph shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0)
- [Graph list children](https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0)
- [SPO pages cannot be iframed](https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sites/cannot-display-sharepoint-pages-in-iframe)

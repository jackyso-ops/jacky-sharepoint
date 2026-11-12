#!/usr/bin/env node
console.log(`
Tunnel the local Vite app (default http://localhost:5173) into monday:

  Option A — monday apps-cli (preferred when using monday tooling)
    mapps tunnel:create -p 5173
    # or: monday-code tunnel / apps-cli equivalent for your CLI version

  Option B — ngrok
    ngrok http 5173

Then:
  1. Set Developer Center → Item view feature URL = https://<tunnel-host>/
  2. Set Entra redirect URI = https://<tunnel-host>/auth/microsoft/callback
  3. Put the same values in .env:
       PUBLIC_BASE_URL=https://<tunnel-host>
       ENTRA_REDIRECT_URI=https://<tunnel-host>/auth/microsoft/callback
  4. Restart npm run dev after .env changes

Vite proxies /api and /auth to the Express server (PORT, default 3001).
`);

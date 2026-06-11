# Outlook Unsubscribe Add-in

A cross-platform Office.js Outlook add-in (Outlook on the web, Windows, and Mac) that safely unsubscribes you from marketing email, moves handled messages to Deleted Items, and can scan the current folder for senders you have not unsubscribed from yet.

## Features

- **One-click unsubscribe** — parses `List-Unsubscribe` / `List-Unsubscribe-Post` and sends the RFC 8058 one-click POST through the backend.
- **Fallbacks** — `mailto:` unsubscribe links and unsubscribe links found in the email body when headers are missing.
- **Move to Deleted Items** — after a successful unsubscribe the email is moved to Deleted Items via Microsoft Graph.
- **Folder scan** — scans the folder of the open message (Inbox, Archive, etc.) for messages with unsubscribe support, hides senders you already unsubscribed from (tracked in Postgres), and lets you unsubscribe + delete each one.
- **Safety** — the backend only accepts HTTPS unsubscribe URLs, blocks private/local network targets, refuses unsafe redirects, applies an 8-second timeout, and redacts recipient tokens from audit logs.

## Architecture

One Railway service runs everything:

- **Express server** (`server/index.ts`) — serves the built taskpane (static `dist/`), the unsubscribe proxy API, the unsubscribed-senders API, and `/api/config`.
- **Postgres** (Railway plugin, optional in dev) — `unsubscribed_senders` table records which senders each user has unsubscribed from. Falls back to in-memory storage when `DATABASE_URL` is unset.
- **Microsoft Graph** — the taskpane uses MSAL Nested App Authentication (NAA) to get `Mail.ReadWrite` tokens for moving messages and listing folder contents. Requires a Microsoft Entra app registration (below).

### Environment variables

| Variable | Purpose |
| --- | --- |
| `PORT` | Server port (Railway sets this automatically). |
| `DATABASE_URL` | Postgres connection string (Railway Postgres plugin). Optional; in-memory fallback without it. |
| `AAD_CLIENT_ID` | Microsoft Entra app (client) ID. Without it, unsubscribe still works but move-to-deleted and folder scan are disabled. |
| `ALLOWED_ORIGINS` | Optional comma-separated extra CORS origins. Not needed when the taskpane is served from the same service. |
| `NODE_ENV` | Set to `development` for local HTTPS dev mode; leave unset/production in Railway. |

## Microsoft Entra app registration (required for move + scan)

1. Go to [entra.microsoft.com](https://entra.microsoft.com) → **Identity > Applications > App registrations > New registration**.
2. Name: `Outlook Unsubscribe Add-in`. Supported account types: **Accounts in this organizational directory only** (or multi-tenant if you want).
3. After creating, open **Authentication > Add a platform > Single-page application** and add both redirect URIs:
   - `brk-multihub://<your-railway-domain>` (e.g. `brk-multihub://mail-unsub-production.up.railway.app`) — required for Nested App Authentication
   - `https://<your-railway-domain>/taskpane.html` — fallback popup auth
4. Under **API permissions**, add **Microsoft Graph > Delegated > Mail.ReadWrite** (and keep `User.Read`). Grant admin consent.
5. Copy the **Application (client) ID** and set it as `AAD_CLIENT_ID` on the Railway service.

## Deploy on Railway

1. Create a Railway project from this repo (`spinman-ITS/mail-unsub`). Railway builds with `npm install && npm run build` and starts with `npm start`.
2. Add the **Postgres** plugin and reference its `DATABASE_URL` on the service.
3. Set `AAD_CLIENT_ID` (see above).
4. Generate a public domain for the service, then generate the production manifest:

```bash
npm run manifest:prod -- <your-railway-domain>
```

5. Sideload `manifest.prod.xml` in Outlook (Add-ins > My add-ins > Add a custom add-in > Add from file), or deploy it org-wide through the Microsoft 365 Admin Center (Settings > Integrated apps).

## Local development

```bash
npm install
npm run certs:install
npm test
npm run dev:all
```

- Add-in UI: `https://localhost:3003/taskpane.html`
- Local API: `https://localhost:8787` (HTTPS via Office dev certs; `NODE_ENV=development` is set by `dev:all`)
- Recent attempt logs: `https://localhost:8787/api/logs`
- Sideload `manifest.xml` (localhost URLs) for local testing.

When opened directly in a browser, the page shows **Local preview mode**; message inspection only works when Outlook loads the page through the sideloaded manifest.

## How unsubscribe decisions are made

1. RFC 8058 one-click (`List-Unsubscribe` HTTPS URL + `List-Unsubscribe-Post: List-Unsubscribe=One-Click`) — POSTed via the backend.
2. `mailto:` unsubscribe from headers — opens a prefilled email.
3. Unsubscribe links found in the email body — opened via the backend (GET).
4. After any successful HTTPS unsubscribe: the sender is recorded in Postgres and the message is moved to Deleted Items via Graph.

The folder scan lists the 50 most recent messages in the current folder, keeps those with `List-Unsubscribe` support, dedupes by sender, and hides senders already recorded for your mailbox.

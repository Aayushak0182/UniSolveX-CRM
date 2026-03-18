# UniSolveX CRM + WhatsApp (Meta Cloud API)

This project now includes:
- CRM frontend (`index.html` + `app.js`)
- WhatsApp webhook server (`backend/server.mjs`) for Meta Cloud API
- Real-time message push to CRM using WebSocket (`/ws`)

## 1. Install dependencies

```bash
npm install
cd backend && npm install
```

## 2. Configure environment

Copy `.env.example` to `.env` and set:

```env
PORT=3001
WHATSAPP_VERIFY_TOKEN=your_verify_token
WHATSAPP_APP_SECRET=your_meta_app_secret
```

## 3. Run app

Terminal 1:

```bash
cd backend
npm start
```

Terminal 2:

```bash
npm run dev
```

Frontend runs on Vite (usually `http://localhost:5173`), webhook server on `http://localhost:3001`.

## Firebase Hosting

Firebase Hosting serves the `public/` folder (see `firebase.json`). The static CRM files are mirrored there (`public/index.html`, `public/login.html`, `public/app.js`, `public/styles.css`) for deployment.

## 4. Meta Cloud API webhook settings

In Meta App Dashboard -> WhatsApp -> Configuration:

- Callback URL: `https://<your-public-domain>/webhook`
- Verify token: same as `.env` `WHATSAPP_VERIFY_TOKEN`
- Subscribe to `messages`

For production, set your Meta **App Secret** in `WHATSAPP_APP_SECRET` (or `META_APP_SECRET`) so the server can verify webhook request signatures (`X-Hub-Signature-256`).

For local development, expose port `3001` with a tunnel (like ngrok or Cloudflare Tunnel).

## 5. CRM behavior

- New incoming WhatsApp message auto-creates/updates contact in CRM.
- Contact name shown is WhatsApp profile name (`contacts[].profile.name`) when available.
- Clicking that contact opens its WhatsApp chat in CRM.
- Messages appear in real time through WebSocket.

## Render deployment

This repo includes `render.yaml` for backend deployment on Render.

- Push the full repo to GitHub.
- In Render, create a new Blueprint service from that GitHub repo, or connect the repo and let Render read `render.yaml`.
- The service deploys from `backend/`, so Render installs `backend/package.json` and starts `backend/server.mjs`.
- Health check path: `/health`
- Set env vars in Render dashboard: `WHATSAPP_VERIFY_TOKEN`, `WHATSAPP_APP_SECRET` (recommended), and optionally `WHATSAPP_ACCESS_TOKEN` + `WHATSAPP_PHONE_NUMBER_ID` (for sending messages)

## Notes

- Current storage is in-memory (resets on server restart).
- If required, add DB persistence (MongoDB/PostgreSQL) in `backend/server.mjs`.

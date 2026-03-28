# Deploying Task Track (free tier)

This app can run **fully in the browser** (empty `API_BASE_URL` in [config.js](config.js)) or with a **hosted API** on Render + MongoDB Atlas.

## 1. MongoDB Atlas (free M0)

1. Create a cluster at [MongoDB Atlas](https://www.mongodb.com/atlas).
2. Database Access: add a database user (username/password).
3. Network Access: add `0.0.0.0/0` (required for Render’s dynamic IPs on the free tier) or tighten later.
4. Connect → Drivers → copy the connection string, replace `<password>`, and set database name (e.g. `tasktrack`).

## 2. Render (free web service)

1. New **Web Service** from your GitHub repo (or deploy manually).
2. **Root directory**: `server`
3. **Build**: `npm install` — **Start**: `npm start`
4. Environment variables:

| Variable | Example / notes |
|----------|------------------|
| `MONGODB_URI` | Atlas connection string |
| `JWT_SECRET` | Long random string |
| `FRONTEND_ORIGIN` | `https://aminmansuri123.github.io` (no path). For multiple origins, use a comma-separated list. |
| `MASTER_EMAIL` | Default `mansuri.amin1@gmail.com` if omitted |
| `MASTER_PASSWORD` | Master login password (only in Render; never commit) |
| `NODE_ENV` | `production` |

5. After deploy, copy the service URL (e.g. `https://tasktrack-api.onrender.com`).

**Cold starts:** Free Render apps sleep after idle; the first request can take ~30–60 seconds.

## 3. GitHub Pages (frontend)

1. Push [config.js](config.js) with `window.API_BASE_URL = 'https://your-service.onrender.com';` (no trailing slash), or use [config.example.js](config.example.js) as a template.
2. Enable Pages on the repo (e.g. branch `main`, folder `/` or `/docs`).
3. Project site URL is typically `https://<user>.github.io/<repo>/`.

**CORS:** `FRONTEND_ORIGIN` on Render must **exactly** match the browser origin (scheme + host, no path), e.g. `https://aminmansuri123.github.io`.

## 4. Cookies and HTTPS

In production, the API sets an **httpOnly** auth cookie with `SameSite=None` and `Secure`. The frontend must be served over **HTTPS** (GitHub Pages does this automatically).

## 5. Master account

If `MASTER_PASSWORD` is set, the server ensures a user with `MASTER_EMAIL` exists, with `isMaster: true`. That account can use **Settings → Master password reset** (hosted mode only) or `POST /api/master/users/:userId/password` with a master session.

## 6. Optional: Blueprint

[render.yaml](render.yaml) at the **repository root** can be used with Render Blueprint; set secret env vars in the Render dashboard.

## 7. Local API testing

```bash
cd server
cp .env.example .env
# Edit .env — set MONGODB_URI, JWT_SECRET, FRONTEND_ORIGIN=http://127.0.0.1:5500
npm install
npm start
```

Serve the frontend with any static server (e.g. VS Code Live Server) so the origin matches `FRONTEND_ORIGIN`.

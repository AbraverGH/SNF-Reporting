# AR Collections Explorer — hosted version

This turns the dashboard into a small full-stack app: one shared link, one
shared dataset, viewable by anyone with the link. No database to set up —
the parsed data is stored as a single JSON file in `data/dataset.json`.

**How it works**
- Anyone who opens the link sees whatever was last published — no drag & drop needed to *view*.
- To *update* the data, drop the new workbook in as before. A banner appears asking
  for an admin key before it publishes for everyone (so a random visitor with the
  link can't overwrite your data).
- If you ever open `public/index.html` directly as a file (no server), it silently
  falls back to the original local-only, drag-and-drop-per-person behavior.

## 1. Run it locally first

```bash
npm install
ADMIN_KEY=pick-a-real-password npm start
```

Visit `http://localhost:3000`. Drop the workbook in, enter your `ADMIN_KEY` in the
publish banner, and confirm `http://localhost:3000/api/dataset` returns data.

## 2. Deploy it somewhere with a *persistent disk*

This is the one thing that matters: `data/dataset.json` must live on storage that
survives restarts and redeploys, or you'll lose the published data every time the
app restarts. Three straightforward options, easiest first:

### Option A — Railway (recommended, simplest)
1. Push this folder to a GitHub repo.
2. On [railway.app](https://railway.app), "New Project" → "Deploy from GitHub repo".
3. Add a **Volume**, mounted at `/app/data` (Railway → your service → Settings → Volumes).
4. Add an environment variable `ADMIN_KEY` set to a real password.
5. Railway auto-detects the Node app and runs `npm start`. You'll get a public URL —
   that's the link to share.

### Option B — Render
1. Push to GitHub, create a new **Web Service** on [render.com](https://render.com) from that repo.
2. Build command: `npm install`. Start command: `npm start`.
3. Add a **Disk** (Render → your service → Disks), mounted at `/opt/render/project/src/data`.
   (Render's free tier does not support persistent disks — you'll need a paid instance for this step.)
4. Add environment variable `ADMIN_KEY`.

### Option C — Your own server / VPS
1. Copy this folder to the server (`git clone` or `scp`).
2. `npm install --production`
3. Run it with a process manager so it survives reboots, e.g.
   `npm install -g pm2 && ADMIN_KEY=... pm2 start server.js --name ar-explorer`
4. Put it behind a reverse proxy (nginx/Caddy) for HTTPS and a real domain.
5. The `data/` folder is just a regular folder on disk here — no extra setup needed,
   just make sure it's not in a path that gets wiped on deploy.

## 3. Publishing updates

Whoever has the `ADMIN_KEY` can update the shared data at any time: open the link,
drop in the new workbook, enter the key in the banner, click **Publish**. Everyone
else's browser will show the new data the next time they load the page.

## 4. Backing up the data

`data/dataset.json` is the entire dataset. You can download it any time with:

```bash
curl https://your-app-url/api/dataset -o backup.json
```

## Notes / things to harden later if this grows

- The admin key is a single shared password, checked on every upload. Fine for a
  small team; swap in real auth (e.g. a login page) if this needs to scale beyond that.
- There's no upload history / versioning — publishing replaces the previous dataset.
  Keep periodic backups (see above) if you want to be able to roll back.
- For a larger team or heavier usage, swapping `data/dataset.json` for a real
  database (Postgres, e.g. a free Supabase/Neon instance) is a natural next step —
  the API shape (`GET/POST /api/dataset`) wouldn't need to change on the frontend.

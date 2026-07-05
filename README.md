# 📞 AutoLead Showcaser
### LAN-based Auto Lead Dialing System — Node.js

---

## Features
- **Admin Panel**: Upload Excel/CSV files of phone numbers
- **Auto Distribution**: Numbers assigned to agents one-by-one, never repeated
- **Agent Dialer**: Clean UI showing one number at a time — no skip button
- **Daily Auto-Reset**: Counts reset at midnight automatically
- **Live Stats**: Admin sees remaining, dialed, agents active — all real-time
- **LAN Network**: Runs on your local network, accessible from any device

---

## Setup & Run

### 1. Install Node.js
Download from: https://nodejs.org (LTS version)

### 2. Extract project folder

### 3. Install dependencies
```
cd autolead
npm install
```

### 4. Start the server
```
npm start
```

### 5. Find your LAN IP
- **Windows**: Run `ipconfig` → look for IPv4 Address (e.g. 192.168.1.10)
- **Mac/Linux**: Run `ifconfig` or `ip addr`

### 6. Share URLs with your team
| Role  | URL |
|-------|-----|
| Admin | http://YOUR-LAN-IP:3000/admin |
| Agent | http://YOUR-LAN-IP:3000/agent |

---

## How to Use

### Admin
1. Open `/admin` in your browser
2. Upload an Excel (.xlsx) or CSV file — Column A must contain phone numbers
3. Watch the stats update as agents dial
4. Upload multiple files — they all go into one pool

### Agent
1. Open `/agent` on their device (phone or PC)
2. Enter their name and log in
3. Press **Start Dialing** — a number appears
4. Dial the number, then press **✓ Dialed — Next Number**
5. Repeat! Daily count is shown on screen

---

## Excel/CSV Format
```
Column A
----------
9876543210
9123456789
8800001234
...
```
Header row (if any) is auto-skipped.

---

## Auto-Reset
Every midnight, daily dial counts reset to 0 automatically.
Numbers are re-available the next day (unless deleted).

---

## Port Change
Edit `server.js` line: `const PORT = 3000;`

---

## 💾 Your Data is Safe When You Update

All numbers, agents, leads, DND list, and EIDs are stored **outside the project folder** —
by default in a hidden folder in your OS user profile:
- **Windows**: `C:\Users\<you>\.autolead-crm`
- **Mac/Linux**: `~/.autolead-crm`

This means you can safely `git pull`, re-download the ZIP, or delete and re-clone this
repo to update the code — your data folder lives elsewhere and is never touched.

If you're upgrading from an older copy of this project that *did* store data inside the
project folder, the server automatically migrates it into the new location the first
time you start it — just check the startup log, it'll say `📦 Migrated existing...`.

**Daily backups**: a dated snapshot of `state.json` is also kept automatically in
`.autolead-crm/backups/` (last 14 days), as a safety net against an accidental
"Clear All" / "Hard Reset" or a corrupted file.

**Custom location**: set the `AUTOLEAD_DATA_DIR` environment variable before starting
the server if you'd rather store data somewhere specific (e.g. a backed-up drive, or a
mounted volume if you ever move this to a cloud host):
```
AUTOLEAD_DATA_DIR=/path/to/your/storage node server.js
```

---

## ☁️ Deploying on Railway (or Render/Heroku/Fly)

**Important:** on a container host, the filesystem itself is thrown away on every
deploy and every restart — including the "outside the project folder" home directory
this app uses by default. That trick only protects you on a real machine (your PC, a
LAN server, a VPS) where the disk actually sticks around. On Railway it does not,
unless you attach a **Volume**.

1. In your Railway project, select this service → **Volumes** tab (or ⌘K → "New Volume") → attach a volume to the service.
2. Set its **mount path** to `/data`.
3. Go to the service's **Variables** tab and add:
   ```
   AUTOLEAD_DATA_DIR=/data
   ```
4. Redeploy.

That's it — the app already migrates any old in-project data into `AUTOLEAD_DATA_DIR`
automatically on first boot, so nothing else needs to change. If you ever forget this
step, the server will now print a loud `🚨 DATA LOSS RISK` warning in the deploy logs
instead of failing silently.


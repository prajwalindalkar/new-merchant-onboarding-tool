# PhonePe Merchant Onboarding — Document Generator

Generates **Partner Resolution** and **BO Declaration** `.docx` files for partnership firm merchant onboarding.

---

## Quick Start (Local)

```bash
npm install
npm start
```
Open **http://localhost:3000** in your browser.

---

## How to Host (Free Options)

### Option A — Render.com (Recommended, Free)
1. Push this folder to a GitHub repo
2. Go to https://render.com → New → Web Service
3. Connect your GitHub repo
4. Build command: `npm install`
5. Start command: `node server.js`
6. Done — you get a permanent URL like `https://yourapp.onrender.com`

### Option B — Railway.app (Free tier)
1. Push to GitHub
2. Go to https://railway.app → New Project → Deploy from GitHub
3. It auto-detects Node.js and deploys
4. Go to Settings → Generate Domain

### Option C — Replit
1. Go to https://replit.com → New Repl → Import from GitHub
2. Paste your repo URL
3. Click Run — you get a live URL

---

## Features

- ✅ **5-step form** — Firm Details → Partners → Resolution → BO Declaration → Generate
- ✅ **Auto-save** — All data saved to browser localStorage; safe to close and come back
- ✅ **Clear button** — Wipe all saved data when starting a new merchant
- ✅ **Validation** — Won't generate without all required fields; PAN format checked; share % must total 100%; present partners must have >50% combined share; BO date must match resolution date
- ✅ **Document generation** — Server-side using docx.js; 100% reliable Word format
- ✅ **Exact template** — Every word matches the PhonePe-approved format

---

## Files

```
merchant-tool/
├── server.js        ← Express server + docx generation logic
├── package.json
└── public/
    └── index.html   ← The form UI
```

---

## Notes

- Print both documents on the **firm's letterhead**
- Affix **signatures** and **firm seal** after printing
- The BO declaration date **must match** the resolution date
- Partner ownership % **must match** the Partnership Deed exactly

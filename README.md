# Grain Elevator Data Transformer

A browser-based tool that converts raw scale ticket reports (Excel) into Balance API–ready CSVs — no server, no upload, all processing happens locally in the browser.

**Live demo:** [your-app.netlify.app](https://your-app.netlify.app) *(update after deploying)*

---

## What it does

Upload an `.xlsx` or `.xls` scale ticket report and the tool produces four ready-to-import CSVs:

| Output file | Contents |
|---|---|
| `IN_SCALE_TICKET_YYYYMMDD.csv` | Inbound regular tickets (corn + soybeans) |
| `OUT_SCALE_TICKET_YYYYMMDD.csv` | Outbound regular tickets (corn + soybeans) |
| `TRANSFER_INTERNAL_YYYYMMDD.csv` | Internal co-op transfers (BOL-matched) |
| `TRANSFER_EXT_PARTNER_YYYYMMDD.csv` | External partner transfers |
| `WARNINGS_YYYYMMDD.csv` | Data quality issues (if any) |

It also surfaces a live warnings dashboard for: duplicate import IDs, missing required fields, invalid quantities, and unmatched/blank BOL numbers.

---

## Configuration

Two enterprise IDs are set as defaults in `src/App.jsx`. Swap them for your actual IDs before deploying:

```js
// Co-op enterprise ID (used on inbound/outbound and as source on transfers)
target_enterprise_id: 'grain-coop'   // → replace with your enterprise ID

// External partner enterprise ID
'target.enterprise_id': 'ext-partner'  // → replace with your partner's ID
'target.id': 'ext-partner-01'          // → replace with your partner's location ID
```

---

## Run locally

Requires Node.js 18+.

```bash
npm install
npm run dev
```

Open [http://localhost:5173](http://localhost:5173).

---

## Deploy to Netlify

### Option A — Netlify UI (easiest)

1. Push this repo to GitHub.
2. Go to [netlify.com](https://netlify.com) → **Add new site → Import an existing project**.
3. Connect your GitHub account and select this repo.
4. Build settings are auto-detected from `netlify.toml`:
   - **Build command:** `npm run build`
   - **Publish directory:** `dist`
5. Click **Deploy site**. Done.

### Option B — Netlify CLI

```bash
npm install -g netlify-cli
npm run build
netlify deploy --prod --dir=dist
```

---

## Push to GitHub (first time)

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/fpineiro23/grain-elevator-data-transformer.git
git push -u origin main
```

---

## Tech stack

- [React 18](https://react.dev/) + [Vite](https://vitejs.dev/)
- [Tailwind CSS](https://tailwindcss.com/)
- [SheetJS (xlsx)](https://sheetjs.com/) — client-side Excel parsing
- [Lucide React](https://lucide.dev/) — icons

All file processing is done entirely in the browser — no data leaves the user's machine.

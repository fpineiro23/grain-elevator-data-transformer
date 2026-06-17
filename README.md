# Grain Elevator Data Transformer

A browser-based tool that converts raw scale ticket reports (Excel) into Balance API–ready CSVs. No server, no data upload — all processing happens locally in your browser.

**Live demo:** [grain-elevator-data-transformer.netlify.app](https://grain-elevator-data-transformer.netlify.app/)

---

## Try it out

A sample file is included in the repo for testing:

1. Go to [grain-elevator-data-transformer.netlify.app](https://grain-elevator-data-transformer.netlify.app/)
2. Download [`public/dummy_ticket_report.xlsx`](./public/dummy_ticket_report.xlsx)
3. Drop it into the app or click **Choose file**
4. Review the output tabs and download your CSVs

The sample file contains anonymized grain elevator data (477 rows) covering inbound, outbound, internal transfers, and external partner transfers across fictional co-op locations.

---

## What it produces

Upload an `.xlsx` or `.xls` scale ticket report and the tool outputs up to five files:

| Output file | Contents |
|---|---|
| `IN_SCALE_TICKET_YYYYMMDD.csv` | Inbound regular tickets (corn + soybeans) |
| `OUT_SCALE_TICKET_YYYYMMDD.csv` | Outbound regular tickets (corn + soybeans) |
| `TRANSFER_INTERNAL_YYYYMMDD.csv` | Internal co-op transfers (BOL-matched) |
| `TRANSFER_EXT_PARTNER_YYYYMMDD.csv` | External partner transfers |
| `WARNINGS_YYYYMMDD.csv` | Data quality issues, if any |

Non-corn/soybean rows are excluded automatically. The warnings dashboard flags duplicate import IDs, missing required fields, invalid quantities, and unmatched or blank BOL numbers.

---

## Configuration

Enterprise IDs default to generic placeholders in `src/App.jsx`. Swap them for your actual IDs before deploying:

```js
// Co-op enterprise ID
target_enterprise_id: 'grain-coop'      // → your enterprise ID

// External partner enterprise ID and location
'target.enterprise_id': 'ext-partner'   // → partner's enterprise ID
'target.id': 'ext-partner-01'           // → partner's location ID
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

### Option A — Netlify UI

1. Push this repo to GitHub.
2. Go to [netlify.com](https://netlify.com) → **Add new site → Import an existing project**.
3. Connect GitHub and select this repo.
4. Build settings are auto-detected from `netlify.toml` (`npm run build` / `dist`).
5. Click **Deploy site**.

### Option B — Netlify CLI

```bash
npm install -g netlify-cli
npm run build
netlify deploy --prod --dir=dist
```

---

## Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/fpineiro23/grain-elevator-data-transformer.git
git push -u origin main
```

---

## Tech stack

- [React 18](https://react.dev/) + [Vite](https://vitejs.dev/)
- [Tailwind CSS](https://tailwindcss.com/)
- [SheetJS (xlsx)](https://sheetjs.com/) — client-side Excel parsing
- [Lucide React](https://lucide.dev/) — icons

All file processing runs entirely in the browser. No data is sent to any server.

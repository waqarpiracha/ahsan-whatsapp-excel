# Ahsan — WhatsApp → Excel importer (v1)

Twice-a-day batch tool for the **"EV charging solutions by Comet"** WhatsApp
group. Ahsan exports the chat, uploads the `.txt` here **together with his
current EV Master Sheet** (`.xlsx`), picks the export's date format, and gets
back his sheet with new survey rows appended.

## What the parser extracts (templates from the real export)

| Message pattern | → Row |
|---|---|
| `*Survey Report* / *Resurvey Report* / Survey done` + `Power cable X ft/m`, `Grounding cable Y ft/m`, `Earth bore …` | Survey row, Status=Pending |
| `Survey Required (BRAND)` + name / address / phones (positional) | Request row, Status=Pending |
| `EV charger installation complete` / `EV charger <brand> installed …` | Install row, Status=Closed |
| Cable lengths with a client but no survey keyword | Row, Status=Pending |

Field mapping: message date→Date, message time→**Time** (new col), brand→Vehicle
Detal, client number→Client Cell # (normalized to 10 digits), client
name→Client Name, address→Address, city→City (detected, never guessed),
power cable→Power Cable (M) with **ft→m ×0.3048** (original kept in Remarks),
ground cable→Grounding Cable (M), bore→Earth bore Req.
(Required/Done/Available/Not required), remote area→**Remote/Local** (new col,
blank otherwise), installation status→Status. Unparseable-but-data-bearing
messages go to the **Review** sheet — never guessed into rows.

## Run locally

```bash
cd ~/workspace/whatsapp-excel-live/ahsan-importer
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python server.py   # http://localhost:5000
```

## Deploy (same Render account as Noman's app)

Push this folder's repo to GitHub, then Render dashboard → New → Web Service →
Public Git repository → Runtime Docker, Plan Free, Health check path `/`.
`render.yaml` carries the same settings as code.

## Test results (2026-10-17, real export: 16,804 messages)

- **1824 rows parsed**: 775 survey · 694 requests · 325 installs · 30 cable-only
- 804 messages → Review sheet (data-bearing but no client/cable/address)
- 14,176 ignored (system, media, deleted, chatter)
- Workbook: Time col inserted after Date, Remote/Local after Remarks, group
  merges rebuilt (D2:L2 / M2:T2 / U2:X2 / Y2:Z2), original 8 rows intact,
  dates in DD/MM/YYYY, stale pre-filled Sr. No cleared.

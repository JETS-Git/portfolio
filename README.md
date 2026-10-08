# Phil Johnson RN — E-learning Portfolio

Interactive portfolio of nursing e-learning projects (design decision records) plus a
credentials and CPD section. Live at https://jets-git.github.io/portfolio/ and on Render.
Every push to `main` redeploys both automatically.

## Files

| File | What it holds |
|---|---|
| `index.html` | The page. Project case studies live in the `CASES` array near the bottom. |
| `data/cpd.json` | Credentials and the public learning log. The page reads this file. |
| `certificates/` | Certificate PDFs you choose to make public. |
| `tools/import_cpd.py` | Pulls new entries from the CPD log export into `data/cpd.json`. |

## Adding new CPD (monthly routine)

1. Log the activity in the Marr Mooditj CPD form as usual.
2. Export the CPD log to Excel.
3. Run `python tools/import_cpd.py "path/to/export.xlsx"` (needs `pip install openpyxl`).
   New entries arrive as drafts (`"show": false`) and existing entries are never changed.
4. In `data/cpd.json`, give each draft a clean `title` and `provider`, set `"show": true`
   for the ones to publish, and remove `"review"`.
5. Optionally copy the certificate PDF into `certificates/` and set `"certificate"`.
6. Commit and push. GitHub Pages and Render update within a few minutes.

Never commit the Excel export: it contains your work email and SharePoint links (`.gitignore` blocks it).

## Adding a ticket or qualification

Add an object to `credentials` in `data/cpd.json`:

```json
{"name": "Advanced Life Support (ALS2)", "issuer": "Provider name", "kind": "Ticket",
 "awarded": "2026-11-01", "expires": "2027-11-01", "certificate": "certificates/als.pdf"}
```

The page marks it Current, Due (within 60 days of expiry) or Expired automatically.

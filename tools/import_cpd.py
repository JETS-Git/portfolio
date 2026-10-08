#!/usr/bin/env python3
"""Import new entries from the Marr Mooditj CPD log export into data/cpd.json.

Usage:
    python tools/import_cpd.py "Staff CPD Activity Log - Educators.xlsx"

What it does
- Reads the Microsoft Forms Excel export of the CPD log.
- Adds any log entry whose ID is not already in data/cpd.json as a DRAFT
  ("show": false, "review": true). Drafts never appear on the website.
- Never changes entries you have already curated, so your edits are safe.
- Never copies your work email or SharePoint links into the public file.

After importing, open data/cpd.json, tidy each draft's title/provider,
set "show": true for the ones you want public, and delete "review".
"""
import json
import re
import sys
import urllib.parse
from datetime import date, datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Install openpyxl first:  pip install openpyxl")

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data" / "cpd.json"

PLACEHOLDER_HINTS = ("I've left this summary", "paste the article", "I couldn't verify")


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def find_col(headers, *starts):
    for i, h in enumerate(headers):
        if any(norm(h).startswith(s) for s in starts):
            return i
    return None


def guess_title(url):
    if not url:
        return ""
    first = str(url).split(";")[0]
    m = re.search(r"file=([^&]+)", first)
    name = urllib.parse.unquote(m.group(1) if m else first.rstrip("/").split("/")[-1])
    name = re.sub(r"\.(pdf|docx?|xlsx?|png|jpe?g)$", "", name, flags=re.I)
    name = re.sub(r"_?Phillip Johnson( \d+)?$", "", name).replace("_", " ")
    return re.sub(r"\s+", " ", name).strip(" -")


def as_date(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v or "")[:10]


def as_hours(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def main(xlsx):
    data = json.loads(DATA.read_text()) if DATA.exists() else {"credentials": [], "cpd": []}
    known = {e.get("log_id") for e in data["cpd"]}

    ws = openpyxl.load_workbook(xlsx, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    h = rows[0]
    c = {
        "id": find_col(h, "id"),
        "cat": find_col(h, "activity category"),
        "evidence": find_col(h, "upload evidence"),
        "date": find_col(h, "when did it occur"),
        "hours": find_col(h, "how many hours"),
        "reflection": find_col(h, "reflection"),
        "units": find_col(h, "associated units"),
    }
    which = [i for i, x in enumerate(h) if norm(x).endswith("which activity?")]
    if c["id"] is None or c["date"] is None:
        sys.exit("This doesn't look like the CPD log export (no ID / date column).")

    added, warnings = [], []
    for r in rows[1:]:
        if r[c["id"]] is None:
            continue
        log_id = int(r[c["id"]])
        if log_id in known:
            continue
        hours = as_hours(r[c["hours"]])
        reflection = re.sub(r"\s+", " ", str(r[c["reflection"]] or "")).strip()
        units = [u.strip() for u in str(r[c["units"]] or "").split(";") if u.strip()]
        entry = {
            "log_id": log_id,
            "show": False,
            "review": True,
            "title": guess_title(r[c["evidence"]]),
            "provider": "",
            "date": as_date(r[c["date"]]),
            "hours": hours,
            "category": str(r[c["cat"]] or "").replace("Category ", ""),
            "type": next((str(r[i]) for i in which if r[i]), ""),
            "units": units,
            "reflection": reflection,
            "certificate": None,
        }
        data["cpd"].append(entry)
        added.append(entry)
        if hours is None:
            warnings.append(f"ID {log_id}: hours value '{r[c['hours']]}' is not a number")
        if any(p.lower() in reflection.lower() for p in PLACEHOLDER_HINTS):
            warnings.append(f"ID {log_id}: reflection looks like placeholder text, rewrite it")

    data["cpd"].sort(key=lambda e: e["date"], reverse=True)
    DATA.parent.mkdir(exist_ok=True)
    DATA.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n")

    print(f"{len(added)} new draft(s) added to {DATA.relative_to(ROOT)}")
    for e in added:
        print(f"  ID {e['log_id']:>3}  {e['date']}  {e['hours'] or '?':>5} h  {e['title'] or '(no title)'}")
    for w in warnings:
        print("  ! " + w)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])

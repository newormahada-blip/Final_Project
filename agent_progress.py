#!/usr/bin/env python3
"""Writes agent_progress.json for the Arnona_Agent site — no addresses, only:
  total   : number of addresses in Adresses.xlsx
  blocks  : the neighborhood runs of Adresses.xlsx, in file order:
            {"neighborhood", "start" (inclusive), "end" (exclusive), "count"}
The routine's position is NOT here: the site reads it from ./index ("index = N", 0-based into the
same list), so a routine push of ./index never needs a bot commit after it.

Run it only when Adresses.xlsx changes (after `storage.py push Adresses.xlsx`):
  python3 agent_progress.py        # Adresses.xlsx local, else from the private bucket
or run the "agent progress" workflow by hand (needs the SUPABASE_SERVICE_KEY repo secret).
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "agent_progress.json"


def blocks_from_addresses():
    import openpyxl
    import storage
    path = storage.ensure_local("Adresses.xlsx")
    ws = openpyxl.load_workbook(path, read_only=True).active
    # Same filter as the routine: skip the header row and rows with no address.
    hoods = [((r[1] if len(r) > 1 else "") or "").strip()
             for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]
    blocks = []
    for i, h in enumerate(hoods):
        if blocks and blocks[-1]["neighborhood"] == h:
            blocks[-1]["end"] = i + 1
            blocks[-1]["count"] += 1
        else:
            blocks.append({"neighborhood": h, "start": i, "end": i + 1, "count": 1})
    return len(hoods), blocks


def main(argv):
    total, blocks = blocks_from_addresses()
    data = {"total": total, "blocks": blocks}
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"agent_progress.json: total={total} blocks={len(blocks)}")


if __name__ == "__main__":
    main(sys.argv[1:])

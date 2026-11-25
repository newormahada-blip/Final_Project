#!/usr/bin/env python3
"""Writes agent_progress.json for the Arnona_Agent site — no addresses, only:
  index   : the routine's position (from ./index), 0-based into the address list
  total   : number of addresses in Adresses.xlsx
  blocks  : the neighborhood runs of Adresses.xlsx, in file order:
            {"neighborhood", "start" (inclusive), "end" (exclusive), "count"}

  python3 agent_progress.py            # blocks from Adresses.xlsx (local, else from storage)
  python3 agent_progress.py --index-only   # keep total/blocks from the existing JSON, refresh index
"""
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "agent_progress.json"


def read_index():
    m = re.search(r"\d+", (ROOT / "index").read_text(encoding="utf-8"))
    return int(m.group()) if m else 0


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
    if "--index-only" in argv and OUT.exists():
        old = json.loads(OUT.read_text(encoding="utf-8"))
        total, blocks = old["total"], old["blocks"]
    else:
        total, blocks = blocks_from_addresses()
    data = {
        "index": read_index(),
        "total": total,
        "blocks": blocks,
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"agent_progress.json: index={data['index']} total={total} blocks={len(blocks)}")


if __name__ == "__main__":
    main(sys.argv[1:])

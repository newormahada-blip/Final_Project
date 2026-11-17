#!/usr/bin/env python3
"""עסקים מוכרים — בודק אם עסק כבר נמצא בעבר (במערכת או בריצה קודמת של הרוטינה).

  check "<שם העסק>" "<כתובת>"        → KNOWN (כבר קיים, לא להכניס לדוח) / NEW (חדש)
  check-report "<דוח>.xlsx"           → בודק כל שורה בדוח ומדפיס KNOWN/NEW לכל אחת
  add "<דוח>.xlsx" --source "<מקור>"  → מוסיף לקובץ את העסקים מהדוח (רק מי שלא קיים כבר)
  sync                                → מוסיף לקובץ את העסקים שבמערכת (האתר) וכותב את רשימת סוגי העסק

בכל בדיקה (check / check-report / add) נבדקים גם כל העסקים שנמצאים עכשיו במערכת —
מתוך reports/agent_snapshot.json בריפו Arnona_Agent, שמתעדכן כל 3 שעות מ-Supabase.
כך עסק שהועלה לאתר אחרי הריצה הקודמת (או שנמצא בריצה שנקטעה) לא ידווח שוב.

ההתאמה: אותה כתובת (רחוב + מספר בית) וגם שם דומה — אחרי ניקוי תארים (עו"ד, ד"ר, בע"מ...),
סימני פיסוק וסדר מילים. כך "בן צבי נרדה" ו-"נרדה בן-צבי, עו\"ד" באותה כתובת נחשבים אותו עסק.
"""
import argparse
import json
import os
import time
import urllib.request
import re
import sys
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
KNOWN_FILE = os.path.join(HERE, "known_businesses.xlsx")
BUSINESS_TYPES_FILE = os.path.join(HERE, "business_types.txt")
SNAPSHOT_URL = "https://raw.githubusercontent.com/urbandetective007/Arnona_Agent/main/reports/agent_snapshot.json"
SNAPSHOT_CACHE = "/tmp/agent_snapshot.json"
SYSTEM_SOURCE = "מערכת (אתר)"
HEADERS = ["שם העסק", "כתובת", "שכונה", "סוג עסק", "קישור", "מקור", "תאריך הוספה"]

# מילים שאינן חלק מזהה של השם
NAME_STOPWORDS = {
    "עוד", "עו״ד", "עורך", "עורכת", "עורכי", "דין", "ונוטריון", "נוטריון", "משרד", "משרדי",
    "דר", "ד״ר", "דוקטור", "פרופ", "פרופסור", "מר", "גב", "הרב", "רב",
    "בעמ", "בע״מ", "חברה", "ושות", "ושותפים", "שותפים", "קבוצת", "ירושלים",
    "the", "ltd", "inc", "co", "dr", "adv",
}
ADDRESS_DROP = {"ירושלים", "jerusalem", "רחוב", "רח", "ישראל", "israel"}
STREET_PREFIX = {"שדרות": "שד", "שדרת": "שד", "שד": "שד", "דרך": "דרך", "סמטת": "סמ", "סמ": "סמ", "כיכר": "ככר", "ככר": "ככר"}


def _clean(text):
    t = str(text or "").lower()
    t = t.replace('"', "").replace("״", "").replace("'", "").replace("׳", "").replace("`", "")
    t = re.sub(r"[-–—_.,;:()\[\]/\\|+]", " ", t)
    return t.split()


def address_key(address):
    """('<רחוב מנורמל>', '<מספר בית>') — או None אם אין מספר בית."""
    raw = str(address or "")
    raw = re.sub(r"\b\d{7}\b", " ", raw)  # מיקוד
    raw = re.sub(r"(\d+)\s*/\s*\d+", r"\1", raw)  # 1/2 → 1
    raw = re.sub(r"(\d+)\s*[אבגדה]\b", r"\1", raw)  # 12א → 12
    tokens = [t for t in _clean(raw) if t not in ADDRESS_DROP]
    numbers = [t for t in tokens if t.isdigit()]
    if not numbers:
        return None
    words = [STREET_PREFIX.get(t, t) for t in tokens if not t.isdigit()]
    return (" ".join(sorted(words)), numbers[0])


ARABIC_STOPWORDS = {"الدكتور", "دكتور", "د", "المحامي", "محامي", "مكتب", "شركة", "القدس", "ال"}
ARABIC_LETTERS = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه", "ـ": ""})
ARABIC_DIACRITICS = re.compile(r"[ً-ْٰ]")


def _norm_token(t):
    """מאחד תעתיקים שונים של אותו שם ערבי: "מוחמד" = "מחמד", "אלחוסייני" = "חוסייני", "أحمد" = "احمد"."""
    if re.search(r"[؀-ۿ]", t):
        t = ARABIC_DIACRITICS.sub("", t).translate(ARABIC_LETTERS)
        return t[2:] if t.startswith("ال") and len(t) > 4 else t
    if t.startswith("אל") and len(t) >= 5:
        t = t[2:]
    return t[0] + re.sub(r"[וי]", "", t[1:]) if len(t) > 2 else t


def name_tokens(name):
    tokens = (t for t in _clean(name) if t not in NAME_STOPWORDS and t not in ARABIC_STOPWORDS and t not in {"אל", "א"})
    return {_norm_token(t) for t in tokens if len(t) > 1}


def names_match(a, b):
    ta, tb = name_tokens(a), name_tokens(b)
    if not ta or not tb:
        return False
    if ta <= tb or tb <= ta:
        return True
    return len(ta & tb) / len(ta | tb) >= 0.5


def load_known():
    if not os.path.exists(KNOWN_FILE):
        return []
    ws = openpyxl.load_workbook(KNOWN_FILE, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    idx = {h: i for i, h in enumerate(rows[0]) if h}
    out = []
    for r in rows[1:]:
        rec = {h: (r[i] if i < len(r) else None) for h, i in idx.items()}
        if rec.get("שם העסק"):
            out.append(rec)
    return out


def load_snapshot():
    """העסקים וסוגי העסק שבמערכת עכשיו. נשמר ב-/tmp לחצי שעה, כדי לא להוריד בכל בדיקה.
    אם אי אפשר להוריד — מדפיס אזהרה (לכתוב אותה בטיוטה) ומחזיר None."""
    try:
        if os.path.exists(SNAPSHOT_CACHE) and time.time() - os.path.getmtime(SNAPSHOT_CACHE) < 1800:
            with open(SNAPSHOT_CACHE, encoding="utf-8") as f:
                return json.load(f)
        with urllib.request.urlopen(SNAPSHOT_URL, timeout=30) as r:
            data = r.read()
        snap = json.loads(data.decode("utf-8"))
        with open(SNAPSHOT_CACHE, "wb") as f:
            f.write(data)
        return snap
    except Exception as e:  # noqa: BLE001
        print(f"⚠️ SYNC FAILED — העסקים שבמערכת לא נטענו ({e}). הבדיקה לפי הקובץ בלבד. לכתוב את זה בטיוטה.",
              file=sys.stderr)
        return None


def system_records(known):
    """עסקים שבמערכת ואינם בקובץ — כרשומות במבנה של הקובץ (מקור: מערכת (אתר))."""
    snap = load_snapshot()
    # מפתח לכל כתובת: רחוב + מספר, ובלי מספר בית — הכתובת המנוקה כולה (אחרת עסק בלי מספר היה נוסף שוב בכל סנכרון).
    akey = lambda a: address_key(a) or ("", " ".join(t for t in _clean(a) if t not in ADDRESS_DROP))
    by_address = {}
    for rec in known:
        by_address.setdefault(akey(rec.get("כתובת")), []).append(rec["שם העסק"])
    out = []
    for b in (snap or {}).get("businesses", []):
        name, address = (b.get("name") or "").strip(), (b.get("address") or "").strip()
        if not name or not address:
            continue
        k = akey(address)
        if any(names_match(name, other) for other in by_address.get(k, [])):
            continue
        by_address.setdefault(k, []).append(name)
        out.append({"שם העסק": name, "כתובת": address, "שכונה": None, "סוג עסק": b.get("type") or None,
                    "קישור": None, "מקור": SYSTEM_SOURCE, "תאריך הוספה": date.today().isoformat()})
    return out


def find_match(name, address, known):
    """מחזיר (רשומה תואמת, 'same_address' / 'name_only') או (None, None)."""
    key = address_key(address)
    name_only = None
    for rec in known:
        rkey = address_key(rec.get("כתובת"))
        if key and rkey and key == rkey and names_match(name, rec["שם העסק"]):
            return rec, "same_address"
        if name_only is None and name_tokens(name) and name_tokens(name) == name_tokens(rec["שם העסק"]) and len(name_tokens(name)) >= 2:
            name_only = rec
    if name_only:
        return name_only, "name_only"
    return None, None


def save_known(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "עסקים מוכרים"
    ws.sheet_view.rightToLeft = True
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    for r in rows:
        ws.append([r.get(h) for h in HEADERS])
    for col, width in zip("ABCDEFG", (38, 32, 18, 26, 45, 28, 13)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(KNOWN_FILE)


def read_report(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    idx = {str(h).strip(): i for i, h in enumerate(rows[0]) if h}
    if "שם העסק" not in idx or "כתובת" not in idx:
        sys.exit(f"❌ {path}: אין עמודות 'שם העסק' ו-'כתובת' בשורה הראשונה")
    get = lambda r, h: (str(r[idx[h]]).strip() if h in idx and idx[h] < len(r) and r[idx[h]] is not None else None)
    out = []
    for r in rows[1:]:
        name, addr = get(r, "שם העסק"), get(r, "כתובת")
        if name and addr:
            link = next((get(r, h) for h in ("קישור 1", "קישור", "קישור 2", "קישור 3") if get(r, h)), None)
            out.append({"שם העסק": name, "כתובת": addr, "שכונה": get(r, "שכונה"),
                        "סוג עסק": get(r, "סוג עסק") or get(r, "סוג העסק"), "קישור": link})
    return out


def report_line(name, address, rec, kind):
    if kind == "same_address":
        return f"KNOWN | {name} | {address} | כבר קיים: {rec['שם העסק']} | {rec['כתובת']} | מקור: {rec.get('מקור')}"
    if kind == "name_only":
        return (f"NEW (שם זהה בכתובת אחרת — לבדוק) | {name} | {address} | "
                f"קיים: {rec['שם העסק']} | {rec['כתובת']} | מקור: {rec.get('מקור')}")
    return f"NEW | {name} | {address}"


def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("check")
    c.add_argument("name")
    c.add_argument("address")
    cr = sub.add_parser("check-report")
    cr.add_argument("report")
    a = sub.add_parser("add")
    a.add_argument("report")
    a.add_argument("--source", required=True)
    sub.add_parser("sync")
    args = p.parse_args()

    known = load_known()
    if args.cmd == "sync":
        extra = system_records(known)
        if extra:
            save_known(known + extra)
        snap = load_snapshot() or {}
        types = snap.get("business_types", [])
        if types:
            with open(BUSINESS_TYPES_FILE, "w", encoding="utf-8") as f:
                f.write("\n".join(types) + "\n")
        print(f"✅ SYNC: {len(snap.get('businesses', []))} עסקים במערכת (נכון ל-{snap.get('generated_at', '?')}), "
              f"{len(extra)} נוספו לקובץ (סה\"כ {len(known) + len(extra)}). {len(types)} סוגי עסק → {BUSINESS_TYPES_FILE}")
        return
    # מה שבמערכת נבדק תמיד, גם אם עוד לא נשמר בקובץ.
    known = known + system_records(known)
    if args.cmd == "check":
        rec, kind = find_match(args.name, args.address, known)
        print(report_line(args.name, args.address, rec, kind))
        sys.exit(1 if kind == "same_address" else 0)
    if args.cmd == "check-report":
        n_known = 0
        for b in read_report(args.report):
            rec, kind = find_match(b["שם העסק"], b["כתובת"], known)
            n_known += kind == "same_address"
            print(report_line(b["שם העסק"], b["כתובת"], rec, kind))
        print(f"סה\"כ KNOWN: {n_known}")
        sys.exit(1 if n_known else 0)
    if args.cmd == "add":
        added = 0
        for b in read_report(args.report):
            if find_match(b["שם העסק"], b["כתובת"], known)[1] == "same_address":
                continue
            b["מקור"] = args.source
            b["תאריך הוספה"] = date.today().isoformat()
            known.append(b)
            added += 1
        save_known(known)
        print(f"✅ נוספו {added} עסקים. סה\"כ בקובץ: {len(known)}")


if __name__ == "__main__":
    main()

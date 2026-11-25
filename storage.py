#!/usr/bin/env python3
"""Private data files for the Game routine and the tools, kept in Supabase Storage
(private bucket `private-data`, folder `final_project/`) instead of this public repo.

Key:
  - SUPABASE_SERVICE_KEY set  -> sent in the `apikey` header (and as Bearer when it is a JWT).
  - not set                   -> a placeholder is sent; the environment's network proxy
                                  replaces it with the real key (cloud sessions with a network secret).
SUPABASE_URL overrides the project URL.

Each object's key is ASCII (Supabase rejects Hebrew and ':'); final_project/manifest.json
maps every repo path to its key and sha256.

Usage:
  python3 storage.py pull [PATH ...]   # download to the same repo paths (default: the routine's files)
  python3 storage.py push PATH ...     # upload local files, update the manifest
  python3 storage.py list
"""
import hashlib
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
URL = os.environ.get("SUPABASE_URL", "https://mcsygsqfyuaexxxwsgem.supabase.co").rstrip("/")
BUCKET = "private-data"
PREFIX = "final_project/"
PLACEHOLDER = "sb_secret_placeholder"

# What the Game routine needs at the start of a run.
ROUTINE_FILES = [
    "Adresses.xlsx",
    "mhb/candidates.json",
    "known_businesses/known_businesses.xlsx",
]


def _headers(extra=None):
    key = os.environ.get("SUPABASE_SERVICE_KEY") or PLACEHOLDER
    h = {"apikey": key}
    if key.startswith("eyJ"):          # legacy service_role JWT also needs Authorization
        h["Authorization"] = "Bearer " + key
    h.update(extra or {})
    return h


def _obj_url(key):
    return f"{URL}/storage/v1/object/{BUCKET}/{urllib.parse.quote(key)}"


def _request(method, key, data=None, content_type=None):
    extra = {"content-type": content_type, "x-upsert": "true"} if data is not None else None
    req = urllib.request.Request(_obj_url(key), data=data, method=method, headers=_headers(extra))
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return r.read()
    except urllib.error.HTTPError as e:
        body = e.read()[:300].decode("utf-8", "replace")
        raise SystemExit(f"storage: {method} {key} -> HTTP {e.code} {body}\n"
                         "Check SUPABASE_SERVICE_KEY or the environment's network secret for "
                         f"{urllib.parse.urlparse(URL).hostname} (header apikey).")


def object_key(path):
    k = (path.replace("excel outputs/", "excel_outputs/")
             .replace("דוח נכסים לתאריך:", "report_")
             .replace("כתובות ייחודיות ", "unique_addresses_")
             .replace(" (1)", "_1").replace(" ", "_"))
    if not re.fullmatch(r"[A-Za-z0-9._/-]+", k):
        raise SystemExit(f"storage: no ASCII key for {path!r}; add a rule to object_key()")
    return PREFIX + k


def load_manifest():
    return json.loads(_request("GET", PREFIX + "manifest.json"))


def save_manifest(man):
    data = json.dumps(man, ensure_ascii=False, indent=1).encode("utf-8")
    _request("POST", PREFIX + "manifest.json", data, "application/json")


def download(path, dest=None):
    """Download repo path `path` to `dest` (default: the same path in the repo). Returns the Path."""
    entry = {e["path"]: e for e in load_manifest()}.get(path)
    key = entry["key"] if entry else object_key(path)
    data = _request("GET", key)
    if entry and hashlib.sha256(data).hexdigest() != entry["sha256"]:
        raise SystemExit(f"storage: checksum mismatch for {path}")
    out = Path(dest) if dest else ROOT / path
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    return out


def ensure_local(path):
    """Local copy if it exists (old checkout), else download it. Returns the Path."""
    p = ROOT / path
    return p if p.exists() else download(path)


def upload(path, src=None):
    data = Path(src or ROOT / path).read_bytes()
    key = object_key(path)
    _request("POST", key, data, "application/octet-stream")
    sha = hashlib.sha256(data).hexdigest()
    if hashlib.sha256(_request("GET", key)).hexdigest() != sha:
        raise SystemExit(f"storage: upload of {path} did not verify")
    man = [e for e in load_manifest() if e["path"] != path]
    man.append({"path": path, "key": key, "sha256": sha, "bytes": len(data), "source": "storage.py push"})
    save_manifest(man)
    return sha


def main(argv):
    if not argv or argv[0] not in ("pull", "push", "list"):
        print(__doc__)
        return 2
    cmd, paths = argv[0], argv[1:]
    if cmd == "list":
        for e in load_manifest():
            print(f'{e["bytes"]:>10}  {e["path"]}')
    elif cmd == "pull":
        for p in paths or ROUTINE_FILES:
            print("pulled", download(p).relative_to(ROOT))
    else:
        if not paths:
            raise SystemExit("storage: push needs at least one path")
        for p in paths:
            print("pushed", p, upload(p)[:12])
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

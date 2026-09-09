"""
NHGIS extract submit / poll / download
=====================================
Reads nhgis_extract_spec.json (written by nhgis_search.py), submits it to the
IPUMS API v2 as one or two extracts, waits for them to finish, downloads the
ZIPs, unzips them, and writes a manifest with extract numbers and file lists.

Setup
-----
    pip install requests python-dotenv
    .env in this directory:  NHGIS_API_KEY=your_key   (account.ipums.org/api_keys)

Usage
-----
    python nhgis_pull.py                   # submit both extracts, wait, download
    python nhgis_pull.py --dry-run         # print the request bodies, submit nothing
    python nhgis_pull.py --skip-historical # primary ACS extract only
    python nhgis_pull.py --resume 123456   # skip submission, poll/download extract #123456
    python nhgis_pull.py --out data/nhgis  # download directory (default)

Reference: https://developer.ipums.org/docs/v2/workflows/create_extracts/nhgis_data/
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import zipfile
from pathlib import Path

import requests

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

SCRIPT_DIR = Path(__file__).parent
if load_dotenv:
    load_dotenv(SCRIPT_DIR / ".env")

API = "https://api.ipums.org"
PARAMS = {"collection": "nhgis", "version": "2"}
SPEC_PATH = SCRIPT_DIR / "nhgis_extract_spec.json"
POLL_SECONDS = 20
MAX_WAIT_MINUTES = 90


def headers() -> dict:
    key = os.environ.get("NHGIS_API_KEY") or os.environ.get("IPUMS_API_KEY")
    if not key:
        sys.exit("ERROR: NHGIS_API_KEY not set (put it in .env or export it).")
    return {"Authorization": key}


def load_bodies(spec_path: Path, skip_historical: bool) -> list[dict]:
    spec = json.loads(spec_path.read_text())
    hist = spec.pop("_historical_extract_separate", None)
    bodies = [spec]
    if hist and not skip_historical:
        hist = {k: v for k, v in hist.items() if not k.startswith("_") and k not in ("note", "california_instances")}
        bodies.append(hist)
    return bodies


def submit(body: dict) -> int:
    r = requests.post(f"{API}/extracts", params=PARAMS, headers={**headers(), "Content-Type": "application/json"},
                      json=body, timeout=60)
    if r.status_code >= 400:
        sys.exit(f"ERROR {r.status_code} submitting '{body.get('description')}':\n{r.text}")
    number = r.json()["number"]
    print(f"  submitted '{body.get('description')}' -> extract #{number}")
    return number


def status(number: int) -> dict:
    r = requests.get(f"{API}/extracts/{number}", params=PARAMS, headers=headers(), timeout=60)
    r.raise_for_status()
    return r.json()


def wait(number: int) -> dict:
    deadline = time.time() + MAX_WAIT_MINUTES * 60
    while time.time() < deadline:
        info = status(number)
        st = info.get("status")
        print(f"  extract #{number}: {st}")
        if st == "completed":
            return info
        if st in ("failed", "canceled"):
            sys.exit(f"ERROR: extract #{number} {st}")
        time.sleep(POLL_SECONDS)
    sys.exit(f"ERROR: extract #{number} not finished after {MAX_WAIT_MINUTES} min; rerun with --resume {number}")


def download(info: dict, out: Path) -> list[Path]:
    out.mkdir(parents=True, exist_ok=True)
    number = info["number"]
    got = []
    for kind, link in (info.get("downloadLinks") or {}).items():
        if kind == "codebookPreview":
            continue
        url = link["url"] if isinstance(link, dict) else link
        dest = out / f"nhgis{number:04d}_{kind}.zip"
        print(f"  downloading {kind} -> {dest.name}")
        with requests.get(url, headers=headers(), stream=True, timeout=600) as r:
            r.raise_for_status()
            with open(dest, "wb") as f:
                for chunk in r.iter_content(1 << 20):
                    f.write(chunk)
        got.append(dest)
        with zipfile.ZipFile(dest) as z:
            z.extractall(out / dest.stem)
        # GIS zips contain zipped shapefiles; unzip one level further
        for inner in (out / dest.stem).rglob("*.zip"):
            with zipfile.ZipFile(inner) as z:
                z.extractall(inner.parent)
    return got


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", default=str(SPEC_PATH))
    ap.add_argument("--out", default=str(SCRIPT_DIR / "data" / "nhgis"))
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--skip-historical", action="store_true")
    ap.add_argument("--resume", type=int, nargs="*", help="extract number(s) already submitted")
    args = ap.parse_args()
    out = Path(args.out)

    bodies = load_bodies(Path(args.spec), args.skip_historical)
    if args.dry_run:
        for b in bodies:
            print(json.dumps(b, indent=2))
        return

    numbers = args.resume or [submit(b) for b in bodies]
    manifest = {"submitted_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()), "extracts": []}
    for n in numbers:
        info = wait(n)
        files = download(info, out)
        manifest["extracts"].append({
            "number": n, "description": info.get("description"),
            "datasets": info.get("datasets"), "shapefiles": info.get("shapefiles"),
            "geographicExtents": info.get("geographicExtents"),
            "files": [str(p.relative_to(out)) for p in files],
        })
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2))
    print(f"\nDone. Files in {out}/ ; manifest.json records extract numbers for the appendix.")
    print("Reminder: NHGIS keeps download links for two weeks; the extract numbers can be resubmitted from Extracts History.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Geocode GVA incident addresses using the U.S. Census Bureau batch geocoder.

API docs: https://geocoding.geo.census.gov/geocoder/Geocoding_Services_API.html

Reads the GVA export CSV, cleans addresses (e.g. "700 block of SE 12th Ave"
-> "700 SE 12th Ave"), submits them in batches to the Census
locations/addressbatch endpoint, and writes a new CSV with Latitude/Longitude
columns suitable for "XY Table To Point" in ArcGIS Pro.

Usage:
    python3 geocode_census.py [input.csv]

The output is written next to the input as <input-name>_geocoded.csv.
"""

import csv
import io
import re
import sys
import time
from pathlib import Path

import requests

HERE = Path(__file__).parent
DEFAULT_INPUT = HERE / "export-75757714-e721-45eb-8b80-0ccd56bb6957.csv"
INPUT_CSV = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_INPUT
OUTPUT_CSV = INPUT_CSV.with_name(INPUT_CSV.stem + "_geocoded.csv")

BATCH_URL = "https://geocoding.geo.census.gov/geocoder/locations/addressbatch"
BENCHMARK = "Public_AR_Current"
CHUNK_SIZE = 500          # addresses per request (API max is 10,000)
MAX_RETRIES = 3
TIMEOUT_SECONDS = 300


def clean_street(address: str) -> str:
    """Normalize a GVA street string so the Census geocoder can parse it."""
    addr = (address or "").strip()
    if not addr or addr.upper() in {"N/A", "NA", "UNKNOWN"}:
        return ""
    # "700 block of SE 12th Ave" -> "700 SE 12th Ave"
    addr = re.sub(r"^(\d+)\s+blo?ck\s+(?:of\s+)?", r"\1 ", addr, flags=re.I)
    # "block of Main St" (no number) -> "Main St"
    addr = re.sub(r"^blo?ck\s+(?:of\s+)?", "", addr, flags=re.I)
    return addr.strip()


def read_input(path: Path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return reader.fieldnames, list(reader)


def build_batch_payload(rows_chunk):
    """Build the no-header CSV the batch endpoint expects:
    Unique ID, Street address, City, State, ZIP"""
    buf = io.StringIO()
    writer = csv.writer(buf)
    for idx, row in rows_chunk:
        writer.writerow([
            idx,
            clean_street(row.get("Address", "")),
            (row.get("City Or County") or "").strip(),
            (row.get("State") or "").strip(),
            "",  # no ZIP in the GVA export
        ])
    return buf.getvalue()


def geocode_chunk(payload_csv: str) -> dict:
    """POST one batch to the Census geocoder; return {id: result_row}."""
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(
                BATCH_URL,
                files={"addressFile": ("addresses.csv", payload_csv, "text/csv")},
                data={"benchmark": BENCHMARK},
                timeout=TIMEOUT_SECONDS,
            )
            resp.raise_for_status()
            break
        except requests.RequestException as exc:
            last_error = exc
            wait = 10 * attempt
            print(f"  request failed ({exc}); retry {attempt}/{MAX_RETRIES} in {wait}s")
            time.sleep(wait)
    else:
        raise RuntimeError(f"Batch failed after {MAX_RETRIES} retries: {last_error}")

    # Response columns: id, input address, match status, match type,
    # matched address, "lon,lat", tigerline id, side
    results = {}
    for fields in csv.reader(io.StringIO(resp.text)):
        if not fields:
            continue
        rec_id = fields[0]
        status = fields[2] if len(fields) > 2 else "No_Match"
        result = {"status": status, "match_type": "", "matched_address": "",
                  "lon": "", "lat": ""}
        if status == "Match" and len(fields) >= 6:
            result["match_type"] = fields[3]
            result["matched_address"] = fields[4]
            lon_lat = fields[5].split(",")
            if len(lon_lat) == 2:
                result["lon"], result["lat"] = lon_lat[0].strip(), lon_lat[1].strip()
        results[rec_id] = result
    return results


def main():
    fieldnames, rows = read_input(INPUT_CSV)
    print(f"Read {len(rows)} rows from {INPUT_CSV.name}")

    indexed = list(enumerate(rows))
    geocodable = [(i, r) for i, r in indexed if clean_street(r.get("Address", ""))]
    skipped = len(indexed) - len(geocodable)
    print(f"{len(geocodable)} rows have a usable address ({skipped} skipped)")

    all_results = {}
    chunks = [geocodable[i:i + CHUNK_SIZE] for i in range(0, len(geocodable), CHUNK_SIZE)]
    for n, chunk in enumerate(chunks, 1):
        print(f"Geocoding batch {n}/{len(chunks)} ({len(chunk)} addresses)...")
        all_results.update(geocode_chunk(build_batch_payload(chunk)))

    out_fields = fieldnames + ["Cleaned_Address", "Match_Status", "Match_Type",
                               "Matched_Address", "Longitude", "Latitude"]
    matched = 0
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=out_fields)
        writer.writeheader()
        for idx, row in indexed:
            result = all_results.get(str(idx), {})
            out = dict(row)
            out["Cleaned_Address"] = clean_street(row.get("Address", ""))
            out["Match_Status"] = result.get("status", "Not_Submitted")
            out["Match_Type"] = result.get("match_type", "")
            out["Matched_Address"] = result.get("matched_address", "")
            out["Longitude"] = result.get("lon", "")
            out["Latitude"] = result.get("lat", "")
            if out["Latitude"]:
                matched += 1
            writer.writerow(out)

    total = len(rows)
    print(f"\nDone: {matched}/{total} rows geocoded "
          f"({matched / total:.1%}) -> {OUTPUT_CSV}")
    print("Coordinates are WGS84 (lat/lon). In ArcGIS Pro use "
          "'XY Table To Point' with X=Longitude, Y=Latitude, GCS_WGS_1984.")


if __name__ == "__main__":
    main()

"""
NHGIS Metadata Search for Thesis Determinants  (v2, Sept 2026)
================================================================
Searches the full IPUMS NHGIS catalog via the IPUMS API v2 metadata
endpoints, matches every determinant in the Combined Determinant Roadmap
against it, and writes the results into a copy of the Determinants workbook.
It also resolves the curated table list for the locked ACS vintage into a
ready-to-submit extract definition (nhgis_extract_spec.json).

What changed from v1 (April 2026)
---------------------------------
* Vocabulary.  v1 matched literature phrases ("poverty rate", "unemployment
  rate", "crowding", "broadband") as raw substrings.  NHGIS descriptions use
  Census phrasing ("Poverty Status in the Past 12 Months", "Employment Status",
  "Occupants per Room", "Internet Subscriptions"), so those hit zero tables.
  v2 adds a Census-vocabulary layer per determinant on top of the spec keywords.
* Word boundaries.  v1 substring matching let "ICE" hit every table containing
  "service" or "price" (416 tables).  v2 matches on word boundaries.
* Geography.  v1 inferred geography from the dataset description and dropped
  any dataset whose description lacked "tract"/"block group", which discarded
  the entire 2020 DHC ("Blocks & Larger Areas") and 1970 counts.  v2 uses a
  rule table for description phrasing and, when an API key is present, fetches
  each relevant dataset's real geogLevels and geographicInstances.
* Nothing is excluded up front.  The whole catalog is searched; results are
  partitioned into the locked vintage, other modern datasets (2010+),
  historical tract datasets (1940-1970), and time series tables, so a wider
  match can never hide the right ACS table.  v1 preferred time series and cut
  every list to three rows.
* Vintage.  The lock is now ACS 2018-2022 (2020 tracts).  v1 used 2010-2014 and
  2019-2023.  Availability in neighbouring vintages is still reported.
* Curated codes.  The Sept 3 table list, plus documented extras, is encoded per
  determinant; the script validates each code against the catalog, resolves
  its a/b dataset, and writes the extract JSON.
* Caches are refreshable (--refresh) and their age is reported.  The API key is
  only required when something must be fetched.

Setup
-----
    pip install requests pandas openpyxl python-dotenv
    .env in this directory:  NHGIS_API_KEY=your_key   (key from account.ipums.org/api_keys)

Usage
-----
    python nhgis_search.py                 # match + write Excel + extract JSON
    python nhgis_search.py --refresh       # re-download catalog caches first
    python nhgis_search.py --vintage 2020_2024 --extent 060
    python nhgis_search.py --no-details    # skip per-dataset API calls (offline)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import difflib
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from dotenv import load_dotenv
except ImportError:  # dotenv is optional; an exported env var works too
    load_dotenv = None

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).parent
if load_dotenv:
    load_dotenv(SCRIPT_DIR / ".env")

API_BASE = "https://api.ipums.org/metadata"      # v2 form: /metadata/<endpoint>?collection=nhgis&version=2
API_PARAMS = {"collection": "nhgis", "version": "2"}

CACHE_DATASETS = SCRIPT_DIR / "nhgis_datasets_cache.json"
CACHE_DATA_TABLES = SCRIPT_DIR / "nhgis_data_tables_cache.json"
CACHE_TIME_SERIES = SCRIPT_DIR / "nhgis_time_series_cache.json"
CACHE_DATASET_DETAIL = SCRIPT_DIR / "nhgis_dataset_detail_cache.json"

KEYWORD_SPEC_PATH = SCRIPT_DIR / "Determinant_Keyword_Spec.xlsx"
DETERMINANTS_PATH = SCRIPT_DIR / "Determinants.xlsx"
OUTPUT_PATH = SCRIPT_DIR / "Determinants_NHGIS_Matched_v2.xlsx"
EXTRACT_SPEC_PATH = SCRIPT_DIR / "nhgis_extract_spec.json"

# Locked design (Sept 1 2026): 2020 tracts, ACS 2018-2022, California.
DEFAULT_VINTAGE = "2018_2022"
DEFAULT_EXTENT = "060"                     # NHGIS state code = FIPS x 10 (CA = 06 -> 060)
DEFAULT_GEOG = "tract"
DEFAULT_SHAPEFILE = "us_tract_2020_tl2020"
COMPARE_VINTAGES = ["2020_2024", "2019_2023", "2017_2021", "2016_2020", "2015_2019", "2010_2014"]

TARGET_GEOGS = {"tract", "blck_grp", "block"}
CACHE_STALE_DAYS = 60

# Historical tract-level datasets worth surfacing separately (HOLC-era controls).
HISTORICAL_TRACT_DATASETS = {
    "1940_tPH_Major", "1940_tPH_NYC", "1950_tPH_Major", "1960_tPH",
    "1970_Cnt2", "1970_Cnt3", "1970_Cnt4Pa", "1970_Cnt4Pb", "1970_Cnt4H",
}

# ---------------------------------------------------------------------------
# Census-vocabulary layer.  Keyed by the "#" column of the Keyword Spec.
# These are phrases as NHGIS/Census actually write table titles.  They are
# matched on word boundaries, case-insensitively, in description and universe.
# Weight 2 = strong signal, 1 = supporting.
# ---------------------------------------------------------------------------
CENSUS_TERMS: dict[int, list[tuple[str, int]]] = {
    3:  [("negro population", 2), ("population by race", 1), ("nonwhite", 1)],
    4:  [("tenure", 2), ("renter occupied", 2), ("owner occupied", 2)],
    5:  [("poverty status", 2), ("ratio of income to poverty level", 2), ("poverty", 1)],
    6:  [("poverty status", 1), ("employment status", 1), ("public assistance income", 1),
         ("educational attainment", 1), ("household type", 1), ("median household income", 1)],
    7:  [("gini index", 2), ("household income in the past 12 months", 2), ("income inequality", 2)],
    8:  [("educational attainment", 2), ("years of school completed", 2), ("school enrollment", 1)],
    9:  [("employment status", 2), ("labor force status", 2), ("unemployed", 2), ("work experience", 1)],
    10: [("own children under 18 years", 2), ("family type", 2), ("household type", 1),
         ("female householder, no spouse", 2), ("living arrangements", 1), ("single parent", 2)],
    11: [("year householder moved into unit", 2), ("geographical mobility", 2),
         ("residence 1 year ago", 2), ("movers", 1)],
    12: [("occupants per room", 2), ("persons per room", 2), ("rooms", 1)],
    13: [("vacancy status", 2), ("occupancy status", 2), ("vacant", 2), ("vacant dwelling units", 2)],
    14: [("vacancy status", 1), ("vacant", 1)],
    15: [("internet subscriptions", 2), ("internet subscription", 2), ("presence of a computer", 1),
         ("computers", 1)],
    16: [("hispanic or latino origin by race", 2), ("race", 2), ("black or african american", 2),
         ("negro population", 1)],
    17: [("hispanic or latino origin by race", 1), ("race", 1)],
    18: [("hispanic or latino origin by race", 1), ("race", 1)],
    19: [("hispanic or latino origin by race", 1), ("household income in the past 12 months", 1)],
    24: [("public assistance income", 2), ("supplemental security income", 1),
         ("food stamps", 1), ("snap", 1)],
    25: [("nativity", 2), ("foreign born", 2), ("foreign-born", 2), ("place of birth", 2),
         ("citizenship status", 1), ("limited english speaking", 2), ("language spoken at home", 1),
         ("year of entry", 1)],
    28: [("units in structure", 2), ("housing units", 1), ("year structure built", 1)],
    32: [("group quarters", 1)],
    38: [("area", 0)],
    44: [("food stamps", 2), ("snap", 2)],
    45: [("sex by age by educational attainment", 2)],
    46: [("marital status", 2), ("never married", 2)],
    47: [("poverty status", 2), ("ratio of income to poverty level", 1)],
    48: [("interest, dividends, or net rental income", 2), ("retirement income", 1)],
    49: [("sex by educational attainment", 2), ("educational attainment", 1)],
    50: [("sex by age by educational attainment", 2)],
    52: [("marital status", 2), ("separated", 2)],
    53: [("veteran status by educational attainment", 2), ("veteran status", 1)],
    54: [("sex by educational attainment", 2), ("graduate or professional degree", 2)],
    56: [("sex by age by educational attainment", 2)],
    57: [("number of workers in family", 2), ("presence of own children under 18 years by family type by employment status", 2),
         ("work experience", 1)],
    58: [("means of transportation to work", 2), ("vehicles available", 1), ("commuting", 1)],
}

# ---------------------------------------------------------------------------
# Curated table codes per determinant (# in Keyword Spec).  Codes from the
# Sept 3 2026 memo plus documented extras.  Every code is validated against the
# catalog at run time; a code missing from the locked vintage is reported, not
# silently dropped.  "compute" notes mean the determinant is derived from the
# listed tables rather than read directly.
# ---------------------------------------------------------------------------
CURATED: dict[int, dict] = {
    3:  {"codes": [], "hist": ["1940_tPH_Major:NT4", "1940_tPH_Major:NT2", "1940_tPH_Major:NT27",
                               "1940_tPH_Major:NT29", "1940_tPH_Major:NT32"],
         "note": "HOLC grade comes from Mapping Inequality + 2020 crosswalk, not NHGIS. "
                 "1940 tract tables (Bogue file) give the 1930s-era Black presence control Markley requires; "
                 "check which of the eight CA cities 1940_tPH_Major covers via geographicInstances."},
    4:  {"codes": ["B25003", "B25008"], "note": "renter share = B25003_003/B25003_001"},
    5:  {"codes": ["B17001", "C17002", "B17017"], "note": "poverty rate from B17001; % <150% FPL from C17002"},
    6:  {"codes": ["B17001", "B23025", "B19057", "B11003", "B15003", "B19013"],
         "note": "compute: PCA/z-score composite of the Layer-2 components"},
    7:  {"codes": ["B19083", "B19001", "B19013"], "note": "Gini direct from B19083 (ACS5b at tract)"},
    8:  {"codes": ["B15003", "B15002"], "note": "% 25+ without HS diploma = sum(B15003_002..016)/B15003_001"},
    9:  {"codes": ["B23025", "B23001"], "note": "unemployment rate = B23025_005/B23025_003 (civilian LF)"},
    10: {"codes": ["B11003", "B11005", "B23008", "B09002"], "note": "single-parent w/ children from B11003 or B11005"},
    11: {"codes": ["B25038", "B07003", "B07001", "B25026"],
         "note": "instability = share of householders moved in since 2017 (B25038) or movers (B07003, ACS5b)"},
    12: {"codes": ["B25014"], "note": "crowding = share with >1.00 occupants per room"},
    13: {"codes": ["B25002", "B25004"], "note": "vacancy rate = B25002_003/B25002_001; 'other vacant' from B25004"},
    14: {"codes": ["B25004"], "note": "vacant LAND is municipal data; B25004 'other vacant' is only a housing proxy"},
    15: {"codes": ["B28002", "B28003"], "note": "no broadband = 1 - (B28002 broadband any type / total)"},
    16: {"codes": ["B03002", "B02001"], "note": "% Black, % Hispanic from B03002 (non-Hispanic Black = _004)"},
    17: {"codes": ["B03002"], "note": "compute: city-level dissimilarity from tract B03002 (metro attribute, not tract)"},
    18: {"codes": ["B03002"], "note": "compute: isolation/exposure indices from tract B03002 within city"},
    19: {"codes": ["B03002", "B19001"], "note": "compute: ICE race-income = (white high-income - Black low-income)/total"},
    24: {"codes": ["B19057", "B22010"], "note": "tract proxy only; county welfare spending is Census of Governments"},
    25: {"codes": ["B05002", "C16002", "B05006"], "note": "% foreign-born from B05002 (ACS5b); linguistic isolation C16002"},
    28: {"codes": ["B25024", "B25034"], "note": "units in structure as residential-intensity proxy; land use is municipal"},
    38: {"codes": [], "shapefile": True, "note": "ALAND/AWATER from TIGER 2020 or shapefile geometry, not a table"},
    42: {"codes": ["B03002"], "note": "compute: Black isolation index within city from tract B03002"},
    43: {"codes": ["B03002"], "note": "compute: Black-white dissimilarity within city from tract B03002"},
    44: {"codes": ["B22010", "B22003", "B22002"], "note": "SNAP households = B22010_002/B22010_001"},
    45: {"codes": ["B15001"], "note": "B15001 (ACS5b): sex x age x attainment, 65+ rows"},
    46: {"codes": ["B12001"], "note": "never married = (male_003 + female_012)/total"},
    47: {"codes": ["B17001", "B17021"], "note": "official poverty rate"},
    48: {"codes": ["B19054"], "note": "households with interest/dividend/rental income"},
    49: {"codes": ["B15002"], "note": "men 25+ < 9th grade = B15002 male rows 003-004"},
    50: {"codes": ["B15001"], "note": "men 25-34 HS+ from B15001"},
    51: {"codes": [], "shapefile": True, "note": "centroid longitude from tract geometry"},
    52: {"codes": ["B12001"], "note": "separated = male_005 + female_014"},
    53: {"codes": ["B21003"], "note": "veteran x educational attainment (ACS5b)"},
    54: {"codes": ["B15002"], "note": "men 25+ graduate/professional degree"},
    56: {"codes": ["B15001"], "note": "men 65+ bachelor's+"},
    57: {"codes": ["B23009", "B23007"], "note": "families w/ children and no workers from B23009"},
    58: {"codes": ["B08301"], "note": "car/truck/van share of workers"},
    59: {"codes": [], "shapefile": True, "note": "centroid latitude from tract geometry"},
}

# Determinants that are conceptual, survey-based, or come from non-Census sources.
# They are still searched, but the verdict says so instead of 'not_found'.
EXTERNAL_SOURCE: dict[int, str] = {
    1: "framing only", 2: "framing only",
    20: "survey (PHDCN-style); no census analogue", 21: "IRS BMF / nonprofit registries",
    22: "survey", 23: "Opportunity Atlas (Chetty) tract file", 26: "NLCD / NAIP canopy rasters", 27: "municipal zoning / land-use layers",
    29: "municipal land use / business registries", 30: "municipal land use",
    31: "OpenStreetMap street network", 32: "HUD public-housing / LIHTC point files",
    33: "GVA incidents + spatial weights", 34: "GVA incidents (prior period)",
    35: "police drug-arrest data", 36: "ABC liquor licenses / OSM POIs", 37: "OSM / business directories",
    39: "individual survey (exclude)", 40: "individual survey (exclude)", 41: "outcome mismatch (exclude)",
    55: "PRISM / NOAA climate normals",
}


# ---------------------------------------------------------------------------
# API access
# ---------------------------------------------------------------------------
def api_headers() -> dict:
    key = os.environ.get("NHGIS_API_KEY") or os.environ.get("IPUMS_API_KEY")
    if not key:
        sys.exit("ERROR: NHGIS_API_KEY not set and a fetch is required. "
                 "Create .env with NHGIS_API_KEY=... (account.ipums.org/api_keys) or run with cached files.")
    # IPUMS documents a bare key in the Authorization header; the v1 'Bearer ' prefix also worked.
    return {"Authorization": key}


def _get(url: str, params: dict | None = None):
    import requests
    for attempt in range(4):
        resp = requests.get(url, headers=api_headers(), params=params, timeout=60)
        if resp.status_code == 429:
            time.sleep(15 * (attempt + 1))        # rate limit: 100 req/min
            continue
        resp.raise_for_status()
        return resp.json()
    sys.exit(f"ERROR: rate-limited repeatedly on {url}")


def cache_age_days(path: Path) -> float | None:
    if not path.exists():
        return None
    return (time.time() - path.stat().st_mtime) / 86400


def fetch_all_pages(endpoint: str, cache_file: Path, refresh: bool) -> list:
    age = cache_age_days(cache_file)
    if age is not None and not refresh:
        flag = "  (STALE, consider --refresh)" if age > CACHE_STALE_DAYS else ""
        print(f"  [cache] {cache_file.name}: {age:.0f} days old{flag}")
        with open(cache_file) as f:
            return json.load(f)

    print(f"  [API] fetching {endpoint} ...")
    url = f"{API_BASE}/{endpoint}"
    params = {**API_PARAMS, "pageSize": 500}         # 500 is the documented maximum
    records, page = [], 0
    while url:
        page += 1
        data = _get(url, params if page == 1 else None)
        chunk = data.get("data", [])
        records.extend(chunk)
        print(f"    page {page}: {len(chunk)} (total {len(records)})")
        url = (data.get("links") or {}).get("nextPage")
        if url:
            time.sleep(0.7)
    with open(cache_file, "w") as f:
        json.dump(records, f)
    return records


def fetch_dataset_details(names: list[str], refresh: bool, enabled: bool) -> dict:
    """Per-dataset metadata: real geogLevels, geographicInstances, dataTables."""
    cache = {}
    if CACHE_DATASET_DETAIL.exists():
        with open(CACHE_DATASET_DETAIL) as f:
            cache = json.load(f)
    if not enabled:
        return cache
    missing = [n for n in names if n not in cache or refresh]
    for i, name in enumerate(missing, 1):
        print(f"  [API] dataset detail {i}/{len(missing)}: {name}")
        cache[name] = _get(f"{API_BASE}/datasets/{name}", API_PARAMS)
        time.sleep(0.7)
    if missing:
        with open(CACHE_DATASET_DETAIL, "w") as f:
            json.dump(cache, f)
    return cache


# ---------------------------------------------------------------------------
# Geography inference (fallback when no dataset detail is cached)
# ---------------------------------------------------------------------------
def infer_geogs(ds_name: str, desc: str, group: str) -> set[str]:
    d = (desc or "").lower()
    g = (group or "").lower()
    out: set[str] = set()
    if "block group" in d or "blocks &" in d or "blocks:" in d or "[blocks" in d or "block groups" in d:
        out |= {"block", "blck_grp", "tract", "county", "state"}
        if "block group" in d and "blocks" not in d:
            out.discard("block")
    if "tract" in d:
        out |= {"tract", "county", "state"}
    if "county" in d or "counties" in d:
        out |= {"county", "state"}
    if "acs" in ds_name.lower() and "1-year" in d:
        out |= {"county", "state"}           # 1-year ACS never reaches tract
    if re.match(r"^(1980|1990)_STF", ds_name):
        out |= {"block", "blck_grp", "tract", "county", "state"}
    if ds_name in {"2010_SF1a", "2020_DHCa", "2010_PL94171", "2020_PL94171", "1980_PL94171"}:
        out |= {"block", "blck_grp", "tract"}
    if not out:
        out = {"county", "state"} if ("count" in g or "census" in g) else set()
    return out


def geogs_from_detail(detail: dict) -> set[str]:
    return {g.get("name", "").lower() for g in detail.get("geogLevels", []) if g.get("name")}


# ---------------------------------------------------------------------------
# Catalog
# ---------------------------------------------------------------------------
def dataset_vintage(ds_name: str) -> str:
    m = re.match(r"^(\d{4}_\d{4})_ACS5", ds_name)
    return m.group(1) if m else ""


def dataset_era(ds_name: str) -> str:
    if re.match(r"^20[12]\d", ds_name):
        return "modern"
    if ds_name in HISTORICAL_TRACT_DATASETS:
        return "historical_tract"
    return "other"


def build_catalog(datasets, data_tables, time_series, details: dict) -> pd.DataFrame:
    ds_lookup = {d["name"]: d for d in datasets}
    rows = []
    for t in data_tables:
        ds = t.get("datasetName", "")
        parent = ds_lookup.get(ds, {})
        geogs = geogs_from_detail(details[ds]) if ds in details else \
            infer_geogs(ds, parent.get("description", ""), parent.get("group", ""))
        rows.append({
            "kind": "data_table", "code": t.get("name", ""), "dataset": ds,
            "group": parent.get("group", ""), "dataset_desc": parent.get("description", ""),
            "description": t.get("description") or "", "universe": t.get("universe") or "",
            "nvars": t.get("nVariables"), "geogs": geogs, "vintage": dataset_vintage(ds),
            "era": dataset_era(ds), "years": [],
        })
    for t in time_series:
        rows.append({
            "kind": "time_series", "code": t.get("name", ""), "dataset": "",
            "group": "Time Series", "dataset_desc": t.get("geographicIntegration", ""),
            "description": t.get("description") or "", "universe": "",
            "nvars": len(t.get("timeSeries", [])),
            "geogs": {g.get("name", "").lower() for g in t.get("geogLevels", [])},
            "vintage": "", "era": "time_series",
            "years": [y.get("name", "") for y in t.get("years", [])],
        })
    df = pd.DataFrame(rows)
    df["text"] = (df["description"] + " || " + df["universe"]).str.lower()
    print(f"  Catalog: {len(data_tables)} data tables + {len(time_series)} time series tables")
    print(f"  Datasets with real geogLevels from API: {len(details)}")
    return df


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------
def parse_keywords(cell) -> list[str]:
    if pd.isna(cell) or not str(cell).strip():
        return []
    out = []
    for kw in str(cell).split(";"):
        kw = re.sub(r"\(.*?\)", "", kw).strip()     # drop parenthetical qualifiers like "(alone)"
        if len(kw) >= 3:
            out.append(kw)
    return out


def term_regex(term: str) -> re.Pattern:
    words = [re.escape(w) for w in re.split(r"[\s\-/]+", term.strip()) if w]
    return re.compile(r"\b" + r"[\s\-/,]*".join(words) + r"\b", re.I)


_RX: dict[str, re.Pattern] = {}


def rx(term: str) -> re.Pattern:
    if term not in _RX:
        _RX[term] = term_regex(term)
    return _RX[term]


def score_catalog(catalog: pd.DataFrame, terms: list[tuple[str, int]], exclusions: list[str]) -> pd.DataFrame:
    """Return matching rows with a score and the list of terms that hit."""
    scores = pd.Series(0, index=catalog.index, dtype=int)
    hits: dict[int, list[str]] = defaultdict(list)
    for term, w in terms:
        if w <= 0:
            continue
        m = catalog["text"].str.contains(rx(term), regex=True, na=False)
        scores[m] += w
        for idx in catalog.index[m]:
            hits[idx].append(term)
    out = catalog[scores > 0].copy()
    out["score"] = scores[scores > 0]
    out["hits"] = [", ".join(hits[i]) for i in out.index]
    for ex in exclusions:
        out = out[~out["text"].str.contains(rx(ex), regex=True, na=False)]
    return out


def has_target_geog(geogs: set) -> bool:
    return bool(geogs & TARGET_GEOGS)


def fmt_rows(df: pd.DataFrame, limit: int, with_dataset=True) -> str:
    lines = []
    for _, r in df.head(limit).iterrows():
        desc = r["description"][:110]
        tag = f" in {r['dataset']}" if with_dataset and r["dataset"] else ""
        yrs = f" [{r['years'][0]}-{r['years'][-1]}]" if r["kind"] == "time_series" and r["years"] else ""
        lines.append(f"{r['code']}{tag} | {desc}{yrs} | s={r['score']}")
    extra = len(df) - limit
    if extra > 0:
        lines.append(f"... +{extra} more")
    return "\n".join(lines)


def vintage_availability(catalog: pd.DataFrame, codes: list[str], vintages: list[str]) -> str:
    if not codes:
        return ""
    by_v = {}
    for v in vintages:
        sub = catalog[(catalog["vintage"] == v) & (catalog["code"].isin(codes))]
        by_v[v] = f"{len(set(sub['code']))}/{len(codes)}"
    return ", ".join(f"{v} {n}" for v, n in by_v.items())


def match_determinant(num: int | None, spec: pd.Series | None, catalog: pd.DataFrame,
                      vintage: str) -> dict:
    terms: list[tuple[str, int]] = []
    exclusions: list[str] = []
    if spec is not None:
        terms += [(k, 2) for k in parse_keywords(spec.get("Primary Keywords (exact terms)"))]
        terms += [(k, 1) for k in parse_keywords(spec.get("Secondary Keywords (broader/related terms)"))]
        exclusions = parse_keywords(spec.get("Exclusion Terms (avoid false positives)"))
    if num in CENSUS_TERMS:
        terms += CENSUS_TERMS[num]

    hits = score_catalog(catalog, terms, exclusions) if terms else catalog.iloc[0:0].copy()
    hits = hits.sort_values(["score", "code"], ascending=[False, True])

    locked = hits[(hits["vintage"] == vintage) & hits["geogs"].apply(has_target_geog)]
    modern = hits[(hits["era"] == "modern") & (hits["vintage"] != vintage) & hits["geogs"].apply(has_target_geog)]
    hist = hits[hits["era"] == "historical_tract"]
    tseries = hits[(hits["era"] == "time_series") & hits["geogs"].apply(has_target_geog)]
    other = hits[~hits.index.isin(locked.index) & ~hits.index.isin(modern.index)
                 & ~hits.index.isin(hist.index) & ~hits.index.isin(tseries.index)]

    cur = CURATED.get(num, {}) if num is not None else {}
    curated_codes = cur.get("codes", [])
    curated_ok, curated_missing = [], []
    for c in curated_codes:
        sub = catalog[(catalog["vintage"] == vintage) & (catalog["code"] == c)]
        if sub.empty:
            curated_missing.append(c)
        else:
            curated_ok.append(f"{c} in {sub.iloc[0]['dataset']} | {sub.iloc[0]['description'][:90]}")
    for h in cur.get("hist", []):
        ds, code = h.split(":")
        sub = catalog[(catalog["dataset"] == ds) & (catalog["code"] == code)]
        curated_ok.append(f"{code} in {ds} | {sub.iloc[0]['description'] if not sub.empty else 'NOT IN CATALOG'}")

    if curated_codes and not curated_missing:
        verdict = "curated_confirmed"
    elif curated_codes:
        verdict = "curated_partial"
    elif cur.get("hist"):
        verdict = "curated_historical"
    elif cur.get("shapefile"):
        verdict = "from_shapefile"
    elif num in EXTERNAL_SOURCE:
        verdict = "external_source"
    elif not locked.empty:
        verdict = "keyword_locked_vintage"
    elif not modern.empty:
        verdict = "keyword_other_modern"
    elif not tseries.empty or not hist.empty:
        verdict = "keyword_time_series_or_historical"
    else:
        verdict = "not_found"

    geogs_found = sorted({g for gs in locked["geogs"] for g in gs} & TARGET_GEOGS) if not locked.empty else []
    codes_for_avail = curated_codes or list(dict.fromkeys(locked["code"].head(5)))

    note = cur.get("note", "")
    if num in EXTERNAL_SOURCE:
        note = (f"source: {EXTERNAL_SOURCE[num]}. " + note).strip()
    if curated_missing:
        note = f"MISSING in {vintage}: {', '.join(curated_missing)}. " + note

    return {
        "verdict": verdict,
        "curated": "\n".join(curated_ok),
        "locked": fmt_rows(locked, 12),
        "modern": fmt_rows(modern, 8),
        "hist": fmt_rows(hist, 6),
        "tseries": fmt_rows(tseries, 6, with_dataset=False),
        "other_n": len(other),
        "geo": ", ".join(geogs_found) if geogs_found else ("n/a" if verdict in ("external_source", "from_shapefile") else "none in locked vintage"),
        "avail": vintage_availability(catalog, codes_for_avail, [vintage] + COMPARE_VINTAGES),
        "top_hits": locked["hits"].head(3).str.cat(sep=" / ") if not locked.empty else "",
        "note": note,
        "n_locked": len(locked), "n_modern": len(modern), "n_hist": len(hist), "n_ts": len(tseries),
    }


# ---------------------------------------------------------------------------
# Spec lookup (exact then fuzzy name match between roadmap and keyword spec)
# ---------------------------------------------------------------------------
def build_spec_lookup(ks: pd.DataFrame) -> dict[str, pd.Series]:
    return {str(r["Determinant"]).strip().lower(): r for _, r in ks.iterrows()
            if pd.notna(r.get("Determinant"))}


def find_spec(name: str, lookup: dict) -> pd.Series | None:
    key = name.strip().lower()
    if key in lookup:
        return lookup[key]
    close = difflib.get_close_matches(key, list(lookup), n=1, cutoff=0.7)
    return lookup[close[0]] if close else None


# ---------------------------------------------------------------------------
# Extract spec (curated codes -> API v2 request body)
# ---------------------------------------------------------------------------
def build_extract_spec(catalog: pd.DataFrame, vintage: str, extent: str, geog: str, shapefile: str,
                       details: dict) -> tuple[dict, pd.DataFrame]:
    code_to_dets: dict[str, list[int]] = defaultdict(list)
    for num, cur in CURATED.items():
        for c in cur.get("codes", []):
            code_to_dets[c].append(num)
    rows, datasets = [], defaultdict(set)
    for code in sorted(code_to_dets):
        sub = catalog[(catalog["vintage"] == vintage) & (catalog["code"] == code)]
        if sub.empty:
            rows.append({"code": code, "dataset": "MISSING", "description": "", "geogs": "",
                         "determinants": ", ".join(map(str, code_to_dets[code]))})
            continue
        r = sub.iloc[0]
        datasets[r["dataset"]].add(code)
        rows.append({"code": code, "dataset": r["dataset"], "description": r["description"],
                     "geogs": ", ".join(sorted(r["geogs"] & TARGET_GEOGS)) or "?",
                     "determinants": ", ".join(map(str, code_to_dets[code]))})
    # Historical (1940) tables, keyed by dataset name
    for num, cur in CURATED.items():
        for h in cur.get("hist", []):
            ds, code = h.split(":")
            datasets[ds].add(code)
            sub = catalog[(catalog["dataset"] == ds) & (catalog["code"] == code)]
            rows.append({"code": code, "dataset": ds, "description": sub.iloc[0]["description"] if not sub.empty else "NOT IN CATALOG",
                         "geogs": "tract", "determinants": str(num)})

    spec = {
        "description": f"ca_{geog}_acs{vintage}_curated",
        "datasets": {ds: {"dataTables": sorted(codes), "geogLevels": [geog]} for ds, codes in sorted(datasets.items())},
        "shapefiles": [shapefile],
        "geographicExtents": [extent],
        "dataFormat": "csv_header",
        "breakdownAndDataTypeLayout": "single_file",
    }
    # 1940 tract tables carry no extent selection; the API will reject an extent on them, so split them out.
    hist = {ds: v for ds, v in spec["datasets"].items() if ds in HISTORICAL_TRACT_DATASETS}
    if hist:
        for ds in hist:
            del spec["datasets"][ds]
        spec["_historical_extract_separate"] = {
            "description": "ca_tract_1940_holc_controls", "datasets": hist,
            "dataFormat": "csv_header", "breakdownAndDataTypeLayout": "single_file",
            "note": "Submit as its own extract. Check details[ds]['geographicInstances'] for CA cities.",
        }
    # Warn about geographic coverage of 1940 data if details are present
    for ds in hist:
        inst = details.get(ds, {}).get("geographicInstances") or []
        if inst:
            ca = [i["description"] for i in inst if "CA" in i.get("description", "") or "California" in i.get("description", "")]
            spec["_historical_extract_separate"]["california_instances"] = ca
    return spec, pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
VERDICT_COLORS = {
    "curated_confirmed": "C6EFCE", "curated_historical": "D9EAD3", "keyword_locked_vintage": "E2F0D9", "curated_partial": "FFEB9C",
    "keyword_other_modern": "FFF2CC", "keyword_time_series_or_historical": "FCE4D6",
    "from_shapefile": "DDEBF7", "external_source": "EDEDED", "not_found": "F8CBAD",
}
RESULT_COLUMNS = [
    ("NHGIS Match Type", "verdict", 24),
    ("Curated Tables (locked vintage)", "curated", 60),
    ("Keyword Matches in Locked Vintage", "locked", 70),
    ("Keyword Matches in Other Modern Datasets", "modern", 55),
    ("1940-1970 Tract Datasets", "hist", 45),
    ("Time Series Tables (tract)", "tseries", 45),
    ("Geographic Levels", "geo", 18),
    ("Vintage Availability (curated or top-5)", "avail", 40),
    ("Keywords That Hit", "top_hits", 30),
    ("Notes / How to Build", "note", 60),
]


def write_results(results: dict[str, dict], extract_rows: pd.DataFrame, spec: dict,
                  src: Path, out: Path, vintage: str):
    wb = openpyxl.load_workbook(src)
    ws = wb["Combined Determinant Roadmap"]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    start = headers.index(RESULT_COLUMNS[0][0]) + 1 if RESULT_COLUMNS[0][0] in headers else ws.max_column + 1
    for i, (h, _, w) in enumerate(RESULT_COLUMNS):
        c = ws.cell(1, start + i, h.replace("locked vintage", vintage).replace("Locked Vintage", vintage))
        c.font, c.fill = HEADER_FONT, HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(start + i)].width = w
    written = 0
    for row in range(2, ws.max_row + 1):
        name = str(ws.cell(row, 2).value or "").strip()
        res = results.get(name.lower())
        if not res:
            continue
        fill = PatternFill("solid", fgColor=VERDICT_COLORS.get(res["verdict"], "FFFFFF"))
        for i, (_, key, _) in enumerate(RESULT_COLUMNS):
            c = ws.cell(row, start + i, res[key])
            c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
        written += 1

    # Extract Spec sheet
    if "Extract Spec" in wb.sheetnames:
        del wb["Extract Spec"]
    es = wb.create_sheet("Extract Spec")
    cols = ["code", "dataset", "description", "geogs", "determinants"]
    widths = [12, 20, 80, 18, 20]
    for j, (col, w) in enumerate(zip(cols, widths), 1):
        c = es.cell(1, j, col)
        c.font, c.fill = HEADER_FONT, HEADER_FILL
        es.column_dimensions[get_column_letter(j)].width = w
    for i, r in enumerate(extract_rows.itertuples(index=False), 2):
        for j, v in enumerate(r, 1):
            es.cell(i, j, v)
    es.cell(len(extract_rows) + 3, 1, "API request body (also written to nhgis_extract_spec.json):")
    es.cell(len(extract_rows) + 4, 1, json.dumps(spec, indent=2)).alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(out)
    print(f"  [Excel] {written} determinant rows + Extract Spec sheet -> {out.name}")


def print_summary(results: dict[str, dict]):
    counts = Counter(r["verdict"] for r in results.values())
    print("\n" + "=" * 64 + "\nNHGIS MATCH SUMMARY\n" + "=" * 64)
    for v in VERDICT_COLORS:
        if counts.get(v):
            print(f"  {v:<36} {counts[v]:3d}")
    weak = [n for n, r in results.items() if r["verdict"] == "not_found"]
    if weak:
        print("\nNOT FOUND anywhere in NHGIS (review or mark external):")
        for n in sorted(weak):
            print(f"  - {n}")
    missing = [(n, r["note"]) for n, r in results.items() if r["verdict"] == "curated_partial"]
    if missing:
        print("\nCurated codes missing from the locked vintage:")
        for n, note in missing:
            print(f"  - {n}: {note.split('.')[0]}")
    print("=" * 64)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="NHGIS metadata search for thesis determinants (v2)")
    ap.add_argument("--refresh", action="store_true", help="re-download catalog caches")
    ap.add_argument("--vintage", default=DEFAULT_VINTAGE, help="locked ACS 5-year vintage, e.g. 2018_2022")
    ap.add_argument("--extent", default=DEFAULT_EXTENT, help="NHGIS state extent code (CA=060)")
    ap.add_argument("--geog", default=DEFAULT_GEOG, help="geographic level for the extract (tract, blck_grp)")
    ap.add_argument("--shapefile", default=DEFAULT_SHAPEFILE)
    ap.add_argument("--no-details", action="store_true", help="skip per-dataset API calls (offline)")
    ap.add_argument("--no-excel", action="store_true")
    ap.add_argument("--output", default=str(OUTPUT_PATH))
    args = ap.parse_args()

    print(f"\n=== NHGIS Metadata Search v2  |  vintage {args.vintage}  |  extent {args.extent}  |  {args.geog} ===\n")

    print("[1] Catalog")
    datasets = fetch_all_pages("datasets", CACHE_DATASETS, args.refresh)
    data_tables = fetch_all_pages("data_tables", CACHE_DATA_TABLES, args.refresh)
    time_series = fetch_all_pages("time_series_tables", CACHE_TIME_SERIES, args.refresh)

    # Real geography for the datasets that matter: locked vintage a/b, comparison vintages,
    # decennial 2010/2020, and the 1940-1970 tract sets.
    detail_names = [d["name"] for d in datasets if
                    dataset_vintage(d["name"]) in [args.vintage] + COMPARE_VINTAGES
                    or d["name"] in HISTORICAL_TRACT_DATASETS
                    or re.match(r"^(2010_SF1|2020_DHC|2020_PL|2010_PL)", d["name"])]
    details = fetch_dataset_details(detail_names, args.refresh, enabled=not args.no_details)

    print("\n[2] Building catalog")
    catalog = build_catalog(datasets, data_tables, time_series, details)
    if catalog[catalog["vintage"] == args.vintage].empty:
        sys.exit(f"ERROR: no tables for vintage {args.vintage}; run --refresh or check the name")

    print("\n[3] Loading Keyword Spec + Roadmap")
    ks = pd.read_excel(KEYWORD_SPEC_PATH, sheet_name="Keyword Spec")
    dr = pd.read_excel(DETERMINANTS_PATH, sheet_name="Combined Determinant Roadmap")
    lookup = build_spec_lookup(ks)
    print(f"  {len(ks)} spec rows, {len(dr)} roadmap rows")

    print("\n[4] Matching")
    results: dict[str, dict] = {}
    for _, r in dr.iterrows():
        name = str(r.get("Determinant", "")).strip()
        if not name or name == "nan" or name.startswith("---"):
            continue
        spec = find_spec(name, lookup)
        num = int(spec["#"]) if spec is not None and pd.notna(spec.get("#")) else None
        res = match_determinant(num, spec, catalog, args.vintage)
        results[name.lower()] = res
        print(f"  [{res['verdict']:<34}] {name[:52]:<52} locked={res['n_locked']:<3} modern={res['n_modern']:<4} "
              f"hist={res['n_hist']:<3} ts={res['n_ts']}")

    print("\n[5] Extract spec")
    spec, extract_rows = build_extract_spec(catalog, args.vintage, args.extent, args.geog, args.shapefile, details)
    with open(EXTRACT_SPEC_PATH, "w") as f:
        json.dump(spec, f, indent=2)
    n_missing = int((extract_rows["dataset"] == "MISSING").sum())
    print(f"  {len(extract_rows)} curated tables across {list(spec['datasets'])} -> {EXTRACT_SPEC_PATH.name}"
          + (f"  ({n_missing} MISSING)" if n_missing else ""))

    if not args.no_excel:
        print("\n[6] Excel")
        write_results(results, extract_rows, spec, DETERMINANTS_PATH, Path(args.output), args.vintage)

    print_summary(results)


if __name__ == "__main__":
    main()

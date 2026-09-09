# Structural Inequality and Gun Violence in California

The reproducible data pipeline for an undergraduate honors paper that builds a
framework linking **structural inequality** to **specific social vulnerabilities**,
and tests that framework empirically using **gun violence in California** as the
case study, at the census-tract level.

Barrett, The Honors College and the W. P. Carey School of Business, Arizona State
University. Supervised by Prof. Asish Satpathy. Expected completion November 2026.

---

## The argument in brief

The paper makes two contributions, kept deliberately separate.

**The framework.** Determinants of structural inequality feed neighborhood
inequality and deprivation, which express themselves through two parallel
branches — spatial and built-environment on one side, social and institutional on
the other — and converge on social vulnerability, which in turn manifests as
specific vulnerabilities. It is written to be reusable across vulnerability
domains rather than tailored to any one of them.

| Layer | Content |
|---|---|
| 1 | Root structural inequality — segregation, redlining, class stratification |
| 2 | Neighborhood inequality and deprivation |
| 3 | Two parallel branches: spatial / built-environment, and social / institutional |
| 4 | Social vulnerability |
| 5 | Specific vulnerabilities — gun violence, homelessness, health complications |

Two rules govern how the layers combine. **No layer may be skipped**: a root cause
cannot produce a layer-5 vulnerability without routing through neighborhood
inequality, then a branch, then social vulnerability. And **the chain feeds back** —
gun violence contributes to incarceration, which contributes to single-parent
households, which reinforces layers 1 and 2.

Gun violence is the *case study*, not the terminal outcome. The layer-5 nodes other
than gun violence are illustrative; they are not operationalized or modeled here.

**The empirical test.** A California single-state test of a pooled
logistic-regression model previously run across 35 cities nationally. Logistic
regression is the baseline and stays the baseline: explainability is the point, and
a model whose coefficients cannot be read back onto the framework layers would
defeat the purpose.

---

## What is in this repository

Everything needed to **rebuild the dataset from scratch on a clean machine**. No
derived data is committed; the scripts fetch it.

```
03 Analysis and Code/
  nhgis_search.py            Searches the full IPUMS NHGIS catalog, matches every
                             determinant against it, and resolves the curated table
                             list into a ready-to-submit extract definition.
  nhgis_extract_spec.json    That extract definition. The single source of truth for
                             which tables, which geography, which vintage.
  nhgis_pull.py              Submits the spec to the IPUMS API v2, polls until the
                             extracts finish, downloads and unzips them, and writes
                             a manifest recording the extract numbers.
  make_appendix.py           Regenerates the reproducibility appendix (Excel +
                             Markdown) from the spec, the catalog cache and the
                             determinant workbook.
  geocode_census.py          Geocodes Gun Violence Archive incident addresses via the
                             U.S. Census batch geocoder, cleaning "700 block of SE
                             12th Ave" into something the geocoder accepts, and emits
                             lat/long columns for a spatial join.

requirements.txt             Python dependencies.
.env.example                 Template for the one credential the pipeline needs.
```

## What is deliberately not here

This repository is public, and a research folder contains a good deal that a public
repository has no business holding. Four categories are excluded by design, not by
oversight:

- **The data itself.** The NHGIS pull alone is 1.7 GB, of which a single national
  tract shapefile is 738 MB — past GitHub's 100 MB per-file limit several times
  over. It is fully regenerable from `nhgis_extract_spec.json`, which is why the
  spec is committed and the output is not.
- **Drafts and figures.** The paper is unsubmitted and pre-defense, and the drafts
  carry an advisor's written edits.
- **Literature.** The article and book PDFs are copyrighted and not mine to
  redistribute. Sources are cited in the paper.
- **Meeting notes.** Raw transcripts of conversations with real people.

The `.gitignore` is written deny-by-default — it ignores everything and re-includes
named files — so that none of the above can be swept in by an absent-minded
`git add -A`.

---

## Setup

Requires Python 3.9 or newer (3.11+ recommended) and a free IPUMS API key.

```bash
git clone https://github.com/ahashemee/structural-inequality-gun-violence-ca.git
cd structural-inequality-gun-violence-ca

python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

Register for a key at [account.ipums.org/api_keys](https://account.ipums.org/api_keys),
then:

```bash
cp .env.example "03 Analysis and Code/.env"
```

and put your key in it. The scripts read `.env` from their own directory, so run
them from `03 Analysis and Code/`.

## Rebuilding the dataset

```bash
cd "03 Analysis and Code"

# Inspect the request bodies without submitting anything
python nhgis_pull.py --dry-run

# Submit both extracts, poll until they finish, download and unzip
python nhgis_pull.py --out ../data/nhgis
```

The pull takes roughly ten to thirty minutes depending on IPUMS queue depth, and
writes `manifest.json` recording the extract numbers. Useful flags:

| Flag | Effect |
|---|---|
| `--dry-run` | Print the request bodies, submit nothing |
| `--skip-historical` | Primary ACS extract only, skipping the 1940 tract tables |
| `--resume 123456` | Skip submission; poll and download an extract already queued |
| `--out PATH` | Download directory |

IPUMS keeps download links for two weeks. After that the extract numbers can be
resubmitted from your Extracts History rather than rebuilt.

To re-derive the curated table list from the live NHGIS catalog instead of trusting
the committed spec:

```bash
python nhgis_search.py --refresh
```

---

## Data sources

| Source | Used for | Access |
|---|---|---|
| IPUMS NHGIS, Version 21.0 | ACS 2018–2022 5-year tract tables; 1940 tract tables; 2020 tract boundaries | [nhgis.org](https://www.nhgis.org) — free account, [doi:10.18128/D050.V21.0](https://doi.org/10.18128/D050.V21.0) |
| Gun Violence Archive | Incident-level gun violence records | [gunviolencearchive.org](https://www.gunviolencearchive.org) — export capped at 2,000 rows per query |
| Mapping Inequality (Digital Scholarship Lab, University of Richmond) | HOLC residential security grades | [dsl.richmond.edu/panorama/redlining](https://dsl.richmond.edu/panorama/redlining) |
| U.S. Census Bureau Geocoder | Address → coordinates for incident records | [geocoding.geo.census.gov](https://geocoding.geo.census.gov) — no key required |

**Extract parameters.** ACS 2018–2022 5-year, 2020 census tracts, `dataFormat:
csv_header`, `breakdownAndDataTypeLayout: single_file`. The primary extract covers
32 tables from `2018_2022_ACS5a` and 12 from `2018_2022_ACS5b`; the secondary covers
NT2, NT4, NT27, NT29 and NT32 from `1940_tPH_Major`. Join key is GISJOIN → GEOID:
drop the leading `G` and the two zero pads, so `G0600370207400` → `06037207400`.

## Known issues

Three things a rerun will reproduce, documented so nobody has to rediscover them:

1. **`geographicExtents` does not filter shapefiles.** The spec sets extent `060`
   (California), and NHGIS applies it to the tabular datasets only. The shapefile
   comes back national — roughly 85,000 tracts and 738 MB unzipped, against
   California's 9,129. Clip to California before any spatial join, or narrow the
   shapefile selection if you resubmit.
2. **The 1940 extract is national by design.** Its block carries no
   `geographicExtents` key, so it returns every state with 1940 tract coverage. The
   California subset must be taken downstream, and it is small: 838 rows, drawn only
   from Los Angeles (589), San Francisco (119), Alameda (117) and Contra Costa (13).
   No other California county has 1940 tract data. That is a hard ceiling on the
   historical-redlining strand rather than a pull error.
3. **`manifest.json` writes null metadata.** `description`, `datasets`, `shapefiles`
   and `geographicExtents` all come back empty. `nhgis_pull.py` reads them from the
   top level of the `GET /extracts/{n}` response, which is not where the API returns
   them. Extract numbers and file lists are correct, so the manifest is usable for
   resubmission; it just does not yet carry full provenance.

---

## Citation

If this framework or pipeline is useful to you:

> Hashemee, A. (2026). *Structural Inequality and Gun Violence in California*.
> Undergraduate honors paper, Barrett, The Honors College, Arizona State University.

Data citation for the NHGIS extracts follows IPUMS terms: Manson, S., Schroeder, J.,
Van Riper, D., et al. *IPUMS National Historical Geographic Information System:
Version 21.0* [dataset]. Minneapolis, MN: IPUMS. 2026.
[doi:10.18128/D050.V21.0](https://doi.org/10.18128/D050.V21.0)

## Status

Active. Data collection and harmonization through September 2026, analysis through
mid-October, writing through November. The code here is research code: it is written
to be read and rerun, not packaged as a library.

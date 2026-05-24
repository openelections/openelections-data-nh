[![Build Status](https://github.com/openelections/openelections-data-nh/actions/workflows/data_tests.yml/badge.svg?branch=master)](https://github.com/openelections/openelections-data-nh/actions/workflows/data_tests.yml?query=branch%3Amaster)

# OpenElections Data New Hampshire

This repository contains pre-processed election results from New Hampshire, formatted to be ingested into the OpenElections [processing pipeline](http://docs.openelections.net/guide/). It contains mostly CSV files converted from PDF tables. Interested in contributing? We have a bunch of [easy tasks](https://github.com/openelections/openelections-data-nh/labels/easy%20task) for you to tackle.

Here is what a [finished CSV file (from Ohio)](https://github.com/openelections/openelections-data-oh/blob/master/2000/20001107__oh__general__president.csv) looks like. Note that each row represents a single result for a single candidate, even if the data has multiple candidates in a single row. Also, vote totals do not contain commas or other formatting.

For extracting text from PDF tables, we recommend [Tabula](http://tabula.technology/), which can be installed and run locally on OSX, Windows or Linux.

If you're familiar with git and Github, clone this repository and get started. If not, you can still help: leave a comment on a task you'd like to work on, or just convert any of the files into CSV and send the result to openelections@gmail.com.

## The modern parser framework (`oe_nh/`)

Everything from 2022 onward is produced by the framework under `oe_nh/`.
Older years (`2012/`, `2014/`, …, `2020/`) are pre-existing CSVs from
earlier contributors and use different conventions; treat them as
read-only history.

### Quick start: re-generate an existing year's CSVs

All `uv` commands below run in a terminal, from the project root
(the directory that contains this `README.md`).

```bash
uv sync --all-groups                       # install deps (one-time)
uv run pytest                              # framework unit tests, all should pass
uv run python -m oe_nh.cli --year 2024     # build every office found in raw/2024/general/
# -> writes 2024/20241105__nh__general__<office>__precinct.csv files
#    prints a pre-flight scan, then a per-office summary with row counts
```

To re-build just one office (for debugging), pass `--office`:

```bash
uv run python -m oe_nh.cli --year 2024 --office governor
```

`--year` and `--office` auto-derive their available choices from the
`oe_nh/jobs/nh_<year>.py` registry, so adding a new year just makes it
show up.

### Adding a new election year

NH SoS publishes new election workbooks roughly six weeks after the
election. To add a year (say 2026):

1. **Download** the source workbooks from the SoS site
   (<https://www.sos.nh.gov/elections>). Under the "Elections" box
   there will be a link for "[year] Election Results"; click that and
   then the "[year] General Election Results" link (unless you're
   working on primaries). One workbook per office, except State
   Representative which is 10 files (one per county). See
   [scripts/fetch-raw.md](scripts/fetch-raw.md) for the per-office
   download links and source-file shape details.

2. **Drop** the downloaded files into `raw/2026/general/`. You don't
   have to rename them — the framework recognizes both canonical
   short-form names and the longer names the SoS publishes
   (`2026-ge-house-belknap_1.xls` and `house-belknap.xls` both work).
   The canonical short forms for reference:

   | Office | Canonical filename(s) |
   | --- | --- |
   | President | `president.xls[x]` |
   | Governor | `governor.xls[x]` |
   | US Senate | `us-senate.xls[x]` |
   | Congressional | `congressional-1.xls[x]`, `congressional-2.xls[x]` |
   | Executive Council | `executive-council.xls[x]` |
   | State Senate | `state-senate.xls[x]` |
   | State Representative | `house-belknap.xls[x]`, `house-carroll.xls[x]`, … (one per NH county) |

   The build command's pre-flight scan tells you which files matched
   which office (and which weren't recognized), so any naming mistake
   is surfaced before parsing starts.

3. **Register the year** by copying `oe_nh/jobs/nh_2024.py` to
   `oe_nh/jobs/nh_2026.py` and editing two values at the top:

   ```python
   GENERAL_FOLDER = "raw/2026/general"
   GENERAL_DATE   = "20261103"        # the election date, YYYYMMDD
   ```

   Each Job stub in the file is already shape-agnostic —
   auto-discovery picks the right parser for each office (see
   [scripts/fetch-raw.md](scripts/fetch-raw.md) for the dispatch
   table). If a particular file needs non-canonical config knobs (e.g.
   the SoS publishes a workbook with the header on row 4 instead of
   row 2), override in that Job's `files=` block.

4. **Build all six offices in one shot:**

   ```bash
   uv run python -m oe_nh.cli --year 2026
   ```

   You'll get a pre-flight scan ("found these files, will build these
   offices, ignoring these unknowns"), per-office build lines, and a
   trailing summary with ✅/⚠️/❌ status and total row counts. CSVs
   land under `2026/`. Anything missing or surprising is reported
   once, in the summary — no need to scroll through logs.

5. **Commit** the new raw files (`raw/2026/general/*.xls*`) AND the
   generated CSVs (`2026/*.csv`). Both belong in version control: raw
   files so the build is reproducible, CSVs because they're the
   published product OpenElections consumes.

### Output conventions

- **Schema** is `county,precinct,office,district,party,candidate,votes` — one row per (precinct, candidate) pair.
- **Office names** are exactly: `President`, `Governor`, `US Senate`, `Congressional`, `Executive Council`, `State Senate`, `State Representative`.
- **Floterial House districts** are emitted with an `F` suffix on the district column — e.g. Belknap's two-seat floterial that overlays Districts 1–7 appears as `district="8F"`. The SoS source file marks these with either `F` or `FL` on the district header (`District No. 8 (2) F` or `District No. 14 (1) FL`); both normalize to `F`.
- **Recount columns** that the SoS sometimes interleaves alongside certified counts (inline `Recount` columns in some 2022 House districts; whole `RECOUNT FIGURES` duplicate district sections in 2024 Strafford) are **dropped**. We ship certified counts only. To get recount data, go to the SoS source files.
- **`BLC` columns** in a handful of 2022 Rockingham House districts (an auxiliary ballot-related count we haven't been able to confirm the meaning of) are also **dropped**. If you know what BLC stands for and want it surfaced, open an issue.
- **Write-Ins / Undervotes / Overvotes** (in 2024 SoS files) are emitted as candidate rows with empty `party`. The `candidate` value is the literal column header.

### Architecture in one paragraph

`oe_nh/cli.py` is the orchestrator: it loads `oe_nh/jobs/nh_<year>.py`,
runs all matching Jobs (or just the one matching `--office`), calls
`oe_nh/discovery.py` to find raw files and build per-shape configs,
runs each through `oe_nh/parser.py`, and writes CSVs via
`oe_nh/writer.py`. `discovery.py` strips common SoS filename
decorations (year prefix, election prefix, `_N` revision suffix,
`-district-N-M` suffix) before matching, so you can drop SoS-named
files in `raw/` without renaming. Each NH-SoS reporting shape has its
own purpose-named Parser + Config dataclass (`CongressionalParser`,
`StatewideByCountyParser`, `ExecutiveCouncilParser`,
`StateSenateParser`, `StateRepresentativeParser`); the
`parse_workbook(path, config)` factory dispatches on config type. Add a
new shape by writing a new Parser + Config and adding a branch to the
factory plus an entry to `discovery._DISPATCH`.

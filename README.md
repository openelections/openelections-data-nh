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
uv run pytest                              # 57 framework tests, all should pass
uv run python -m oe_nh.cli \
    --year 2024 --election general --office governor
# -> writes 2024/20241105__nh__general__governor__precinct.csv
```

`--year` and `--office` show the available choices in `--help`; they
auto-derive from the `oe_nh/jobs/nh_<year>.py` registry, so adding a
new year shows up automatically.

### Adding a new election year

NH SoS publishes new election workbooks roughly six weeks after the
election. To add a year (say 2026):

1. **Download** the source workbooks from the SoS site
   (<https://www.sos.nh.gov/elections>). One workbook per office,
   except State Representative which is 10 files (one per county).
   See [scripts/fetch-raw.md](scripts/fetch-raw.md) for the per-office
   download links and source-file shape details.

2. **Rename** each downloaded file to its canonical short-form name and
   drop it under `raw/2026/general/`. The naming convention:

   | Office | Filename(s) |
   | --- | --- |
   | President | `president.xls[x]` |
   | Governor | `governor.xls[x]` |
   | US Senate | `us-senate.xls[x]` |
   | Congressional | `congressional-1.xls[x]`, `congressional-2.xls[x]` |
   | Executive Council | `executive-council.xls[x]` |
   | State Senate | `state-senate.xls[x]` |
   | State Representative | `house-belknap.xls[x]`, `house-carroll.xls[x]`, … (one per NH county) |

   `.xls` or `.xlsx` — the framework sniffs the file's magic bytes, so
   extension can lie. Drop the SoS's `_N` revision suffix and any
   year/election prefix. **The canonical names above are required for
   auto-discovery** — anything else will be silently skipped.

3. **Register the year** by copying `oe_nh/jobs/nh_2024.py` to
   `oe_nh/jobs/nh_2026.py` and editing two things:

   ```python
   _GENERAL = "raw/2026/general"
   _DATE    = "20261103"           # the election date, YYYYMMDD
   ```

   That's it. Each Job stub is already shape-agnostic — auto-discovery
   picks the right parser for each office (see
   [scripts/fetch-raw.md](scripts/fetch-raw.md) for the dispatch
   table). If a particular file needs non-canonical config knobs (e.g.
   the SoS publishes a workbook with the header on row 4 instead of
   row 2), override in the Job's `files=` block.

4. **Generate** each office's CSV:

   ```bash
   for office in president governor us-senate congressional \
                 executive-council state-senate state-representative; do
       uv run python -m oe_nh.cli --year 2026 --election general --office "$office"
   done
   ```

   The parser logs each `path -> ConfigClassName` pair to stderr before
   parsing, so you can sanity-check what auto-discovery picked. CSVs
   land under `2026/`.

5. **Verify** with the OpenElections data tests:

   ```bash
   scripts/run-data-tests.sh
   ```

   Pre-existing failures in 2014–2020 CSVs are tracked as legacy data
   issues; any new failure in a 2026 CSV needs investigation.

6. **Commit** the new raw files (`raw/2026/general/*.xls*`) AND the
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
picks the Job matching `--election`/`--office`, calls
`oe_nh/discovery.py` to find raw files and build per-shape configs,
runs each through `oe_nh/parser.py`, and writes the CSV via
`oe_nh/writer.py`. Each NH-SoS reporting shape has its own
purpose-named Parser + Config dataclass (`CongressionalParser`,
`StatewideByCountyParser`, `ExecutiveCouncilParser`,
`StateSenateParser`, `StateRepresentativeParser`); the
`parse_workbook(path, config)` factory dispatches on config type. Add a
new shape by writing a new Parser + Config and adding a branch to the
factory.

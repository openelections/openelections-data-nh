# openelections-data-nh — Claude project notes

## What this repo does

Pre-processes New Hampshire election results into the OpenElections CSV
format. Output CSVs are ingested by the OpenElections [processing
pipeline](http://docs.openelections.net/guide/) (note: that domain may
not resolve; sister-state repos in `github.com/openelections/` are the
practical reference for current output conventions).

## Layout at a glance

| Path | What's there |
|---|---|
| [oe_nh/](oe_nh/) | Modern parser framework (this is where new work goes) |
| [raw/`<year>`/`<election>`/](raw/) | Committed source `.xls` / `.xlsx` files from sos.nh.gov |
| [`<year>`/](2024/) | Output CSVs (also: pre-existing 2000-2020 CSVs from earlier contributors) |
| [scripts/fetch-raw.md](scripts/fetch-raw.md) | Manual procedure for downloading new SoS files |
| [scripts/run-data-tests.sh](scripts/run-data-tests.sh) | Run the four OpenElections data tests locally |
| [2012/code/](2012/code/), [2014/](2014/), [2016/parser.py](2016/parser.py) | Pre-existing one-off scrapers — historical, not actively maintained |
| [tests/](tests/) | unit + Hypothesis property tests for the new framework |

## How to run the parser

```bash
uv sync --all-groups            # install deps
uv run pytest                   # run the test suite
uv run python -m oe_nh.cli --year 2024 --election general --office president
scripts/run-data-tests.sh       # validate produced CSVs against OE tests
```

## Architecture in 30 seconds

`oe_nh/cli.py` is the orchestrator. It loads year-specific job registries
(`oe_nh/jobs/nh_<year>.py`), finds the matching `Job` for the requested
election/office, runs `parse_workbook()` for each raw file the Job
references, and writes the result CSV via `oe_nh/writer.py`.

The interesting code is in [oe_nh/parser.py](oe_nh/parser.py). Two shapes
to know about:

1. **Single-sheet workbook** (Congressional CD1/CD2): towns down column 0,
   candidates in `header_row`, vote matrix below. Set `header_row` and
   optionally `lookup_county_from_town=True` if the file has no county
   column.

2. **Multi-sheet workbook with section scanning** (President, Governor,
   US Senate): `multi_sheet=True` enables row-by-row scanning for
   county-name section headers. The same code path covers 2024
   (1 section per sheet, 11 sheets) and 2022 (multiple sections per
   sheet, with Summary+Belknap stacked on sheet 0 and Strafford+Sullivan
   stacked on the last sheet).

A row is a section header iff cell 0 (after stripping `" County"`) is
a known NH county AND cell 1 is a non-numeric candidate label. The
second check distinguishes a section header from a Summary block's
data row that happens to start with a county name.

Edge cases (State House's multi-district-per-file shape) are designed
to become `Parser` subclasses; none exist yet.

## What's covered, what's deferred

**Shipped and merged upstream:**

- `claude/uv-setup`: pyproject.toml, uv.lock, Py2 print fixes, scripts/run-data-tests.sh
- `claude/narrow-exceptions`: tightened bare except blocks in 2012 scrapers

**Pushed, PR open (PR #2 on tclancy's fork to upstream):**

- `claude/nh-rewrite`: framework + 2022/2024 General CSVs for Pres, Gov,
  US Senate, Congressional. Maintainer feedback: "looks pretty
  reasonable; how hard to extend to state-level races?"

**In progress (next session):**

- **State-level races for 2022 + 2024**: Executive Council, State
  Senate, State House. Raw files dropped under `raw/<year>/general/`
  but NOT yet renamed to the convention (still have `<year>-ge-`
  prefixes and `_N` revision suffixes from SoS). Discovery so far:
  - Exec Council: ONE file containing all 5 districts (multi-section
    by district inside, like our county sections)
  - State Senate: ONE file containing all 24 districts (same shape)
  - State House: 10 files, one per county. Each contains multiple
    districts (multi-member, varied) — district markers look like
    "Belknap 2 (4)" meaning district 2, 4-seat.
  - Likely implementable as a `district_marker_pattern` knob on the
    existing section scanner — not a separate subclass — since the
    iteration shape is the same as our existing county sections, just
    keyed on district name.
  - One open question Tom flagged: convention should files be
    `state-house-belknap.xls` (matches `state-senate.xls`) or
    `house-belknap.xls` (matches SoS's internal naming)?

**Deferred (longer-term):**

- Primaries (Presidential Primary 2024, State Primary 2022 + 2024) —
  per the original spec but lower priority than General
- Pre-existing 2014/2016/2018/2020 CSV data quality issues (triaged
  but not fixed — `2016/parser.py` line 92 etc. emit `None` for county
  in statewide sections; 2018 precinct file has duplicated Scattering
  rows; etc.)
- Smarter auto-discovery: per-office `ParserConfig` templates so a
  Congressional Job becomes a one-liner (currently each Job lists files
  explicitly because each office needs different config knobs)

## Conventions worth knowing

- **Branch names:** `claude/<topic>` per Tom's global CLAUDE.md
- **Raw file naming:** `<office-slug>.xls[x]` (single statewide file)
  or `<office-slug>-<location>.xls[x]` (county slug or district digits).
  See [scripts/fetch-raw.md](scripts/fetch-raw.md). SoS files are
  manually renamed when committed (the SoS site has unstable Drupal
  `_N` revision suffixes).
- **Output CSV schema:** `[county, precinct, office, district, party,
  candidate, votes]` — matches 2018-2020 files in this repo and is
  consistent with the modern OpenElections direction.
- **The OE data tests are git-pinned in CI** — see
  `.github/workflows/data_tests.yml`. Local runner reads the pin out
  of that yaml so they can't drift.

## Surprises and lore

- **NH SoS publishes mixed `.xls` and `.xlsx` in the same election** —
  WorkbookReader sniffs magic bytes rather than trusting the extension.
- **Sheet 0 of multi-sheet workbooks is the county summary** (gets
  silently skipped via `skip_sheet_markers` config).
- **Each county sheet ends with a `TOTALS` row** that would otherwise be
  treated as a precinct. Default `skip_town_values` includes
  `{TOTALS, Totals, Total}`.
- **Two unincorporated Coos townships have legacy abbreviated names**
  in `town_to_county.py` ('At. & Gil. Academy Grant',
  'Thompson & Meserve's Pur.'); 2024 SoS uses slightly different
  abbreviations. Aliases in `PRECINCT_ALIASES` map between them.
- **"us-senator" vs "us-senate"** — SoS files use the former; we use
  the latter as the canonical office slug.
- **2018 + 2020 precinct CSVs combine all offices into one file**.
  Our work uses one CSV per office (matching 2012 + modern sister-state
  conventions). Both styles are reasonable; the framework can emit
  either by tweaking the orchestrator.

## Where to look first

If the OpenElections maintainers respond to the [in-flight PR](https://github.com/tclancy/openelections-data-nh/pull/new/claude/nh-rewrite),
their feedback determines what's next. Otherwise the highest-value
follow-ups are probably:

1. **Add Executive Council** (5 statewide districts, simple shape).
2. **Build the State House subclass** — biggest expansion of the
   framework, unlocks all the down-ballot data.
3. **Combine outputs into per-election precinct.csv files** if upstream
   prefers that to per-office files.

## Related branches (local-only, not pushed)

- `claude/nh-rewrite-design` — original design doc from the brainstorm
  at the start of the session. Lives at
  `docs/superpowers/specs/2026-05-23-nh-parser-rewrite-design.md` on
  that branch only. Useful for "why was this decided" archeology.

# Fetching raw NH election workbooks

The NH Secretary of State publishes election results as Excel workbooks on
[sos.nh.gov](https://www.sos.nh.gov/elections). The download links are
manually-generated and the file names embed Drupal revision suffixes
(`_0`, `_1`, etc.) that change when the SoS replaces a file.

This repository expects the raw `.xls` / `.xlsx` files to be **committed**
under `raw/<year>/<election>/`. The downloader is intentionally a manual
human step — automation would mean either reverse-engineering the SoS
anti-bot layer or burning credits on a real-browser tool every time the
parsers run. Manually downloading new files once per election cycle is fine.

## File naming convention

Strip every SoS file down to a canonical short-form name when committing:

| Office | Source layout | Committed filename(s) |
| --- | --- | --- |
| President | one statewide workbook | `president.xls[x]` |
| Governor | one statewide workbook | `governor.xls[x]` |
| US Senate | one statewide workbook | `us-senate.xls[x]` |
| Congressional | one workbook per CD | `congressional-1.xls[x]`, `congressional-2.xls[x]` |
| Executive Council | one statewide workbook (5 sheets, one per district) | `executive-council.xls[x]` |
| State Senate | one statewide workbook (15–16 sheets, mix of one-and many-district) | `state-senate.xls[x]` |
| State Representative | one workbook per county | `house-belknap.xls[x]`, `house-carroll.xls[x]`, … (10 files) |

The extension can be `.xls` or `.xlsx`. The framework dispatches on file
content (magic bytes) so the extension can mislabel the actual format.

Drop the `_N` revision suffix the SoS adds. Drop the `<year>-ge-` prefix
the SoS adds. The short canonical names above are what every Job points at.

Auto-discovery (`Job.auto_discover=True`) only finds files matching
`<office_slug>.xls[x]` or `<office_slug>-<location>.xls[x]` and only
produces `CongressionalConfig`s. For shapes that need other configs
(`StatewideByCountyConfig`, `ExecutiveCouncilConfig`, etc.) list the file
explicitly in `Job.files=`.

## Procedure

1. Visit the year landing page:

   | Year | URL |
   | ---- | --- |
   | 2024 | <https://www.sos.nh.gov/2024-election-results> |
   | 2022 | <https://www.sos.nh.gov/elections/2022-election-results> |

2. Each year page links to one or more election sub-pages: Presidential
   Primary (presidential years only), State Primary, and General. Visit
   each sub-page in turn.

3. Download every workbook linked under the offices we cover (President,
   US Senate, Governor, Congressional, Executive Council, State Senate,
   State Representative).

4. Place each downloaded file under `raw/<year>/<election>/` using a name
   from the canonical-name table above. Example layout for 2024 General:

   ```text
   raw/2024/general/
     president.xls
     governor.xls
     us-senate.xls            # (2024 had no US Senate; example only)
     congressional-1.xlsx
     congressional-2.xlsx
     executive-council.xls
     state-senate.xls
     house-belknap.xls
     house-carroll.xlsx
     ... (8 more counties)
   ```

5. Register a `Job` for each office in `oe_nh/jobs/nh_<year>.py`.
   Auto-discovery is the default: for known office slugs (the seven covered
   above), the framework finds the right files in `folder` and builds the
   right Config dataclass automatically. Most Jobs need no `files=` block:

   ```python
   from oe_nh.jobs import Job

   _GENERAL = "raw/2024/general"

   JOBS: list[Job] = [
       Job(office_slug="president", office_name="President",
           election="general", date="20241105",
           output_basename="general__president__precinct",
           folder=_GENERAL),
       Job(office_slug="executive-council", office_name="Executive Council",
           election="general", date="20241105",
           output_basename="general__executive__council__precinct",
           folder=_GENERAL),
       # ... and so on for the other 5 offices
   ]
   ```

   The dispatch table in `oe_nh/discovery.py` knows which Config to build
   per office slug. To override (e.g. a non-canonical filename or a custom
   header_row), import the specific Config and supply it explicitly:

   ```python
   from oe_nh.parser import StatewideByCountyConfig

   Job(office_slug="president", office_name="President",
       election="general", date="20241105",
       output_basename="general__president__precinct",
       folder=_GENERAL,
       files=[("weird-president-name.xls",
               StatewideByCountyConfig(office="President", header_row=4))])
   ```

   Explicit entries in `files=` win over auto-discovered ones at the same
   filename; entries with new filenames are added.

6. Run the parser:

   ```bash
   uv run python -m oe_nh.cli --year 2024 --election general --office governor
   ```

   `--year` and `--office` choices are auto-derived from the registered
   year modules + Jobs, so adding a new year/office shows up in `--help`
   without touching `cli.py`.

7. The CSV lands under `<year>/<date>__nh__<output_basename>.csv`.

8. Run the data tests:

   ```bash
   scripts/run-data-tests.sh
   ```

   Pre-existing failures in 2014/2016/2018/2020 CSVs are tracked
   separately (legacy data, not generated by this framework). Any new
   failure in a 2022/2024 CSV needs investigation.

## File shapes

NH SoS files come in five shapes. Pick the matching config dataclass:

### 1. `CongressionalConfig` — single sheet, towns down col 0

One workbook per district; one sheet; one row per town. Source has no
county column, so opt into town→county backfill:

```python
CongressionalConfig(
    office="Congressional", district="1", header_row=2,
    lookup_county_from_town=True,
)
```

### 2. `StatewideByCountyConfig` — multi-sheet, county sections

One workbook covers the whole state. Most sheets contain one county's
town-level data. Some sheets combine multiple sections — the NH 2022
Governor and US Senate files merge Summary+Belknap into sheet 0 and
Strafford+Sullivan into the last sheet.

A section header is detected by scanning every row in each sheet for a
first cell that matches a known NH county (with or without `" County"`
suffix) AND has a candidate-like label in cell 1. `"Summary By Counties"`
is in the default skip set. Each section ends at the next section header
(or end of sheet); `TOTALS` rows are filtered.

```python
StatewideByCountyConfig(office="President", header_row=2)
```

### 3. `ExecutiveCouncilConfig` — multi-sheet, one district per sheet

Five sheets named `council 1` through `council 5`. District is parsed
from the sheet name. All defaults are pinned to the canonical SoS
Executive Council layout, so usually a bare `ExecutiveCouncilConfig()`
is enough.

### 4. `StateSenateConfig` — multi-sheet, district sections per sheet

15–16 sheets. Most hold one district (`senate 1` … `senate 9`), some
bundle 2–3 districts back-to-back (`senate 10 and 11`, `Senate 14 - 16`,
…). Sections are marked by a `State Senate District N` row; the
candidate header sits one row below. A stray `Sheet1` tab (2024) is
silently skipped. Inline `Recount` and `BLC` columns are dropped.
Usually `StateSenateConfig()` works as-is.

### 5. `StateRepresentativeConfig` — per-county file, many districts per sheet

One workbook per county (10 files). Each sheet has many sections marked
by `District No. N (M) [F|FL]` in col 0, where M is the seat count and
the optional F/FL marks a floterial district (normalized to `NF` in
output). Multi-seat districts have stacked candidate stripes
(continuation rows with blank col 0 followed by more candidates for the
same district). Inline `Recount`/`BLC` columns are dropped. Sections
marked `RECOUNT FIGURES` (2024 Strafford) are skipped entirely so we
ship certified counts only.

Build the per-county Job entries with the shared helper:

```python
from oe_nh.jobs import Job, house_files

Job(
    office_slug="state-representative", office_name="State Representative",
    election="general", date="20241105",
    output_basename="general__state__representative__precinct",
    folder="raw/2024/general",
    files=house_files(xlsx_counties=frozenset({
        "carroll", "grafton", "merrimack", "rockingham", "sullivan",
    })),  # NH SoS mixes .xls and .xlsx year-to-year
    auto_discover=False,
)
```

### Special-case rows (apply to every shape)

- `TOTALS` / `Totals` / `Total` precinct rows are skipped by default.
- Internal whitespace in precinct names is collapsed
  (`"Concord  Ward 1"` → `"Concord Ward 1"`). Same for candidate names
  (`"WRITE-IN   Kathy DesRoches"` → `"WRITE-IN Kathy DesRoches"`).
- `lookup_county_from_town=True` understands ward suffixes
  (`"Dover - Ward 3"` → looks up `"Dover"` → `Strafford`) and a small
  alias map in `oe_nh/mappings/town_to_county.py` for unincorporated
  townships and SoS spelling variants.

If a workbook arrives in a shape that doesn't match any of the five
above, add a new Parser + Config pair in `oe_nh/parser.py` and a new
branch in `parse_workbook`. The existing parsers are good templates.

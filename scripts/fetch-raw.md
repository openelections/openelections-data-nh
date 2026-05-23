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

The framework auto-discovers files in a job's `folder` (e.g.
`raw/2024/general/`) based on this convention:

```text
<office_slug>.xls[x]              # single statewide file
<office_slug>-<location>.xls[x]   # many files, one per county or district
```

Where:

- `<office_slug>` is one of: `president`, `us-senate`, `governor`, `congressional`
- `<location>` is either a county slug (`belknap`, `carroll`, `cheshire`, `coos`,
  `grafton`, `hillsborough`, `merrimack`, `rockingham`, `strafford`, `sullivan`)
  or a district identifier (`1`, `2`, `cd-1`, etc. — the framework extracts
  the leading digits)

The extension can be `.xls` or `.xlsx`. The framework dispatches on file
content (magic bytes) so the extension can mislabel the actual format.

## Procedure

1. Visit the year landing page:

   | Year | URL |
   | ---- | --- |
   | 2024 | <https://www.sos.nh.gov/2024-election-results> |
   | 2022 | <https://www.sos.nh.gov/elections/2022-election-results> |

2. Each year page links to one or more election sub-pages: Presidential
   Primary (presidential years only), State Primary, and General. Visit
   each sub-page in turn.

3. Download every workbook linked under the offices we care about
   (President, US Senate, Governor, Congressional District 1 & 2).

4. Place each downloaded file under `raw/<year>/<election>/` using a name
   that matches the convention. **Drop the `_N` revision suffix** that the
   SoS site appends. Example layout for 2024 General:

   ```text
   raw/2024/general/
     us-senate.xlsx
     governor-belknap.xls
     governor-carroll.xls
     ... (one per county)
     president-belknap.xls
     ... (one per county)
     congressional-1.xlsx     # CD1, possibly statewide single file
     congressional-2.xlsx     # CD2
   ```

   If the actual SoS file is split a different way (e.g. President as one
   statewide file, no per-county split), use that shape — the framework
   handles both. Just keep the office slug as the filename prefix.

5. Register a `Job` for each office in `oe_nh/jobs/nh_<year>.py`. A complete
   example for 2024 General Governor:

   ```python
   from oe_nh.jobs import Job

   JOBS: list[Job] = [
       Job(
           office_slug="governor",
           office_name="Governor",
           election="general",
           date="20241105",
           output_basename="general__governor__precinct",
           folder="raw/2024/general",
       ),
       # ... more Jobs for president, us-senate, congressional
   ]
   ```

   Because `auto_discover` defaults to `True`, every file in
   `raw/2024/general/` that matches `governor.xls[x]` or
   `governor-<county>.xls[x]` is picked up automatically. **No need to list
   each file by hand.**

   If a particular file doesn't match the convention (e.g. SoS published a
   weird filename you don't want to rename), add it explicitly via `files`:

   ```python
   Job(
       office_slug="president",
       office_name="President",
       election="general",
       date="20241105",
       output_basename="general__president__precinct",
       folder="raw/2024/general",
       files=[
           # Override the discovered config for this one file:
           ("president-strafford.xls",
            ParserConfig(office="President", county="Strafford", header_row=4)),
           # Or add a file that doesn't match the auto-discovery convention:
           ("special-recount-grafton-7.xlsx",
            ParserConfig(office="President", county="Grafton")),
       ],
   )
   ```

   The explicit list is merged with auto-discovery; entries in `files`
   override the discovered config when the filename matches.

6. Run the parser:

   ```bash
   uv run python -m oe_nh.cli --year 2024 --election general --office governor
   ```

7. The CSV lands under `<year>/<date>__nh__<output_basename>.csv`, matching
   the existing repo convention.

8. Run the data tests against the new CSV to catch anything obviously wrong:

   ```bash
   scripts/run-data-tests.sh
   ```

## Notes on file shapes

NH SoS files come in three flavors. Examples are from 2024 General:

### 1. Single-sheet by-district workbook (Congressional)

One workbook per CD; one sheet inside; one row per town. Example:
`congressional-1.xlsx` (CD1), `congressional-2.xlsx` (CD2). The source
data has no county column, so the Job opts into town→county backfill:

```python
ParserConfig(
    office="Congressional",
    district="1",
    header_row=2,
    lookup_county_from_town=True,
)
```

### 2. Multi-sheet workbook with one or more county sections per sheet (President, Governor, US Senate)

One workbook covers the whole state. Most sheets contain one county's
town-level data. Some sheets combine multiple sections — the NH 2022
Governor and US Senator files merge Summary+Belknap into sheet 0 and
Strafford+Sullivan into the last sheet, separated by section-header
rows.

A section header is detected by scanning every row in each sheet: any
row whose first cell (after stripping a trailing `" County"`) matches
a known NH county is treated as the start of a new section, and the
same row is also that section's candidate-header. `"Summary By
Counties"` is in the default `skip_sheet_markers` set and silently
skipped. Each section ends at the next section header (or end of sheet);
internal `TOTALS` / `Totals` rows are filtered by `skip_town_values`.

A data row inside the Summary block (e.g. `["Belknap", 20499, ...]`)
is NOT mistaken for a section header because the cell to the right of
the county name is numeric rather than a candidate label.

```python
ParserConfig(
    office="President",
    header_row=2,
    multi_sheet=True,
)
```

### 3. Single-sheet statewide (rare)

If SoS ever publishes a statewide race as one sheet with no per-county
breakdown, the default `Parser` shape works directly: towns down column
0, candidates across the header row. No multi-sheet, no backfill needed.

### Special-case rows

- The parser skips precinct rows whose value is `"TOTALS"`, `"Totals"`,
  or `"Total"` by default — every NH SoS sheet has a county-totals row
  at the bottom.
- Internal whitespace in precinct names is collapsed (`"Concord  Ward 1"`
  → `"Concord Ward 1"`).
- `lookup_county_from_town=True` understands ward suffixes
  (`"Dover - Ward 3"` → looks up `"Dover"` → `Strafford`) and a small
  alias map for unincorporated townships.

If a workbook doesn't match any of the three shapes (e.g. multi-district
files like State House, where one sheet contains multiple districts
separated by header rows), it needs a `Parser` subclass — see the design
doc under `docs/superpowers/specs/`.

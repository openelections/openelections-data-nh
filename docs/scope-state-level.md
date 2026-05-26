# Scope: 2022 + 2024 General — State-Level Offices

**Status:** Draft for Tom's review · 2026-05-24
**Branch:** `claude/nh-state-level`
**Builds on:** `claude/nh-rewrite` (open PR upstream)

## Goal

Extend the `oe_nh` parser framework to produce OpenElections CSVs for
the three remaining 2022 + 2024 General offices: **Executive Council**,
**State Senate**, **State Representative**. Six new CSVs total.

## Updated hypothesis vs. CLAUDE.md handoff

The handoff guessed all three offices would fit the same "section
scanner with a `district_marker_pattern` knob." After eyeballing the
raw files, that hypothesis is partially falsified. The three offices
have **three distinct shapes**:

| Office | Shape |
| --- | --- |
| Executive Council | One sheet per district, 5 sheets, no internal scanning |
| State Senate | One sheet per district *most* of the time, but 6 of 15 sheets bundle 2–3 districts back-to-back |
| State Representative | 10 files (one per county), each sheet has many districts back-to-back with continuation rows |

Trying to share one Parser class with knobs across all three would
push `ParserConfig` to ~20 fields, most mutually exclusive. Instead
we restructure the framework into a small family of named, purpose-
oriented parsers, dispatched by an explicit `shape` field in each
Job's config. See "Architecture: parser factory" below.

## Architecture: parser factory

Three phases. Phases 1 and 2 are in scope of this branch. Phase 3 is
a follow-on.

### Phase 1 — Subclass refactor of existing code (no behavior change)

Extract the two existing code paths into named, purpose-oriented
parsers, each with its own small config dataclass:

| Parser | What it parses |
| --- | --- |
| `CongressionalParser` | Single sheet, towns down col 0, candidates across; county looked up from town. Current use: Congressional CD1/CD2. |
| `StatewideByCountyParser` | Multi-sheet workbook with one sheet per county (plus an optional summary sheet to skip). Current use: President, Governor, US Senate. Name reflects role rather than office because the shape serves 1:N offices. |

Both inherit from a thin `Parser` base. A `parse_workbook(path, config)`
factory dispatches on `config.shape` (a `Literal[...]` discriminator).

Hard constraint: zero behavior change. Validate by running the
existing test suite + `scripts/run-data-tests.sh` + byte-diffing
every existing output CSV before/after.

### Phase 2 — Add three new parsers for state-level work

| Parser | What it parses |
| --- | --- |
| `ExecutiveCouncilParser` | Multi-sheet workbook, one district per sheet; district read from sheet name. |
| `StateSenateParser` | Like Exec Council, but some sheets bundle multiple districts back-to-back. Within-sheet scan keys on `r'^State Senate District\s+(\d+)'`; section's header row is the *next* row, not the marker row. Skip stray `Sheet1`. |
| `StateRepresentativeParser` | Per-county file, single sheet, many districts back-to-back with continuation rows for multi-seat. Handles: district+seat marker `r'^District No\.\s+(\d+)\s*\((\d+)\)'` in col 0; 2024-only date row at row 2; 2022 interleaved `Recount` columns (dropped); 2024-only `F` floterial suffix (TBD pending question 2 below); 2024 trailing `Undervotes`/`Overvotes`/`Write-Ins` columns. |

Each is its own ~80-line class with its own small config dataclass.
No flags-tunneling-flags.

### Phase 3 (deferred) — Opt-in content sniffing

After Phase 2, once five shapes exist, add `detect_shape(path) -> shape_id`
that peeks at a workbook and guesses. Job authors can pin `shape=...`
to override. Detection MUST log what it picked. Sniffing isn't worth
writing with fewer than ~5 shapes to differentiate.

## File shapes — what I actually saw

### Executive Council (`executive-council.xls`, 5 sheets)

- Sheet names: `council 1` … `council 5` — one district per sheet
- Row 0: `'State of New Hampshire - General  Election'`
- Row 1: `'Executive Council - District No. 1'`
- Row 2: header. Col 0 = serial date (44873.0 in 2022, 45601.0 in 2024). Cols 1+ = candidate strings like `'Joseph D. Kenney, r'`. 2024 also has `'Undervotes'`, `'Overvotes'`, `'Write-Ins'` columns. 2022 has `'Scatter'`.
- Rows 3+: town name + vote counts, ending with no totals row.

District is parseable from sheet name (`r'council (\d+)'`) **and**
from row 1 — either works. Sheet name is cleaner.

No section scanning needed.

### State Senate (`state-senate.xls`, 15–16 sheets)

- Sheet names: `senate 1` … `senate 9`, then `senate 10 and 11`, `senate 12 and 13`, `Senate 14 - 16`, `senate 17-18`, `senate 19-21`, `senate 22-24`
- 2024 file has a trailing `Sheet1` to skip
- Single-district sheet layout: same as Exec Council (district label row 1, header row 2, data row 3+)
- Multi-district sheet layout (e.g., `senate 10 and 11`):
  - Row 0: state title
  - Row 1: `'State Senate District 10'` ← section start
  - Row 2: header row
  - Rows 3–N: data
  - Row N+1: `'Totals'` row
  - Row N+2: blank
  - Row N+3: `'State Senate District 11'` ← next section start
  - Row N+4: header (col 0 sometimes just a space, not a date)
  - …

Section header rule: cell 0 matches `r'^State Senate District\s+(\d+)'`.
Unlike the existing county scanner, the *candidate* header is on the
**next** row, not the same row.

### State Representative (`house-<county>.xls`, 10 files)

- Single sheet per file, sheet name like `'rbelk rep'`
- Row 0: state title
- Row 1: `'State Representatives - BELKNAP County'`
- Row 2 (2024): date row, candidate cells blank → district sections start at row 3
- Row 2 (2022): first district section starts here directly (no separate date row)
- Each district section:
  - Header row: col 0 = `'District No. 1 (1)'` (district number + seats), cols 1+ = candidate strings like `'Ploszaj, r'`. 2024 adds `'Undervotes'`/`'Overvotes'`/`'Write-Ins'` in trailing cols.
  - Data rows: town/precinct + vote counts
  - `'Totals'` row
- Multi-seat districts (M > 2) have **continuation rows** where col 0 is blank and cols 1+ list additional candidates for the same district, then per-town data, then totals
- 2022 has interleaved `'Recount'` columns in some multi-seat districts (e.g., Belknap district 6: candidate, Recount, candidate, Recount, …)
- 2024 has occasional weird suffixes like `'District No. 8 (2) F'` — unclear what the F means; need to verify before code

Office name in output is `'State Representative'` per 2012 precedent.
District is the local integer (1..N within county); county column
disambiguates.

## Implementation order

1. **Phase 1 — Refactor.** Extract `CongressionalParser` and
   `StatewideByCountyParser` from the existing `Parser`. Per-shape
   config dataclasses. Factory in `parse_workbook` dispatches on a new
   `shape` discriminator. Validate by running the test suite +
   `scripts/run-data-tests.sh` + byte-diffing every existing CSV in
   `2022/` and `2024/` before vs after — must be identical. Commit
   when clean.
2. **Phase 2a — Executive Council.** Add `ExecutiveCouncilParser`,
   its config dataclass, and Job entries for 2022 and 2024. Output:
   `<year>/<date>__nh__general__executive__council__precinct.csv`.
3. **Phase 2b — State Senate.** Add `StateSenateParser`, its config
   dataclass, and Job entries. Output:
   `<year>/<date>__nh__general__state__senate__precinct.csv`.
4. **Phase 2c — State Representative.** Add `StateRepresentativeParser`,
   its config dataclass, and Job entries (one per county file, or one
   fan-out Job — judgment call when we get there). Output:
   `<year>/<date>__nh__general__state__representative__precinct.csv`.
   Resolves: `F` floterial suffix (decided before coding); 2022
   recount columns (dropped); continuation rows; trailing
   undervotes/overvotes/write-ins.
5. **Bookkeeping.** Update `scripts/fetch-raw.md` to document the
   short-form naming convention for the three new offices.

Phase 3 (content sniffing) is a separate effort, after Phase 2c
ships.

## Resolved decisions

- **CSV output:** one CSV per office (matches current 2024 work).
- **2022 House `Recount` columns:** dropped from output; we ship
  certified counts only.
- **Subclass vs. config knob for new shapes:** subclass, with
  per-shape config dataclass.
- **Raw file naming:** short form (`executive-council.xls`,
  `state-senate.xls`, `house-<county>.xls[x]`); 22 files renamed.
- **Parser naming:** office-named (`ExecutiveCouncilParser` etc.),
  with `StatewideByCountyParser` as the one exception because it
  serves 1:N offices.

## Open question still to resolve

- **`F` suffix on `District No. 8 (2) F` in 2024 House.** Best
  guess: floterial. Will dig (SoS docs and/or compare to 2018/2020
  CSVs) before starting Phase 2c. Tom offered to pull older files
  for comparison if needed.

## Out of scope

- Primaries (2022 + 2024 state primary, 2024 presidential primary) —
  deferred per existing handoff
- Pre-existing 2014/2016/2018/2020 data quality fixes — separate effort
- Auto-discovery of per-office ParserConfig templates — separate effort
- Re-architecting Jobs to fan out by county for the House files (vs.
  10 hand-written entries) — judgment call once Step 3 is real

## Reply to the OE maintainer

Once Step 1 lands locally I can give a concrete answer to the "how
hard to extend to state-level races" question. Tentative one-liner:
"Exec Council and State Senate were straightforward extensions of the
existing multi-sheet shape; State House needed a small subclass for
multi-seat continuation rows and floterial districts. ~XXX lines of
new code, no changes to existing parsers."

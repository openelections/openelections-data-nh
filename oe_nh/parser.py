"""Turn a workbook + a ParserConfig into normalized OpenElections rows.

The default Parser handles the common shape that the NH SoS uses for statewide
single-district races (President, US Senate, Governor) and per-CD Congressional
races: town/precinct names down the first column, candidate names in a header
row, vote counts in the matrix below.

Many SoS workbooks split town-level data across multiple sheets (one per
county) — and sometimes a single sheet contains multiple county "sections"
back-to-back (e.g. NH 2022 Governor merges Summary + Belknap into sheet 0,
and Strafford + Sullivan into the last sheet). `multi_sheet=True` plus
`parse_workbook` handle both shapes: every sheet is scanned for section
boundaries, marked by a row whose first cell contains a known NH county
name (with or without `' County'` suffix). Each section has its own
candidate-header row.

Edge cases (multi-district-per-file like State House) become subclasses that
override the small set of hook methods at the bottom of Parser.
"""

from __future__ import annotations

import dataclasses
import pathlib
import re
import sys
from dataclasses import dataclass
from typing import Iterator

from oe_nh.mappings.counties import NH_COUNTIES
from oe_nh.mappings.town_to_county import county_for_precinct
from oe_nh.workbook import WorkbookReader


@dataclass(frozen=True)
class NormalizedRow:
    county: str
    precinct: str
    office: str
    district: str
    party: str
    candidate: str
    votes: int


@dataclass
class ParserConfig:
    """Describes the shape of one workbook so the Parser can iterate it.

    Most knobs have sensible defaults for the common NH SoS shape.
    """

    office: str
    """Office name to emit in every row (e.g. 'President', 'US Senate')."""

    sheet_index: int = 0
    """Which sheet to read."""

    header_row: int = 3
    """Zero-indexed row containing candidate names."""

    town_col: int = 0
    """Zero-indexed column containing the town/precinct name."""

    candidate_cols_start: int = 1
    """First data column. Cells `[header_row, candidate_cols_start:]` are candidate names."""

    county: str | None = None
    """If the workbook covers one county, supply it here. None means it's a column
    (typically set per-section in multi_sheet mode)."""

    district: str = ""
    """Empty for statewide; '1'/'2' for Congressional; etc."""

    party_from_candidate: bool = True
    """If True, candidate cells like 'Smith, R' split into ('Smith', 'R')."""

    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})
    """Town cell values that mean 'skip this row'. The defaults cover the
    common NH SoS pattern of a county-totals row at the bottom of each section."""

    skip_empty_votes: bool = True
    """If True, cells with empty/whitespace votes do not emit a row."""

    lookup_county_from_town: bool = False
    """If True AND `county` is empty, look up the row's precinct in the NH
    town->county map and use that. Useful for by-district workbooks
    (Congressional, US Senate) where the source data doesn't include county."""

    multi_sheet: bool = False
    """If True, `parse_workbook` iterates every sheet in the workbook and
    scans each sheet for county-name section headers (with `' County'` suffix
    optional). Each section has its own candidate-header row."""

    section_marker_col: int = 0
    """In multi_sheet mode, the column containing the section-header marker.
    Defaults to 0 (the leftmost column)."""

    skip_sheet_markers: frozenset[str] = frozenset({"Summary By Counties"})
    """Values at `section_marker_col` that mean 'silently skip the section
    that starts here'. The canonical case is the per-state summary block."""

    stop_row: int | None = None
    """If set, the Parser yields rows up to (but not including) this row index.
    Set internally by `parse_workbook` to bound each section's iteration at
    the next section's start. Callers typically don't set this directly."""


_NUMERIC_RE = re.compile(r"^-?\d+(?:\.0+)?$")


def _coerce_votes(value) -> int | None:
    """Return an int vote count or None if the value should be treated as 'no vote'."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        # NH SoS sometimes emits float-looking integers. Reject true fractions.
        if value.is_integer():
            return int(value)
        return None
    s = str(value).strip()
    if not s or s == "--" or s == "-":
        return None
    if _NUMERIC_RE.match(s):
        return int(float(s))
    return None


class Parser:
    """The default 'towns down, candidates across' parser.

    Subclass for shapes that don't fit (e.g. files with district markers between
    blocks of rows) and override the `_should_skip_row` hook.
    """

    def __init__(self, config: ParserConfig, reader: WorkbookReader):
        self._config = config
        self._reader = reader
        self._candidates = self._read_candidate_row()

    def __iter__(self) -> Iterator[NormalizedRow]:
        cfg = self._config
        end_row = cfg.stop_row if cfg.stop_row is not None else self._reader.nrows
        for row in range(cfg.header_row + 1, end_row):
            yield from self._rows_at(row)

    def _read_candidate_row(self) -> list[tuple[str, str]]:
        """Return list of (candidate, party) for each data column."""
        row = self._reader.row_values(self._config.header_row)
        out: list[tuple[str, str]] = []
        for col in range(self._config.candidate_cols_start, len(row)):
            label = str(row[col]).strip()
            if not label:
                out.append(("", ""))
                continue
            out.append(self._split_candidate(label))
        return out

    def _split_candidate(self, label: str) -> tuple[str, str]:
        if self._config.party_from_candidate and "," in label:
            name, _, party = label.partition(",")
            return name.strip(), party.strip().upper()
        return label, ""

    def _rows_at(self, row: int) -> Iterator[NormalizedRow]:
        cfg = self._config
        town_value = self._reader.cell_value(row, cfg.town_col)
        # Collapse internal whitespace runs ("Concord  Ward 1" -> "Concord Ward 1")
        # so downstream consumers don't have to guess about source-data noise.
        town = " ".join(str(town_value).split()) if town_value is not None else ""
        if self._should_skip_row(row, town):
            return
        county = cfg.county or ""
        if not county and cfg.lookup_county_from_town:
            county = county_for_precinct(town) or ""
        for i, (candidate, party) in enumerate(self._candidates):
            if not candidate:
                continue
            col = cfg.candidate_cols_start + i
            if col >= self._reader.ncols:
                break
            raw = self._reader.cell_value(row, col)
            votes = _coerce_votes(raw)
            if votes is None and cfg.skip_empty_votes:
                continue
            yield NormalizedRow(
                county=county,
                precinct=town,
                office=cfg.office,
                district=cfg.district,
                party=party,
                candidate=candidate,
                votes=votes if votes is not None else 0,
            )

    def _should_skip_row(self, row: int, town: str) -> bool:
        cfg = self._config
        if not town:
            return True
        if town in cfg.skip_town_values:
            return True
        return False


@dataclass(frozen=True)
class _Section:
    """One county-section within a sheet."""
    header_row: int        # row index of the section header (cell 0 = county marker, cells 1+ = candidates)
    county: str | None     # canonical NH county name, or None for "skip this section"


def parse_workbook(path: pathlib.Path, config: ParserConfig) -> Iterator[NormalizedRow]:
    """Top-level entry point. Handles single-sheet and multi-sheet/multi-section workbooks.

    Single-sheet (`config.multi_sheet=False`): opens `config.sheet_index` and
    yields rows with the config's `header_row`.

    Multi-sheet: iterates every sheet in the workbook. Each sheet is scanned
    row-by-row for section headers (rows whose first cell, after stripping
    `' County'` suffix, matches a known NH county name). Each section gets
    its own header row and (optional) stop row. The summary-by-counties
    section is skipped via `skip_sheet_markers`.
    """
    if not config.multi_sheet:
        reader = WorkbookReader(path, sheet_index=config.sheet_index)
        yield from Parser(config, reader)
        return

    for sheet_index in range(WorkbookReader.sheet_count(path)):
        reader = WorkbookReader(path, sheet_index=sheet_index)
        sections = _find_sections(reader, config)
        for section, next_start in _with_bounds(sections, reader.nrows):
            if section.county is None:
                continue  # skip (summary section, or unrecognized marker)
            sub_config = dataclasses.replace(
                config,
                county=section.county,
                sheet_index=sheet_index,
                header_row=section.header_row,
                stop_row=next_start,
            )
            yield from Parser(sub_config, reader)


def _find_sections(reader: WorkbookReader, config: ParserConfig) -> list[_Section]:
    """Scan a sheet row-by-row, identifying section header rows.

    A row is a section header iff:
    - cell at `section_marker_col` is a known skip marker (e.g. "Summary By Counties"), OR
    - cell at `section_marker_col`, stripped of trailing `" County"` whitespace,
      matches a known NH county AND the next cell (column 1) contains a non-numeric
      string (i.e. a candidate label, not a vote count).

    The second condition distinguishes a real section header from a row inside
    the summary section that happens to have a county name in column 0
    (e.g. row 3 of the summary block: `['Belknap', 20499, ...]`).
    """
    sections: list[_Section] = []
    marker_col = config.section_marker_col
    if reader.ncols <= marker_col:
        return sections

    for row in range(reader.nrows):
        cell0 = reader.cell_value(row, marker_col)
        label = "" if cell0 is None else str(cell0).strip()
        if not label:
            continue
        if label in config.skip_sheet_markers:
            sections.append(_Section(header_row=row, county=None))
            continue
        candidate = label.removesuffix(" County").strip()
        if candidate not in NH_COUNTIES:
            continue
        # Likely county name in cell 0; confirm by checking column 1 is a non-numeric string.
        if not _looks_like_header_row(reader, row, config):
            continue
        sections.append(_Section(header_row=row, county=candidate))
    return sections


def _looks_like_header_row(reader: WorkbookReader, row: int, config: ParserConfig) -> bool:
    """True if `row` looks like a section header (candidate labels in cells 1+)
    rather than a data row (vote counts in cells 1+)."""
    next_col = config.section_marker_col + 1
    if next_col >= reader.ncols:
        return False
    value = reader.cell_value(row, next_col)
    if value is None or value == "":
        return False
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return False
    return isinstance(value, str) and bool(value.strip())


def _with_bounds(sections: list[_Section], total_rows: int) -> Iterator[tuple[_Section, int]]:
    """Yield each section paired with the row at which the next section starts
    (or `total_rows` for the last section)."""
    for i, section in enumerate(sections):
        next_start = sections[i + 1].header_row if i + 1 < len(sections) else total_rows
        yield section, next_start

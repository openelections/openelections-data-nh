"""Parse NH SoS election workbooks into normalized OpenElections rows.

Each NH office's reporting shape gets its own purpose-named parser class
plus a small per-shape config dataclass. The top-level `parse_workbook`
factory dispatches on the config type. This keeps each shape's logic
locally readable instead of branching one giant class on flags.

Public API:

- `parse_workbook(path, config) -> Iterator[NormalizedRow]` — top-level entry
- `NormalizedRow` — output row
- Config dataclasses, one per shape:
  - `CongressionalConfig` — single sheet, towns down col 0, candidates across.
    Used for whole-district races shipped as one workbook (Congressional CD1/CD2).
  - `StatewideByCountyConfig` — multi-sheet workbook, one sheet per county,
    plus an optional summary sheet to skip. Used for statewide single-race
    elections (President, Governor, US Senate).
- Parser classes, one per shape (`CongressionalParser`, `StatewideByCountyParser`)
- `ParserConfig` — type alias = union of all config dataclasses, for type hints
"""

from __future__ import annotations

import pathlib
import re
from dataclasses import dataclass
from typing import Iterator, Union

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


# ---------------------------------------------------------------------------
# Per-shape config dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CongressionalConfig:
    """Single sheet, towns down col 0, candidates across.

    Used for races where one workbook holds the whole race's results
    (e.g. NH Congressional District 1, District 2). The Parser reads
    the candidate header row, then yields one NormalizedRow per
    (town, candidate) pair below it.
    """

    office: str
    """Office name written into every row (e.g. 'Congressional')."""

    sheet_index: int = 0
    """Which sheet to read."""

    header_row: int = 3
    """Zero-indexed row containing candidate names."""

    town_col: int = 0
    """Zero-indexed column containing the town/precinct name."""

    candidate_cols_start: int = 1
    """First data column. Cells `[header_row, candidate_cols_start:]` are candidate names."""

    county: str | None = None
    """If the workbook covers one county, supply it here. None + lookup_county_from_town
    means look up county from the town map; None + neither means emit empty county."""

    district: str = ""
    """Empty for statewide; '1'/'2' for Congressional; etc."""

    party_from_candidate: bool = True
    """If True, candidate cells like 'Smith, R' split into ('Smith', 'R')."""

    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})
    """Town cell values that mean 'skip this row'."""

    skip_empty_votes: bool = True
    """If True, cells with empty/whitespace votes do not emit a row."""

    lookup_county_from_town: bool = False
    """If True AND `county` is empty, look up the row's precinct in the NH
    town->county map and use that. Useful for by-district workbooks where
    the source data doesn't include county."""

    drop_column_labels: frozenset[str] = frozenset()
    """Candidate-header cell values that trigger dropping the column from
    the output (case-insensitive). NH SoS files sometimes interleave
    auxiliary columns alongside real candidates — e.g. 'Recount' (recount
    counts in 2022 House files; we ship certified counts only) or 'BLC'
    (unknown ballot-related auxiliary in some 2022 House files)."""


@dataclass(frozen=True)
class StatewideByCountyConfig:
    """Multi-sheet workbook, one sheet per county, with within-sheet section scanning.

    Used for statewide single-race elections (President, Governor, US Senate)
    where the SoS publishes one tab per county, plus an optional summary
    tab. Some tabs hold multiple county sections back-to-back (NH 2022
    Governor stacks Summary + Belknap on sheet 0, and Strafford + Sullivan
    on the last sheet).

    Each county section gets its own candidate-header row read independently;
    candidates may differ between counties in primaries.
    """

    office: str
    """Office name written into every row."""

    header_row: int = 3
    """Zero-indexed row of the first candidate header in each section.
    For 2024 NH SoS workbooks this is row 2; the framework reads the
    actual header row per-section using the section marker as an anchor."""

    town_col: int = 0
    """Zero-indexed column containing the town/precinct name."""

    candidate_cols_start: int = 1
    """First data column within each section."""

    district: str = ""
    """Statewide races leave this empty."""

    party_from_candidate: bool = True
    """If True, 'Smith, R' splits into ('Smith', 'R')."""

    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})

    skip_empty_votes: bool = True

    section_marker_col: int = 0
    """Column containing the county-name section header. Defaults to leftmost."""

    skip_sheet_markers: frozenset[str] = frozenset({"Summary By Counties"})
    """Values at section_marker_col that mean 'silently skip the section
    that starts here'. Canonical case: per-state summary block."""


@dataclass(frozen=True)
class ExecutiveCouncilConfig:
    """Multi-sheet workbook, one district per sheet.

    Used for NH Executive Council general elections: five sheets named
    ``council 1`` ... ``council 5``, each holding one district's town-by-town
    results. The district number is parsed from the sheet name. Each sheet
    has a date cell in col 0 of the header row, then candidates across.
    Town column is looked up against the NH town->county map so each row
    carries the right county.
    """

    office: str = "Executive Council"
    """Office name written into every row."""

    header_row: int = 2
    """Zero-indexed row containing candidate names. NH SoS uses row 2:
    row 0 = state title, row 1 = district label, row 2 = header."""

    town_col: int = 0
    """Zero-indexed column containing the town/precinct name."""

    candidate_cols_start: int = 1
    """First data column. The header row's col 0 holds a date, ignored."""

    district_from_sheet_name: re.Pattern = re.compile(
        r"council\s+(\d+)", re.IGNORECASE
    )
    """Pattern applied to each sheet's name. The first capture group is the
    district number. Sheets that don't match are silently skipped."""

    party_from_candidate: bool = True
    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})
    skip_empty_votes: bool = True


@dataclass(frozen=True)
class StateSenateConfig:
    """Multi-sheet workbook with one or more district sections per sheet.

    Used for NH State Senate general elections. Most sheets hold one
    district (``senate 1`` ... ``senate 9``), but a handful bundle 2–3
    districts back-to-back (``senate 10 and 11``, ``Senate 14 - 16``,
    etc.). Each district begins with a marker row whose first cell
    matches ``district_section_marker``; the candidate header is the
    NEXT row, and data continues until the following marker or the end
    of the sheet. Sheets named in ``skip_sheet_names`` (e.g. an empty
    'Sheet1' tab in 2024) are silently skipped.
    """

    office: str = "State Senate"
    """Office name written into every row."""

    town_col: int = 0
    """Zero-indexed column containing the town/precinct name."""

    candidate_cols_start: int = 1
    """First data column. The header row's col 0 holds a date (or blank), ignored."""

    district_section_marker: re.Pattern = re.compile(
        r"^State Senate District\s+(\d+)", re.IGNORECASE
    )
    """Pattern matched against cell 0 of each row. The first capture group
    is the district number. The candidate header row sits at marker_row + 1."""

    skip_sheet_names: frozenset[str] = frozenset({"Sheet1"})
    """Sheet names that should be silently skipped (e.g. empty
    leftover 'Sheet1' in some workbooks)."""

    drop_column_labels: frozenset[str] = frozenset({"Recount", "BLC"})
    """Candidate-header labels to drop from output (case-insensitive).
    Some 2022 State Senate sheets interleave Recount columns with cert
    counts; we ship cert only."""

    party_from_candidate: bool = True
    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})
    skip_empty_votes: bool = True


@dataclass(frozen=True)
class StateRepresentativeConfig:
    """Per-county workbook with many District sections per sheet.

    Used for NH State Representative (House) general elections. Each county
    publishes one workbook (one sheet) with many districts back-to-back.
    A district section begins with ``District No. N (M) [F|FL]`` in col 0,
    where N is the district number, M is the seat count, and the optional
    F/FL marks a floterial district (normalized to ``NF`` in output).

    Within a multi-seat district, candidate stripes are stacked: the marker
    row IS the first stripe's header (col 0 = marker, cols 1+ = candidate
    labels); per-town data rows follow; then a Totals row; then another
    stripe header (col 0 blank, cols 1+ = the next batch of candidates).

    Quirks the parser drops:
    - 2022 files mix 'Recount' columns inline with candidates → column
      labeled exactly 'Recount' is dropped from output.
    - 2024 files duplicate some districts with 'RECOUNT FIGURES' suffix
      on the marker → that section is skipped entirely (we ship the
      certified counts, which appear in the first occurrence of the
      district).
    """

    office: str = "State Representative"
    """Office name written into every row."""

    county: str = ""
    """County is determined per-file by the Job (one workbook per county)."""

    district_marker: re.Pattern = re.compile(
        r"^District\s+No\.?\s*(\d+)\s*\(\d+\)\s*(FL?)?\s*(RECOUNT\s+FIGURES)?\s*$",
        re.IGNORECASE,
    )
    """Pattern for district section headers.
    Groups: (district_number, floterial_marker_or_None, recount_marker_or_None)."""

    floterial_suffix: str = "F"
    """Appended to the district number when the marker has F/FL.
    e.g. district number '8' + floterial → ``8F``."""

    drop_column_labels: frozenset[str] = frozenset({"Recount", "BLC"})
    """Candidate-header labels to drop from output (case-insensitive).
    Covers 2022 inline 'Recount' columns; 2024 inline 'RECOUNT' columns;
    and 'BLC' auxiliary columns in some 2022 Rockingham House districts."""

    town_col: int = 0
    candidate_cols_start: int = 1
    party_from_candidate: bool = True
    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})
    skip_empty_votes: bool = True


# Type alias for "any parser config", useful in shared callsites (jobs, cli).
ParserConfig = Union[
    CongressionalConfig,
    StatewideByCountyConfig,
    ExecutiveCouncilConfig,
    StateSenateConfig,
    StateRepresentativeConfig,
]


# ---------------------------------------------------------------------------
# Vote coercion
# ---------------------------------------------------------------------------


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
        if value.is_integer():
            return int(value)
        return None
    s = str(value).strip()
    if not s or s == "--" or s == "-":
        return None
    if _NUMERIC_RE.match(s):
        return int(float(s))
    return None


# ---------------------------------------------------------------------------
# CongressionalParser — the matrix-parsing primitive
# ---------------------------------------------------------------------------


class CongressionalParser:
    """Parses a single rectangular block of rows: header row + data rows below.

    The default for single-sheet, whole-race workbooks (NH Congressional
    CD1/CD2). Also reused internally by `StatewideByCountyParser` for each
    county section it finds — the `stop_row` constructor kwarg lets a caller
    bound iteration to a single section within a multi-section sheet.
    """

    def __init__(
        self,
        config: CongressionalConfig,
        reader: WorkbookReader,
        *,
        stop_row: int | None = None,
    ):
        self._config = config
        self._reader = reader
        self._stop_row = stop_row
        self._candidates = self._read_candidate_row()

    def __iter__(self) -> Iterator[NormalizedRow]:
        cfg = self._config
        end_row = self._stop_row if self._stop_row is not None else self._reader.nrows
        for row in range(cfg.header_row + 1, end_row):
            yield from self._rows_at(row)

    def _read_candidate_row(self) -> list[tuple[str, str]]:
        """Return list of (candidate, party) for each data column.

        Cells whose label matches `drop_column_labels` (case-insensitive)
        become ("", "") so the column emits no rows."""
        row = self._reader.row_values(self._config.header_row)
        drop_set = {label.casefold() for label in self._config.drop_column_labels}
        out: list[tuple[str, str]] = []
        for col in range(self._config.candidate_cols_start, len(row)):
            label = str(row[col]).strip()
            if not label:
                out.append(("", ""))
                continue
            if label.casefold() in drop_set:
                out.append(("", ""))
                continue
            out.append(self._split_candidate(label))
        return out

    def _split_candidate(self, label: str) -> tuple[str, str]:
        # Collapse internal whitespace runs in candidate names so labels
        # like 'WRITE-IN   Kathy DesRoches' (3 spaces between WRITE-IN and
        # the name) come out clean — matches what we already do for towns.
        if self._config.party_from_candidate and "," in label:
            name, _, party = label.partition(",")
            return " ".join(name.split()), party.strip().upper()
        return " ".join(label.split()), ""

    def _rows_at(self, row: int) -> Iterator[NormalizedRow]:
        cfg = self._config
        town_value = self._reader.cell_value(row, cfg.town_col)
        # Collapse internal whitespace runs ("Concord  Ward 1" -> "Concord Ward 1")
        # so downstream consumers don't have to guess about source-data noise.
        town = " ".join(str(town_value).split()) if town_value is not None else ""
        if self._should_skip_row(town):
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

    def _should_skip_row(self, town: str) -> bool:
        if not town:
            return True
        if town in self._config.skip_town_values:
            return True
        return False


# ---------------------------------------------------------------------------
# StatewideByCountyParser — scans sheets + sections, reuses CongressionalParser per section
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _Section:
    """One county section within a sheet."""
    header_row: int        # row index of the section header (cell 0 = county marker, cells 1+ = candidates)
    county: str | None     # canonical NH county name, or None for "skip this section"


class StatewideByCountyParser:
    """Parses multi-sheet workbooks with one or more county sections per sheet.

    For each sheet, scans rows for section headers: a row whose first cell
    (after stripping trailing ' County') matches a known NH county AND whose
    second cell is a non-numeric string (i.e. a candidate label, not a vote
    count). Each section is then parsed by an internal CongressionalParser
    bounded to that section's row range.

    The per-state summary block at the top of the first sheet is detected
    via `skip_sheet_markers` and skipped silently.
    """

    def __init__(self, config: StatewideByCountyConfig, path: pathlib.Path):
        self._config = config
        self._path = path

    def __iter__(self) -> Iterator[NormalizedRow]:
        for sheet_index in range(WorkbookReader.sheet_count(self._path)):
            reader = WorkbookReader(self._path, sheet_index=sheet_index)
            sections = _find_county_sections(reader, self._config)
            for section, next_start in _with_bounds(sections, reader.nrows):
                if section.county is None:
                    continue
                section_config = self._config_for_section(section.county, section.header_row)
                yield from CongressionalParser(section_config, reader, stop_row=next_start)

    def _config_for_section(self, county: str, header_row: int) -> CongressionalConfig:
        """Synthesize a CongressionalConfig for one county section.

        StatewideByCountyParser delegates section-level parsing to
        CongressionalParser. This translates between the two configs.
        """
        cfg = self._config
        return CongressionalConfig(
            office=cfg.office,
            sheet_index=0,  # ignored; reader is already bound
            header_row=header_row,
            town_col=cfg.town_col,
            candidate_cols_start=cfg.candidate_cols_start,
            county=county,
            district=cfg.district,
            party_from_candidate=cfg.party_from_candidate,
            skip_town_values=cfg.skip_town_values,
            skip_empty_votes=cfg.skip_empty_votes,
            lookup_county_from_town=False,
        )


def _find_county_sections(
    reader: WorkbookReader, config: StatewideByCountyConfig
) -> list[_Section]:
    """Scan a sheet row-by-row, identifying county-section header rows.

    A row is a section header iff:
    - cell at `section_marker_col` is a known skip marker (e.g. "Summary By Counties"), OR
    - cell at `section_marker_col`, stripped of trailing ' County', matches a known
      NH county name AND the next cell (column 1) contains a non-numeric string
      (i.e. a candidate label, not a vote count).

    The second clause's non-numeric check distinguishes a true section header
    from a row inside the summary block that happens to lead with a county name
    (e.g. ['Belknap', 20499, ...]).
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
        if not _looks_like_header_row(reader, row, marker_col):
            continue
        sections.append(_Section(header_row=row, county=candidate))
    return sections


def _looks_like_header_row(reader: WorkbookReader, row: int, marker_col: int) -> bool:
    """True if `row` looks like a section header (candidate labels in cells 1+)
    rather than a data row (vote counts in cells 1+)."""
    next_col = marker_col + 1
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


# ---------------------------------------------------------------------------
# ExecutiveCouncilParser — one district per sheet, district from sheet name
# ---------------------------------------------------------------------------


class ExecutiveCouncilParser:
    """Parses multi-sheet workbooks where each sheet is one whole district.

    Used for NH Executive Council general elections. For each sheet whose
    name matches `config.district_from_sheet_name`, the district number is
    captured and the rest of the sheet is parsed via CongressionalParser
    with `lookup_county_from_town=True` (since each row's town spans many
    counties across districts).
    """

    def __init__(self, config: ExecutiveCouncilConfig, path: pathlib.Path):
        self._config = config
        self._path = path

    def __iter__(self) -> Iterator[NormalizedRow]:
        for sheet_index in range(WorkbookReader.sheet_count(self._path)):
            reader = WorkbookReader(self._path, sheet_index=sheet_index)
            district = self._district_for_sheet(reader.sheet_name)
            if district is None:
                continue
            section_config = self._congressional_config_for_district(district)
            yield from CongressionalParser(section_config, reader)

    def _district_for_sheet(self, sheet_name: str) -> str | None:
        match = self._config.district_from_sheet_name.search(sheet_name)
        return match.group(1) if match else None

    def _congressional_config_for_district(self, district: str) -> CongressionalConfig:
        cfg = self._config
        return CongressionalConfig(
            office=cfg.office,
            sheet_index=0,  # ignored; reader is already bound
            header_row=cfg.header_row,
            town_col=cfg.town_col,
            candidate_cols_start=cfg.candidate_cols_start,
            county=None,
            district=district,
            party_from_candidate=cfg.party_from_candidate,
            skip_town_values=cfg.skip_town_values,
            skip_empty_votes=cfg.skip_empty_votes,
            lookup_county_from_town=True,
        )


# ---------------------------------------------------------------------------
# StateSenateParser — multi-sheet with within-sheet district scanning
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _DistrictSection:
    """One district section within a sheet (State Senate shape)."""
    marker_row: int   # row where 'State Senate District N' lives (cell 0)
    header_row: int   # marker_row + 1; row with candidate names in cells 1+
    district: str     # captured district number as a string


class StateSenateParser:
    """Parses multi-sheet workbooks where each sheet holds one or more districts.

    Used for NH State Senate. For each sheet (except those in
    `skip_sheet_names`), scans every row for the district marker pattern.
    Each match opens a section bounded by the next marker (or end of sheet)
    and is parsed by an internal CongressionalParser whose header is at
    marker_row + 1.
    """

    def __init__(self, config: StateSenateConfig, path: pathlib.Path):
        self._config = config
        self._path = path

    def __iter__(self) -> Iterator[NormalizedRow]:
        for sheet_index in range(WorkbookReader.sheet_count(self._path)):
            reader = WorkbookReader(self._path, sheet_index=sheet_index)
            if reader.sheet_name in self._config.skip_sheet_names:
                continue
            sections = self._find_district_sections(reader)
            for section, next_start in _with_district_bounds(sections, reader.nrows):
                cong_config = self._congressional_config_for_district(section)
                yield from CongressionalParser(cong_config, reader, stop_row=next_start)

    def _find_district_sections(self, reader: WorkbookReader) -> list[_DistrictSection]:
        sections: list[_DistrictSection] = []
        marker_re = self._config.district_section_marker
        for row in range(reader.nrows):
            cell = reader.cell_value(row, 0)
            label = "" if cell is None else str(cell).strip()
            if not label:
                continue
            match = marker_re.match(label)
            if match is None:
                continue
            sections.append(_DistrictSection(
                marker_row=row,
                header_row=row + 1,
                district=match.group(1),
            ))
        return sections

    def _congressional_config_for_district(
        self, section: _DistrictSection
    ) -> CongressionalConfig:
        cfg = self._config
        return CongressionalConfig(
            office=cfg.office,
            sheet_index=0,  # ignored; reader is already bound
            header_row=section.header_row,
            town_col=cfg.town_col,
            candidate_cols_start=cfg.candidate_cols_start,
            county=None,
            district=section.district,
            party_from_candidate=cfg.party_from_candidate,
            skip_town_values=cfg.skip_town_values,
            skip_empty_votes=cfg.skip_empty_votes,
            lookup_county_from_town=True,
            drop_column_labels=cfg.drop_column_labels,
        )


def _with_district_bounds(
    sections: list[_DistrictSection], total_rows: int
) -> Iterator[tuple[_DistrictSection, int]]:
    """Yield each district section paired with the next section's marker row
    (or `total_rows` for the last). Using the next marker_row (not header_row)
    as the bound ensures the next section's marker is excluded from the
    current section's data iteration."""
    for i, section in enumerate(sections):
        next_start = sections[i + 1].marker_row if i + 1 < len(sections) else total_rows
        yield section, next_start


# ---------------------------------------------------------------------------
# StateRepresentativeParser — per-county, many districts, multi-seat stripes
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _HouseDistrictSection:
    """One State Rep district section within a sheet."""
    marker_row: int   # row with 'District No. N (M) [F|FL]'
    district_num: str # captured district number as a string
    floterial: bool   # True if the marker carried F or FL
    skip: bool        # True if RECOUNT FIGURES — section discarded


class StateRepresentativeParser:
    """Parses per-county State House workbooks with stacked candidate stripes.

    Each file is one county, one sheet, many districts. Districts are
    separated by ``District No. N (M)`` markers. Within a district section,
    candidates may be split across multiple "stripes": the marker row
    itself is the first stripe's header, subsequent stripes begin where
    col 0 is blank and cols 1+ hold candidate-like (non-numeric) labels.
    Each stripe has its own per-town data rows.
    """

    def __init__(self, config: StateRepresentativeConfig, path: pathlib.Path):
        self._config = config
        self._path = path

    def __iter__(self) -> Iterator[NormalizedRow]:
        reader = WorkbookReader(self._path, sheet_index=0)
        sections = self._find_district_sections(reader)
        for i, section in enumerate(sections):
            section_end = sections[i + 1].marker_row if i + 1 < len(sections) else reader.nrows
            if section.skip:
                continue
            yield from self._parse_section(reader, section, section_end)

    def _find_district_sections(self, reader: WorkbookReader) -> list[_HouseDistrictSection]:
        sections: list[_HouseDistrictSection] = []
        for row in range(reader.nrows):
            cell = reader.cell_value(row, 0)
            label = "" if cell is None else str(cell).strip()
            if not label:
                continue
            match = self._config.district_marker.match(label)
            if match is None:
                continue
            sections.append(_HouseDistrictSection(
                marker_row=row,
                district_num=match.group(1),
                floterial=bool(match.group(2)),
                skip=bool(match.group(3)),
            ))
        return sections

    def _parse_section(
        self,
        reader: WorkbookReader,
        section: _HouseDistrictSection,
        section_end: int,
    ) -> Iterator[NormalizedRow]:
        district = self._format_district(section)
        stripes = self._find_stripes(reader, section.marker_row, section_end)
        for stripe_header, stripe_end in stripes:
            yield from self._parse_stripe(reader, district, stripe_header, stripe_end)

    def _format_district(self, section: _HouseDistrictSection) -> str:
        if section.floterial:
            return f"{section.district_num}{self._config.floterial_suffix}"
        return section.district_num

    def _find_stripes(
        self, reader: WorkbookReader, section_start: int, section_end: int
    ) -> list[tuple[int, int]]:
        """Return [(header_row, end_row_exclusive), ...] for each candidate stripe.

        First stripe header is `section_start` (the district marker row,
        whose cols 1+ hold the first batch of candidates). Subsequent
        stripe headers are rows where col 0 is blank/whitespace AND cols
        1+ contain non-numeric strings (candidate labels). Each stripe
        ends at the next stripe's header or `section_end`.
        """
        headers = [section_start]
        for row in range(section_start + 1, section_end):
            cell0 = reader.cell_value(row, 0)
            label0 = "" if cell0 is None else str(cell0).strip()
            if label0:
                continue
            if self._row_has_candidate_labels(reader, row):
                headers.append(row)
        return [
            (h, headers[i + 1] if i + 1 < len(headers) else section_end)
            for i, h in enumerate(headers)
        ]

    def _row_has_candidate_labels(self, reader: WorkbookReader, row: int) -> bool:
        """True if cells 1+ contain at least one non-empty non-numeric value.

        Used to distinguish a continuation stripe header (candidate names)
        from a blank-cell-0 data row (vote counts, which shouldn't normally
        appear with a blank town anyway, but be defensive)."""
        for col in range(self._config.candidate_cols_start, reader.ncols):
            val = reader.cell_value(row, col)
            if val is None or val == "":
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                return False
            if isinstance(val, str) and val.strip():
                return True
        return False

    def _parse_stripe(
        self,
        reader: WorkbookReader,
        district: str,
        header_row: int,
        stripe_end: int,
    ) -> Iterator[NormalizedRow]:
        candidates = self._read_stripe_candidates(reader, header_row)
        if not candidates:
            return
        cfg = self._config
        for row in range(header_row + 1, stripe_end):
            town_value = reader.cell_value(row, cfg.town_col)
            town = " ".join(str(town_value).split()) if town_value is not None else ""
            if not town or town in cfg.skip_town_values:
                continue
            for col, candidate, party in candidates:
                if col >= reader.ncols:
                    break
                raw = reader.cell_value(row, col)
                votes = _coerce_votes(raw)
                if votes is None and cfg.skip_empty_votes:
                    continue
                yield NormalizedRow(
                    county=cfg.county,
                    precinct=town,
                    office=cfg.office,
                    district=district,
                    party=party,
                    candidate=candidate,
                    votes=votes if votes is not None else 0,
                )

    def _read_stripe_candidates(
        self, reader: WorkbookReader, header_row: int
    ) -> list[tuple[int, str, str]]:
        """Return [(col_index, candidate_name, party), ...] for this stripe.

        Skips blank cells and any cell whose label is in `drop_column_labels`
        (case-insensitive). Collapses internal whitespace in candidate
        names so things like 'WRITE-IN   Kathy DesRoches' come out tidy."""
        row = reader.row_values(header_row)
        cfg = self._config
        drop_set = {label.casefold() for label in cfg.drop_column_labels}
        out: list[tuple[int, str, str]] = []
        for col in range(cfg.candidate_cols_start, len(row)):
            label = str(row[col]).strip() if row[col] is not None else ""
            if not label:
                continue
            if label.casefold() in drop_set:
                continue
            if cfg.party_from_candidate and "," in label:
                name, _, party = label.partition(",")
                out.append((col, " ".join(name.split()), party.strip().upper()))
            else:
                out.append((col, " ".join(label.split()), ""))
        return out


# ---------------------------------------------------------------------------
# Top-level factory
# ---------------------------------------------------------------------------


def parse_workbook(path: pathlib.Path, config: ParserConfig) -> Iterator[NormalizedRow]:
    """Top-level entry point. Dispatches on the config's type.

    Each shape has its own Parser class and Config dataclass. This factory
    chooses the right Parser for the supplied Config. Adding a new shape is
    a matter of writing a new Parser + Config and adding a branch here.
    """
    if isinstance(config, CongressionalConfig):
        reader = WorkbookReader(path, sheet_index=config.sheet_index)
        yield from CongressionalParser(config, reader)
        return
    if isinstance(config, StatewideByCountyConfig):
        yield from StatewideByCountyParser(config, path)
        return
    if isinstance(config, ExecutiveCouncilConfig):
        yield from ExecutiveCouncilParser(config, path)
        return
    if isinstance(config, StateSenateConfig):
        yield from StateSenateParser(config, path)
        return
    if isinstance(config, StateRepresentativeConfig):
        yield from StateRepresentativeParser(config, path)
        return
    raise TypeError(
        f"parse_workbook: unknown config type {type(config).__name__}. "
        f"Expected one of: CongressionalConfig, StatewideByCountyConfig, "
        f"ExecutiveCouncilConfig, StateSenateConfig, StateRepresentativeConfig."
    )

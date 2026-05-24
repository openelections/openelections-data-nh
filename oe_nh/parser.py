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

    party_from_candidate: bool = True
    skip_town_values: frozenset[str] = frozenset({"TOTALS", "Totals", "Total"})
    skip_empty_votes: bool = True


# Type alias for "any parser config", useful in shared callsites (jobs, cli).
ParserConfig = Union[
    CongressionalConfig,
    StatewideByCountyConfig,
    ExecutiveCouncilConfig,
    StateSenateConfig,
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
    raise TypeError(
        f"parse_workbook: unknown config type {type(config).__name__}. "
        f"Expected one of: CongressionalConfig, StatewideByCountyConfig, "
        f"ExecutiveCouncilConfig, StateSenateConfig."
    )

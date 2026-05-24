"""Tests for the multi-sheet / multi-section path through parse_workbook."""

from __future__ import annotations

import pathlib

import openpyxl

from oe_nh.parser import ParserConfig, parse_workbook


def _multi_sheet_xlsx(path: pathlib.Path, sheets: list[tuple[str, list[list]]]) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets:
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


def _county_section(county_label: str) -> list[list]:
    """A 4-row mini-section: header (county + candidate labels), two data rows, TOTALS row."""
    return [
        [county_label, "Smith, R", "Jones, D", "Write-Ins"],
        ["TownA", 100, 200, 5],
        ["TownB", 50, 75, 1],
        ["TOTALS", 150, 275, 6],
    ]


def _county_sheet(county: str) -> list[list]:
    """A 6-row mock sheet: title noise on rows 0-1, then a single county section."""
    return [
        ["", "State of New Hampshire", "", ""],
        ["", "Office", "", ""],
    ] + _county_section(f"{county} County")


def _summary_then_county_sheet(county: str) -> list[list]:
    """Mirrors the NH 2022 shape where sheet 0 contains the Summary block
    followed immediately by the first county's town-level data."""
    return [
        ["", "State of New Hampshire", "", ""],
        ["", "Office", "", ""],
        # Summary section
        ["Summary By Counties", "Smith, R", "Jones, D", "Write-Ins"],
        ["Belknap", 1000, 2000, 50],
        ["Carroll", 1100, 2200, 60],
        ["TOTALS", 2100, 4200, 110],
        # Followed by the county section
    ] + _county_section(f"{county} County")


def test_one_section_per_sheet(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _multi_sheet_xlsx(p, [
        ("belknap", _county_sheet("Belknap")),
        ("carroll", _county_sheet("Carroll")),
    ])
    cfg = ParserConfig(office="President", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    assert sorted({r.county for r in rows}) == ["Belknap", "Carroll"]
    # 2 counties * 2 towns * (2 candidates + Write-Ins) = 12; TOTALS skipped.
    assert len(rows) == 12


def test_skip_summary_section(tmp_path: pathlib.Path) -> None:
    """Sheet 0 has Summary block followed by Belknap; Summary is silently skipped."""
    p = tmp_path / "x.xlsx"
    _multi_sheet_xlsx(p, [
        ("sheet0", _summary_then_county_sheet("Belknap")),
    ])
    cfg = ParserConfig(office="President", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    # Only Belknap rows should be emitted, not the per-county summary roll-up
    assert all(r.county == "Belknap" for r in rows)
    # Two towns x 3 candidate-like columns = 6 rows
    assert len(rows) == 6


def test_multiple_sections_in_one_sheet(tmp_path: pathlib.Path) -> None:
    """Mirrors the NH 2022 'strafford and sullivan gov' shape."""
    p = tmp_path / "x.xlsx"
    contents = [
        ["", "State of New Hampshire", "", ""],
        ["", "Office", "", ""],
    ] + _county_section("Strafford") + [
        [],  # blank row separator
    ] + _county_section("Sullivan County")
    _multi_sheet_xlsx(p, [("combined", contents)])

    cfg = ParserConfig(office="Governor", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    assert sorted({r.county for r in rows}) == ["Strafford", "Sullivan"]
    # Each section: 2 towns x 3 candidate-like columns = 6 rows -> 12 total
    assert len(rows) == 12


def test_county_section_data_rows_not_treated_as_headers(tmp_path: pathlib.Path) -> None:
    """A data row whose first cell happens to be a county name (the Summary
    block's per-county totals: 'Belknap', 'Carroll', ...) must NOT be picked
    up as a section header. Distinguished by column 1 being numeric."""
    p = tmp_path / "x.xlsx"
    _multi_sheet_xlsx(p, [
        ("sheet0", _summary_then_county_sheet("Belknap")),
    ])
    cfg = ParserConfig(office="President", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    # If the Summary's 'Belknap' data row were mistakenly treated as a section
    # header, we'd get duplicated Belknap output. Confirm we only see Belknap
    # rows once (precinct=TownA or TownB), not "precinct=TOTALS" or any of
    # the summary-row county-name "precincts".
    precincts = {r.precinct for r in rows}
    assert precincts == {"TownA", "TownB"}


def test_skips_county_totals_row(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _multi_sheet_xlsx(p, [("c", _county_sheet("Belknap"))])
    cfg = ParserConfig(office="X", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    assert "TOTALS" not in {r.precinct for r in rows}


def test_county_label_with_or_without_suffix(tmp_path: pathlib.Path) -> None:
    """Both 'Hillsborough' and 'Hillsborough County' work as section markers."""
    p = tmp_path / "x.xlsx"
    contents = [
        ["", "State of New Hampshire", "", ""],
        ["", "Office", "", ""],
    ] + _county_section("Hillsborough") + [
        [],
    ] + _county_section("Coos County")
    _multi_sheet_xlsx(p, [("x", contents)])

    cfg = ParserConfig(office="X", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    assert sorted({r.county for r in rows}) == ["Coos", "Hillsborough"]


def test_county_label_trailing_whitespace(tmp_path: pathlib.Path) -> None:
    """'Strafford   ' (trailing whitespace) still matches."""
    p = tmp_path / "x.xlsx"
    _multi_sheet_xlsx(p, [("x", [
        ["", "noise", "", ""],
        ["", "noise", "", ""],
    ] + _county_section("Strafford   ")
    )])
    cfg = ParserConfig(office="X", multi_sheet=True)
    rows = list(parse_workbook(p, cfg))
    assert all(r.county == "Strafford" for r in rows)


def test_single_sheet_unchanged(tmp_path: pathlib.Path) -> None:
    """multi_sheet=False (Congressional-style) still goes through the simple path."""
    p = tmp_path / "x.xlsx"
    _multi_sheet_xlsx(p, [
        ("only", [
            ["", ""],
            ["", ""],
            ["", ""],
            ["Town", "Smith, R"],
            ["Albany", 100],
        ]),
    ])
    cfg = ParserConfig(office="X", county="C", header_row=3)
    rows = list(parse_workbook(p, cfg))
    assert len(rows) == 1
    assert rows[0].precinct == "Albany"
    assert rows[0].candidate == "Smith"

"""Unit tests for CongressionalParser + CongressionalConfig."""

from __future__ import annotations

import pathlib

import openpyxl

from oe_nh.parser import (
    CongressionalConfig,
    CongressionalParser,
    NormalizedRow,
    _coerce_votes,
)
from oe_nh.workbook import WorkbookReader


def _xlsx(path: pathlib.Path, rows: list[list]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)


def _read(path: pathlib.Path, config: CongressionalConfig) -> list[NormalizedRow]:
    return list(CongressionalParser(config, WorkbookReader(path)))


def test_basic_two_candidates(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["Town", "Trump, R", "Biden, D"],
        ["Albany", 142, 202],
        ["Bartlett", 730, 1094],
    ])
    rows = _read(p, CongressionalConfig(office="President", county="Carroll"))
    assert rows == [
        NormalizedRow("Carroll", "Albany", "President", "", "R", "Trump", 142),
        NormalizedRow("Carroll", "Albany", "President", "", "D", "Biden", 202),
        NormalizedRow("Carroll", "Bartlett", "President", "", "R", "Trump", 730),
        NormalizedRow("Carroll", "Bartlett", "President", "", "D", "Biden", 1094),
    ]


def test_blank_rows_skipped(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["Town", "Trump, R"],
        ["Albany", 100],
        ["", ""],            # blank town -> skip
        ["Bartlett", 50],
    ])
    rows = _read(p, CongressionalConfig(office="President", county="Carroll"))
    assert [r.precinct for r in rows] == ["Albany", "Bartlett"]


def test_skip_town_values(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", ""],
        ["", ""],
        ["", ""],
        ["Town", "Trump, R"],
        ["Albany", 100],
        ["County Total", 999],   # noise we want to drop
        ["Bartlett", 50],
    ])
    rows = _read(
        p,
        CongressionalConfig(
            office="President",
            county="Carroll",
            skip_town_values=frozenset({"County Total"}),
        ),
    )
    assert [r.precinct for r in rows] == ["Albany", "Bartlett"]


def test_party_in_candidate_label(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", ""], ["", ""], ["", ""],
        ["Town", "Smith"],   # no party in label
        ["Albany", 100],
    ])
    rows = _read(p, CongressionalConfig(office="X", county="Y"))
    assert rows[0].candidate == "Smith"
    assert rows[0].party == ""


def test_empty_votes_skipped_by_default(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", "", ""], ["", "", ""], ["", "", ""],
        ["Town", "Trump, R", "Biden, D"],
        ["Albany", 100, ""],
    ])
    rows = _read(p, CongressionalConfig(office="X", county="Y"))
    assert [(r.candidate, r.votes) for r in rows] == [("Trump", 100)]


def test_empty_votes_emitted_as_zero_when_configured(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", "", ""], ["", "", ""], ["", "", ""],
        ["Town", "Trump, R", "Biden, D"],
        ["Albany", 100, ""],
    ])
    rows = _read(p, CongressionalConfig(office="X", county="Y", skip_empty_votes=False))
    assert [(r.candidate, r.votes) for r in rows] == [("Trump", 100), ("Biden", 0)]


def test_district_set_on_every_row(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", ""], ["", ""], ["", ""],
        ["Town", "Smith, R"],
        ["Albany", 100],
    ])
    rows = _read(p, CongressionalConfig(office="Congressional", county="Carroll", district="1"))
    assert rows[0].district == "1"


def test_county_empty_string_when_none(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", ""], ["", ""], ["", ""],
        ["Town", "Smith, R"],
        ["Albany", 100],
    ])
    rows = _read(p, CongressionalConfig(office="X"))
    assert rows[0].county == ""


def test_float_int_values_coerced(tmp_path: pathlib.Path) -> None:
    """xlsx often serializes ints as float."""
    p = tmp_path / "x.xlsx"
    _xlsx(p, [
        ["", ""], ["", ""], ["", ""],
        ["Town", "Smith, R"],
        ["Albany", 100.0],
    ])
    rows = _read(p, CongressionalConfig(office="X"))
    assert rows[0].votes == 100
    assert isinstance(rows[0].votes, int)


def test_coerce_votes_unit() -> None:
    """Standalone coverage for the _coerce_votes helper."""
    assert _coerce_votes(0) == 0
    assert _coerce_votes(42) == 42
    assert _coerce_votes(42.0) == 42
    assert _coerce_votes(42.5) is None      # true float not accepted
    assert _coerce_votes("100") == 100
    assert _coerce_votes("100.0") == 100
    assert _coerce_votes("") is None
    assert _coerce_votes("--") is None
    assert _coerce_votes("-") is None
    assert _coerce_votes("not a number") is None
    assert _coerce_votes(None) is None
    assert _coerce_votes(True) is None      # bool is not a vote

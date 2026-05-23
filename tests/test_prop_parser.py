"""Property-based tests for the Parser.

We synthesize a SheetShape, materialize it as an xlsx, parse it, and assert
invariants about the resulting rows. The xls path is exercised separately
via real-file fixtures once raw NH SoS files are committed.
"""

from __future__ import annotations

import pathlib

import openpyxl
from hypothesis import given, settings
from hypothesis import strategies as st

from oe_nh.parser import NormalizedRow, Parser, ParserConfig
from oe_nh.workbook import WorkbookReader
from tests.strategies import SynthesizedSheet, synthesized_sheets


def _materialize(sheet: SynthesizedSheet, path: pathlib.Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    for _ in range(sheet.header_row):
        ws.append([""])
    ws.append(["Town"] + [sheet.header_label(name, party) for name, party in sheet.candidates])
    for town, row in zip(sheet.towns, sheet.votes):
        ws.append([town] + row)
    wb.save(path)


def _parse(sheet: SynthesizedSheet, path: pathlib.Path) -> list[NormalizedRow]:
    reader = WorkbookReader(path)
    config = ParserConfig(
        office="Office",
        county="County",
        header_row=sheet.header_row,
        candidate_cols_start=1,
    )
    return list(Parser(config, reader))


@given(sheet=synthesized_sheets())
@settings(deadline=None)  # I/O makes per-example timing noisy
def test_parser_emits_one_row_per_town_per_candidate(tmp_path_factory, sheet):
    path = tmp_path_factory.mktemp("p") / "x.xlsx"
    _materialize(sheet, path)

    rows = _parse(sheet, path)

    expected_count = len(sheet.towns) * len(sheet.candidates)
    assert len(rows) == expected_count


@given(sheet=synthesized_sheets())
@settings(deadline=None)
def test_parser_preserves_vote_counts(tmp_path_factory, sheet):
    path = tmp_path_factory.mktemp("p") / "x.xlsx"
    _materialize(sheet, path)

    rows = _parse(sheet, path)
    by_key = {(r.precinct, r.candidate): r.votes for r in rows}
    for ti, town in enumerate(sheet.towns):
        for ci, (cand, _) in enumerate(sheet.candidates):
            assert by_key[(town, cand)] == sheet.votes[ti][ci]


@given(sheet=synthesized_sheets())
@settings(deadline=None)
def test_parser_emits_non_negative_votes(tmp_path_factory, sheet):
    path = tmp_path_factory.mktemp("p") / "x.xlsx"
    _materialize(sheet, path)
    rows = _parse(sheet, path)
    for r in rows:
        assert r.votes >= 0


@given(sheet=synthesized_sheets())
@settings(deadline=None)
def test_parser_always_sets_county_and_office(tmp_path_factory, sheet):
    path = tmp_path_factory.mktemp("p") / "x.xlsx"
    _materialize(sheet, path)
    rows = _parse(sheet, path)
    for r in rows:
        assert r.county == "County"
        assert r.office == "Office"
        assert r.precinct  # never blank — strategy filters blank-after-strip


@given(sheet=synthesized_sheets())
@settings(deadline=None)
def test_parser_roundtrip_preserves_party(tmp_path_factory, sheet):
    """If a candidate label is 'Name, PARTY' on input, party should come out PARTY."""
    path = tmp_path_factory.mktemp("p") / "x.xlsx"
    _materialize(sheet, path)
    rows = _parse(sheet, path)

    # Build expected map from sheet; party '' means no comma in label.
    expected = {name: (party.upper() if party else "") for name, party in sheet.candidates}
    for r in rows:
        assert r.party == expected[r.candidate]

"""Tests for CSVWriter."""

from __future__ import annotations

import csv
import pathlib

from oe_nh.parser import NormalizedRow
from oe_nh.writer import HEADERS, write_csv


def test_writes_header_and_rows(tmp_path: pathlib.Path) -> None:
    rows = [
        NormalizedRow("Carroll", "Albany", "President", "", "R", "Trump", 142),
        NormalizedRow("Carroll", "Albany", "President", "", "D", "Biden", 202),
    ]
    out = tmp_path / "20241105__nh__general__president__precinct.csv"
    count = write_csv(out, rows)
    assert count == 2

    with open(out) as fh:
        reader = csv.reader(fh)
        all_rows = list(reader)
    assert all_rows[0] == HEADERS
    assert all_rows[1] == ["Carroll", "Albany", "President", "", "R", "Trump", "142"]
    assert all_rows[2] == ["Carroll", "Albany", "President", "", "D", "Biden", "202"]


def test_creates_parent_directories(tmp_path: pathlib.Path) -> None:
    rows = [NormalizedRow("C", "P", "O", "", "", "Smith", 1)]
    out = tmp_path / "deep" / "nested" / "out.csv"
    write_csv(out, rows)
    assert out.exists()


def test_empty_rows_writes_only_header(tmp_path: pathlib.Path) -> None:
    out = tmp_path / "empty.csv"
    count = write_csv(out, [])
    assert count == 0
    with open(out) as fh:
        rows = list(csv.reader(fh))
    assert rows == [HEADERS]

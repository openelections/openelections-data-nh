"""Unit tests for oe_nh.discovery."""

from __future__ import annotations

import pathlib

import openpyxl

from oe_nh.discovery import discover_files, merge
from oe_nh.parser import CongressionalConfig


def _make(path: pathlib.Path) -> None:
    """Create a minimal valid xlsx so the file exists with a real extension."""
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".xlsx":
        wb = openpyxl.Workbook()
        wb.save(path)
    else:
        path.write_bytes(b"\xd0\xcf\x11\xe0fake")


def test_discover_empty_folder(tmp_path: pathlib.Path) -> None:
    assert discover_files(tmp_path, "president", "President") == []


def test_discover_missing_folder(tmp_path: pathlib.Path) -> None:
    assert discover_files(tmp_path / "nope", "president", "President") == []


def test_discover_single_statewide_file(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "us-senate.xlsx")
    out = discover_files(tmp_path, "us-senate", "US Senate")
    assert out == [("us-senate.xlsx", CongressionalConfig(office="US Senate"))]


def test_discover_by_county(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "president-belknap.xlsx")
    _make(tmp_path / "president-carroll.xlsx")
    _make(tmp_path / "president-cheshire.xlsx")
    out = discover_files(tmp_path, "president", "President")
    assert out == [
        ("president-belknap.xlsx", CongressionalConfig(office="President", county="Belknap")),
        ("president-carroll.xlsx", CongressionalConfig(office="President", county="Carroll")),
        ("president-cheshire.xlsx", CongressionalConfig(office="President", county="Cheshire")),
    ]


def test_discover_district_from_number(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "congressional-1.xlsx")
    _make(tmp_path / "congressional-2.xlsx")
    out = discover_files(tmp_path, "congressional", "Congressional")
    assert out == [
        ("congressional-1.xlsx", CongressionalConfig(office="Congressional", district="1")),
        ("congressional-2.xlsx", CongressionalConfig(office="Congressional", district="2")),
    ]


def test_discover_district_with_prefix(tmp_path: pathlib.Path) -> None:
    """`cd-1.xlsx` should still extract district='1'."""
    _make(tmp_path / "congressional-cd-1.xlsx")
    out = discover_files(tmp_path, "congressional", "Congressional")
    assert out == [
        ("congressional-cd-1.xlsx", CongressionalConfig(office="Congressional", district="1")),
    ]


def test_discover_ignores_non_workbooks(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "president-belknap.xlsx")
    (tmp_path / "notes.txt").write_text("ignore me")
    (tmp_path / "president-belknap.pdf").write_bytes(b"%PDF-")
    out = discover_files(tmp_path, "president", "President")
    assert [name for name, _ in out] == ["president-belknap.xlsx"]


def test_discover_ignores_unrelated_office(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "governor-belknap.xlsx")
    _make(tmp_path / "us-senate.xlsx")
    out = discover_files(tmp_path, "president", "President")
    assert out == []


def test_discover_mixed_xls_xlsx(tmp_path: pathlib.Path) -> None:
    """Convention is case-insensitive on extension and supports both."""
    _make(tmp_path / "president-belknap.xls")
    _make(tmp_path / "president-carroll.xlsx")
    out = discover_files(tmp_path, "president", "President")
    assert [name for name, _ in out] == ["president-belknap.xls", "president-carroll.xlsx"]


def test_discover_unknown_location_falls_back(tmp_path: pathlib.Path) -> None:
    """Unknown location segments don't crash; we just emit an empty county/district."""
    _make(tmp_path / "president-mystery.xlsx")
    out = discover_files(tmp_path, "president", "President")
    assert out == [("president-mystery.xlsx", CongressionalConfig(office="President"))]


def test_merge_no_overlap_concats() -> None:
    discovered = [("a.xls", CongressionalConfig(office="X"))]
    explicit = [("b.xls", CongressionalConfig(office="X", county="Y"))]
    out = merge(discovered, explicit)
    assert out == [
        ("a.xls", CongressionalConfig(office="X")),
        ("b.xls", CongressionalConfig(office="X", county="Y")),
    ]


def test_merge_explicit_overrides_discovered() -> None:
    """If the same filename appears in both, the explicit config wins."""
    discovered = [("a.xls", CongressionalConfig(office="X", county="WrongCounty"))]
    explicit = [("a.xls", CongressionalConfig(office="X", county="RightCounty"))]
    out = merge(discovered, explicit)
    assert out == [("a.xls", CongressionalConfig(office="X", county="RightCounty"))]


def test_merge_preserves_discovered_order() -> None:
    discovered = [
        ("a.xls", CongressionalConfig(office="X")),
        ("b.xls", CongressionalConfig(office="X")),
    ]
    explicit = [("c.xls", CongressionalConfig(office="X"))]
    out = merge(discovered, explicit)
    assert [name for name, _ in out] == ["a.xls", "b.xls", "c.xls"]

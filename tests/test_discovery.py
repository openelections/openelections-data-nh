"""Unit tests for oe_nh.discovery (per-shape dispatch + legacy fallback)."""

from __future__ import annotations

import pathlib

import openpyxl
import pytest

from oe_nh.discovery import _normalize_sos_stem, discover_files, merge
from oe_nh.parser import (
    CongressionalConfig,
    ExecutiveCouncilConfig,
    StateRepresentativeConfig,
    StateSenateConfig,
    StatewideByCountyConfig,
)


def _make(path: pathlib.Path) -> None:
    """Create a minimal valid xlsx so the file exists with a real extension."""
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".xlsx":
        wb = openpyxl.Workbook()
        wb.save(path)
    else:
        path.write_bytes(b"\xd0\xcf\x11\xe0fake")


# ---------------------------------------------------------------------------
# Statewide-shape dispatch (president, governor, us-senate, executive-council,
# state-senate) — exactly one canonical file per office.
# ---------------------------------------------------------------------------


def test_discover_president_statewide(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "president.xls")
    out = discover_files(tmp_path, "president", "President")
    assert out == [(
        "president.xls",
        StatewideByCountyConfig(office="President", header_row=2),
    )]


def test_discover_governor_statewide(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "governor.xlsx")
    out = discover_files(tmp_path, "governor", "Governor")
    assert out == [(
        "governor.xlsx",
        StatewideByCountyConfig(office="Governor", header_row=2),
    )]


def test_discover_us_senate_statewide(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "us-senate.xlsx")
    out = discover_files(tmp_path, "us-senate", "US Senate")
    assert out == [(
        "us-senate.xlsx",
        StatewideByCountyConfig(office="US Senate", header_row=2),
    )]


def test_discover_executive_council(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "executive-council.xls")
    out = discover_files(tmp_path, "executive-council", "Executive Council")
    assert out == [(
        "executive-council.xls",
        ExecutiveCouncilConfig(office="Executive Council"),
    )]


def test_discover_state_senate(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "state-senate.xls")
    out = discover_files(tmp_path, "state-senate", "State Senate")
    assert out == [(
        "state-senate.xls",
        StateSenateConfig(office="State Senate"),
    )]


def test_statewide_dispatch_ignores_per_county_files(tmp_path: pathlib.Path) -> None:
    """A statewide-only office should not pick up per-county-looking files."""
    _make(tmp_path / "president.xls")
    _make(tmp_path / "president-belknap.xls")
    out = discover_files(tmp_path, "president", "President")
    assert [name for name, _ in out] == ["president.xls"]


# ---------------------------------------------------------------------------
# Congressional dispatch (one file per district, district captured from filename)
# ---------------------------------------------------------------------------


def test_discover_congressional_districts(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "congressional-1.xlsx")
    _make(tmp_path / "congressional-2.xlsx")
    out = discover_files(tmp_path, "congressional", "Congressional")
    assert out == [
        ("congressional-1.xlsx", CongressionalConfig(
            office="Congressional", district="1", header_row=2,
            lookup_county_from_town=True,
        )),
        ("congressional-2.xlsx", CongressionalConfig(
            office="Congressional", district="2", header_row=2,
            lookup_county_from_town=True,
        )),
    ]


def test_discover_congressional_district_from_prefix(tmp_path: pathlib.Path) -> None:
    """`congressional-cd-1.xlsx` should still extract district='1'."""
    _make(tmp_path / "congressional-cd-1.xlsx")
    out = discover_files(tmp_path, "congressional", "Congressional")
    assert out == [(
        "congressional-cd-1.xlsx",
        CongressionalConfig(
            office="Congressional", district="1", header_row=2,
            lookup_county_from_town=True,
        ),
    )]


# ---------------------------------------------------------------------------
# State Representative dispatch (one file per county, prefix is "house-")
# ---------------------------------------------------------------------------


def test_discover_state_representative_by_county(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "house-belknap.xls")
    _make(tmp_path / "house-carroll.xlsx")
    _make(tmp_path / "house-cheshire.xls")
    out = discover_files(tmp_path, "state-representative", "State Representative")
    assert out == [
        ("house-belknap.xls", StateRepresentativeConfig(
            office="State Representative", county="Belknap",
        )),
        ("house-carroll.xlsx", StateRepresentativeConfig(
            office="State Representative", county="Carroll",
        )),
        ("house-cheshire.xls", StateRepresentativeConfig(
            office="State Representative", county="Cheshire",
        )),
    ]


def test_discover_state_representative_ignores_unknown_county(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "house-mars.xls")
    out = discover_files(tmp_path, "state-representative", "State Representative")
    assert out == []


def test_discover_state_representative_ignores_non_house_files(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "house-belknap.xls")
    _make(tmp_path / "executive-council.xls")
    out = discover_files(tmp_path, "state-representative", "State Representative")
    assert [name for name, _ in out] == ["house-belknap.xls"]


# ---------------------------------------------------------------------------
# Generic discovery behaviors (apply regardless of dispatch path)
# ---------------------------------------------------------------------------


def test_discover_empty_folder(tmp_path: pathlib.Path) -> None:
    assert discover_files(tmp_path, "president", "President") == []


def test_discover_missing_folder(tmp_path: pathlib.Path) -> None:
    assert discover_files(tmp_path / "nope", "president", "President") == []


def test_discover_ignores_non_workbook_extensions(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "president.xls")
    (tmp_path / "notes.txt").write_text("ignore me")
    (tmp_path / "president.pdf").write_bytes(b"%PDF-")
    out = discover_files(tmp_path, "president", "President")
    assert [name for name, _ in out] == ["president.xls"]


def test_discover_xls_and_xlsx_extensions(tmp_path: pathlib.Path) -> None:
    """Both .xls and .xlsx are recognized; only office_slug pattern matters."""
    _make(tmp_path / "house-belknap.xls")
    _make(tmp_path / "house-carroll.xlsx")
    out = discover_files(tmp_path, "state-representative", "State Representative")
    assert [name for name, _ in out] == ["house-belknap.xls", "house-carroll.xlsx"]


# ---------------------------------------------------------------------------
# Legacy fallback for unknown office slugs
# ---------------------------------------------------------------------------


def test_unknown_slug_falls_back_to_legacy_single_file(tmp_path: pathlib.Path) -> None:
    """An office_slug not in _DISPATCH falls back to the legacy
    CongressionalConfig-only discovery."""
    _make(tmp_path / "future-office.xls")
    out = discover_files(tmp_path, "future-office", "Future Office")
    assert out == [(
        "future-office.xls",
        CongressionalConfig(office="Future Office"),
    )]


def test_unknown_slug_falls_back_to_legacy_per_county(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "future-office-belknap.xls")
    out = discover_files(tmp_path, "future-office", "Future Office")
    assert out == [(
        "future-office-belknap.xls",
        CongressionalConfig(office="Future Office", county="Belknap"),
    )]


# ---------------------------------------------------------------------------
# merge() — unchanged from before; explicit configs override discovered ones
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# SoS-shape normalization (strip year prefix, election prefix, _N revision,
# -district-N-M suffix, and the standalone "district" word in compound names)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem,expected", [
    # Already canonical → unchanged
    ("governor", "governor"),
    ("house-belknap", "house-belknap"),
    ("congressional-1", "congressional-1"),
    ("executive-council", "executive-council"),
    # Year prefix
    ("2024-governor", "governor"),
    ("2022-governor", "governor"),
    # Election prefix (ge = General, sp = State Primary, pp = Presidential Primary)
    ("2024-ge-house-belknap", "house-belknap"),
    ("2022-ge-governor", "governor"),
    # _N revision suffix
    ("house-belknap_1", "house-belknap"),
    ("2024-ge-house-belknap_2", "house-belknap"),
    # Multi-district statewide suffix (Exec Council 1-5, State Senate 1-24)
    ("executive-council-district-1-5", "executive-council"),
    ("2022-executive-council-district-1-5_0", "executive-council"),
    ("state-senate-district-1-24", "state-senate"),
    ("2022-ge-state-senate-district-1-24_1", "state-senate"),
    # Standalone "district" word in single-district compounds
    ("congressional-district-1", "congressional-1"),
    ("congressional-district-2", "congressional-2"),
    # All decorations at once
    ("2024-ge-state-senate-district-1-24_4", "state-senate"),
    # Unknown stems pass through (discovery decides whether they match)
    ("ballots-cast", "ballots-cast"),
    ("notes", "notes"),
])
def test_normalize_sos_stem(stem: str, expected: str) -> None:
    assert _normalize_sos_stem(stem) == expected


# ---------------------------------------------------------------------------
# Discovery picks up SoS-shaped filenames (not just canonical)
# ---------------------------------------------------------------------------


def test_discover_governor_accepts_sos_form(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "2024-ge-governor.xls")
    out = discover_files(tmp_path, "governor", "Governor")
    assert out == [(
        "2024-ge-governor.xls",
        StatewideByCountyConfig(office="Governor", header_row=2),
    )]


def test_discover_state_senate_accepts_sos_form(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "2022-ge-state-senate-district-1-24_1.xls")
    out = discover_files(tmp_path, "state-senate", "State Senate")
    assert out == [(
        "2022-ge-state-senate-district-1-24_1.xls",
        StateSenateConfig(office="State Senate"),
    )]


def test_discover_executive_council_accepts_sos_form(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "2022-executive-council-district-1-5_0.xls")
    out = discover_files(tmp_path, "executive-council", "Executive Council")
    assert out == [(
        "2022-executive-council-district-1-5_0.xls",
        ExecutiveCouncilConfig(office="Executive Council"),
    )]


def test_discover_house_accepts_sos_form(tmp_path: pathlib.Path) -> None:
    _make(tmp_path / "2024-ge-house-belknap_2.xls")
    _make(tmp_path / "house-carroll.xlsx")
    out = discover_files(tmp_path, "state-representative", "State Representative")
    assert out == [
        ("2024-ge-house-belknap_2.xls", StateRepresentativeConfig(
            office="State Representative", county="Belknap",
        )),
        ("house-carroll.xlsx", StateRepresentativeConfig(
            office="State Representative", county="Carroll",
        )),
    ]


def test_discover_congressional_accepts_district_word(tmp_path: pathlib.Path) -> None:
    """`congressional-district-1.xlsx` (which Tom encountered) should match."""
    _make(tmp_path / "congressional-district-1.xlsx")
    _make(tmp_path / "congressional-district-2.xlsx")
    out = discover_files(tmp_path, "congressional", "Congressional")
    assert out == [
        ("congressional-district-1.xlsx", CongressionalConfig(
            office="Congressional", district="1", header_row=2,
            lookup_county_from_town=True,
        )),
        ("congressional-district-2.xlsx", CongressionalConfig(
            office="Congressional", district="2", header_row=2,
            lookup_county_from_town=True,
        )),
    ]

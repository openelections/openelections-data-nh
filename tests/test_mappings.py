"""Smoke tests for the town_to_county mapping."""

from __future__ import annotations

from oe_nh.mappings.town_to_county import TOWN_TO_COUNTY


VALID_COUNTIES = frozenset({
    "Belknap", "Carroll", "Cheshire", "Coos", "Grafton",
    "Hillsborough", "Merrimack", "Rockingham", "Strafford", "Sullivan",
})


def test_all_counties_are_valid() -> None:
    seen = set(TOWN_TO_COUNTY.values())
    assert seen == VALID_COUNTIES, f"unexpected counties: {seen - VALID_COUNTIES}"


def test_no_blank_keys() -> None:
    for town in TOWN_TO_COUNTY:
        assert town.strip(), f"blank town key: {town!r}"


def test_known_towns() -> None:
    assert TOWN_TO_COUNTY["Manchester"] == "Hillsborough"
    assert TOWN_TO_COUNTY["Concord"] == "Merrimack"
    assert TOWN_TO_COUNTY["Portsmouth"] == "Rockingham"


def test_size_at_or_above_seed() -> None:
    """Seeded from 2012/code/county.pkl with 259 entries."""
    assert len(TOWN_TO_COUNTY) >= 259

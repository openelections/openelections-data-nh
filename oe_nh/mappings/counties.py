"""Canonical NH counties and a slug<->name lookup.

NH has had the same 10 counties since the 19th century. Stable list.
"""

from __future__ import annotations


# Canonical names as they should appear in CSV output.
NH_COUNTIES: tuple[str, ...] = (
    "Belknap",
    "Carroll",
    "Cheshire",
    "Coos",
    "Grafton",
    "Hillsborough",
    "Merrimack",
    "Rockingham",
    "Strafford",
    "Sullivan",
)


def _slug(name: str) -> str:
    return name.lower().replace(" ", "-")


COUNTY_BY_SLUG: dict[str, str] = {_slug(c): c for c in NH_COUNTIES}


def county_from_slug(slug: str) -> str | None:
    """Return the canonical county name for a slug (case-insensitive)."""
    return COUNTY_BY_SLUG.get(slug.strip().lower())

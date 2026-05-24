"""Auto-discover raw workbook files under a Job's folder and build the right
per-shape config for each one.

The canonical naming convention (relative to a Job's `folder`):

    president.xls[x]                  StatewideByCountyConfig (one statewide file)
    governor.xls[x]                   StatewideByCountyConfig
    us-senate.xls[x]                  StatewideByCountyConfig
    executive-council.xls[x]          ExecutiveCouncilConfig
    state-senate.xls[x]               StateSenateConfig
    congressional-<N>.xls[x]          CongressionalConfig (one per CD)
    house-<county>.xls[x]             StateRepresentativeConfig (one per county)

A Job whose `office_slug` is one of the keys above triggers the matching
dispatch entry. Discovery iterates the Job's folder, picks files matching
the slug's filename pattern, and instantiates the right Config dataclass
with sensible defaults that match NH SoS reality (e.g. `header_row=2`).

Unknown office slugs fall back to legacy CongressionalConfig discovery
(the original behavior) so the existing convention still works for
hypothetical future single-sheet offices.
"""

from __future__ import annotations

import pathlib
import re
from dataclasses import dataclass
from typing import Callable

from oe_nh.mappings.counties import county_from_slug
from oe_nh.parser import (
    CongressionalConfig,
    ExecutiveCouncilConfig,
    ParserConfig,
    StateRepresentativeConfig,
    StateSenateConfig,
    StatewideByCountyConfig,
)


_EXTS = (".xls", ".xlsx")
_DIGITS = re.compile(r"\d+")

# Decorations the NH SoS puts on filenames that we strip before matching:
_SOS_YEAR_PREFIX = re.compile(r"^\d{4}-")
_SOS_ELECTION_PREFIX = re.compile(r"^(?:ge|gn|sp|pp)-")
_SOS_REVISION_SUFFIX = re.compile(r"_\d+$")
# Multi-district statewide files end in `-district-N-M` (e.g.
# `executive-council-district-1-5` for Exec Council's 5 districts, or
# `state-senate-district-1-24` for State Senate's 24).
_SOS_DISTRICT_RANGE_SUFFIX = re.compile(r"-district-\d+-\d+$")
# Single-district files sometimes spell it `-district-N` (e.g.
# `congressional-district-1`); compress to `-N`.
_SOS_DISTRICT_WORD = re.compile(r"-district-(\d+)$")


def _normalize_sos_stem(stem: str) -> str:
    """Strip common NH SoS upstream decorations from a filename stem to expose
    the canonical form the dispatch table expects.

    Examples (left = SoS form, right = canonical):
      `2024-ge-house-belknap_1`             -> `house-belknap`
      `2022-executive-council-district-1-5_0` -> `executive-council`
      `2022-ge-state-senate-district-1-24_1`  -> `state-senate`
      `congressional-district-1`            -> `congressional-1`
      `governor`                            -> `governor`  (no change)

    Idempotent — already-canonical stems pass through unchanged.
    """
    s = _SOS_YEAR_PREFIX.sub("", stem)
    s = _SOS_ELECTION_PREFIX.sub("", s)
    s = _SOS_REVISION_SUFFIX.sub("", s)
    s = _SOS_DISTRICT_RANGE_SUFFIX.sub("", s)
    s = _SOS_DISTRICT_WORD.sub(r"-\1", s)
    return s


@dataclass(frozen=True)
class _OfficeDispatch:
    """How to find files and build configs for one office slug.

    `display_name` is the human-readable office name that ends up in the
    `office` column of every output row (e.g. 'State Representative').

    `filename_pattern` is one of:
    - 'statewide': matches `<office_slug>.xls[x]` only.
    - 'congressional': matches `<office_slug>-<digits>.xls[x]`, captures the
       leading digits as the district.
    - 'house-county': matches `house-<county_slug>.xls[x]`, captures the
       canonical county name.

    `config_factory(office_name, location)` builds the Config for one match.
    `location` is the captured digits or county; the empty string for statewide.
    """
    display_name: str
    filename_pattern: str
    config_factory: Callable[[str, str], ParserConfig]


_DISPATCH: dict[str, _OfficeDispatch] = {
    "president": _OfficeDispatch(
        display_name="President",
        filename_pattern="statewide",
        config_factory=lambda office_name, _loc: StatewideByCountyConfig(
            office=office_name, header_row=2,
        ),
    ),
    "governor": _OfficeDispatch(
        display_name="Governor",
        filename_pattern="statewide",
        config_factory=lambda office_name, _loc: StatewideByCountyConfig(
            office=office_name, header_row=2,
        ),
    ),
    "us-senate": _OfficeDispatch(
        display_name="US Senate",
        filename_pattern="statewide",
        config_factory=lambda office_name, _loc: StatewideByCountyConfig(
            office=office_name, header_row=2,
        ),
    ),
    "congressional": _OfficeDispatch(
        display_name="Congressional",
        filename_pattern="congressional",
        config_factory=lambda office_name, district: CongressionalConfig(
            office=office_name, district=district, header_row=2,
            lookup_county_from_town=True,
        ),
    ),
    "executive-council": _OfficeDispatch(
        display_name="Executive Council",
        filename_pattern="statewide",
        config_factory=lambda office_name, _loc: ExecutiveCouncilConfig(
            office=office_name,
        ),
    ),
    "state-senate": _OfficeDispatch(
        display_name="State Senate",
        filename_pattern="statewide",
        config_factory=lambda office_name, _loc: StateSenateConfig(
            office=office_name,
        ),
    ),
    "state-representative": _OfficeDispatch(
        display_name="State Representative",
        filename_pattern="house-county",
        config_factory=lambda office_name, county: StateRepresentativeConfig(
            office=office_name, county=county,
        ),
    ),
}


def office_display_name(slug: str) -> str | None:
    """Canonical human-readable name for an office slug, or None if unknown."""
    dispatch = _DISPATCH.get(slug)
    return dispatch.display_name if dispatch else None


def registered_office_slugs() -> list[str]:
    """All office slugs known to the dispatch table, in registration order."""
    return list(_DISPATCH.keys())


def discover_files(
    folder: pathlib.Path,
    office_slug: str,
    office_name: str,
) -> list[tuple[str, ParserConfig]]:
    """Return [(filename, ParserConfig)] for every file in `folder` matching the
    office's canonical convention.

    Filenames are leaf names (not paths); callers compose with `folder` to get
    a real path. Output is sorted for stable, reviewable Job output.

    Unknown office slugs fall back to the legacy single-sheet
    CongressionalConfig discovery."""
    if not folder.is_dir():
        return []

    dispatch = _DISPATCH.get(office_slug)
    if dispatch is None:
        return _discover_legacy(folder, office_slug, office_name)

    out: list[tuple[str, ParserConfig]] = []
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in _EXTS:
            continue
        canonical_stem = _normalize_sos_stem(path.stem)
        location = _match_filename(canonical_stem, office_slug, dispatch.filename_pattern)
        if location is None:
            continue
        config = dispatch.config_factory(office_name, location)
        out.append((path.name, config))
    return out


def _match_filename(stem: str, office_slug: str, pattern: str) -> str | None:
    """Return the captured location (or '' for statewide-no-location) if `stem`
    matches the pattern, else None. `stem` is the canonical form (post-
    `_normalize_sos_stem`)."""
    if pattern == "statewide":
        return "" if stem == office_slug else None

    if pattern == "congressional":
        prefix = f"{office_slug}-"
        if not stem.startswith(prefix):
            return None
        location = stem[len(prefix):]
        match = _DIGITS.search(location)
        return match.group(0) if match else None

    if pattern == "house-county":
        prefix = "house-"
        if not stem.startswith(prefix):
            return None
        county = county_from_slug(stem[len(prefix):])
        return county  # None if not a recognized county slug

    return None


def _discover_legacy(
    folder: pathlib.Path, office_slug: str, office_name: str
) -> list[tuple[str, ParserConfig]]:
    """Original single-sheet discovery for office slugs not in _DISPATCH.

    Matches `<office_slug>.xls[x]` (statewide single file) or
    `<office_slug>-<location>.xls[x]` (location = county slug or digits).
    Always builds a CongressionalConfig. Kept so adding a brand-new office
    slug without updating _DISPATCH still produces *something* parseable.
    """
    out: list[tuple[str, ParserConfig]] = []
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in _EXTS:
            continue
        config = _classify_legacy(path.stem, office_slug, office_name)
        if config is not None:
            out.append((path.name, config))
    return out


def _classify_legacy(
    stem: str, office_slug: str, office_name: str
) -> CongressionalConfig | None:
    """Original CongressionalConfig-only classifier."""
    if stem == office_slug:
        return CongressionalConfig(office=office_name)

    prefix = f"{office_slug}-"
    if not stem.startswith(prefix):
        return None
    location = stem[len(prefix):]

    county = county_from_slug(location)
    if county is not None:
        return CongressionalConfig(office=office_name, county=county)

    digits = _DIGITS.search(location)
    if digits is not None:
        return CongressionalConfig(office=office_name, district=digits.group(0))

    return CongressionalConfig(office=office_name)


def merge(
    discovered: list[tuple[str, ParserConfig]],
    explicit: list[tuple[str, ParserConfig]],
) -> list[tuple[str, ParserConfig]]:
    """Merge discovered + explicit file lists, with explicit taking precedence.

    Dedupe key is the filename. Output preserves discovered order, then any
    explicit entries that weren't in the discovered set.
    """
    explicit_by_name = {name: cfg for name, cfg in explicit}
    out: list[tuple[str, ParserConfig]] = []
    seen: set[str] = set()
    for name, cfg in discovered:
        out.append((name, explicit_by_name.get(name, cfg)))
        seen.add(name)
    for name, cfg in explicit:
        if name not in seen:
            out.append((name, cfg))
    return out

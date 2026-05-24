"""Auto-discover raw workbook files under a Job's folder by naming convention.

Convention (relative to a Job's `folder`):

    <office_slug>.xls            single statewide file (e.g. us-senate.xlsx)
    <office_slug>-<location>.xls one-of-many files, location encodes county or district
                                 (e.g. president-belknap.xls, congressional-1.xlsx)

The location segment is interpreted as:
- a county slug if it matches one of the 10 NH counties (Belknap, Carroll, ...)
- otherwise a district identifier (the leading digits are extracted, so
  `cd-1`, `1`, `district-1`, `cd1` all yield district='1')
"""

from __future__ import annotations

import pathlib
import re

from oe_nh.mappings.counties import county_from_slug
from oe_nh.parser import CongressionalConfig, ParserConfig


_EXTS = (".xls", ".xlsx")
_DIGITS = re.compile(r"\d+")


def discover_files(
    folder: pathlib.Path,
    office_slug: str,
    office_name: str,
) -> list[tuple[str, ParserConfig]]:
    """Return [(filename, ParserConfig)] for every file in `folder` matching the convention.

    Filenames are leaf names (not paths); callers compose with `folder` to get
    a real path. Filenames are sorted for stable output. Auto-discovery only
    produces `CongressionalConfig`s — multi-sheet shapes (StatewideByCountyConfig
    etc.) must be set explicitly in the Job's `files=`.
    """
    if not folder.is_dir():
        return []

    out: list[tuple[str, ParserConfig]] = []
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in _EXTS:
            continue
        config = _classify(path.stem, office_slug, office_name)
        if config is not None:
            out.append((path.name, config))
    return out


def _classify(stem: str, office_slug: str, office_name: str) -> CongressionalConfig | None:
    """Turn a filename stem into a CongressionalConfig if it matches the convention."""
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

    # Unknown location segment — emit with an empty county/district. The
    # parser will still produce rows; the user can override via Job.files
    # if they want a smarter mapping for this file.
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

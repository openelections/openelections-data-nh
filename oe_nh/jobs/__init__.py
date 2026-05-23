"""Job registries describe which raw files map to which ParserConfig, per (year, election, office).

Each year's module defines a `JOBS: list[Job]`. The orchestrator imports the
relevant module, finds the job matching the requested (election, office), and
runs it.

Convention over configuration: a Job points at a `folder` (e.g. `raw/2024/general`)
and the framework auto-discovers files in that folder named according to the
convention `<office_slug>.xls[x]` (single statewide file) or
`<office_slug>-<location>.xls[x]` (many files, where `<location>` is a county
slug or district number). Anything explicit in `files=[...]` is added or
overrides discovery at the same filename.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from oe_nh.parser import ParserConfig


@dataclass(frozen=True)
class Job:
    """One CSV worth of work: many raw files in, one output CSV out."""

    office_slug: str
    """CLI identifier: 'president', 'us-senate', 'governor', 'congressional'.
    Also the filename prefix the auto-discoverer looks for under `folder`."""

    office_name: str
    """Human-readable office name written into every CSV row (`office` column)."""

    election: str
    """CLI identifier: 'presidential-primary', 'state-primary', 'general'."""

    date: str
    """YYYYMMDD election date; used in the output filename."""

    output_basename: str
    """Filename portion after the date, e.g. 'general__president__precinct'."""

    folder: str
    """Folder containing the raw .xls/.xlsx files, relative to repo root.
    E.g. 'raw/2024/general'. The auto-discoverer scans this folder."""

    files: list[tuple[str, ParserConfig]] = field(default_factory=list)
    """Explicit (filename-relative-to-folder, ParserConfig) pairs.

    Takes precedence over auto-discovery: if `files` lists a filename that the
    discoverer would also find, the explicit ParserConfig is used. Useful for
    one-off files that don't match the naming convention, or when the default
    config knobs need to be tweaked for a specific file.
    """

    auto_discover: bool = True
    """If True (default), scan `folder` for files matching the convention and
    add them to the file list (deduped against `files`)."""

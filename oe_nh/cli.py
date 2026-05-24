"""Orchestrator: turn an election year's raw workbooks into OpenElections CSVs.

Two subcommands:

    uv run python -m oe_nh.cli new-year 2026
        # prompts for the General election date, creates raw/2026/general/

    uv run python -m oe_nh.cli --year 2024
        # builds every office found in raw/2024/general/ and prints a summary

    uv run python -m oe_nh.cli --year 2024 --office governor
        # builds just Governor (useful for debugging a single office)

There are no year-specific Python modules. Jobs are derived at runtime from
the contents of `raw/<year>/<election>/` plus the election date stored in
`raw/<year>/.dates.json` (created automatically by `new-year` or prompted
on first build).
"""

from __future__ import annotations

import argparse
import json
import pathlib
import re
import sys
from dataclasses import dataclass
from typing import Iterable

from oe_nh.discovery import (
    discover_files,
    merge,
    office_display_name,
    registered_office_slugs,
)
from oe_nh.jobs import Job
from oe_nh.parser import NormalizedRow, parse_workbook
from oe_nh.writer import write_csv


REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
_VALID_ELECTIONS = ("presidential-primary", "state-primary", "general")
_DATES_FILENAME = ".dates.json"
_WORKBOOK_EXTS = (".xls", ".xlsx")
_DATE_INPUT_RE = re.compile(r"^(\d{4})-?(\d{2})-?(\d{2})$")


# ---------------------------------------------------------------------------
# Registry: years and election dates
# ---------------------------------------------------------------------------


def _raw_dir() -> pathlib.Path:
    return REPO_ROOT / "raw"


def _year_dir(year: int) -> pathlib.Path:
    return _raw_dir() / str(year)


def _dates_path(year: int) -> pathlib.Path:
    return _year_dir(year) / _DATES_FILENAME


def _registered_years() -> list[int]:
    """Years that have a raw/<year>/.dates.json file."""
    if not _raw_dir().is_dir():
        return []
    years: list[int] = []
    for child in _raw_dir().iterdir():
        if not child.is_dir():
            continue
        if not child.name.isdigit():
            continue
        if not (child / _DATES_FILENAME).is_file():
            continue
        years.append(int(child.name))
    return sorted(years)


def _load_dates(year: int) -> dict[str, str]:
    """Load `raw/<year>/.dates.json`. Returns {} if absent."""
    path = _dates_path(year)
    if not path.exists():
        return {}
    return json.loads(path.read_text())


def _save_dates(year: int, dates: dict[str, str]) -> None:
    """Write `raw/<year>/.dates.json` (creates parent dirs as needed)."""
    path = _dates_path(year)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(dates, indent=2) + "\n")


def _normalize_date_input(raw: str) -> str | None:
    """Accept YYYY-MM-DD or YYYYMMDD; return YYYYMMDD or None if malformed."""
    match = _DATE_INPUT_RE.match(raw.strip())
    return match.group(1) + match.group(2) + match.group(3) if match else None


def _prompt_for_date(year: int, election: str) -> str:
    """Interactively prompt the user for an election date and persist it.
    Returns the YYYYMMDD string."""
    while True:
        print(
            f"\n❓  I need the date of the {year} "
            f"{election.replace('-', ' ')} election to build CSVs.",
            file=sys.stderr,
        )
        raw = input("    Enter date (YYYY-MM-DD): ").strip()
        date = _normalize_date_input(raw)
        if date is None:
            print(
                f"    \"{raw}\" doesn't look like a date. Try again, e.g. 2026-11-03.",
                file=sys.stderr,
            )
            continue
        dates = _load_dates(year)
        dates[election] = date
        _save_dates(year, dates)
        print(
            f"    ✅  Saved to raw/{year}/{_DATES_FILENAME} — won't ask again.\n",
            file=sys.stderr,
        )
        return date


# ---------------------------------------------------------------------------
# Job derivation: raw/<year>/<election>/ + dates -> list[Job]
# ---------------------------------------------------------------------------


def _output_basename(election: str, office_slug: str) -> str:
    """Construct the canonical OpenElections output filename basename.

    Election in the basename is `primary` for any *-primary election,
    `general` otherwise. Office slug's hyphens become double underscores.
    """
    election_for_output = "primary" if election.endswith("-primary") else "general"
    office_part = office_slug.replace("-", "__")
    return f"{election_for_output}__{office_part}__precinct"


def _jobs_for_year(year: int, election: str, *, prompt_for_missing_date: bool) -> list[Job]:
    """Derive Jobs for one (year, election) by scanning raw/<year>/<election>/.

    For each registered office slug, builds a Job pointing at that folder.
    The Job's auto_discover finds files in the folder.

    If the election's date is not in raw/<year>/.dates.json:
    - prompt_for_missing_date=True → ask the user interactively, persist
    - prompt_for_missing_date=False → return [] (caller decides what to do)
    """
    election_dir = _year_dir(year) / election
    if not election_dir.is_dir():
        return []
    dates = _load_dates(year)
    date = dates.get(election)
    if date is None:
        if not prompt_for_missing_date:
            return []
        date = _prompt_for_date(year, election)
    folder_rel = str(election_dir.relative_to(REPO_ROOT))
    return [
        Job(
            office_slug=slug,
            office_name=office_display_name(slug) or slug,
            election=election,
            date=date,
            output_basename=_output_basename(election, slug),
            folder=folder_rel,
        )
        for slug in registered_office_slugs()
    ]


# ---------------------------------------------------------------------------
# File resolution + execution
# ---------------------------------------------------------------------------


def resolve_files(job: Job) -> list[tuple[pathlib.Path, "ParserConfig"]]:  # noqa: F821
    """Materialize the final (path, ParserConfig) list for a Job."""
    folder = REPO_ROOT / job.folder
    if job.auto_discover:
        discovered = discover_files(folder, job.office_slug, job.office_name)
    else:
        discovered = []
    merged = merge(discovered, job.files)
    return [(folder / name, cfg) for name, cfg in merged]


def _output_path(year: int, job: Job) -> pathlib.Path:
    filename = f"{job.date}__nh__{job.output_basename}.csv"
    return REPO_ROOT / str(year) / filename


@dataclass
class _JobResult:
    job: Job
    raw_files: list[tuple[pathlib.Path, object]]
    output_path: pathlib.Path | None
    rows_written: int
    status: str  # "ok" | "no-files" | "error"
    error: str | None = None


def _run_job(year: int, job: Job) -> _JobResult:
    raw_files = resolve_files(job)
    out = _output_path(year, job)
    if not raw_files:
        return _JobResult(
            job=job, raw_files=[], output_path=None, rows_written=0,
            status="no-files",
        )
    try:
        rows: list[NormalizedRow] = []
        for path, config in raw_files:
            if not path.exists():
                raise FileNotFoundError(f"Missing raw file: {path}")
            rows.extend(parse_workbook(path, config))
        count = write_csv(out, rows)
        return _JobResult(
            job=job, raw_files=raw_files, output_path=out, rows_written=count,
            status="ok",
        )
    except Exception as exc:  # noqa: BLE001 — broad on purpose; surfaced in summary
        return _JobResult(
            job=job, raw_files=raw_files, output_path=out, rows_written=0,
            status="error", error=f"{type(exc).__name__}: {exc}",
        )


# ---------------------------------------------------------------------------
# Pre-flight scan + summary report (matches building-cli-for-humans skill)
# ---------------------------------------------------------------------------


def _scan_folder(folder: pathlib.Path) -> set[str]:
    if not folder.is_dir():
        return set()
    return {
        p.name for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in _WORKBOOK_EXTS
    }


def _expected_pattern_for(job: Job) -> str:
    slug = job.office_slug
    if slug == "state-representative":
        return "house-<county>.xls[x]"
    if slug == "congressional":
        return f"{slug}-<N>.xls[x]"
    return f"{slug}.xls[x]"


def _print_preflight(year: int, election: str, jobs: list[Job]) -> None:
    folder = REPO_ROOT / jobs[0].folder if jobs else None
    if folder is None or not folder.is_dir():
        print(f"\n📁 raw folder not found for {year} {election}", file=sys.stderr)
        return

    all_files = _scan_folder(folder)
    matched_files: set[str] = set()
    print(
        f"\n📁 Scanning {folder.relative_to(REPO_ROOT)} — "
        f"found {len(all_files)} workbook file(s):",
        file=sys.stderr,
    )
    for job in jobs:
        raw = resolve_files(job)
        matched_files.update(p.name for p, _ in raw)
        if raw:
            files_str = ", ".join(sorted(p.name for p, _ in raw))
            print(f"   ✅  {job.office_name}: {files_str}", file=sys.stderr)
        else:
            print(
                f"   ⚠️   {job.office_name}: no matching files "
                f"(looked for {_expected_pattern_for(job)})",
                file=sys.stderr,
            )

    unknown = sorted(all_files - matched_files)
    for name in unknown:
        print(f"   ❓  {name}: not a known office; ignoring", file=sys.stderr)


def _print_summary(results: list[_JobResult]) -> int:
    print("\n📝 Build summary:", file=sys.stderr)
    success = 0
    no_files = 0
    errors = 0
    total_rows = 0
    for r in results:
        if r.status == "ok":
            success += 1
            total_rows += r.rows_written
            rel = r.output_path.relative_to(REPO_ROOT) if r.output_path else "?"
            print(
                f"   ✅  {r.job.office_name:<22} → {r.rows_written:>5} rows → {rel}",
                file=sys.stderr,
            )
        elif r.status == "no-files":
            no_files += 1
            print(
                f"   ⚠️   {r.job.office_name:<22} → skipped (no raw files for "
                f"{_expected_pattern_for(r.job)} in {r.job.folder})",
                file=sys.stderr,
            )
        else:
            errors += 1
            print(
                f"   ❌  {r.job.office_name:<22} → {r.error}",
                file=sys.stderr,
            )

    total = len(results)
    # Skips on offices the year doesn't have are normal; don't poison exit code.
    print(
        f"\nBuilt {success}/{total} office(s), {total_rows:,} total rows. "
        f"{no_files} skipped, {errors} error(s).",
        file=sys.stderr,
    )
    return 1 if errors else 0


# ---------------------------------------------------------------------------
# new-year subcommand: scaffold raw/<year>/general/ and persist the date
# ---------------------------------------------------------------------------


def _new_year(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="oe_nh.cli new-year",
        description="Scaffold a new election year.",
    )
    parser.add_argument(
        "year", type=int,
        help="The election year, e.g. 2026.",
    )
    parser.add_argument(
        "--general", default=None, metavar="YYYY-MM-DD",
        help="Date of the General Election. If omitted, you'll be prompted.",
    )
    args = parser.parse_args(argv)
    year = args.year

    print(f"\n🏛   Setting up NH {year}\n", file=sys.stderr)

    raw_general = _year_dir(year) / "general"
    output_dir = REPO_ROOT / str(year)
    raw_general.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"   ✅  Created {raw_general.relative_to(REPO_ROOT)}/", file=sys.stderr)
    print(f"   ✅  Created {output_dir.relative_to(REPO_ROOT)}/", file=sys.stderr)

    # Resolve the General date — flag arg, existing dates.json, or prompt.
    dates = _load_dates(year)
    if args.general is not None:
        date = _normalize_date_input(args.general)
        if date is None:
            sys.stderr.write(
                f"\n❌  --general value \"{args.general}\" is not a valid date "
                f"(use YYYY-MM-DD).\n"
            )
            return 2
        dates["general"] = date
        _save_dates(year, dates)
        print(f"   ✅  Saved General date {date} to raw/{year}/{_DATES_FILENAME}", file=sys.stderr)
    elif "general" not in dates:
        _prompt_for_date(year, "general")  # writes dates.json internally
    else:
        print(
            f"   ℹ️   General date already on file: {dates['general']} "
            f"(in raw/{year}/{_DATES_FILENAME})",
            file=sys.stderr,
        )

    print(
        f"\nNext steps:\n"
        f"  1. Download workbooks from https://www.sos.nh.gov/{year}-election-results\n"
        f"  2. Drop them into raw/{year}/general/ (any SoS-shaped filename works)\n"
        f"  3. Run: uv run python -m oe_nh.cli --year {year}\n",
        file=sys.stderr,
    )
    return 0


# ---------------------------------------------------------------------------
# build subcommand (the default — invoked when no subcommand given)
# ---------------------------------------------------------------------------


def _office_slugs_for(year: int, election: str) -> list[str]:
    """All office slugs registered globally — independent of the year. (Used
    for validating --office; whether files exist is reported in the
    pre-flight scan, not in argument validation.)"""
    return registered_office_slugs()


def _parse_build_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="oe_nh.cli",
        description="Build OpenElections CSVs for an election year.",
    )
    parser.add_argument(
        "--year", type=int, required=True,
        help="Election year (e.g. 2022, 2024). Run `new-year` first to register a new year.",
    )
    parser.add_argument(
        "--election", default="general",
        help=f"Election type. One of: {', '.join(_VALID_ELECTIONS)}. Default: general.",
    )
    parser.add_argument(
        "--office", default=None,
        help="Office slug to filter to (e.g. governor). Omit to build all offices.",
    )
    return parser.parse_args(argv)


def _validate_build(args: argparse.Namespace) -> tuple[int, str, str | None]:
    years = _registered_years()
    if args.year not in years:
        sys.stderr.write(
            f"\n❌  Year {args.year} is not registered.\n"
            f"    Available years: {', '.join(str(y) for y in years) or '(none)'}.\n"
            f"    To add a new year, run: "
            f"uv run python -m oe_nh.cli new-year {args.year}\n\n"
        )
        sys.exit(2)

    if args.election not in _VALID_ELECTIONS:
        sys.stderr.write(
            f"\n❌  Election \"{args.election}\" is not recognized.\n"
            f"    Valid values: {', '.join(_VALID_ELECTIONS)}.\n\n"
        )
        sys.exit(2)

    if args.office is not None:
        slugs = _office_slugs_for(args.year, args.election)
        if args.office not in slugs:
            sys.stderr.write(
                f"\n❌  Office \"{args.office}\" is not a registered office.\n"
                f"    Available offices: {', '.join(slugs) or '(none)'}.\n\n"
            )
            sys.exit(2)

    return args.year, args.election, args.office


def _build(argv: list[str]) -> int:
    args = _parse_build_args(argv)
    year, election, office_filter = _validate_build(args)

    all_jobs = _jobs_for_year(year, election, prompt_for_missing_date=True)
    if office_filter is not None:
        all_jobs = [j for j in all_jobs if j.office_slug == office_filter]

    if not all_jobs:
        sys.stderr.write(
            f"❌  No raw files found under raw/{year}/{election}/. "
            f"Drop your workbooks there and try again.\n"
        )
        sys.exit(2)

    # Single-office mode: keep output tight
    if office_filter is not None:
        result = _run_job(year, all_jobs[0])
        if result.status == "ok":
            print(
                f"✅  Wrote {result.rows_written:,} rows to "
                f"{result.output_path.relative_to(REPO_ROOT)}",
                file=sys.stderr,
            )
            return 0
        if result.status == "no-files":
            sys.stderr.write(
                f"⚠️   No raw files found for {result.job.office_name} "
                f"(looked for {_expected_pattern_for(result.job)} in "
                f"{result.job.folder}).\n"
            )
            return 1
        sys.stderr.write(f"❌  {result.error}\n")
        return 1

    # Multi-office mode: pre-flight + run + summary
    print(
        f"\n🏛   Building NH {year} {election.replace('-', ' ').title()} CSVs",
        file=sys.stderr,
    )
    _print_preflight(year, election, all_jobs)
    print("\n🔨 Building...", file=sys.stderr)
    results = [_run_job(year, job) for job in all_jobs]
    return _print_summary(results)


# ---------------------------------------------------------------------------
# Main: dispatch on subcommand
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if argv and argv[0] == "new-year":
        return _new_year(argv[1:])
    return _build(argv)


if __name__ == "__main__":
    sys.exit(main())

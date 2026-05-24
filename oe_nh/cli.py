"""Orchestrator: turn an election year's raw workbooks into OpenElections CSVs.

Typical use:

    uv run python -m oe_nh.cli --year 2024
        # builds every office found in raw/2024/general/ and prints a summary

    uv run python -m oe_nh.cli --year 2024 --office governor
        # builds just Governor (useful for debugging a single office)

Outputs CSVs under `<year>/<date>__nh__<output_basename>.csv`, one per office.
"""

from __future__ import annotations

import argparse
import importlib
import pathlib
import re
import sys
from dataclasses import dataclass
from typing import Iterable

from oe_nh.discovery import discover_files, merge
from oe_nh.jobs import Job
from oe_nh.parser import NormalizedRow, parse_workbook
from oe_nh.writer import write_csv


REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
_YEAR_MODULE_RE = re.compile(r"^nh_(\d{4})\.py$")
_VALID_ELECTIONS = ("presidential-primary", "state-primary", "general")


# ---------------------------------------------------------------------------
# Registry helpers
# ---------------------------------------------------------------------------


def _registered_years() -> list[int]:
    """Discover all `nh_<year>.py` modules under oe_nh/jobs/ at runtime."""
    pkg = importlib.import_module("oe_nh.jobs")
    pkg_dir = pathlib.Path(pkg.__file__).parent  # type: ignore[arg-type]
    years: list[int] = []
    for path in pkg_dir.iterdir():
        match = _YEAR_MODULE_RE.match(path.name)
        if match:
            years.append(int(match.group(1)))
    return sorted(years)


def _load_jobs(year: int) -> list[Job]:
    module = importlib.import_module(f"oe_nh.jobs.nh_{year}")
    return list(module.JOBS)


def _office_slugs_for(year: int, election: str) -> list[str]:
    return sorted({j.office_slug for j in _load_jobs(year) if j.election == election})


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
    raw_files: list[tuple[pathlib.Path, object]]  # (path, config) pairs
    output_path: pathlib.Path | None
    rows_written: int
    status: str  # "ok" | "no-files" | "error"
    error: str | None = None


def _run_job(year: int, job: Job) -> _JobResult:
    """Run one Job end-to-end. Captures success/failure into a _JobResult so
    the caller can summarize across many Jobs."""
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
# Pre-flight scan + summary report
# ---------------------------------------------------------------------------


_WORKBOOK_EXTS = (".xls", ".xlsx")


def _scan_folder(folder: pathlib.Path) -> set[str]:
    """Return the set of workbook filenames in `folder` (empty if missing)."""
    if not folder.is_dir():
        return set()
    return {
        p.name for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in _WORKBOOK_EXTS
    }


def _print_preflight(year: int, election: str, jobs: list[Job]) -> None:
    """Print one summary of what was found in raw/<year>/<election>/ vs what
    the registered Jobs expect. Helps the user see, before any parsing
    starts, whether they're missing input files."""
    folder = REPO_ROOT / jobs[0].folder if jobs else None
    if folder is None or not folder.is_dir():
        print(f"\n📁 raw folder not found for {year} {election}", file=sys.stderr)
        return

    all_files = _scan_folder(folder)
    matched_files: set[str] = set()
    print(
        f"\n📁 Scanning {folder.relative_to(REPO_ROOT)} — found {len(all_files)} workbook file(s):",
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
                f"   ⚠️   {job.office_name}: no matching files (looked for "
                f"{_expected_pattern_for(job)})",
                file=sys.stderr,
            )

    unknown = sorted(all_files - matched_files)
    for name in unknown:
        print(f"   ❓  {name}: not a known office; ignoring", file=sys.stderr)


def _expected_pattern_for(job: Job) -> str:
    """Human-readable description of what filenames would have matched."""
    slug = job.office_slug
    if slug == "state-representative":
        return "house-<county>.xls[x]"
    if slug == "congressional":
        return f"{slug}-<N>.xls[x]"
    return f"{slug}.xls[x]"


def _print_summary(results: list[_JobResult]) -> int:
    """Print the trailing summary and return an exit code (0 if all succeeded
    or were merely missing files; 1 if any Job errored)."""
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
    print(
        f"\nBuilt {success}/{total} office(s), {total_rows:,} total rows. "
        f"{no_files} skipped, {errors} error(s).",
        file=sys.stderr,
    )
    return 1 if errors else 0


# ---------------------------------------------------------------------------
# Argument parsing — validate manually so errors are friendly and fire once
# ---------------------------------------------------------------------------


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=__doc__.split("\n")[0],
    )
    parser.add_argument(
        "--year", type=int, required=True,
        help="Election year (e.g. 2022, 2024).",
    )
    parser.add_argument(
        "--election", default="general",
        help=f"Election type. One of: {', '.join(_VALID_ELECTIONS)}. Default: general.",
    )
    parser.add_argument(
        "--office", default=None,
        help="Office slug to filter to (e.g. governor). Omit to build all "
             "offices found in the year's raw folder.",
    )
    return parser.parse_args(argv)


def _validate(args: argparse.Namespace) -> tuple[int, str, str | None]:
    """Validate args, exit with friendly message on failure. Returns
    (year, election, office_or_None)."""
    years = _registered_years()
    if args.year not in years:
        sys.stderr.write(
            f"\n❌  Year {args.year} is not registered.\n"
            f"    Available years: {', '.join(str(y) for y in years) or '(none)'}.\n"
            f"    To add a new year, see the README's "
            f"\"Adding a new election year\" section.\n\n"
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
                f"\n❌  Office \"{args.office}\" is not registered for "
                f"{args.year} {args.election}.\n"
                f"    Available offices: {', '.join(slugs) or '(none)'}.\n\n"
            )
            sys.exit(2)

    return args.year, args.election, args.office


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    year, election, office_filter = _validate(args)

    all_jobs = [j for j in _load_jobs(year) if j.election == election]
    if office_filter is not None:
        all_jobs = [j for j in all_jobs if j.office_slug == office_filter]
        if len(all_jobs) > 1:
            sys.stderr.write(
                f"❌  Ambiguous: {len(all_jobs)} jobs match "
                f"election={election!r}, office={office_filter!r}.\n"
            )
            sys.exit(2)

    if not all_jobs:
        sys.stderr.write(
            f"❌  No jobs registered for {year} {election}"
            f"{f' / {office_filter}' if office_filter else ''}.\n"
        )
        sys.exit(2)

    # Single-office mode: keep output tight (no pre-flight scan, no summary)
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


if __name__ == "__main__":
    sys.exit(main())

"""Orchestrator: pick a job from a year's registry, run it, write the CSV."""

from __future__ import annotations

import argparse
import importlib
import pathlib
import re
import sys
from typing import Iterable

from oe_nh.discovery import discover_files, merge
from oe_nh.jobs import Job
from oe_nh.parser import NormalizedRow, parse_workbook
from oe_nh.writer import write_csv


REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent

_YEAR_MODULE_RE = re.compile(r"^nh_(\d{4})\.py$")


def _registered_years() -> list[int]:
    """Discover all `nh_<year>.py` modules under oe_nh/jobs/ at runtime.

    This lets `--year` and `--office` choices stay in lock-step with the
    registry without manual maintenance: drop a new nh_<year>.py file in
    place and it shows up in --help automatically.
    """
    pkg = importlib.import_module("oe_nh.jobs")
    pkg_dir = pathlib.Path(pkg.__file__).parent  # type: ignore[arg-type]
    years: list[int] = []
    for path in pkg_dir.iterdir():
        match = _YEAR_MODULE_RE.match(path.name)
        if match:
            years.append(int(match.group(1)))
    return sorted(years)


def _all_office_slugs() -> list[str]:
    """Union of office_slugs across every registered year module."""
    offices: set[str] = set()
    for year in _registered_years():
        for job in _load_jobs(year):
            offices.add(job.office_slug)
    return sorted(offices)


def _load_jobs(year: int) -> list[Job]:
    module = importlib.import_module(f"oe_nh.jobs.nh_{year}")
    return list(module.JOBS)


def _find_job(jobs: list[Job], election: str, office: str) -> Job:
    matches = [j for j in jobs if j.election == election and j.office_slug == office]
    if not matches:
        raise SystemExit(f"No job for election={election!r}, office={office!r}")
    if len(matches) > 1:
        raise SystemExit(f"Ambiguous: {len(matches)} jobs match election={election!r}, office={office!r}")
    return matches[0]


def resolve_files(job: Job) -> list[tuple[pathlib.Path, "ParserConfig"]]:  # noqa: F821
    """Materialize the final (path, ParserConfig) list for a Job.

    Combines auto-discovered files (if enabled) with explicit `job.files`,
    deduping by filename. Returns absolute paths.
    """
    folder = REPO_ROOT / job.folder
    if job.auto_discover:
        discovered = discover_files(folder, job.office_slug, job.office_name)
    else:
        discovered = []
    merged = merge(discovered, job.files)
    return [(folder / name, cfg) for name, cfg in merged]


def _parsed_rows(job: Job) -> Iterable[NormalizedRow]:
    resolved = resolve_files(job)
    if not resolved:
        raise SystemExit(
            f"No raw files found for job (folder={job.folder!r}, office_slug={job.office_slug!r}). "
            f"Check scripts/fetch-raw.md."
        )
    for path, config in resolved:
        if not path.exists():
            raise SystemExit(f"Missing raw file: {path}")
        yield from parse_workbook(path, config)


def _output_path(year: int, job: Job) -> pathlib.Path:
    filename = f"{job.date}__nh__{job.output_basename}.csv"
    return REPO_ROOT / str(year) / filename


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Parse NH election workbooks into OpenElections CSVs.")
    parser.add_argument("--year", type=int, required=True, choices=_registered_years())
    parser.add_argument("--election", required=True,
                        choices=["presidential-primary", "state-primary", "general"])
    parser.add_argument("--office", required=True, choices=_all_office_slugs())
    # us-senate is the canonical slug; the NH SoS files sometimes use 'us-senator'.
    args = parser.parse_args(argv)

    jobs = _load_jobs(args.year)
    job = _find_job(jobs, args.election, args.office)
    out = _output_path(args.year, job)
    rows = list(_parsed_rows(job))
    count = write_csv(out, rows)
    print(f"Wrote {count} rows to {out.relative_to(REPO_ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Emit OpenElections-format CSVs."""

from __future__ import annotations

import csv
import pathlib
from typing import Iterable

from oe_nh.parser import NormalizedRow


HEADERS = ["county", "precinct", "office", "district", "party", "candidate", "votes"]


def write_csv(path: pathlib.Path, rows: Iterable[NormalizedRow]) -> int:
    """Write rows to path in OpenElections CSV format. Returns the row count."""
    path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with open(path, "w", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(HEADERS)
        for row in rows:
            writer.writerow([
                row.county,
                row.precinct,
                row.office,
                row.district,
                row.party,
                row.candidate,
                row.votes,
            ])
            count += 1
    return count

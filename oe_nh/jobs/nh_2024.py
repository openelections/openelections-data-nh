"""2024 NH election jobs.

Add Job entries here as raw files arrive under raw/2024/. See
scripts/fetch-raw.md for the naming convention and download procedure.
"""

from __future__ import annotations

from oe_nh.jobs import Job
from oe_nh.parser import ParserConfig


_GENERAL_FOLDER = "raw/2024/general"


# All four 2024 General jobs follow the same shape: header on row 2,
# party embedded in candidate label, Undervotes/Overvotes/Write-Ins
# emitted as candidate rows with empty party.
#
# Pres + Gov are multi-sheet workbooks (one summary sheet + one per county).
# Congressional is a single sheet covering the whole district.
JOBS: list[Job] = [
    Job(
        office_slug="president",
        office_name="President",
        election="general",
        date="20241105",
        output_basename="general__president__precinct",
        folder=_GENERAL_FOLDER,
        files=[
            ("president.xls", ParserConfig(
                office="President",
                header_row=2,
                multi_sheet=True,
            )),
        ],
        auto_discover=False,
    ),
    Job(
        office_slug="governor",
        office_name="Governor",
        election="general",
        date="20241105",
        output_basename="general__governor__precinct",
        folder=_GENERAL_FOLDER,
        files=[
            ("governor.xls", ParserConfig(
                office="Governor",
                header_row=2,
                multi_sheet=True,
            )),
        ],
        auto_discover=False,
    ),
    Job(
        office_slug="congressional",
        office_name="Congressional",
        election="general",
        date="20241105",
        output_basename="general__congressional__precinct",
        folder=_GENERAL_FOLDER,
        files=[
            ("congressional-1.xlsx", ParserConfig(
                office="Congressional",
                district="1",
                header_row=2,
                lookup_county_from_town=True,
            )),
            ("congressional-2.xlsx", ParserConfig(
                office="Congressional",
                district="2",
                header_row=2,
                lookup_county_from_town=True,
            )),
        ],
        auto_discover=False,
    ),
]

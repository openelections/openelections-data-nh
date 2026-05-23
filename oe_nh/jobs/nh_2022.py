"""2022 NH election jobs.

2022 was a midterm: no Presidential race. US Senate (Hassan), Governor
(Sununu), and both Congressional Districts were on the ballot.

The Governor and US Senate workbooks use a multi-section-per-sheet shape:
sheet 0 combines the per-state summary with Belknap's town-level data,
and the final sheet combines Strafford + Sullivan in two stacked
sections. `multi_sheet=True` and the section scanner handle both.
"""

from __future__ import annotations

from oe_nh.jobs import Job
from oe_nh.parser import ParserConfig


_GENERAL_FOLDER = "raw/2022/general"


JOBS: list[Job] = [
    Job(
        office_slug="us-senate",
        office_name="US Senate",
        election="general",
        date="20221108",
        output_basename="general__us__senate__precinct",
        folder=_GENERAL_FOLDER,
        files=[
            ("us-senate.xls", ParserConfig(
                office="US Senate",
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
        date="20221108",
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
        date="20221108",
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

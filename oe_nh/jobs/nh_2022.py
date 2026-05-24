"""2022 NH election jobs.

2022 was a midterm: no Presidential race. US Senate (Hassan), Governor
(Sununu), Executive Council (5 districts), State Senate (24 districts),
State Representative (per-county), and both Congressional Districts were
on the ballot.

Each Job lets auto-discovery find raw files and build the right Config
based on the office_slug (see oe_nh/discovery.py). Override via `files=`
only when a specific file needs non-canonical config knobs.
"""

from __future__ import annotations

from oe_nh.jobs import Job


_GENERAL = "raw/2022/general"
_DATE = "20221108"


JOBS: list[Job] = [
    Job(
        office_slug="us-senate", office_name="US Senate",
        election="general", date=_DATE,
        output_basename="general__us__senate__precinct",
        folder=_GENERAL,
    ),
    Job(
        office_slug="governor", office_name="Governor",
        election="general", date=_DATE,
        output_basename="general__governor__precinct",
        folder=_GENERAL,
    ),
    Job(
        office_slug="congressional", office_name="Congressional",
        election="general", date=_DATE,
        output_basename="general__congressional__precinct",
        folder=_GENERAL,
    ),
    Job(
        office_slug="executive-council", office_name="Executive Council",
        election="general", date=_DATE,
        output_basename="general__executive__council__precinct",
        folder=_GENERAL,
    ),
    Job(
        office_slug="state-senate", office_name="State Senate",
        election="general", date=_DATE,
        output_basename="general__state__senate__precinct",
        folder=_GENERAL,
    ),
    Job(
        office_slug="state-representative", office_name="State Representative",
        election="general", date=_DATE,
        output_basename="general__state__representative__precinct",
        folder=_GENERAL,
    ),
]

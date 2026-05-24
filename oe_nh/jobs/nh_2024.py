"""2024 NH election jobs.

Each Job lets auto-discovery find raw files and build the right Config
based on the office_slug (see oe_nh/discovery.py). Add a new office by
copy-pasting one of these stubs with new slug/name/output_basename.
"""

from __future__ import annotations

from oe_nh.jobs import Job


_GENERAL = "raw/2024/general"
_DATE = "20241105"


JOBS: list[Job] = [
    Job(
        office_slug="president", office_name="President",
        election="general", date=_DATE,
        output_basename="general__president__precinct",
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

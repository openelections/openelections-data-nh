"""Pytest configuration. Registers Hypothesis profiles per the property-testing skill."""

from __future__ import annotations

import os

from hypothesis import settings as hypothesis_settings


hypothesis_settings.register_profile("ci", max_examples=100)
hypothesis_settings.register_profile("dev", max_examples=30)
hypothesis_settings.load_profile(os.environ.get("HYPOTHESIS_PROFILE", "ci"))

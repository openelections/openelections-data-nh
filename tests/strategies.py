"""Reusable Hypothesis strategies for synthesizing workbook-shaped data.

The synthetic workbooks aren't byte-for-byte real NH SoS files, but they cover
the structural variations the Parser needs to handle: arbitrary towns,
arbitrary candidate counts, missing cells, mixed numeric/blank values, etc.
"""

from __future__ import annotations

from dataclasses import dataclass

from hypothesis import strategies as st


# Restrict text to alphanumerics + a few separators so we don't fight Excel's
# unicode quirks. Real NH data uses ASCII + apostrophes + spaces.
_SAFE_CHAR = st.characters(
    whitelist_categories=("Lu", "Ll", "Nd"),
    whitelist_characters=" .'-",
    blacklist_characters="\x00\r\n\t",
)
# Exclude values that ParserConfig.skip_town_values defaults to, so the
# strategy doesn't generate "towns" that the parser would correctly skip.
_RESERVED_TOWN_NAMES = {"TOTALS", "Totals", "Total"}

safe_text = (
    st.text(alphabet=_SAFE_CHAR, min_size=1, max_size=20)
    .map(str.strip)
    .filter(lambda s: bool(s) and s not in _RESERVED_TOWN_NAMES)
)

party_codes = st.sampled_from(["R", "D", "LIB", "IND", "G", "A.D.", ""])
office_names = st.sampled_from(["President", "US Senate", "Governor", "Congressional District"])
vote_count = st.integers(min_value=0, max_value=999_999)


@dataclass
class SynthesizedSheet:
    """A workbook shape we can write to disk and parse back."""

    candidates: list[tuple[str, str]]   # (name, party); party "" means no party
    towns: list[str]
    votes: list[list[int]]              # votes[town_index][candidate_index]
    header_row: int                     # how many blank rows above the candidate row

    def header_label(self, candidate: str, party: str) -> str:
        return f"{candidate}, {party}" if party else candidate


@st.composite
def synthesized_sheets(draw, max_towns: int = 8, max_candidates: int = 5) -> SynthesizedSheet:
    n_candidates = draw(st.integers(min_value=1, max_value=max_candidates))
    n_towns = draw(st.integers(min_value=1, max_value=max_towns))
    header_row = draw(st.integers(min_value=0, max_value=3))

    # Sample distinct candidate names by drawing more and dedup'ing.
    name_pool = draw(st.lists(safe_text, min_size=n_candidates * 2, max_size=n_candidates * 2, unique=True))
    candidates = [
        (name_pool[i], draw(party_codes))
        for i in range(n_candidates)
    ]
    town_pool = draw(st.lists(safe_text, min_size=n_towns * 2, max_size=n_towns * 2, unique=True))
    towns = town_pool[:n_towns]
    votes = [
        [draw(vote_count) for _ in range(n_candidates)]
        for _ in range(n_towns)
    ]
    return SynthesizedSheet(
        candidates=candidates,
        towns=towns,
        votes=votes,
        header_row=header_row,
    )

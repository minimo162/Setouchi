"""Utilities for normalising section headings.

This module implements a tiny portion of the design document's
``Section Heading Normalisation`` (section 16.3).  A heading is mapped to
one of the predefined ``section_key`` values by fuzzy matching against
known aliases.  When multiple candidates exceed the threshold the
``SECTION_PRIORITY`` list resolves conflicts.
"""
from __future__ import annotations

from difflib import SequenceMatcher
from typing import Dict, Iterable, Optional

# Predefined dictionary of section keys and their heading aliases.
SECTION_ALIASES: Dict[str, Iterable[str]] = {
    "business_overview": ["business overview", "business description"],
    "risk_factors": ["risk factors", "risks"],
    "management_analysis": [
        "management analysis",
        "management's discussion and analysis",
        "management's discussion & analysis",
        "md&a",
    ],
    "sustainability": ["sustainability", "esg"],
    "r_and_d": ["research and development", "research & development", "r&d"],
    "corporate_governance": ["corporate governance", "corporate governance report"],
}

# Priority list for resolving conflicts when multiple section keys match
# with equal scores.  Earlier entries take precedence.
SECTION_PRIORITY = [
    "management_analysis",
    "business_overview",
    "risk_factors",
    "sustainability",
    "r_and_d",
    "corporate_governance",
]


def _best_alias_score(text: str, aliases: Iterable[str]) -> float:
    """Return the best fuzzy match score for ``text`` among ``aliases``."""
    text_l = text.lower()
    best = 0.0
    for alias in aliases:
        score = SequenceMatcher(None, text_l, alias.lower()).ratio()
        if score > best:
            best = score
    return best


def normalise_section_heading(
    heading: str, *, threshold: float = 0.8
) -> Optional[str]:
    """Return the ``section_key`` for ``heading`` if confidently matched.

    Parameters
    ----------
    heading:
        Raw heading string obtained from a report.
    threshold:
        Minimum similarity score required to accept a match.  Default is
        ``0.8`` as referenced in the design document.

    Returns
    -------
    str | None
        The canonical section key, or ``None`` when no alias meets the
        threshold.
    """
    if not heading:
        return None

    scores = []
    for key, aliases in SECTION_ALIASES.items():
        best = _best_alias_score(heading, aliases)
        if best >= threshold:
            scores.append((best, key))

    if not scores:
        return None

    # Determine the maximum score and gather all keys achieving it.
    max_score = max(score for score, _ in scores)
    candidates = [key for score, key in scores if score == max_score]
    if len(candidates) == 1:
        return candidates[0]

    # Resolve ties using priority order.
    for key in SECTION_PRIORITY:
        if key in candidates:
            return key

    return candidates[0]


__all__ = ["normalise_section_heading", "SECTION_ALIASES", "SECTION_PRIORITY"]

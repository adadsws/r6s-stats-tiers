"""Shared visible-tier rules for all R6 workbook outputs."""

from typing import Dict, Tuple


VISIBLE_TIER_ORDER: Tuple[str, ...] = ("S", "A", "B", "C", "D", "F")
TIER_COLORS: Dict[str, str] = {
    "S": "E74C3C",
    "A": "F39C12",
    "B": "F1C40F",
    "C": "2ECC71",
    "D": "3498DB",
    "F": "7F8C8D",
}
RAW_SCORE_TIERS = {100: "S", 85: "A", 70: "B", 55: "C", 40: "D", 20: "F", 0: "boof"}


def display_tier(raw_tier: str) -> str:
    if not isinstance(raw_tier, str) or not raw_tier.strip():
        raise ValueError("unknown raw tier: %r" % (raw_tier,))
    normalized = raw_tier.strip()
    if normalized.casefold() == "boof":
        return VISIBLE_TIER_ORDER[-1]
    visible = normalized.upper()
    if visible not in VISIBLE_TIER_ORDER:
        raise ValueError("unknown raw tier: %s" % raw_tier)
    return visible


def display_tier_for_score(score: int) -> str:
    if isinstance(score, bool) or score not in RAW_SCORE_TIERS:
        raise ValueError("unknown score: %r" % (score,))
    return display_tier(RAW_SCORE_TIERS[score])


SCORE_TO_DISPLAY_TIER = {
    score: display_tier(raw_tier) for score, raw_tier in RAW_SCORE_TIERS.items()
}

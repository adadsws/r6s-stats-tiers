"""Shared visual semantics for XLSX and PDF leaderboard reports."""

from typing import Mapping, Tuple


FONT_FAMILY = "Microsoft YaHei"

COLOURS: Mapping[str, str] = {
    "text": "202327",
    "text_strong": "1B1D20",
    "text_muted": "595959",
    "card_fill": "F4F5F7",
    "border": "CDD1D5",
    "missing_fill": "D9D9D9",
    "title_fill": "17191D",
    "sheet_title_fill": "1F4E78",
    "section_fill": "4472C4",
    "patch_header_fill": "5B9BD5",
    "note_fill": "D9EAF7",
    "link": "4472C4",
    "white": "FFFFFF",
}

MISSING_FILL = COLOURS["missing_fill"]

PATCH_DIRECTION_COLOURS: Mapping[str, str] = {
    "增强": "548235",
    "削弱": "C00000",
    "混合": "BF9000",
}

XLSX_FONT_SIZES: Mapping[str, float] = {
    "title": 16,
    "patch_title": 18,
    "band": 12,
    "patch_header": 11,
    "name": 10,
    "body": 9,
    "gadget": 8,
    "source": 8,
    "page": 7,
}

PDF_FONT_SIZES: Mapping[str, float] = {
    "title": 19,
    "band": 11,
    "patch_header": 12,
    "name": 10,
    "body": 9,
    "patch_table": 8,
    "gadget": 6.5,
    "source": 7,
    "page": 7,
}

GADGETS_PER_LINE = 4
PDF_CARD_BODY_ROW_MM = 6
PDF_GADGET_ICON_BOX_MM = PDF_CARD_BODY_ROW_MM
XLSX_CARD_BODY_ROW_PT = 17
XLSX_GADGET_TOKEN_PX: Tuple[int, int] = (24, 22)
XLSX_GADGET_ICON_BOX_PX = 22
XLSX_GADGET_COLUMN_OFFSET_PX = 14


def feature_text(label: str, present: bool) -> str:
    """Return a shared checkmark/dash label for a boolean card feature."""
    return "%s %s" % (label, "✓" if present else "-")


def rpm_text(prefix: str, values: Tuple[int, ...]) -> str:
    """Return a shared RPM label, using a dash for an empty weapon group."""
    payload = "/".join(str(value) for value in values) or "-"
    return "%s %s" % (prefix, payload)


def is_missing_field(text: str) -> bool:
    """Recognize card fields whose complete semantic value is missing."""
    normalized = str(text).strip()
    return normalized == "-" or normalized.endswith(" -")

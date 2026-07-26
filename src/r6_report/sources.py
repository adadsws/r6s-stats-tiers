"""Validated source metadata shared by data collection and workbooks."""

import json
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Mapping, Tuple
from urllib.parse import urlparse


EXPECTED_SCORE_MAP = {
    "S": 100,
    "A": 85,
    "B": 70,
    "C": 55,
    "D": 40,
    "F": 20,
    "boof": 0,
}
PATCH_DIRECTIONS = ("增强", "削弱", "混合")


class SourceDataError(ValueError):
    """Raised when report source metadata is incomplete or inconsistent."""


@dataclass(frozen=True)
class RatingSource:
    creator: str
    title: str
    url: str
    video_id: str
    published: date
    season: str
    covered_patch: str
    covered_through: date
    coverage_basis: str
    final_frame: str
    captured_at: datetime


@dataclass(frozen=True)
class WikiManifest:
    season: str
    season_name: str
    patch: str
    fetched_at: datetime
    sources: Mapping[str, str]
    counts: Mapping[str, int]


@dataclass(frozen=True)
class PatchChange:
    direction: str
    subject: str
    detail: str


@dataclass(frozen=True)
class PatchRecord:
    patch: str
    season: str
    season_name: str
    released: date
    wiki_url: str
    official_url: str
    changes: Tuple[PatchChange, ...]


@dataclass(frozen=True)
class ReportSources:
    rating: RatingSource
    wiki: WikiManifest
    patches: Tuple[PatchRecord, ...]
    patch_index_url: str


def parse_iso_date(value: object, label: str) -> date:
    text = _required_string(value, label)
    try:
        return date.fromisoformat(text)
    except ValueError as error:
        raise SourceDataError("%s must be an ISO date: %s" % (label, text)) from error


def parse_iso_datetime(value: object, label: str) -> datetime:
    text = _required_string(value, label)
    normalized = text[:-1] + "+00:00" if text.endswith("Z") else text
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError as error:
        raise SourceDataError(
            "%s must be an ISO datetime: %s" % (label, text)
        ) from error
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        raise SourceDataError("%s must include a timezone: %s" % (label, text))
    return parsed


def parse_rating_document(
    document: object,
) -> Tuple[RatingSource, Mapping[str, Tuple[str, ...]], Mapping[str, int]]:
    root = _required_mapping(document, "rating document")
    source_data = _required_mapping(root.get("source"), "rating source")
    score_map = _required_mapping(root.get("score_map"), "rating score_map")
    if dict(score_map) != EXPECTED_SCORE_MAP:
        raise SourceDataError("rating score_map does not match the fixed mapping")

    source = RatingSource(
        creator=_required_string(source_data.get("creator"), "source.creator"),
        title=_required_string(source_data.get("title"), "source.title"),
        url=_https_url(source_data.get("url"), "source.url"),
        video_id=_required_string(source_data.get("video_id"), "source.video_id"),
        published=parse_iso_date(source_data.get("published"), "source.published"),
        season=_required_string(source_data.get("season"), "source.season"),
        covered_patch=_required_string(
            source_data.get("covered_patch"), "source.covered_patch"
        ),
        covered_through=parse_iso_date(
            source_data.get("covered_through"), "source.covered_through"
        ),
        coverage_basis=_required_string(
            source_data.get("coverage_basis"), "source.coverage_basis"
        ),
        final_frame=_required_string(
            source_data.get("final_frame"), "source.final_frame"
        ),
        captured_at=parse_iso_datetime(
            source_data.get("captured_at"), "source.captured_at"
        ),
    )
    if source.covered_through < source.published:
        raise SourceDataError("source.covered_through cannot predate publication")

    tier_data = _required_mapping(root.get("tiers"), "rating tiers")
    if set(tier_data) != set(EXPECTED_SCORE_MAP):
        raise SourceDataError("rating tiers must match the fixed score_map keys")
    tiers: Dict[str, Tuple[str, ...]] = {}
    scores: Dict[str, int] = {}
    for tier in EXPECTED_SCORE_MAP:
        values = tier_data.get(tier)
        if not isinstance(values, list):
            raise SourceDataError("rating tier %s must be a list" % tier)
        names = []
        for index, value in enumerate(values, start=1):
            name = _required_string(value, "rating tier %s item %d" % (tier, index))
            if name in scores:
                raise SourceDataError("operator appears in multiple tiers: %s" % name)
            names.append(name)
            scores[name] = EXPECTED_SCORE_MAP[tier]
        tiers[tier] = tuple(names)
    if not scores:
        raise SourceDataError("rating tiers contain no operators")
    return source, tiers, scores


def parse_wiki_manifest(document: object) -> WikiManifest:
    root = _required_mapping(document, "Wiki manifest")
    if root.get("schema_version") != 1:
        raise SourceDataError("Wiki manifest schema_version must be 1")
    source_data = _required_mapping(root.get("sources"), "Wiki manifest sources")
    sources = {
        _required_string(name, "Wiki source name"): _https_url(
            value, "Wiki source %s" % name
        )
        for name, value in source_data.items()
    }
    required_sources = {"operator", "weapon", "weapon_config"}
    if not required_sources.issubset(sources):
        missing = ", ".join(sorted(required_sources - set(sources)))
        raise SourceDataError("Wiki manifest sources missing: %s" % missing)

    count_data = _required_mapping(root.get("counts"), "Wiki manifest counts")
    counts: Dict[str, int] = {}
    for name, value in count_data.items():
        if isinstance(value, bool) or not isinstance(value, int) or value < 1:
            raise SourceDataError("Wiki count must be a positive integer: %s" % name)
        counts[str(name)] = value
    if not required_sources.issubset(counts):
        missing = ", ".join(sorted(required_sources - set(counts)))
        raise SourceDataError("Wiki manifest counts missing: %s" % missing)

    return WikiManifest(
        season=_required_string(root.get("season"), "Wiki season"),
        season_name=_required_string(root.get("season_name"), "Wiki season_name"),
        patch=_required_string(root.get("patch"), "Wiki patch"),
        fetched_at=parse_iso_datetime(root.get("fetched_at"), "Wiki fetched_at"),
        sources=sources,
        counts=counts,
    )


def parse_patch_document(
    document: object,
) -> Tuple[str, Tuple[PatchRecord, ...]]:
    root = _required_mapping(document, "patch document")
    if root.get("schema_version") != 1:
        raise SourceDataError("patch document schema_version must be 1")
    index_url = _https_url(root.get("index_url"), "patch index_url")
    parse_iso_datetime(root.get("generated_at"), "patch generated_at")
    values = root.get("patches")
    if not isinstance(values, list):
        raise SourceDataError("patches must be a list")

    records = []
    seen = set()
    for index, value in enumerate(values, start=1):
        item = _required_mapping(value, "patch item %d" % index)
        patch = _required_string(item.get("patch"), "patch item %d patch" % index)
        if patch in seen:
            raise SourceDataError("duplicate patch: %s" % patch)
        seen.add(patch)
        change_values = item.get("changes")
        if not isinstance(change_values, list):
            raise SourceDataError("patch %s changes must be a list" % patch)
        changes = []
        change_keys = set()
        for change_index, change_value in enumerate(change_values, start=1):
            change = _required_mapping(
                change_value, "patch %s change %d" % (patch, change_index)
            )
            direction = _required_string(
                change.get("direction"), "patch %s change direction" % patch
            )
            if direction not in PATCH_DIRECTIONS:
                raise SourceDataError(
                    "patch %s has unknown direction: %s" % (patch, direction)
                )
            parsed_change = PatchChange(
                direction=direction,
                subject=_required_string(
                    change.get("subject"), "patch %s change subject" % patch
                ),
                detail=_required_string(
                    change.get("detail"), "patch %s change detail" % patch
                ),
            )
            key = (parsed_change.subject, parsed_change.detail)
            if key in change_keys:
                raise SourceDataError(
                    "patch %s contains a duplicate change: %s"
                    % (patch, parsed_change.subject)
                )
            change_keys.add(key)
            changes.append(parsed_change)
        records.append(
            PatchRecord(
                patch=patch,
                season=_required_string(
                    item.get("season"), "patch %s season" % patch
                ),
                season_name=_required_string(
                    item.get("season_name"), "patch %s season_name" % patch
                ),
                released=parse_iso_date(
                    item.get("released"), "patch %s released" % patch
                ),
                wiki_url=_https_url(
                    item.get("wiki_url"), "patch %s wiki_url" % patch
                ),
                official_url=_https_url(
                    item.get("official_url"), "patch %s official_url" % patch
                ),
                changes=tuple(changes),
            )
        )
    records.sort(key=lambda record: (record.released, record.patch))
    return index_url, tuple(records)


def validate_patch_interval(
    rating: RatingSource,
    wiki: WikiManifest,
    patches: Tuple[PatchRecord, ...],
) -> None:
    upper = wiki.fetched_at.date()
    if rating.covered_through > upper:
        raise SourceDataError("rating coverage date is later than Wiki fetched_at")
    for patch in patches:
        if not rating.covered_through < patch.released <= upper:
            raise SourceDataError(
                "patch outside report interval: %s (%s)"
                % (patch.patch, patch.released.isoformat())
            )


def load_report_sources(data_dir: Path) -> ReportSources:
    root = Path(data_dir)
    rating_document = _load_json(root / "athieno" / "latest.json")
    wiki_document = _load_json(root / "wiki" / "manifest.json")
    patch_document = _load_json(root / "patches" / "patches.json")
    rating, _, _ = parse_rating_document(rating_document)
    wiki = parse_wiki_manifest(wiki_document)
    index_url, patches = parse_patch_document(patch_document)
    validate_patch_interval(rating, wiki, patches)
    return ReportSources(
        rating=rating,
        wiki=wiki,
        patches=patches,
        patch_index_url=index_url,
    )


def _load_json(path: Path) -> object:
    try:
        text = Path(path).read_text(encoding="utf-8")
    except OSError as error:
        raise SourceDataError("cannot read source file %s: %s" % (path, error)) from error
    try:
        return json.loads(text)
    except json.JSONDecodeError as error:
        raise SourceDataError("invalid JSON in source file %s: %s" % (path, error)) from error


def _required_mapping(value: object, label: str) -> Mapping[str, object]:
    if not isinstance(value, dict):
        raise SourceDataError("%s must be an object" % label)
    return value


def _required_string(value: object, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise SourceDataError("%s must be a nonempty string" % label)
    return value.strip()


def _https_url(value: object, label: str) -> str:
    text = _required_string(value, label)
    parsed = urlparse(text)
    if parsed.scheme.lower() != "https" or not parsed.netloc:
        raise SourceDataError("%s must be an HTTPS URL: %s" % (label, text))
    return text

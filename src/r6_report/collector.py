"""Collect validated Huiji data, icon, and patch snapshots under inputs/."""

import argparse
import json
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, Mapping, Optional, Sequence, Tuple

from PIL import Image as PillowImage

from . import operator_stats, tier_chart
from .patch_catalog import (
    PATCH_INDEX_URL,
    PatchCatalogError,
    parse_patch_index_html,
    parse_patch_wikitext,
    select_patch_interval,
)
from .sources import (
    ReportSources,
    SourceDataError,
    WikiManifest,
    parse_patch_document,
    parse_rating_document,
    parse_wiki_manifest,
    validate_patch_interval,
)
from .wiki_client import HuijiClient, WikiClientError


TABLE_SPECS = (
    ("operator", "Data:Operator.tabx", operator_stats.OPERATOR_FIELDS),
    ("weapon", "Data:WeaponData.tabx", operator_stats.WEAPON_FIELDS),
    ("weapon_config", "Data:WeaponConfig.tabx", operator_stats.CONFIG_FIELDS),
)


class CollectionError(RuntimeError):
    """Raised when a complete new snapshot cannot be produced."""


def collect_snapshot(
    *,
    data_dir: Path,
    archive_dir: Path,
    temp_dir: Path,
    now: datetime,
    client,
) -> WikiManifest:
    if now.tzinfo is None or now.utcoffset() is None:
        raise CollectionError("collection time must include a timezone")

    data_root = Path(data_dir)
    archive_root = Path(archive_dir)
    temp_root = Path(temp_dir)
    try:
        rating_document = _read_json(data_root / "athieno" / "latest.json")
        rating_source, _, rating_scores = parse_rating_document(rating_document)

        tables: Dict[str, Sequence[Mapping[str, object]]] = {}
        for key, title, fields in TABLE_SPECS:
            tables[key] = client.fetch_tabx(title, fields)
        rows = operator_stats.build_operator_rows(
            tables["operator"],
            tables["weapon"],
            tables["weapon_config"],
        )
        operator_names = {
            row.name
            for side in operator_stats.SIDES
            for row in rows[side]
        }
        operator_keys = {
            operator_stats.operator_key(name)
            for name in operator_names
        }
        if set(rating_scores) != operator_keys:
            missing = sorted(operator_keys - set(rating_scores))
            extra = sorted(set(rating_scores) - operator_keys)
            raise CollectionError(
                "rating and Wiki operators differ; missing=%s extra=%s"
                % (", ".join(missing) or "-", ", ".join(extra) or "-")
            )

        index_html = client.fetch_parsed_html("更新补丁总表")
        index_entries = parse_patch_index_html(index_html)
        current_entries = [
            entry for entry in index_entries if entry.released <= now.date()
        ]
        if not current_entries:
            raise CollectionError("patch index has no release on or before collection time")
        current_patch = current_entries[-1]
        interval_entries = select_patch_interval(
            index_entries,
            rating_source.covered_through,
            now.date(),
        )
        patches = tuple(
            parse_patch_wikitext(
                entry,
                client.fetch_wikitext(entry.wiki_title),
                operator_names,
            )
            for entry in interval_entries
        )

        stage_root = temp_root / ("r6-report-" + uuid.uuid4().hex)
        wiki_stage = stage_root / "wiki"
        icon_stage = stage_root / "icons"
        patch_stage = stage_root / "patches"
        wiki_stage.mkdir(parents=True)
        patch_stage.mkdir(parents=True)

        for key, title, fields in TABLE_SPECS:
            _write_json(
                wiki_stage / ("%s.json" % key),
                {
                    "schema_version": 1,
                    "title": title,
                    "fields": list(fields),
                    "records": list(tables[key]),
                },
            )
        wiki_document = {
            "schema_version": 1,
            "season": current_patch.season,
            "season_name": current_patch.season_name,
            "patch": current_patch.patch,
            "fetched_at": now.isoformat(),
            "sources": {
                key: "https://r6s.huijiwiki.com/wiki/%s" % title
                for key, title, _fields in TABLE_SPECS
            },
            "counts": {
                key: len(tables[key])
                for key, _title, _fields in TABLE_SPECS
            },
        }
        _write_json(wiki_stage / "manifest.json", wiki_document)

        patch_document = {
            "schema_version": 1,
            "index_url": PATCH_INDEX_URL,
            "generated_at": now.isoformat(),
            "patches": [_patch_to_json(patch) for patch in patches],
        }
        _write_json(patch_stage / "patches.json", patch_document)
        _write_json(
            patch_stage / "manifest.json",
            {
                "schema_version": 1,
                "index_url": PATCH_INDEX_URL,
                "generated_at": now.isoformat(),
                "count": len(patches),
                "first_patch": patches[0].patch if patches else None,
                "last_patch": patches[-1].patch if patches else None,
            },
        )

        client.prepare_operator_icons(rows, icon_stage / "operator")
        gadget_items = _gadget_items(rows)
        gadget_paths = client.prepare_gadget_icons(
            gadget_items,
            icon_stage / "gadget",
        )
        weapon_items = _weapon_items(rows)
        weapon_paths = client.prepare_weapon_icons(
            weapon_items,
            icon_stage / "weapon",
        )
        _validate_operator_icons(icon_stage / "operator", operator_names)
        _validate_gadget_icons(gadget_paths, gadget_items, icon_stage / "gadget")
        _validate_weapon_icons(weapon_paths, weapon_items, icon_stage / "weapon")

        wiki_manifest = parse_wiki_manifest(wiki_document)
        _index_url, parsed_patches = parse_patch_document(patch_document)
        validate_patch_interval(rating_source, wiki_manifest, parsed_patches)
        _activate_snapshot(
            stage_root,
            data_root,
            archive_root,
            now,
            ("wiki", "icons", "patches"),
        )
        return wiki_manifest
    except CollectionError:
        raise
    except (
        OSError,
        SourceDataError,
        PatchCatalogError,
        WikiClientError,
        operator_stats.R6StatsError,
        tier_chart.TierChartError,
    ) as error:
        raise CollectionError(str(error)) from error


def _gadget_items(
    rows: Mapping[str, Sequence[operator_stats.OperatorRow]]
) -> Tuple[tier_chart.GadgetItem, ...]:
    items = []
    for side in operator_stats.SIDES:
        for row in rows[side]:
            items.extend(
                tier_chart.parse_gadgets("\n".join(row.secondary_gadgets))
            )
    unique = {}
    for item in items:
        unique.setdefault((item.name, item.quantity), item)
    return tuple(unique.values())


def _weapon_items(
    rows: Mapping[str, Sequence[operator_stats.OperatorRow]]
) -> Tuple[tier_chart.WeaponItem, ...]:
    items = []
    for side in operator_stats.SIDES:
        for row in rows[side]:
            for rate in row.primary_automatic + row.secondary_automatic:
                items.append(
                    tier_chart.WeaponItem(
                        rate.name,
                        tier_chart.weapon_icon_key(rate.name),
                        int(rate.firerate),
                    )
                )
            for rate in row.primary_semiautomatic + row.secondary_shotguns:
                items.append(
                    tier_chart.WeaponItem(
                        rate.name,
                        tier_chart.weapon_icon_key(rate.name),
                    )
                )
    unique = {}
    for item in items:
        unique.setdefault(item.icon_key, item)
    return tuple(unique.values())


def _validate_operator_icons(directory: Path, names: Iterable[str]) -> None:
    for name in names:
        filename = operator_stats.operator_key(name) + ".png"
        for kind in ("white", "badge"):
            path = directory / kind / filename
            if not _valid_image(path):
                raise CollectionError(
                    "missing or invalid operator %s: %s" % (kind, name)
                )


def _validate_gadget_icons(
    paths: Mapping[str, Path],
    items: Iterable[tier_chart.GadgetItem],
    directory: Path,
) -> None:
    root = directory.resolve()
    names = tuple(dict.fromkeys(item.name for item in items))
    for name in names:
        path = paths.get(name)
        if path is None:
            raise CollectionError("missing gadget icon: %s" % name)
        resolved = Path(path).resolve()
        if root != resolved.parent and root not in resolved.parents:
            raise CollectionError("gadget icon is outside snapshot: %s" % name)
        if not _valid_image(resolved):
            raise CollectionError("missing or invalid gadget icon: %s" % name)


def _validate_weapon_icons(
    paths: Mapping[str, Path],
    items: Iterable[tier_chart.WeaponItem],
    directory: Path,
) -> None:
    root = directory.resolve()
    for item in items:
        path = paths.get(item.icon_key)
        if path is None:
            raise CollectionError("missing weapon icon: %s" % item.name)
        resolved = Path(path).resolve()
        if root != resolved.parent and root not in resolved.parents:
            raise CollectionError("weapon icon is outside snapshot: %s" % item.name)
        if not _valid_image(resolved):
            raise CollectionError("missing or invalid weapon icon: %s" % item.name)


def _valid_image(path: Path) -> bool:
    try:
        with PillowImage.open(path) as image:
            image.verify()
        return True
    except (OSError, ValueError):
        return False


def _activate_snapshot(
    stage_root: Path,
    data_root: Path,
    archive_root: Path,
    now: datetime,
    names: Sequence[str],
) -> None:
    data_root.mkdir(parents=True, exist_ok=True)
    archive = archive_root / now.strftime("%Y%m%dT%H%M%S%z")
    previous = []
    activated = []
    try:
        for name in names:
            target = data_root / name
            if target.exists():
                archive.mkdir(parents=True, exist_ok=True)
                archived = archive / name
                target.replace(archived)
                previous.append((target, archived))
        for name in names:
            source = stage_root / name
            target = data_root / name
            source.replace(target)
            activated.append((source, target))
    except OSError as error:
        for source, target in reversed(activated):
            if target.exists() and not source.exists():
                target.replace(source)
        for target, archived in reversed(previous):
            if archived.exists() and not target.exists():
                archived.replace(target)
        raise CollectionError("failed to activate snapshot: %s" % error) from error


def _patch_to_json(patch) -> Mapping[str, object]:
    return {
        "patch": patch.patch,
        "season": patch.season,
        "season_name": patch.season_name,
        "released": patch.released.isoformat(),
        "wiki_url": patch.wiki_url,
        "official_url": patch.official_url,
        "changes": [
            {
                "direction": change.direction,
                "subject": change.subject,
                "detail": change.detail,
            }
            for change in patch.changes
        ],
    }


def _read_json(path: Path) -> object:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except OSError as error:
        raise CollectionError("cannot read %s: %s" % (path, error)) from error
    except json.JSONDecodeError as error:
        raise CollectionError("invalid JSON in %s: %s" % (path, error)) from error


def _write_json(path: Path, document: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(document, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Collect current R6 Wiki data, icons, and patch snapshots."
    )
    parser.add_argument("--inputs-dir", type=Path, default=Path("inputs"))
    parser.add_argument(
        "--archive-dir",
        type=Path,
        default=Path("~archive") / "data-snapshots",
    )
    parser.add_argument("--temp-dir", type=Path, default=Path("~temp"))
    arguments = parser.parse_args(argv)
    try:
        manifest = collect_snapshot(
            data_dir=arguments.inputs_dir,
            archive_dir=arguments.archive_dir,
            temp_dir=arguments.temp_dir,
            now=datetime.now(timezone.utc).astimezone(),
            client=HuijiClient(),
        )
    except CollectionError as error:
        print("错误：%s" % error, file=sys.stderr)
        return 1
    print(
        "数据快照：%s %s · %s"
        % (manifest.season, manifest.patch, manifest.fetched_at.isoformat())
    )
    print("输入目录：%s" % arguments.inputs_dir.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

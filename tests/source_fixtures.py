from datetime import date, datetime, timezone

from r6_report.sources import (
    PatchChange,
    PatchRecord,
    RatingSource,
    ReportSources,
    WikiManifest,
)


def make_report_sources(with_changes=False):
    return ReportSources(
        rating=RatingSource(
            creator="Athieno",
            title="OFFICIAL Operator Tier List in Y11S2",
            url="https://youtu.be/fAjTjhNdJe4",
            video_id="fAjTjhNdJe4",
            published=date(2026, 6, 2),
            season="Y11S2",
            covered_patch="Y11S2",
            covered_through=date(2026, 6, 2),
            coverage_basis="视频最终榜单画面",
            final_frame="33:44",
            captured_at=datetime(2026, 7, 25, 8, 0, tzinfo=timezone.utc),
        ),
        wiki=WikiManifest(
            season="Y11S2",
            season_name="系统覆盖行动",
            patch="Y11S2.2",
            fetched_at=datetime(2026, 7, 25, 9, 0, tzinfo=timezone.utc),
            sources={
                "operator": "https://r6s.huijiwiki.com/wiki/Data:Operator.tabx",
                "weapon": "https://r6s.huijiwiki.com/wiki/Data:WeaponData.tabx",
                "weapon_config": "https://r6s.huijiwiki.com/wiki/Data:WeaponConfig.tabx",
            },
            counts={"operator": 2, "weapon": 3, "weapon_config": 3},
        ),
        patches=(
            PatchRecord(
                patch="Y11S2.1",
                season="Y11S2",
                season_name="系统覆盖行动",
                released=date(2026, 6, 23),
                wiki_url="https://r6s.huijiwiki.com/wiki/Y11S2.1更新补丁",
                official_url="https://www.ubisoft.com/y11s21",
                changes=(
                    (PatchChange("增强", "Alice", "测试增强内容。"),)
                    if with_changes
                    else ()
                ),
            ),
            PatchRecord(
                patch="Y11S2.2",
                season="Y11S2",
                season_name="系统覆盖行动",
                released=date(2026, 7, 14),
                wiki_url="https://r6s.huijiwiki.com/wiki/Y11S2.2更新补丁",
                official_url="https://www.ubisoft.com/y11s22",
                changes=(
                    (
                        PatchChange("削弱", "Bob", "测试削弱内容。"),
                        PatchChange("混合", "Alice", "测试混合内容。"),
                    )
                    if with_changes
                    else ()
                ),
            ),
        ),
        patch_index_url="https://r6s.huijiwiki.com/wiki/更新补丁总表",
    )

import json
import tempfile
import unittest
from pathlib import Path

import _path_setup
from r6_report import sources


def rating_document():
    return {
        "source": {
            "creator": "Athieno",
            "title": "OFFICIAL Operator Tier List in Y11S2",
            "url": "https://youtu.be/fAjTjhNdJe4",
            "video_id": "fAjTjhNdJe4",
            "published": "2026-06-02",
            "season": "Y11S2",
            "covered_patch": "Y11S2",
            "covered_through": "2026-06-02",
            "coverage_basis": "明确补丁",
            "final_frame": "33:44",
            "captured_at": "2026-07-25T10:00:00+08:00",
        },
        "score_map": {
            "S": 100,
            "A": 85,
            "B": 70,
            "C": 55,
            "D": 40,
            "F": 20,
            "boof": 0,
        },
        "tiers": {
            "S": ["ace"],
            "A": ["ash"],
            "B": [],
            "C": [],
            "D": [],
            "F": [],
            "boof": ["rook"],
        },
    }


def wiki_document():
    return {
        "schema_version": 1,
        "season": "Y11S2",
        "season_name": "系统覆盖行动",
        "patch": "Y11S2.2",
        "fetched_at": "2026-07-25T11:00:00+08:00",
        "sources": {
            "operator": "https://r6s.huijiwiki.com/wiki/Data:Operator.tabx",
            "weapon": "https://r6s.huijiwiki.com/wiki/Data:WeaponData.tabx",
            "weapon_config": "https://r6s.huijiwiki.com/wiki/Data:WeaponConfig.tabx",
        },
        "counts": {"operator": 77, "weapon": 100, "weapon_config": 200},
    }


def patch_document():
    return {
        "schema_version": 1,
        "index_url": "https://r6s.huijiwiki.com/wiki/更新补丁总表",
        "generated_at": "2026-07-25T11:00:00+08:00",
        "patches": [
            {
                "patch": "Y11S2.1",
                "season": "Y11S2",
                "season_name": "系统覆盖行动",
                "released": "2026-06-23",
                "wiki_url": "https://r6s.huijiwiki.com/wiki/Y11S2.1更新补丁",
                "official_url": "https://www.ubisoft.com/y11s21",
                "changes": [
                    {
                        "direction": "增强",
                        "subject": "Sens",
                        "detail": "持续时间提高。",
                    }
                ],
            },
            {
                "patch": "Y11S2.2",
                "season": "Y11S2",
                "season_name": "系统覆盖行动",
                "released": "2026-07-14",
                "wiki_url": "https://r6s.huijiwiki.com/wiki/Y11S2.2更新补丁",
                "official_url": "https://www.ubisoft.com/y11s22",
                "changes": [],
            },
        ],
    }


class SourceContractTests(unittest.TestCase):
    def test_loads_complete_rating_source_and_scores(self):
        source, tiers, scores = sources.parse_rating_document(rating_document())

        self.assertEqual(source.covered_patch, "Y11S2")
        self.assertEqual(source.covered_through.isoformat(), "2026-06-02")
        self.assertEqual(tiers["boof"], ("rook",))
        self.assertEqual(scores, {"ace": 100, "ash": 85, "rook": 0})

    def test_rejects_naive_timestamp_duplicate_operator_and_bad_score_map(self):
        naive = rating_document()
        naive["source"]["captured_at"] = "2026-07-25T10:00:00"
        with self.assertRaisesRegex(sources.SourceDataError, "timezone"):
            sources.parse_rating_document(naive)

        duplicate = rating_document()
        duplicate["tiers"]["A"].append("ace")
        with self.assertRaisesRegex(sources.SourceDataError, "multiple tiers"):
            sources.parse_rating_document(duplicate)

        bad_scores = rating_document()
        bad_scores["score_map"]["A"] = 80
        with self.assertRaisesRegex(sources.SourceDataError, "score_map"):
            sources.parse_rating_document(bad_scores)

    def test_loads_wiki_and_patch_documents_and_validates_interval(self):
        wiki = sources.parse_wiki_manifest(wiki_document())
        index_url, patches = sources.parse_patch_document(patch_document())
        rating, _, _ = sources.parse_rating_document(rating_document())

        sources.validate_patch_interval(rating, wiki, patches)

        self.assertEqual(wiki.patch, "Y11S2.2")
        self.assertEqual(index_url, "https://r6s.huijiwiki.com/wiki/更新补丁总表")
        self.assertEqual([patch.patch for patch in patches], ["Y11S2.1", "Y11S2.2"])
        self.assertEqual(patches[0].changes[0].direction, "增强")

    def test_rejects_patch_outside_interval_duplicate_patch_and_non_https_source(self):
        rating, _, _ = sources.parse_rating_document(rating_document())
        wiki = sources.parse_wiki_manifest(wiki_document())

        outside = patch_document()
        outside["patches"][0]["released"] = "2026-06-02"
        _, patches = sources.parse_patch_document(outside)
        with self.assertRaisesRegex(sources.SourceDataError, "outside"):
            sources.validate_patch_interval(rating, wiki, patches)

        duplicate = patch_document()
        duplicate["patches"][1]["patch"] = "Y11S2.1"
        with self.assertRaisesRegex(sources.SourceDataError, "duplicate patch"):
            sources.parse_patch_document(duplicate)

        bad_url = wiki_document()
        bad_url["sources"]["operator"] = "http://example.com/operator"
        with self.assertRaisesRegex(sources.SourceDataError, "HTTPS"):
            sources.parse_wiki_manifest(bad_url)

    def test_load_report_sources_reads_standard_data_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            (data_dir / "athieno").mkdir()
            (data_dir / "wiki").mkdir()
            (data_dir / "patches").mkdir()
            (data_dir / "athieno" / "latest.json").write_text(
                json.dumps(rating_document(), ensure_ascii=False),
                encoding="utf-8",
            )
            (data_dir / "wiki" / "manifest.json").write_text(
                json.dumps(wiki_document(), ensure_ascii=False),
                encoding="utf-8",
            )
            (data_dir / "patches" / "patches.json").write_text(
                json.dumps(patch_document(), ensure_ascii=False),
                encoding="utf-8",
            )

            report = sources.load_report_sources(data_dir)

        self.assertEqual(report.rating.video_id, "fAjTjhNdJe4")
        self.assertEqual(report.wiki.patch, "Y11S2.2")
        self.assertEqual(len(report.patches), 2)


if __name__ == "__main__":
    unittest.main()

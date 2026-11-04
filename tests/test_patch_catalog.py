import unittest
from datetime import date

import _path_setup
from r6_report import patch_catalog


INDEX_HTML = """
<table class="wikitable">
  <tr><th>所属赛季</th><th>补丁版本</th><th>推送日期</th></tr>
  <tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.2更新补丁">Y11S2.2更新补丁</a></td><td>2026-07-14</td></tr>
  <tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.1更新补丁">Y11S2.1更新补丁</a></td><td>2026-06-23</td></tr>
  <tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2更新补丁">Y11S2更新补丁</a></td><td>2026-06-02</td></tr>
</table>
"""


PATCH_WIKITEXT = """
{{Infobox patch
|补丁序号=159
|y=11
|s=2
|n=2
|赛季=系统覆盖行动
|推送日期=2026/7/14
|来源=[https://www.ubisoft.com/zh-tw/game/rainbow-six/siege/news-updates/example Ubisoft]
}}

==平衡性调整==
{{干员改动
|WAMAI
|「磁力销毁系统」
最大总充能数提高至 7 个（原为 6 个）。
新增[[机动护盾]]。
[[冲击手榴弹]]被替换为[[遥控炸药]]。

|JAGER
|调整为3速度与1生命值（原为2速度与2生命值）。
{{wi|416ccarbine}}降低水平和垂直后坐力。

|DOKKAEBI
|冷却时间延长至 14 秒（原为 7 秒）。
}}
"""


class PatchCatalogTests(unittest.TestCase):
    def test_parses_patch_index_and_orders_oldest_to_newest(self):
        entries = patch_catalog.parse_patch_index_html(INDEX_HTML)

        self.assertEqual(
            [entry.patch for entry in entries],
            ["Y11S2", "Y11S2.1", "Y11S2.2"],
        )
        self.assertEqual(entries[-1].season, "Y11S2")
        self.assertEqual(
            entries[-1].wiki_url,
            "https://r6s.huijiwiki.com/wiki/Y11S2.2更新补丁",
        )

    def test_selects_strict_lower_and_inclusive_upper_date_interval(self):
        entries = patch_catalog.parse_patch_index_html(INDEX_HTML)

        selected = patch_catalog.select_patch_interval(
            entries,
            date(2026, 6, 2),
            date(2026, 7, 14),
        )

        self.assertEqual(
            [entry.patch for entry in selected],
            ["Y11S2.1", "Y11S2.2"],
        )

    def test_extracts_nested_operator_changes_and_official_source(self):
        entry = patch_catalog.parse_patch_index_html(INDEX_HTML)[-1]

        record = patch_catalog.parse_patch_wikitext(
            entry,
            PATCH_WIKITEXT,
            {"Wamai", "Jäger", "Dokkaebi"},
        )

        self.assertEqual(
            record.official_url,
            "https://www.ubisoft.com/zh-tw/game/rainbow-six/siege/news-updates/example",
        )
        self.assertEqual(
            [change.subject for change in record.changes],
            ["Wamai", "Jäger", "Dokkaebi"],
        )
        self.assertEqual(
            [change.direction for change in record.changes],
            ["增强", "混合", "混合"],
        )
        self.assertIn("机动护盾", record.changes[0].detail)
        self.assertIn("416ccarbine", record.changes[1].detail)
        self.assertNotIn("[[", record.changes[0].detail)
        self.assertNotIn("{{", record.changes[1].detail)

    def test_preserves_patch_with_no_operator_changes(self):
        entry = patch_catalog.parse_patch_index_html(INDEX_HTML)[1]
        record = patch_catalog.parse_patch_wikitext(
            entry,
            "{{Infobox patch|来源=[https://www.ubisoft.com/no-change Ubisoft]}}",
            {"Ace"},
        )

        self.assertEqual(record.changes, ())

    def test_expands_grouped_operator_argument_into_individual_changes(self):
        entry = patch_catalog.parse_patch_index_html(INDEX_HTML)[1]
        record = patch_catalog.parse_patch_wikitext(
            entry,
            """
{{Infobox patch|来源=[https://www.ubisoft.com/grouped Ubisoft]}}
{{干员改动
|SENTRY MUTE CASTLE DOC KAPKAN JAGER FROST LESION VIGIL GOYO MELUSI ARUNI THUNDERBIRD FENRIR
|防弹摄像头作用范围提高。
}}
""",
            (
                "Sentry", "Mute", "Castle", "Doc", "Kapkan", "Jäger",
                "Frost", "Lesion", "Vigil", "Goyo", "Melusi", "Aruni",
                "Thunderbird", "Fenrir",
            ),
        )

        self.assertEqual(
            [change.subject for change in record.changes],
            [
                "Sentry", "Mute", "Castle", "Doc", "Kapkan", "Jäger",
                "Frost", "Lesion", "Vigil", "Goyo", "Melusi", "Aruni",
                "Thunderbird", "Fenrir",
            ],
        )
        self.assertTrue(
            all("防弹摄像头" in change.detail for change in record.changes)
        )
        self.assertTrue(
            all(change.direction == "增强" for change in record.changes)
        )

    def test_classifies_common_chinese_buff_and_nerf_phrases(self):
        self.assertEqual(
            patch_catalog.classify_direction(
                "Y99S1", "Ace", "侦测范围提升至 22 米。"
            ),
            "增强",
        )
        self.assertEqual(
            patch_catalog.classify_direction(
                "Y99S1", "Ace", "技能冷却时间降低至 18 秒。"
            ),
            "增强",
        )
        self.assertEqual(
            patch_catalog.classify_direction(
                "Y99S1", "Ace", "降低水平和垂直后坐力。"
            ),
            "增强",
        )
        self.assertEqual(
            patch_catalog.classify_direction(
                "Y99S1", "Ace", "效果持续时间缩短至 10 秒。"
            ),
            "削弱",
        )

    def test_rejects_duplicate_index_rows_bad_dates_and_unknown_operator(self):
        duplicate = INDEX_HTML.replace(
            "</table>",
            '<tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.2更新补丁">'
            "Y11S2.2更新补丁</a></td><td>2026-07-14</td></tr></table>",
        )
        with self.assertRaisesRegex(patch_catalog.PatchCatalogError, "duplicate"):
            patch_catalog.parse_patch_index_html(duplicate)

        with self.assertRaisesRegex(patch_catalog.PatchCatalogError, "date"):
            patch_catalog.parse_patch_index_html(
                INDEX_HTML.replace("2026-07-14", "2026/07/14")
            )

        entry = patch_catalog.parse_patch_index_html(INDEX_HTML)[-1]
        with self.assertRaisesRegex(patch_catalog.PatchCatalogError, "unknown operator"):
            patch_catalog.parse_patch_wikitext(
                entry,
                "{{Infobox patch|来源=[https://www.ubisoft.com/x Ubisoft]}}"
                "{{干员改动|UNKNOWN|提高数量。}}",
                {"Ace"},
            )


if __name__ == "__main__":
    unittest.main()

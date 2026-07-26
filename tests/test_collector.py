import json
import subprocess
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

from PIL import Image

import _path_setup
from r6_report import collector
from r6_report import wiki_client


OPERATORS = [
    {
        "id": "op1",
        "name": "Ace",
        "camp": "进攻方",
        "speed": 2,
        "_index": 1,
        "props": "手雷×2;烟雾弹×2",
    },
    {
        "id": "op2",
        "name": "Lesion",
        "camp": "防守方",
        "speed": 2,
        "_index": 2,
        "props": "遥控炸药×1",
    },
]
WEAPONS = [
    {
        "id": "w1",
        "zh_model": "R4-C",
        "firerate": 860,
        "projectile": 1,
        "index": 1,
        "type": "突击步枪",
        "equipment": 1,
    },
    {
        "id": "w2",
        "zh_model": "SMG-11",
        "firerate": 1270,
        "projectile": 1,
        "index": 2,
        "type": "冲锋枪",
        "equipment": 2,
    },
]
CONFIGS = [{"id": "w1", "user": "op1"}, {"id": "w2", "user": "op2"}]

INDEX_HTML = """
<table>
<tr><th>所属赛季</th><th>补丁版本</th><th>推送日期</th></tr>
<tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.2更新补丁">Y11S2.2更新补丁</a></td><td>2026-07-14</td></tr>
<tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.1更新补丁">Y11S2.1更新补丁</a></td><td>2026-06-23</td></tr>
</table>
"""


def rating_document():
    return {
        "source": {
            "creator": "Athieno",
            "title": "Tier List",
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
            "A": ["lesion"],
            "B": [],
            "C": [],
            "D": [],
            "F": [],
            "boof": [],
        },
    }


class FakeHuijiClient:
    def __init__(self, missing_badge=False):
        self.missing_badge = missing_badge

    def fetch_tabx(self, title, fields):
        records = {
            "Data:Operator.tabx": OPERATORS,
            "Data:WeaponData.tabx": WEAPONS,
            "Data:WeaponConfig.tabx": CONFIGS,
        }[title]
        return [dict(record) for record in records]

    def fetch_parsed_html(self, title):
        self.assert_title(title, "更新补丁总表")
        return INDEX_HTML

    def fetch_wikitext(self, title):
        patch = title.replace("更新补丁", "")
        subject = "ACE" if patch == "Y11S2.1" else "LESION"
        return (
            "{{Infobox patch|来源=[https://www.ubisoft.com/%s Ubisoft]}}"
            "{{干员改动|%s|数量提高。}}" % (patch.lower(), subject)
        )

    def prepare_operator_icons(self, rows, directory):
        badge = directory / "badge"
        white = directory / "white"
        badge.mkdir(parents=True)
        white.mkdir(parents=True)
        names = [row.name.lower() for side in ("进攻方", "防守方") for row in rows[side]]
        for name in names:
            Image.new("RGBA", (16, 16), "white").save(white / (name + ".png"))
            if not (self.missing_badge and name == "lesion"):
                Image.new("RGBA", (16, 16), "red").save(badge / (name + ".png"))
        return directory

    def prepare_gadget_icons(self, items, directory):
        directory.mkdir(parents=True)
        paths = {}
        for index, item in enumerate(items):
            if item.name in paths:
                continue
            path = directory / ("gadget-%d.png" % index)
            Image.new("RGBA", (16, 16), "blue").save(path)
            paths[item.name] = path
        return paths

    @staticmethod
    def assert_title(actual, expected):
        if actual != expected:
            raise AssertionError("expected %s, got %s" % (expected, actual))


class CollectorTests(unittest.TestCase):
    def prepare_directories(self, root):
        data_dir = root / "data"
        archive_dir = root / "archive"
        temp_dir = root / "temp"
        (data_dir / "athieno").mkdir(parents=True)
        (data_dir / "athieno" / "latest.json").write_text(
            json.dumps(rating_document(), ensure_ascii=False),
            encoding="utf-8",
        )
        return data_dir, archive_dir, temp_dir

    def test_stages_complete_snapshot_before_replacing_active_data(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data_dir, archive_dir, temp_dir = self.prepare_directories(root)

            manifest = collector.collect_snapshot(
                data_dir=data_dir,
                archive_dir=archive_dir,
                temp_dir=temp_dir,
                now=datetime(2026, 7, 25, 3, 0, tzinfo=timezone.utc),
                client=FakeHuijiClient(),
            )

            self.assertEqual(manifest.patch, "Y11S2.2")
            self.assertTrue((data_dir / "wiki" / "operator.json").is_file())
            self.assertTrue((data_dir / "icons" / "operator" / "badge" / "ace.png").is_file())
            patches = json.loads(
                (data_dir / "patches" / "patches.json").read_text(encoding="utf-8")
            )
            self.assertEqual(
                [patch["patch"] for patch in patches["patches"]],
                ["Y11S2.1", "Y11S2.2"],
            )

    def test_failed_icon_validation_keeps_previous_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data_dir, archive_dir, temp_dir = self.prepare_directories(root)
            (data_dir / "wiki").mkdir()
            old_manifest = b'{"old": true}\n'
            (data_dir / "wiki" / "manifest.json").write_bytes(old_manifest)

            with self.assertRaisesRegex(collector.CollectionError, "badge"):
                collector.collect_snapshot(
                    data_dir=data_dir,
                    archive_dir=archive_dir,
                    temp_dir=temp_dir,
                    now=datetime(2026, 7, 25, 3, 0, tzinfo=timezone.utc),
                    client=FakeHuijiClient(missing_badge=True),
                )

            self.assertEqual(
                (data_dir / "wiki" / "manifest.json").read_bytes(),
                old_manifest,
            )

    def test_rejects_naive_collection_time_before_network_calls(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data_dir, archive_dir, temp_dir = self.prepare_directories(root)
            with self.assertRaisesRegex(collector.CollectionError, "timezone"):
                collector.collect_snapshot(
                    data_dir=data_dir,
                    archive_dir=archive_dir,
                    temp_dir=temp_dir,
                    now=datetime(2026, 7, 25, 3, 0),
                    client=FakeHuijiClient(),
                )


class WikiClientTests(unittest.TestCase):
    def test_retries_empty_and_invalid_json_before_success(self):
        results = iter(
            [
                SimpleNamespace(stdout=""),
                SimpleNamespace(stdout="{"),
                SimpleNamespace(stdout='{"parse":{"text":"<table></table>"}}'),
            ]
        )
        attempts = []
        run_options = []
        sleeps = []

        def run(command, **kwargs):
            attempts.append(command)
            run_options.append(kwargs)
            return next(results)

        client = wiki_client.HuijiClient(
            run_command=run,
            which=lambda name: "curl.exe",
            sleep=sleeps.append,
        )

        self.assertEqual(
            client.fetch_parsed_html("更新补丁总表"),
            "<table></table>",
        )
        self.assertEqual(len(attempts), 3)
        self.assertEqual(sleeps, [1.0, 2.0])
        self.assertTrue(
            all(options["errors"] == "replace" for options in run_options)
        )

    def test_exhausts_four_http_failures(self):
        attempts = []

        def run(command, **kwargs):
            attempts.append(command)
            raise subprocess.CalledProcessError(
                22,
                command,
                stderr="HTTP 403",
            )

        client = wiki_client.HuijiClient(
            run_command=run,
            which=lambda name: "curl.exe",
            sleep=lambda seconds: None,
        )

        with self.assertRaisesRegex(
            wiki_client.WikiClientError,
            "after 4 attempts.*HTTP 403",
        ):
            client.fetch_wikitext("Y11S2.2更新补丁")
        self.assertEqual(len(attempts), 4)


if __name__ == "__main__":
    unittest.main()

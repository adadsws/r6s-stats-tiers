import unittest
from pathlib import Path

import _path_setup


class PackageLayoutTests(unittest.TestCase):
    def test_global_agents_layout_contract(self):
        root = Path(__file__).resolve().parents[1]
        for required in (
            "AGENTS.md",
            "AGENT_CONTEXT.md",
            "README.md",
            "CHANGELOG.md",
        ):
            self.assertTrue((root / required).is_file(), required)
        for required_dir in (
            "inputs",
            "docs/plans",
            "docs/finished_plans",
            "~archive",
        ):
            self.assertTrue((root / required_dir).is_dir(), required_dir)
        for legacy in (
            "data",
            "output",
            "~archived",
            "docs/superpowers/plans",
        ):
            self.assertFalse((root / legacy).exists(), legacy)

        ignore = (root / ".gitignore").read_text(encoding="utf-8")
        for rule in (
            "/~temp/",
            "/~outputs/",
            "/.worktrees/",
            "/.venv/",
            "__pycache__/",
        ):
            self.assertIn(rule, ignore)

    def test_runtime_modules_are_importable_from_r6_report(self):
        from r6_report import leaderboards, operator_stats

        self.assertTrue(callable(operator_stats.main))
        self.assertTrue(callable(leaderboards.main))

    def test_root_has_one_pipeline_batch_file(self):
        root = Path(__file__).resolve().parents[1]
        batch_files = sorted(path.name for path in root.glob("*.bat"))
        self.assertEqual(batch_files, ["run_r6_report.bat"])
        text = (root / batch_files[0]).read_text(encoding="utf-8")
        stages = (
            "-m r6_report.collector",
            "-m r6_report.operator_stats",
            "-m r6_report.leaderboards",
        )
        positions = tuple(text.index(stage) for stage in stages)
        self.assertEqual(positions, tuple(sorted(positions)))
        self.assertIn("if errorlevel 1", text)
        self.assertIn('set "PYTHONUTF8=1"', text)
        self.assertIn('--inputs-dir "%~dp0inputs"', text)
        self.assertIn(
            '--archive-dir "%~dp0~archive\\data-snapshots"',
            text,
        )
        self.assertIn(
            '--output "%~dp0~temp\\r6_operator_stats.xlsx"',
            text,
        )
        self.assertIn('--output-dir "%~dp0~outputs"', text)

    def test_pipeline_batch_uses_crlf_and_git_preserves_it(self):
        root = Path(__file__).resolve().parents[1]
        batch = (root / "run_r6_report.bat").read_bytes()
        self.assertEqual(batch.count(b"\n"), batch.count(b"\r\n"))
        attributes = (root / ".gitattributes").read_text(encoding="utf-8")
        self.assertIn("*.bat text eol=crlf", attributes)

    def test_skill_only_updates_athieno_video_tier_data(self):
        root = Path(__file__).resolve().parents[1]
        skill_path = root / "skills" / "build-r6-operator-report" / "SKILL.md"
        old_skill_path = (
            root / ".codex" / "skills" / "build-r6-operator-report"
        )
        self.assertTrue(skill_path.is_file())
        self.assertFalse(old_skill_path.exists())
        skill = skill_path.read_text(encoding="utf-8")
        for required in (
            "Athieno",
            "YouTube",
            "final_frame",
            "inputs/athieno/latest.json",
        ):
            self.assertIn(required, skill)
        for forbidden in (
            "run_r6_report.bat",
            "Wiki",
            ".xlsx",
            ".pdf",
            "Poppler",
            "Microsoft Excel",
            "LibreOffice",
        ):
            self.assertNotIn(forbidden, skill)

    def test_output_filenames_are_chinese(self):
        from r6_report import leaderboards

        self.assertEqual(
            leaderboards.EXPECTED_OUTPUTS,
            (
                "视频评分榜.xlsx",
                "主武器射速榜.xlsx",
                "速度榜.xlsx",
                "稀有枪械榜.xlsx",
                "次要装备榜.xlsx",
            ),
        )


if __name__ == "__main__":
    unittest.main()

# Guard the repository layout and the project-specific build entrypoint.
import unittest
from pathlib import Path

import _path_setup


class PackageLayoutTests(unittest.TestCase):
    def test_at_most_one_active_plan(self):
        root = Path(__file__).resolve().parents[1]
        active_plans = [
            path
            for path in (root / "docs" / "plans").rglob("*.md")
            if path != root / "docs" / "plans" / "README.md"
        ]
        self.assertLessEqual(len(active_plans), 1, active_plans)

    def test_global_agents_layout_contract(self):
        root = Path(__file__).resolve().parents[1]
        for required in (
            "README.md",
            "README_EN.md",
            "docs/AGENTS.md",
            "docs/AGENT_CONTEXT.md",
            "docs/CHANGELOG.md",
        ):
            self.assertTrue((root / required).is_file(), required)
        for misplaced in ("AGENTS.md", "AGENT_CONTEXT.md", "CHANGELOG.md"):
            self.assertFalse((root / misplaced).exists(), misplaced)
        for required_dir in (
            "data",
            "docs/finished_plans",
            "docs/lessons",
        ):
            self.assertTrue((root / required_dir).is_dir(), required_dir)
        for legacy in (
            "inputs",
            "output",
            "~archived",
            "previews",
            "docs/previews",
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

    def test_rebuild_batch_lives_under_src(self):
        root = Path(__file__).resolve().parents[1]
        self.assertEqual(tuple(root.glob("*.bat")), ())
        batch = root / "src" / "rebuild_leaderboards.bat"
        self.assertTrue(batch.is_file())
        text = batch.read_text(encoding="utf-8")
        stages = (
            "-m r6_report.collector",
            "-m r6_report.operator_stats",
            "-m r6_report.leaderboards",
        )
        positions = tuple(text.index(stage) for stage in stages)
        self.assertEqual(positions, tuple(sorted(positions)))
        self.assertIn("if errorlevel 1", text)
        self.assertIn('set "PYTHONUTF8=1"', text)
        self.assertIn('set "PROJECT_ROOT=%%~fI"', text)
        self.assertIn('cd /d "%PROJECT_ROOT%"', text)
        self.assertIn("仅用于重新采集数据并重建榜单", text)
        self.assertIn("直接使用 docs\\outputs 中的成品", text)
        self.assertIn('--data-dir "%PROJECT_ROOT%\\data"', text)
        self.assertIn(
            '--archive-dir "%PROJECT_ROOT%\\~archive\\data-snapshots"',
            text,
        )
        self.assertIn(
            '--output "%PROJECT_ROOT%\\~temp\\r6_operator_stats.xlsx"',
            text,
        )
        self.assertIn('--output-dir "%PROJECT_ROOT%\\~outputs"', text)

    def test_pipeline_batch_uses_crlf_and_git_preserves_it(self):
        root = Path(__file__).resolve().parents[1]
        batch = (root / "src" / "rebuild_leaderboards.bat").read_bytes()
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
            "data/athieno/latest.json",
        ):
            self.assertIn(required, skill)
        for forbidden in (
            "run_r6_report.bat",
            "rebuild_leaderboards.bat",
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

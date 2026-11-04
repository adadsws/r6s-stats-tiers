import unittest

from tests import _path_setup  # noqa: F401

from r6_report import report_theme as theme


class ReportThemeTests(unittest.TestCase):
    def test_formats_feature_and_rpm_fields_consistently(self):
        self.assertEqual(theme.feature_text("副喷", True), "副喷 ✓")
        self.assertEqual(theme.feature_text("主狙", False), "主狙 -")
        self.assertEqual(theme.rpm_text("副", ()), "副 -")
        self.assertEqual(theme.rpm_text("主", (690, 650)), "主 690/650")

    def test_detects_only_semantic_missing_fields(self):
        self.assertTrue(theme.is_missing_field("-"))
        self.assertTrue(theme.is_missing_field("副喷 -"))
        self.assertTrue(theme.is_missing_field("主 -"))
        self.assertFalse(theme.is_missing_field("Y11S2.1 - Y11S2.2"))

    def test_exposes_all_patch_direction_colours(self):
        self.assertEqual(
            set(theme.PATCH_DIRECTION_COLOURS),
            {"增强", "削弱", "混合"},
        )
        self.assertEqual(theme.MISSING_FILL, "D9D9D9")


if __name__ == "__main__":
    unittest.main()

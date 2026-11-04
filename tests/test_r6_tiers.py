import unittest

import _path_setup
from r6_report import tiers


class VisibleTierTests(unittest.TestCase):
    def test_boof_normalizes_to_current_lowest_visible_letter(self):
        self.assertEqual(tiers.VISIBLE_TIER_ORDER, ("S", "A", "B", "C", "D", "F"))
        self.assertEqual(tiers.display_tier("boof"), "F")
        self.assertEqual(tiers.display_tier("BOOF"), "F")
        self.assertEqual(tiers.display_tier_for_score(0), "F")
        self.assertEqual(tiers.TIER_COLORS["F"], "7F8C8D")

    def test_rejects_unknown_raw_tier_and_score(self):
        with self.assertRaisesRegex(ValueError, "unknown raw tier"):
            tiers.display_tier("Z")
        with self.assertRaisesRegex(ValueError, "unknown score"):
            tiers.display_tier_for_score(99)


if __name__ == "__main__":
    unittest.main()

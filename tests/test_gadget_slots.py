import unittest

import _path_setup

from r6_report.gadget_slots import GadgetSlotError, arrange_gadgets
from r6_report.tier_chart import GadgetItem


class GadgetSlotTests(unittest.TestCase):
    def test_attack_gadgets_use_fixed_seven_slots(self):
        gadgets = (
            GadgetItem("烟雾弹", 2),
            GadgetItem("破片手榴弹", 2),
            GadgetItem("阔剑地雷", 2),
        )

        arranged = arrange_gadgets("进攻方", gadgets)

        self.assertEqual(
            tuple(item.name if item else None for item in arranged),
            (
                "破片手榴弹",
                None,
                None,
                None,
                "阔剑地雷",
                None,
                "烟雾弹",
            ),
        )

    def test_defense_gadgets_use_fixed_seven_slots(self):
        gadgets = (
            GadgetItem("感应警报器", 2),
            GadgetItem("遥控炸药", 1),
            GadgetItem("防弹摄像头", 1),
        )

        arranged = arrange_gadgets("防守方", gadgets)

        self.assertEqual(
            tuple(item.name if item else None for item in arranged),
            (
                "遥控炸药",
                None,
                None,
                None,
                "防弹摄像头",
                None,
                "感应警报器",
            ),
        )

    def test_unknown_gadget_is_rejected(self):
        with self.assertRaisesRegex(
            GadgetSlotError,
            "未定义固定槽位",
        ):
            arrange_gadgets(
                "进攻方",
                (GadgetItem("未知装备", 1),),
            )

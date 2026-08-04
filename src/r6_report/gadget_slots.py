"""定义次要装备图标的固定阵营槽位。"""

from typing import Iterable, Mapping, Optional, Tuple

from .tier_chart import GadgetItem


class GadgetSlotError(ValueError):
    """次要装备无法分配到固定槽位时抛出。"""


GADGET_SLOT_NAMES: Mapping[
    str,
    Tuple[Optional[str], ...],
] = {
    "进攻方": (
        "破片手榴弹",
        "闪光弹",
        "硬突破炸药",
        "爆破炸药",
        "阔剑地雷",
        "电磁脉冲式冲击弹",
        "烟雾弹",
    ),
    "防守方": (
        "遥控炸药",
        "机动护盾",
        "冲击手榴弹",
        "倒刺铁丝网",
        "防弹摄像头",
        "观测工具阻拦器",
        "感应警报器",
    ),
}


def arrange_gadgets(
    side: str,
    gadgets: Iterable[GadgetItem],
) -> Tuple[Optional[GadgetItem], ...]:
    """按阵营固定槽位排列已有的次要装备。"""
    try:
        slot_names = GADGET_SLOT_NAMES[side]
    except KeyError as error:
        raise GadgetSlotError("未知阵营：%s" % side) from error

    by_name = {item.name: item for item in gadgets}
    known_names = {name for name in slot_names if name}
    unknown = sorted(set(by_name) - known_names)
    if unknown:
        raise GadgetSlotError(
            "未定义固定槽位：%s" % "、".join(unknown)
        )
    return tuple(by_name.get(name) if name else None for name in slot_names)

# 榜单级内自身指标优先排序 Implementation Plan

**状态：** 已完成并归档。

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让五张榜单在每个左侧阶级内部先按当前榜单自身原始指标排序，再按视频评分和现有其余维度排序。

**Architecture:** 保留 `group_cards()` 的分档和重复成员逻辑，在共享排序模块增加一个只负责当前榜单原始指标的排序键，并把它放在现有跨榜维度键之前。XLSX 与 PDF 继续共享同一组已排序卡片，不在 renderer 内增加分支。

**Tech Stack:** Python 3.9+、标准库 `unittest`、openpyxl、ReportLab、现有 `r6_report.leaderboards` 生成器。

## Global Constraints

- 级内排序链固定为：当前榜单自身原始指标（降序）→ 视频评分维度 → 其余榜单维度（保持现有全局顺序）→ Wiki 来源顺序。
- 视频评分榜使用原始 `score`；主武器射速榜使用最高主武器 RPM；速度榜使用速度值。
- 稀有枪械榜和次要装备榜在同一当前分类内自身指标相同，不得借用其他当前榜分类的等级提前。
- 保留现有分档阈值、成员关系、卡片内容、版式、输出文件名及固定 `api_version = 1` 约束。
- `output/` 只存本地生成物，不纳入 Git；仓库改动必须提交。

---

### Task 1: 用 TDD 实现当前榜单自身指标优先

**Files:**
- Modify: `tests/test_r6_leaderboards.py:20-310`
- Modify: `src/r6_report/leaderboards.py:277-305`

**Interfaces:**
- Consumes: `OperatorCard.score: int`、`OperatorCard.primary_rpms: Tuple[int, ...]`、`OperatorCard.speed: int`。
- Produces: `_self_metric_sort_key(card: OperatorCard, dimension: str) -> Tuple[int, ...]`，并由 `sort_cards_for_band()` 组合完整排序键。

- [ ] **Step 1: 扩展测试构造器并写主武器 RPM 优先的失败测试**

在 `make_card()` 增加 `score=85` 关键字参数，并将 `OperatorCard(score=85)` 改为 `OperatorCard(score=score)`。随后在 `LeaderboardSortingTests` 增加：

```python
def test_primary_band_uses_raw_rpm_before_video(self):
    cards = [
        make_card(
            "Higher Video Lower RPM",
            1,
            visible_tier="S",
            primary=(860,),
        ),
        make_card(
            "Lower Video Higher RPM",
            2,
            visible_tier="A",
            primary=(900,),
        ),
    ]

    sorted_cards = lb.sort_cards_for_band(
        cards,
        "primary_rpm",
        "进攻方",
    )

    self.assertEqual(
        [card.name for card in sorted_cards],
        ["Lower Video Higher RPM", "Higher Video Lower RPM"],
    )
```

- [ ] **Step 2: 运行定向测试并确认按预期失败**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest tests.test_r6_leaderboards.LeaderboardSortingTests.test_primary_band_uses_raw_rpm_before_video -v
```

Expected: FAIL；实际顺序仍是 `Higher Video Lower RPM` 在前，证明旧实现先比较视频 Tier。

- [ ] **Step 3: 写视频榜原始分数优先的失败测试**

```python
def test_video_band_uses_raw_score_before_primary_rpm(self):
    cards = [
        make_card(
            "Lower Score Higher RPM",
            1,
            visible_tier="S",
            score=90,
            primary=(900,),
        ),
        make_card(
            "Higher Score Lower RPM",
            2,
            visible_tier="S",
            score=95,
            primary=(700,),
        ),
    ]

    sorted_cards = lb.sort_cards_for_band(cards, "video", "进攻方")

    self.assertEqual(
        [card.name for card in sorted_cards],
        ["Higher Score Lower RPM", "Lower Score Higher RPM"],
    )
```

- [ ] **Step 4: 运行第二个定向测试并确认按预期失败**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest tests.test_r6_leaderboards.LeaderboardSortingTests.test_video_band_uses_raw_score_before_primary_rpm -v
```

Expected: FAIL；实际顺序仍由主武器射速档决定。

- [ ] **Step 5: 实现最小的自身指标排序键**

在 `best_dimension_rank()` 后增加：

```python
def _self_metric_sort_key(
    card: OperatorCard,
    dimension: str,
) -> Tuple[int, ...]:
    if dimension == "video":
        return (-card.score,)
    if dimension == "primary_rpm":
        return (-max(card.primary_rpms, default=-1),)
    if dimension == "speed":
        return (-card.speed,)
    return ()
```

把 `sort_cards_for_band()` 内部 `sort_key()` 改为先拼接自身指标键：

```python
def sort_key(card: OperatorCard):
    return _self_metric_sort_key(card, current_dimension) + tuple(
        best_dimension_rank(card, dimension, side)
        for dimension in dimensions
    ) + (card.source_order,)
```

稀有枪械与次要装备返回空元组，确保当前分类内不会按其他当前榜分类成员关系重排。

- [ ] **Step 6: 运行两个定向测试并确认通过**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest `
  tests.test_r6_leaderboards.LeaderboardSortingTests.test_primary_band_uses_raw_rpm_before_video `
  tests.test_r6_leaderboards.LeaderboardSortingTests.test_video_band_uses_raw_score_before_primary_rpm `
  -v
```

Expected: 2 tests，全部 PASS。

- [ ] **Step 7: 运行完整榜单测试，修正与新规则冲突的旧断言**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest tests.test_r6_leaderboards -v
```

Expected: 全部 PASS。现有 `test_primary_band_uses_video_then_speed` 的三张卡 RPM 都是 `900`，继续证明自身指标相同时视频 Tier 优先；现有 `test_video_band_uses_primary_then_speed_before_later_dimensions` 的原始分数都由构造器设为 `85`，继续证明视频分数相同时主武器射速档优先。

- [ ] **Step 8: 增加重复分类榜不借用其他成员关系的保护测试**

在 `LeaderboardSortingTests` 增加以下两个测试；它们保护既有正确语义，防止以后把 `best_dimension_rank()` 误用为分类榜自身键：

```python
def test_rare_band_does_not_promote_other_rare_memberships(self):
    cards = [
        make_card(
            "Higher Video Current Membership Only",
            1,
            visible_tier="S",
            has_semiautomatic=True,
        ),
        make_card(
            "Lower Video Also Higher Membership",
            2,
            visible_tier="A",
            has_semiautomatic=True,
            has_secondary_shotgun=True,
        ),
    ]

    sorted_cards = lb.sort_cards_for_band(cards, "rare", "进攻方")

    self.assertEqual(
        [card.name for card in sorted_cards],
        [
            "Higher Video Current Membership Only",
            "Lower Video Also Higher Membership",
        ],
    )

def test_gadget_band_does_not_promote_other_gadget_memberships(self):
    cards = [
        make_card(
            "Higher Video Current Membership Only",
            1,
            visible_tier="S",
            gadgets=("闪光弹",),
        ),
        make_card(
            "Lower Video Also Higher Membership",
            2,
            visible_tier="A",
            gadgets=("破片手榴弹", "闪光弹"),
        ),
    ]

    sorted_cards = lb.sort_cards_for_band(cards, "gadget", "进攻方")

    self.assertEqual(
        [card.name for card in sorted_cards],
        [
            "Higher Video Current Membership Only",
            "Lower Video Also Higher Membership",
        ],
    )
```

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest `
  tests.test_r6_leaderboards.LeaderboardSortingTests.test_rare_band_does_not_promote_other_rare_memberships `
  tests.test_r6_leaderboards.LeaderboardSortingTests.test_gadget_band_does_not_promote_other_gadget_memberships `
  -v
```

Expected: 2 tests，全部 PASS。

- [ ] **Step 9: 提交排序实现和回归测试**

```powershell
git add -- tests/test_r6_leaderboards.py src/r6_report/leaderboards.py
git commit -m "fix: prioritize leaderboard self metrics"
```

### Task 2: 记录变更、完整验证并重新生成五榜

**Files:**
- Modify: `CHANGELOG.md`
- Generate locally: `output/*.xlsx`
- Generate locally: `output/*.pdf`
- Generate locally: `output/图片版/*/*.png`

**Interfaces:**
- Consumes: Task 1 修改后的 `sort_cards_for_band()` 和现有 `python -m r6_report.leaderboards` CLI。
- Produces: 五份 XLSX、五份 PDF、每份三页 PNG；生成物只保存在被忽略的 `output/`。

- [ ] **Step 1: 在 Changelog 记录排序语义变化**

在 `CHANGELOG.md` 的 `Unreleased > Changed` 增加：

```markdown
- 五张榜单的每个阶级内部改为优先按当前榜单自身原始指标降序排列；主武器射速榜先比较最高主武器 RPM，RPM 相同后再比较视频评分。
```

- [ ] **Step 2: 运行完整测试套件**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest discover -s tests -v
```

Expected: 全部测试 PASS，退出代码 0，且无 traceback。

- [ ] **Step 3: 使用现有数据快照重新生成全部榜单**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m r6_report.leaderboards `
  --data-dir data `
  --input data/r6_operator_stats.xlsx `
  --output-dir output
```

Expected: 退出代码 0，并逐项输出五份 XLSX、五份 PDF 和对应 PNG 页面路径。

- [ ] **Step 4: 验证生成物数量和主武器射速榜工作簿顺序**

先验证文件数量：

```powershell
(Get-ChildItem -LiteralPath output -Filter *.xlsx).Count
(Get-ChildItem -LiteralPath output -Filter *.pdf).Count
(Get-ChildItem -LiteralPath 'output/图片版' -Recurse -Filter *.png).Count
```

Expected: 依次为 `5`、`5`、`15`。随后使用真实数据卡片验证两方每个主武器射速档的排序键：

```powershell
@'
from pathlib import Path

from r6_report.leaderboards import VIDEO_BANDS, group_cards
from r6_report.tier_chart import load_operator_cards

cards = load_operator_cards(Path("data/r6_operator_stats.xlsx"))
for side, side_cards in cards.items():
    groups = group_cards(side_cards, "primary_rpm", side)
    for band, band_cards in groups.items():
        actual = [
            (
                -(max(card.primary_rpms) if card.primary_rpms else -1),
                VIDEO_BANDS.index(card.tier),
            )
            for card in band_cards
        ]
        assert actual == sorted(actual), (side, band, actual)
print("主武器射速榜两方各档排序验证通过")
'@ | python -
```

Expected: 输出 `主武器射速榜两方各档排序验证通过`，退出代码 0。

- [ ] **Step 5: 检查差异与 Git 边界**

Run:

```powershell
git diff --check
git status --short
```

Expected: 只出现计划内源码、测试、Changelog 和执行中的计划文件；`output/` 不出现在 Git 状态中；现有无关 `~archived/output/` 保持未跟踪且不纳入提交。

- [ ] **Step 6: 提交 Changelog**

```powershell
git add -- CHANGELOG.md
git commit -m "docs: record leaderboard sorting change"
```

- [ ] **Step 7: 完成后归档本实施计划**

把本文件从 `docs/superpowers/plans/2026-08-04-leaderboard-self-metric-sorting.md` 移到 `~archived/superpowers-plans/2026-08-04-leaderboard-self-metric-sorting.md`，保留内容不删除，然后单独提交：

```powershell
git add -- docs/superpowers/plans/2026-08-04-leaderboard-self-metric-sorting.md ~archived/superpowers-plans/2026-08-04-leaderboard-self-metric-sorting.md
git commit -m "docs: archive leaderboard sorting plan"
```

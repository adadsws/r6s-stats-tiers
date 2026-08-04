# R6 Primary Automatic RPM Tiers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add attacker and defender charts ranked only by each operator's highest primary automatic-weapon RPM, using fixed four-tier thresholds derived from the combined 77-operator population.

**Architecture:** Keep the existing combined primary/secondary RPM chart unchanged. Add primary-only selector, classifier, sorter, and grouper functions, then parameterize the existing worksheet renderer so both chart variants share the same card layout while using separate titles and tier functions.

**Tech Stack:** Python 3.9+, `openpyxl`, Pillow, standard-library `unittest`, Microsoft Excel COM for final compatibility and visual verification.

## Global Constraints

- Primary-only fixed thresholds are `Ⅰ >= 860`, `Ⅱ 780-859`, `Ⅲ 700-779`, and `Ⅳ < 700 or no primary automatic weapon`.
- Thresholds are calculated from the current combined 77-operator population once and are not recalculated during normal generation.
- The current combined chart keeps its fixed thresholds `1000 / 850 / 750`.
- Equal RPM values must remain in the same tier.
- Operators without a primary automatic weapon remain visible in `Ⅳ` after all numeric RPM values.
- Each tier wraps after five operator cards.
- Athieno video tiers remain compact colored `S/A/B/C/D/F` badges beside operator names.
- Existing operator card facts and gadget icons remain unchanged.
- The final workbook sheet order is `进攻方射速榜`, `防守方射速榜`, `进攻方主手射速榜`, `防守方主手射速榜`, `补丁说明`.
- The project is not a valid Git repository, so commit steps are omitted.

---

### Task 1: Primary-only RPM selection, sorting, and fixed tiers

**Files:**
- Modify: `tests/test_r6_rpm_chart.py`
- Modify: `r6_rpm_chart.py`

**Interfaces:**
- Consumes: `OperatorCard.primary_rpms`, `OperatorCard.source_order`.
- Produces: `highest_primary_automatic_rpm(card) -> Optional[int]`, `primary_rpm_tier_for_rate(rate) -> str`, `primary_rpm_tier_for_card(card) -> str`, `sort_cards_by_primary_rpm(cards) -> List[OperatorCard]`, and `group_cards_by_primary_rpm_tier(cards) -> Mapping[str, Tuple[OperatorCard, ...]]`.

- [ ] **Step 1: Write failing primary-only unit tests**

Add tests equivalent to:

```python
def test_primary_highest_ignores_faster_secondary(self):
    card = make_card("Primary Only", 1, primary=(900, 700), secondary=(1270,))
    self.assertEqual(rpm.highest_primary_automatic_rpm(card), 900)

def test_primary_fixed_tiers_cover_threshold_edges(self):
    cases = (
        (1200, "Ⅰ"), (860, "Ⅰ"), (859, "Ⅱ"), (780, "Ⅱ"),
        (779, "Ⅲ"), (700, "Ⅲ"), (699, "Ⅳ"), (None, "Ⅳ"),
    )
    for rate, expected in cases:
        with self.subTest(rate=rate):
            self.assertEqual(rpm.primary_rpm_tier_for_rate(rate), expected)

def test_primary_groups_sort_by_primary_and_keep_missing_last(self):
    cards = [
        make_card("Secondary Fast", 1, primary=(700,), secondary=(1270,)),
        make_card("Primary Fast", 2, primary=(900,), secondary=()),
        make_card("No Primary", 3, primary=(), secondary=(1200,)),
    ]
    groups = rpm.group_cards_by_primary_rpm_tier(cards)
    self.assertEqual([card.name for card in groups["Ⅱ"]], ["Primary Fast"])
    self.assertEqual(
        [card.name for card in groups["Ⅲ"]], ["Secondary Fast"]
    )
    self.assertEqual([card.name for card in groups["Ⅳ"]], ["No Primary"])
```

- [ ] **Step 2: Run focused tests and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_rpm_chart.RpmSortTests -v
```

Expected: failures report missing primary-only functions.

- [ ] **Step 3: Implement the primary-only functions**

Add constants and functions:

```python
PRIMARY_RPM_TIER_THRESHOLDS = (860, 780, 700)

def highest_primary_automatic_rpm(card):
    return max(card.primary_rpms) if card.primary_rpms else None

def primary_rpm_tier_for_rate(rate):
    validate_rpm_rate(rate)
    if rate is None:
        return "Ⅳ"
    if rate >= 860:
        return "Ⅰ"
    if rate >= 780:
        return "Ⅱ"
    if rate >= 700:
        return "Ⅲ"
    return "Ⅳ"

def primary_rpm_tier_for_card(card):
    return primary_rpm_tier_for_rate(highest_primary_automatic_rpm(card))

def sort_cards_by_primary_rpm(cards):
    return sort_cards_by_rate(cards, highest_primary_automatic_rpm)

def group_cards_by_primary_rpm_tier(cards):
    return group_cards_by_tier(
        cards,
        sort_cards_by_primary_rpm,
        primary_rpm_tier_for_card,
    )
```

Extract shared validation and internal sort/group helpers only where needed to avoid duplicating the existing combined behavior.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run the same `RpmSortTests` command. Expected: all sorting and threshold tests pass.

### Task 2: Add two primary-only chart worksheets

**Files:**
- Modify: `tests/test_r6_rpm_chart.py`
- Modify: `r6_rpm_chart.py`

**Interfaces:**
- Consumes: Task 1's `sort_cards_by_primary_rpm` and `group_cards_by_primary_rpm_tier`.
- Produces: a five-sheet output workbook with two combined charts, two primary-only charts, and patch notes.

- [ ] **Step 1: Write a failing workbook structure and rendering test**

Extend the workbook test to assert:

```python
self.assertEqual(
    workbook.sheetnames,
    [
        "进攻方射速榜",
        "防守方射速榜",
        "进攻方主手射速榜",
        "防守方主手射速榜",
        "补丁说明",
    ],
)
primary_sheet = workbook["进攻方主手射速榜"]
self.assertEqual(primary_sheet["A3"].value, "Ⅰ")
self.assertTrue(primary_sheet["C3"].value.startswith("Primary 1200\n"))
self.assertEqual(primary_sheet.freeze_panes, "B3")
self.assertEqual(
    primary_sheet.cell(primary_sheet.max_row, 1).value,
    notes.DATA_STATUS_TEXT,
)
```

Use fixture cards where a faster secondary RPM would change the combined tier but must not change the primary-only tier. Assert the combined sheet and primary sheet place that card in different bands.

- [ ] **Step 2: Run the workbook test and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_rpm_chart.RpmWorkbookTests -v
```

Expected: failure because the two primary-only worksheets do not exist.

- [ ] **Step 3: Parameterize the existing renderer**

Change `_render_rpm_side_sheet` to accept a chart title and pre-grouped tier mapping:

```python
def _render_rpm_side_sheet(
    sheet,
    title_text,
    groups,
    badge_dir,
    token_paths,
    video_tier_badges,
):
    ...
    title = sheet.cell(1, 1, title_text)
    ...
```

Keep all card, image, merge, freeze, print, and status-row behavior unchanged.

- [ ] **Step 4: Create all four chart sheets in fixed order**

In `write_rpm_workbook`, create:

```python
chart_specs = (
    ("进攻方射速榜", "进攻方 · 自动枪械射速榜", "进攻方", group_cards_by_rpm_tier),
    ("防守方射速榜", "防守方 · 自动枪械射速榜", "防守方", group_cards_by_rpm_tier),
    (
        "进攻方主手射速榜",
        "进攻方 · 主手自动枪械射速榜",
        "进攻方",
        group_cards_by_primary_rpm_tier,
    ),
    (
        "防守方主手射速榜",
        "防守方 · 主手自动枪械射速榜",
        "防守方",
        group_cards_by_primary_rpm_tier,
    ),
)
```

Render each spec, append its status row, and set its print area before adding `补丁说明`.

- [ ] **Step 5: Run workbook and CLI tests and verify GREEN**

Run:

```powershell
python -m unittest tests.test_r6_rpm_chart -v
```

Expected: all RPM chart tests pass and both new sheets contain the required layout.

### Task 3: Document and synchronize the generator

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Modify: `.codex/skills/build-r6-operator-report/SKILL.md`
- Modify: `.codex/skills/build-r6-operator-report/scripts/r6_rpm_chart.py`

**Interfaces:**
- Consumes: the final behavior from Tasks 1 and 2.
- Produces: accurate user documentation and a project Skill script identical to the root script.

- [ ] **Step 1: Update documentation**

Document:

```text
主手榜固定分界：Ⅰ >=860、Ⅱ 780-859、Ⅲ 700-779、Ⅳ <700 或无主手自动枪械。
当前全体四档人数：20 / 20 / 18 / 19。
```

State that the two original combined sheets remain unchanged and the workbook now contains four chart sheets plus `补丁说明`.

- [ ] **Step 2: Update the project Skill workflow**

Add verification requirements for exact five-sheet order, primary-only ranking, fixed thresholds, no-primary placement, and the unchanged combined chart.

- [ ] **Step 3: Synchronize the Skill script**

Copy the completed root `r6_rpm_chart.py` to:

```text
.codex/skills/build-r6-operator-report/scripts/r6_rpm_chart.py
```

- [ ] **Step 4: Verify script synchronization and Skill validity**

Run:

```powershell
Get-FileHash r6_rpm_chart.py
Get-FileHash .codex\skills\build-r6-operator-report\scripts\r6_rpm_chart.py
python <USER_HOME>\.codex\skills\.system\skill-creator\scripts\quick_validate.py .codex\skills\build-r6-operator-report
```

Expected: both SHA-256 values are identical and validation prints `Skill is valid!`.

### Task 4: Generate and verify the final workbook

**Files:**
- Modify: `r6_operator_rpm_chart.xlsx`

**Interfaces:**
- Consumes: the completed generator and current `r6_operator_stats.xlsx`.
- Produces: the final five-sheet RPM workbook.

- [ ] **Step 1: Run the full automated suite**

Run:

```powershell
python -m unittest discover -s tests -v
```

Expected: all tests pass with zero failures and zero errors.

- [ ] **Step 2: Generate the actual workbook**

Run:

```powershell
python r6_rpm_chart.py --input r6_operator_stats.xlsx --output r6_operator_rpm_chart.xlsx
```

Expected: 39 attackers, 38 defenders, and the absolute output path.

- [ ] **Step 3: Audit workbook structure and population**

Reopen the workbook read-only and assert:

```python
workbook.sheetnames == [
    "进攻方射速榜",
    "防守方射速榜",
    "进攻方主手射速榜",
    "防守方主手射速榜",
    "补丁说明",
]
```

Count all primary-only tier assignments across both sides and verify `20 / 20 / 18 / 19`. Verify all four side sheets contain the final data-status row and `补丁说明.freeze_panes is None`.

- [ ] **Step 4: Open and render with Microsoft Excel**

Open `r6_operator_rpm_chart.xlsx` read-only through Excel COM, export all five worksheets to PDF or PNG, and inspect every rendered page. Confirm:

- left tier bands are visible and aligned;
- cards wrap after five operators;
- Athieno badges sit beside operator names;
- no card text, icon, status row, or title overlaps;
- both primary-only sheets use the same visual system as the combined sheets;
- `补丁说明` is legible and unfrozen.

- [ ] **Step 5: Re-run the full suite after final output generation**

Run:

```powershell
python -m unittest discover -s tests -v
```

Expected: all tests still pass with zero failures and zero errors.

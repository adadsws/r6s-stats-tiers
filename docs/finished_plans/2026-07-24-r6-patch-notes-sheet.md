# R6 Patch Notes Sheet Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a shared, source-linked `补丁说明` worksheet to all three generated R6 workbooks and append a latest-data status line to every existing data/chart worksheet.

**Architecture:** Create `r6_patch_notes.py` as the only owner of patch metadata, source URLs, direction colours, score formatting, the patch worksheet renderer, and the shared bottom status row. The three existing generators provide their already-validated operator score mappings and call the shared rendering functions without changing their core data, tier, or RPM layout.

**Tech Stack:** Python 3.9+, openpyxl 3.x, standard-library `dataclasses` and `unittest`, Microsoft Excel COM and PDF rendering for final visual verification.

## Global Constraints

- Each final workbook has its existing two sheets followed by exactly one `补丁说明` sheet.
- Every original sheet ends with `数据状态：除 Athieno Y11S2 视频评分外，本表其他信息均按生成时的灰机 Wiki 最新数据更新。`
- The status row stays outside the statistics workbook's structured Excel table.
- Patch content covers Y11S2.1 and Y11S2.2 changes affecting operator strength, including buffs, nerfs, mixed changes, bulletproof-camera beneficiaries, Solis, and Solid Snake.
- Patch changes never modify `data/athieno_y11s2.json` or the visible Athieno scores.
- Source names and URLs remain clickable and auditable in the workbook.
- Existing operator, tier, RPM, icon, and gadget layouts remain unchanged apart from the appended status row.
- `AGENTS.md` remains ignored and must not be committed.
- Git commands currently fail because the workspace is not recognized as a repository; do not attempt commits until repository metadata is restored.

---

### Task 1: Shared Patch Metadata And Worksheet Renderer

**Files:**
- Create: `r6_patch_notes.py`
- Create: `tests/test_r6_patch_notes.py`

**Interfaces:**
- Consumes: `openpyxl.Workbook`, `openpyxl.worksheet.worksheet.Worksheet`, and `Mapping[str, int]` video scores.
- Produces: `DATA_STATUS_TEXT`, `PatchNotesError`, `required_operator_names()`, `append_data_status_row(sheet, last_column) -> int`, and `add_patch_notes_sheet(workbook, scores_by_name) -> Worksheet`.

- [ ] **Step 1: Write failing metadata and status-row tests**

Create `tests/test_r6_patch_notes.py` with a score fixture covering every name returned by the intended public interface:

```python
import unittest

from openpyxl import Workbook

import r6_patch_notes as notes


class PatchNotesTests(unittest.TestCase):
    def score_map(self):
        return {name: 70 for name in notes.required_operator_names()}

    def test_appends_merged_status_as_last_nonempty_row(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "content"

        row = notes.append_data_status_row(sheet, 9)

        self.assertEqual(row, 3)
        self.assertEqual(sheet.max_row, 3)
        self.assertEqual(sheet["A3"].value, notes.DATA_STATUS_TEXT)
        self.assertIn("A3:I3", {str(item) for item in sheet.merged_cells.ranges})
        self.assertTrue(sheet["A3"].alignment.wrap_text)

    def test_builds_patch_sheet_with_sources_directions_and_scores(self):
        workbook = Workbook()
        sheet = notes.add_patch_notes_sheet(workbook, self.score_map())

        self.assertEqual(sheet.title, "补丁说明")
        self.assertIn("除 Athieno Y11S2 视频评分外", sheet["A2"].value)
        self.assertEqual(
            [sheet.cell(11, column).value for column in range(1, 7)],
            ["方向", "补丁", "日期", "干员/对象", "视频评分", "更新内容"],
        )
        values = [cell.value for row in sheet.iter_rows() for cell in row]
        self.assertIn("Y11S2.1", values)
        self.assertIn("Y11S2.2", values)
        self.assertIn("Wamai", values)
        self.assertIn("Thorn", values)
        self.assertIn("Dokkaebi", values)
        self.assertTrue(any(cell.hyperlink for row in sheet.iter_rows() for cell in row))
```

- [ ] **Step 2: Run the new test and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_patch_notes -v
```

Expected: import failure because `r6_patch_notes.py` does not exist.

- [ ] **Step 3: Implement structured metadata and validation**

Create `r6_patch_notes.py` with immutable data classes and exact shared constants:

```python
from dataclasses import dataclass
from typing import Mapping, Tuple

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter

from r6_tiers import SCORE_TO_DISPLAY_TIER


DATA_STATUS_TEXT = (
    "数据状态：除 Athieno Y11S2 视频评分外，"
    "本表其他信息均按生成时的灰机 Wiki 最新数据更新。"
)
DIRECTION_COLOURS = {
    "增强": "E2F0D9",
    "削弱": "FCE4D6",
    "混合": "FFF2CC",
}


class PatchNotesError(ValueError):
    pass


@dataclass(frozen=True)
class PatchSource:
    label: str
    url: str


@dataclass(frozen=True)
class PatchChange:
    direction: str
    patch: str
    date: str
    subject: str
    detail: str
```

Define the five approved sources and structured changes exactly from the design specification. Expand the bulletproof-camera change into one `PatchChange` per affected operator:

```python
CAMERA_OPERATORS = (
    "Sentry", "Mute", "Castle", "Doc", "Kapkan", "Jäger", "Frost",
    "Lesion", "Vigil", "Goyo", "Melusi", "Aruni", "Thunderbird", "Fenrir",
)
```

Implement validation before rendering:

```python
def required_operator_names() -> Tuple[str, ...]:
    return tuple(sorted({change.subject for change in PATCH_CHANGES}))


def _score_text(name: str, scores: Mapping[str, int]) -> str:
    if name not in scores:
        raise PatchNotesError("missing video score for patch subject: %s" % name)
    score = scores[name]
    tier = SCORE_TO_DISPLAY_TIER.get(score)
    if tier is None:
        raise PatchNotesError("unknown video score for patch subject: %s = %s" % (name, score))
    return "%s / %d" % (tier, score)
```

Reject directions outside `DIRECTION_COLOURS`, duplicate `(patch, subject, detail)` records, non-HTTPS source URLs, and an existing `补丁说明` worksheet.

- [ ] **Step 4: Implement the bottom status row and patch sheet**

Implement:

```python
def append_data_status_row(sheet, last_column: int) -> int:
    status_row = sheet.max_row + 2
    sheet.merge_cells(
        start_row=status_row,
        start_column=1,
        end_row=status_row,
        end_column=last_column,
    )
    cell = sheet.cell(status_row, 1, DATA_STATUS_TEXT)
    cell.fill = PatternFill("solid", fgColor="E7E6E6")
    cell.font = Font(name="Microsoft YaHei", size=9, italic=True, color="595959")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[status_row].height = 30
    return status_row
```

Implement `add_patch_notes_sheet()` with rows 1-3 for title/status/rating warning, rows 5-9 for source labels and full URL hyperlinks, row 11 for the six headers, and row 12 onward for changes. Use widths `(10, 12, 14, 22, 14, 78)`, freeze `A12`, set `auto_filter.ref`, hide gridlines, use landscape fit-to-width printing, and set the print area through the final change row.

- [ ] **Step 5: Add strict failure tests**

Add tests asserting:

```python
with self.assertRaisesRegex(notes.PatchNotesError, "missing video score"):
    notes.add_patch_notes_sheet(Workbook(), {})

workbook = Workbook()
workbook.create_sheet("补丁说明")
with self.assertRaisesRegex(notes.PatchNotesError, "already exists"):
    notes.add_patch_notes_sheet(workbook, self.score_map())
```

Also assert the direction cells for representative `增强`, `削弱`, and `混合` rows use the three configured fills and that all five source URL cells have hyperlinks beginning with `https://`.

- [ ] **Step 6: Run shared-module tests and verify GREEN**

Run:

```powershell
python -m unittest tests.test_r6_patch_notes -v
```

Expected: all patch-note tests pass.

---

### Task 2: Statistics Workbook Integration

**Files:**
- Modify: `r6_operator_stats.py`
- Modify: `tests/test_r6_operator_stats.py`

**Interfaces:**
- Consumes: `append_data_status_row()` and `add_patch_notes_sheet()` from Task 1.
- Produces: `r6_operator_stats.xlsx` with sheets `进攻方`, `防守方`, `补丁说明`.

- [ ] **Step 1: Update the existing workbook test first**

Modify `WorkbookAndCliTests.test_workbook_has_two_structured_sides_and_expected_presentation`:

```python
self.assertEqual(workbook.sheetnames, ["进攻方", "防守方", "补丁说明"])
for sheet in (attackers, defenders):
    self.assertEqual(sheet.cell(sheet.max_row, 1).value, notes.DATA_STATUS_TEXT)
    self.assertEqual(next(iter(sheet.tables.values())).ref, "A1:I2")
    self.assertEqual(sheet.print_area, f"'${sheet.title}'!$A$1:$I${sheet.max_row}")
self.assertEqual(workbook["补丁说明"]["A1"].value, "Y11S2 视频评分后续补丁说明")
```

Import `r6_patch_notes as notes` in the test and provide a complete injected patch score map:

```python
patch_scores = {name: 70 for name in notes.required_operator_names()}
stats.write_workbook(output, rows, ratings, icon_dir, patch_scores=patch_scores)
```

- [ ] **Step 2: Run the statistics workbook test and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_operator_stats.WorkbookAndCliTests.test_workbook_has_two_structured_sides_and_expected_presentation -v
```

Expected: failure because `write_workbook` has no `patch_scores` parameter and creates only two sheets.

- [ ] **Step 3: Integrate shared rendering**

Add imports and extend the writer signature:

```python
from r6_patch_notes import add_patch_notes_sheet, append_data_status_row


def write_workbook(
    path: Path,
    rows: Mapping[str, Sequence[OperatorRow]],
    ratings: Mapping[str, OperatorRating],
    icon_dir: Path,
    patch_scores: Optional[Mapping[str, int]] = None,
) -> None:
```

Capture `last_data_row` before creating each structured table:

```python
last_data_row = sheet.max_row
table_ref = "A1:I%d" % last_data_row
```

Then append the shared status row and set the print area:

```python
status_row = append_data_status_row(sheet, len(HEADERS))
sheet.print_area = "A1:I%d" % status_row
```

After both side sheets are complete, derive full production scores when no test mapping is injected:

```python
scores = (
    dict(patch_scores)
    if patch_scores is not None
    else {
        row.name: int(ratings[operator_key(row.name)].score)
        for camp in SIDES
        for row in rows[camp]
    }
)
add_patch_notes_sheet(workbook, scores)
```

- [ ] **Step 4: Run statistics tests and verify GREEN**

Run:

```powershell
python -m unittest tests.test_r6_operator_stats -v
```

Expected: all statistics tests pass, including unchanged CLI counts.

---

### Task 3: Tier And RPM Workbook Integration

**Files:**
- Modify: `r6_tier_chart.py`
- Modify: `r6_rpm_chart.py`
- Modify: `tests/test_r6_tier_chart.py`
- Modify: `tests/test_r6_rpm_chart.py`

**Interfaces:**
- Consumes: shared status/patch renderers from Task 1 and `OperatorCard.score`.
- Produces: chart workbooks whose first two sheets are unchanged and whose final sheet is `补丁说明`.

- [ ] **Step 1: Update Tier workbook tests first**

Change the expected sheet list and inject patch scores:

```python
patch_scores = {name: 70 for name in notes.required_operator_names()}
chart.write_tier_workbook(
    output, cards, badge_dir, gadget_icons, patch_scores=patch_scores
)
self.assertEqual(
    workbook.sheetnames,
    ["进攻方简图", "防守方简图", "补丁说明"],
)
```

For each original sheet assert:

```python
self.assertEqual(sheet.cell(sheet.max_row, 1).value, notes.DATA_STATUS_TEXT)
self.assertTrue(any(
    str(merged).endswith(f":K{sheet.max_row}")
    for merged in sheet.merged_cells.ranges
))
self.assertIn(str(sheet.max_row), sheet.print_area)
```

Filter tier-label assertions through `chart.TIER_ORDER` so the appended status text is not treated as a tier label.

- [ ] **Step 2: Update RPM workbook tests first**

Inject the same complete patch score map into `write_rpm_workbook`, expect `["进攻方射速榜", "防守方射速榜", "补丁说明"]`, and assert the last nonempty row on each ranking sheet is the shared status text merged through column `O`.

- [ ] **Step 3: Run both layout tests and verify RED**

Run:

```powershell
python -m unittest \
  tests.test_r6_tier_chart.WorkbookLayoutTests.test_renders_two_five_card_tier_sheets_with_images_and_print_settings \
  tests.test_r6_rpm_chart.RpmWorkbookTests.test_renders_sorted_five_card_rows_with_letter_only_tier_boxes -v
```

Expected: both fail because the writer signatures and third sheets do not exist.

- [ ] **Step 4: Integrate Tier workbook rendering**

Extend the signature:

```python
def write_tier_workbook(
    path: Path,
    cards: Mapping[str, Iterable[OperatorCard]],
    operator_icon_dir: Path,
    gadget_icons: Mapping[str, Path],
    patch_scores: Optional[Mapping[str, int]] = None,
) -> None:
```

After `_render_side_sheet()`:

```python
status_row = append_data_status_row(sheet, 1 + CARDS_PER_ROW * 2)
sheet.print_area = f"A1:{get_column_letter(1 + CARDS_PER_ROW * 2)}{status_row}"
```

After both side sheets:

```python
scores = (
    dict(patch_scores)
    if patch_scores is not None
    else {
        card.name: card.score
        for side in SOURCE_SHEETS
        for card in normalized_cards[side]
    }
)
add_patch_notes_sheet(workbook, scores)
```

- [ ] **Step 5: Integrate RPM workbook rendering**

Apply the same optional parameter and score derivation. Append status through column `CARDS_PER_ROW * 3`, update each print area, then add the shared patch sheet before saving.

- [ ] **Step 6: Run Tier and RPM test modules and verify GREEN**

Run:

```powershell
python -m unittest tests.test_r6_tier_chart tests.test_r6_rpm_chart -v
```

Expected: all chart tests pass, original image counts remain unchanged, and each workbook has the third patch sheet.

---

### Task 4: Documentation, Skill Sync, Final Workbooks, And Visual QA

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Modify: `.codex/skills/build-r6-operator-report/SKILL.md`
- Copy: `r6_patch_notes.py` to `.codex/skills/build-r6-operator-report/scripts/r6_patch_notes.py`
- Copy updated: `r6_operator_stats.py`, `r6_tier_chart.py`, `r6_rpm_chart.py`
- Regenerate: `r6_operator_stats.xlsx`
- Regenerate: `r6_operator_tier_chart.xlsx`
- Regenerate: `r6_operator_rpm_chart.xlsx`

**Interfaces:**
- Consumes: all passing generators and tests from Tasks 1-3.
- Produces: documented, validated project and three final Excel workbooks.

- [ ] **Step 1: Update documentation**

Document that every workbook now contains `补丁说明`, that every original sheet ends with the latest-data status line, and that only Athieno scores remain pinned to the video. Add an Unreleased changelog entry naming Y11S2.1/Y11S2.2, source hyperlinks, direction colours, and strength-impacting fixes.

- [ ] **Step 2: Update and sync the project Skill**

Add Skill instructions requiring:

```text
- Verify the final worksheet is exactly `补丁说明`.
- Verify each original worksheet's final nonempty row equals DATA_STATUS_TEXT.
- Verify patch hyperlinks target Huiji Wiki and Ubisoft HTTPS URLs.
- Verify patch rows use the bundled Y11S2.1/Y11S2.2 metadata without re-scoring operators.
```

Mechanically copy the four production scripts into the Skill `scripts/` directory and verify SHA-256 equality.

- [ ] **Step 3: Run the complete automated verification**

Run:

```powershell
python -m unittest discover -s tests -v
py -3 <USER_HOME>\.codex\skills\.system\skill-creator\scripts\quick_validate.py .codex\skills\build-r6-operator-report
```

Expected: all tests pass and the validator prints `Skill is valid!`.

- [ ] **Step 4: Regenerate all final workbooks**

Run:

```powershell
python r6_operator_stats.py --output r6_operator_stats.xlsx
python r6_tier_chart.py --input r6_operator_stats.xlsx --output r6_operator_tier_chart.xlsx
python r6_rpm_chart.py --input r6_operator_stats.xlsx --output r6_operator_rpm_chart.xlsx
```

Expected counts remain 39 attackers, 38 defenders, and 62 unique automatic weapons.

- [ ] **Step 5: Audit workbook structure**

Reopen all three files and assert:

```python
expected = {
    "r6_operator_stats.xlsx": ["进攻方", "防守方", "补丁说明"],
    "r6_operator_tier_chart.xlsx": ["进攻方简图", "防守方简图", "补丁说明"],
    "r6_operator_rpm_chart.xlsx": ["进攻方射速榜", "防守方射速榜", "补丁说明"],
}
```

Check that all six original sheets end with `DATA_STATUS_TEXT`, all three patch sheets contain the six headers and five hyperlinks, source tables exclude the status row, patch sheets have zero images, and existing image totals remain 77/274/274.

- [ ] **Step 6: Perform Excel and visual QA**

Open each workbook read-only with Microsoft Excel COM, export the three `补丁说明` sheets and representative original sheets to PDF/PNG, and visually verify:

- all source links are readable and clickable;
- no patch detail or status text is clipped;
- green/red/yellow direction fills are distinguishable;
- the original chart/card layouts are unchanged;
- the bottom status row is visible and remains the last row.

- [ ] **Step 7: Report completion**

Report the exact automated test count, Skill validation result, Excel compatibility result, and provide standalone links to the three final `.xlsx` files.

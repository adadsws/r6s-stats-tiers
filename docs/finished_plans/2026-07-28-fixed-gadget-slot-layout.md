# Fixed Gadget Slot Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render every operator card with a fixed two-row, four-column gadget area whose seven side-specific gadget types always occupy the same 3+4 slots in PDF and XLSX.

**Architecture:** Add a focused `gadget_slots` module that owns the attack/defense slot maps and converts unordered `GadgetItem` values into an eight-element tuple containing one permanent empty slot. Both renderers consume that tuple, so PDF and XLSX share placement semantics while retaining their existing size units and drawing code.

**Tech Stack:** Python 3.12, Pillow, ReportLab, openpyxl, standard-library `unittest`, pypdf/pdfplumber, Poppler, Microsoft Excel COM, artifact-tool.

## Global Constraints

- Work directly on the current `main`; do not create a branch or worktree.
- Every card uses two rows and four columns even when it has fewer than seven gadgets.
- Attack slots are `破片手榴弹`, `闪光弹`, `硬突破炸药`, empty / `爆破炸药`, `阔剑地雷`, `电磁脉冲式冲击弹`, `烟雾弹`.
- Defense slots are `遥控炸药`, `机动护盾`, `冲击手榴弹`, empty / `倒刺铁丝网`, `防弹摄像头`, `观测工具阻拦器`, `感应警报器`.
- Missing gadgets leave their assigned slot empty; other gadgets never move forward.
- PDF visible icon size remains bounded by 6 mm × 6 mm after transparent-padding crop.
- XLSX tokens remain 24 px × 22 px with a 22 px visible box; the two-row gadget area is 34 pt.
- PDF and XLSX continue to omit gadget names and quantities.
- Impact EMP uses Ubisoft's official `R6S-EMP-Impact-grenade.png` source; its white line art is converted to black without changing the alpha silhouette.
- `output/` and `~temp/` remain ignored and uncommitted.
- Regenerated report artifacts and previews under `docs/` are committed.
- The active plan moves to `~archived/superpowers-plans/` after implementation is complete.

---

### Task 1: Shared fixed-slot mapping

**Files:**
- Create: `src/r6_report/gadget_slots.py`
- Create: `tests/test_gadget_slots.py`

**Interfaces:**
- Consumes: `r6_report.tier_chart.GadgetItem`.
- Produces: `GadgetSlotError(ValueError)`, `GADGET_SLOT_NAMES: Mapping[str, Tuple[Optional[str], ...]]`, and `arrange_gadgets(side: str, gadgets: Iterable[GadgetItem]) -> Tuple[Optional[GadgetItem], ...]`.

- [x] **Step 1: Write failing mapping tests**

Create `tests/test_gadget_slots.py` with `_path_setup`, `unittest`, and literal expectations:

```python
import unittest

import _path_setup

from r6_report.gadget_slots import GadgetSlotError, arrange_gadgets
from r6_report.tier_chart import GadgetItem


class GadgetSlotTests(unittest.TestCase):
    def test_attack_gadgets_use_fixed_three_plus_four_slots(self):
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
                None,
                "阔剑地雷",
                None,
                "烟雾弹",
            ),
        )

    def test_defense_gadgets_use_fixed_three_plus_four_slots(self):
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
```

- [x] **Step 2: Run tests and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_gadget_slots -v
```

Expected: import failure because `r6_report.gadget_slots` does not exist.

- [x] **Step 3: Implement the shared slot module**

Create `src/r6_report/gadget_slots.py`:

```python
"""Define stable side-specific slots for secondary gadget icons."""

from typing import Iterable, Mapping, Optional, Tuple

from .tier_chart import GadgetItem


class GadgetSlotError(ValueError):
    """Raised when a gadget cannot be assigned to a fixed slot."""


GADGET_SLOT_NAMES: Mapping[
    str,
    Tuple[Optional[str], ...],
] = {
    "进攻方": (
        "破片手榴弹",
        "闪光弹",
        "硬突破炸药",
        None,
        "爆破炸药",
        "阔剑地雷",
        "电磁脉冲式冲击弹",
        "烟雾弹",
    ),
    "防守方": (
        "遥控炸药",
        "机动护盾",
        "冲击手榴弹",
        None,
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
    try:
        slot_names = GADGET_SLOT_NAMES[side]
    except KeyError as error:
        raise GadgetSlotError("未知阵营：%s" % side) from error

    items = tuple(gadgets)
    by_name = {item.name: item for item in items}
    unknown = sorted(set(by_name) - {name for name in slot_names if name})
    if unknown:
        raise GadgetSlotError(
            "未定义固定槽位：%s" % "、".join(unknown)
        )
    return tuple(by_name.get(name) if name else None for name in slot_names)
```

- [x] **Step 4: Run the focused and mapping-adjacent tests**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_gadget_slots test_r6_leaderboards.LeaderboardClassificationTests -v
```

Expected: all tests pass.

- [x] **Step 5: Commit the shared mapping**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/gadget_slots.py tests/test_gadget_slots.py
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: define fixed gadget slots"
```

---

### Task 1B: Official Impact EMP line-art source

**Files:**
- Modify: `src/r6_report/tier_chart.py:61-76,269-325`
- Modify: `tests/test_r6_tier_chart.py:147-235`
- Replace: `data/icons/gadget/impact-emp-grenade.png`

**Interfaces:**
- Consumes: Ubisoft official PNG URL confirmed from the Sledge operator loadout page.
- Produces: `GADGET_DIRECT_URLS: Mapping[str, str]` and a cached black line-art PNG retaining the official image alpha channel.

- [x] **Step 1: Write a failing official-source test**

Add a test that prepares only `电磁脉冲式冲击弹` with a fake downloader. The downloader writes a transparent 16 × 16 PNG containing opaque white pixels. Assert:

```python
self.assertEqual(
    calls[0][-1],
    (
        "https://staticctf.ubisoft.com/"
        "J3yJr34U2pZ2Ieem48Dwy9uqj5PNUQTn/"
        "7izurbA5jDmnsmdeBdgKZO/"
        "29bca81243dda4084a92521ac0c03592/"
        "R6S-EMP-Impact-grenade.png"
    ),
)
with Image.open(icons["电磁脉冲式冲击弹"]) as icon:
    rgba = icon.convert("RGBA")
    visible = [
        pixel for pixel in rgba.getdata() if pixel[3] > 0
    ]
self.assertTrue(visible)
self.assertTrue(all(pixel[:3] == (0, 0, 0) for pixel in visible))
```

The test's literal direct URL is:

```text
https://staticctf.ubisoft.com/J3yJr34U2pZ2Ieem48Dwy9uqj5PNUQTn/7izurbA5jDmnsmdeBdgKZO/29bca81243dda4084a92521ac0c03592/R6S-EMP-Impact-grenade.png
```

- [x] **Step 2: Run the test and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_r6_tier_chart.GadgetIconTests.test_impact_emp_uses_official_line_art_source -v
```

Expected: failure because `GADGET_DIRECT_URLS` does not exist and the current path resolves the gray Wiki real-object file.

- [x] **Step 3: Add the official source override and color normalization**

Define:

```python
GADGET_DIRECT_URLS = {
    "电磁脉冲式冲击弹": (
        "https://staticctf.ubisoft.com/"
        "J3yJr34U2pZ2Ieem48Dwy9uqj5PNUQTn/"
        "7izurbA5jDmnsmdeBdgKZO/"
        "29bca81243dda4084a92521ac0c03592/"
        "R6S-EMP-Impact-grenade.png"
    ),
}
```

In `prepare_gadget_icons`, select `GADGET_DIRECT_URLS.get(name)` before calling `resolve_wiki_file_url`. After validating the downloaded temporary file and only for names in `GADGET_DIRECT_URLS`, convert every visible pixel to `(0, 0, 0, alpha)` with Pillow and save back to the temporary PNG before atomic replacement.

- [x] **Step 4: Replace the current cache from the confirmed official file**

Use the already downloaded `~temp/official-impact-emp-grenade.png`, convert its visible white line pixels to black while preserving dimensions and alpha, and write the result to `data/icons/gadget/impact-emp-grenade.png`. Verify the visible alpha bounding box differs from the previous real-object image and all visible RGB pixels are black.

- [x] **Step 5: Run gadget source tests**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_r6_tier_chart.GadgetIconTests -v
```

Expected: all gadget icon tests pass.

- [x] **Step 6: Commit the official source update**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/tier_chart.py tests/test_r6_tier_chart.py data/icons/gadget/impact-emp-grenade.png
git -c safe.directory=<PROJECT_ROOT> commit -m "fix: use official impact emp icon"
```

---

### Task 2: XLSX fixed two-row placement

**Files:**
- Modify: `src/r6_report/leaderboards.py:20-30,527-550,692-713`
- Modify: `tests/test_r6_leaderboards.py:607-674`

**Interfaces:**
- Consumes: `arrange_gadgets(card.side, card.gadgets)` from Task 1.
- Produces: every populated XLSX card has a 34 pt gadget row and gadget image anchors derived from fixed slot indexes `0..7`.

- [x] **Step 1: Replace the seven-gadget regression fixture with real side-specific names**

Update `test_wraps_seven_gadgets_inside_striker_and_sentry_cards` so the attacker and defender receive deliberately scrambled complete sets:

```python
attacker = make_card(
    "Striker",
    1,
    gadgets=(
        "烟雾弹",
        "电磁脉冲式冲击弹",
        "破片手榴弹",
        "爆破炸药",
        "硬突破炸药",
        "阔剑地雷",
        "闪光弹",
    ),
)
defender = make_card(
    "Sentry",
    1,
    side="防守方",
    gadgets=(
        "感应警报器",
        "冲击手榴弹",
        "倒刺铁丝网",
        "遥控炸药",
        "观测工具阻拦器",
        "机动护盾",
        "防弹摄像头",
    ),
)
```

Change the literal expected anchors for both sheets to:

```python
[
    (name_cell.column - 1, 14, 0),
    (name_cell.column, 14, 0),
    (name_cell.column + 1, 14, 0),
    (name_cell.column - 1, 14, 22),
    (name_cell.column, 14, 22),
    (name_cell.column + 1, 14, 22),
    (name_cell.column + 2, 14, 22),
]
```

Add a sparse-card test using an attacker with only `烟雾弹`; assert its sole image is anchored at the fourth information column with row offset 22 px and that the gadget row height is 34 pt.

- [x] **Step 2: Run the XLSX tests and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_r6_leaderboards.LeaderboardWorkbookTests.test_wraps_seven_gadgets_inside_striker_and_sentry_cards test_r6_leaderboards.LeaderboardWorkbookTests.test_sparse_gadget_keeps_its_fixed_xlsx_slot -v
```

Expected: the complete-set test reports old 4+3 input-order anchors, and the sparse smoke icon reports first-row first-column placement instead of second-row fourth-column placement.

- [x] **Step 3: Render XLSX from the shared eight-slot tuple**

Import `arrange_gadgets` in `leaderboards.py`. Replace the chunk-dependent `gadget_lines` calculation with:

```python
sheet.row_dimensions[gadget_row].height = (
    2 * theme.XLSX_CARD_BODY_ROW_PT
)
```

Replace enumeration over `card.gadgets` with:

```python
for gadget_slot, gadget in enumerate(
    arrange_gadgets(card.side, card.gadgets)
):
    if gadget is None:
        continue
    gadget_column = gadget_slot % theme.GADGETS_PER_LINE
    gadget_line = gadget_slot // theme.GADGETS_PER_LINE
    token = _excel_image(
        token_paths[(gadget.name, gadget.quantity)]
    )
    _add_offset_image(
        sheet,
        token,
        first_info + gadget_column,
        gadget_row,
        theme.XLSX_GADGET_TOKEN_PX[0],
        theme.XLSX_GADGET_TOKEN_PX[1],
        theme.XLSX_GADGET_COLUMN_OFFSET_PX,
        gadget_line * theme.XLSX_GADGET_TOKEN_PX[1],
    )
```

- [x] **Step 4: Run XLSX renderer tests**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_r6_leaderboards.LeaderboardWorkbookTests -v
```

Expected: all workbook tests pass, including 3+4 fixed positions and sparse-slot preservation.

- [x] **Step 5: Commit the XLSX implementation**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/leaderboards.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: fix xlsx gadget positions"
```

---

### Task 3: PDF fixed two-row placement

**Files:**
- Modify: `src/r6_report/pdf_leaderboards.py:20-30,277-305`
- Modify: `tests/test_r6_leaderboards.py:686-733`

**Interfaces:**
- Consumes: `arrange_gadgets(card.side, card.gadgets)` from Task 1.
- Produces: a PDF gadget `Table` with exactly two rows and four equal-width columns for every card.

- [x] **Step 1: Change the PDF component test to assert a sparse fixed slot**

Keep the existing padded 40×20 visible icon fixture and the attacker with only `烟雾弹`. Replace the first-cell lookup with literal table assertions:

```python
self.assertEqual(len(gadget_table._cellvalues), 2)
self.assertEqual(
    gadget_table._colWidths,
    [19 * mm, 19 * mm, 19 * mm, 19 * mm],
)
self.assertEqual(gadget_table._cellvalues[0], ["", "", "", ""])
self.assertEqual(gadget_table._cellvalues[1][:3], ["", "", ""])
gadget_icon = gadget_table._cellvalues[1][3]
self.assertIsInstance(gadget_icon, pdf_lb.Image)
```

Retain the 6 mm × 3 mm visible-aspect assertions and centered alignment assertion for `gadget_icon`.

- [x] **Step 2: Run the PDF test and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_row_omits_name_and_uses_equal_slots -v
```

Expected: the current table has one row and places smoke in its first cell.

- [x] **Step 3: Build exactly two PDF rows from shared slots**

Import `arrange_gadgets` in `pdf_leaderboards.py`. Replace the append-and-chunk logic with:

```python
gadget_cells = []
for gadget in arrange_gadgets(card.side, card.gadgets):
    if gadget is None:
        gadget_cells.append("")
        continue
    path = gadget_icons[gadget.name]
    icon_box = theme.PDF_GADGET_ICON_BOX_MM * mm
    image_width, image_height = _fit_image_size(
        path,
        icon_box,
        icon_box,
    )
    icon = Image(
        _cropped_image_source(path),
        width=image_width,
        height=image_height,
    )
    icon.hAlign = "CENTER"
    gadget_cells.append(icon)

gadget_rows = [
    gadget_cells[:theme.GADGETS_PER_LINE],
    gadget_cells[theme.GADGETS_PER_LINE:],
]
```

Pass `gadget_rows` directly to `Table`; do not fall back to a single empty row.

- [x] **Step 4: Run all PDF and CLI tests**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
python -m unittest test_r6_leaderboards.LeaderboardCliTests -v
```

Expected: all PDF/CLI tests pass, including five PDF and five XLSX generation.

- [x] **Step 5: Commit the PDF implementation**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/pdf_leaderboards.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: fix pdf gadget positions"
```

---

### Task 4: Full verification and regenerated artifacts

**Files:**
- Regenerate: `docs/视频评分榜.pdf`
- Regenerate: `docs/视频评分榜.xlsx`
- Regenerate: `docs/主武器射速榜.pdf`
- Regenerate: `docs/主武器射速榜.xlsx`
- Regenerate: `docs/速度榜.pdf`
- Regenerate: `docs/速度榜.xlsx`
- Regenerate: `docs/稀有枪械榜.pdf`
- Regenerate: `docs/稀有枪械榜.xlsx`
- Regenerate: `docs/次要装备榜.pdf`
- Regenerate: `docs/次要装备榜.xlsx`
- Regenerate: `docs/previews/*.png`
- Move: `docs/superpowers/plans/2026-07-28-fixed-gadget-slot-layout.md` to `~archived/superpowers-plans/2026-07-28-fixed-gadget-slot-layout.md`

**Interfaces:**
- Consumes: completed shared mapping and both renderers.
- Produces: committed PDF/XLSX artifacts and previews verified against the design.

- [x] **Step 1: Run the complete automated test suite**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONUTF8='1'
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue
python -m unittest discover -s tests -v
```

Expected: all tests pass with zero failures and no modified `__pycache__` files.

- [x] **Step 2: Confirm no workbook locks and generate into `~temp/`**

Run:

```powershell
Get-ChildItem -LiteralPath docs -Filter '~$*.xlsx' -Force
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m r6_report.leaderboards `
  --data-dir data `
  --input data/r6_operator_stats.xlsx `
  --output-dir ~temp/fixed-gadget-slot-output
```

Expected: the lock-file command prints nothing and generation reports five PDF plus five XLSX paths.

- [x] **Step 3: Verify generated structures before replacing docs**

Use openpyxl and pypdf to assert for all five outputs:

- each workbook contains attack, defense, and patch sheets;
- every populated operator card has a 34 pt gadget row;
- Striker and Sentry each have seven images at the literal 3+4 anchors;
- a sparse smoke-only attacker places its image in second-row fourth-column in the component test;
- every PDF contains three pages;
- PDF card text contains no gadget names or quantities.

Expected: all assertions pass.

- [x] **Step 4: Copy the ten verified files to `docs/` and refresh previews**

Copy only the five `.pdf` and five `.xlsx` files from `~temp/fixed-gadget-slot-output/` into `docs/`. Render PDF pages 1 and 2 at 100 DPI with Poppler into the existing ten README preview names under `docs/previews/`.

- [x] **Step 5: Run spreadsheet semantic checks**

Import all five XLSX files with artifact-tool, scan all 15 worksheets for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, and `#N/A`, then render every worksheet.

Expected: five workbooks import successfully, 15 worksheets render, and the formula-error scan returns zero matches.

- [x] **Step 6: Perform native visual verification**

Open the five XLSX files read-only through Microsoft Excel COM, export every workbook to PDF, and render the exports with Poppler. Also render all five project PDFs with Poppler. Review:

- all 15 project PDF pages;
- all attack/defense XLSX pages and patch-note continuation pages;
- first-row fourth slot is empty on every populated card;
- icons do not compact when adjacent gadget types are missing;
- Striker and Sentry display 3+4 icons;
- all gadget areas are exactly two rows;
- no icon overlaps a card border, neighboring icon, footer, or following band;
- other card fields, gray missing-value fills, patch colors, sorting, and pagination remain unchanged.

Expected: no clipping, overlap, missing icon, compaction, or layout drift.

- [x] **Step 7: Archive the completed implementation plan**

Move this file to:

```text
~archived/superpowers-plans/2026-07-28-fixed-gadget-slot-layout.md
```

Do not delete it and do not add `~archived/output/`.

- [x] **Step 8: Run final verification and commit artifacts**

Run the complete `unittest` suite again, rerun the structural validators, and inspect `git diff --check` plus `git status --short`. Then commit:

```powershell
git -c safe.directory=<PROJECT_ROOT> add docs src tests ~archived/superpowers/plans/2026-07-28-fixed-gadget-slot-layout.md
git -c safe.directory=<PROJECT_ROOT> commit -m "docs: regenerate fixed-slot reports"
```

Expected: the commit succeeds on `main`; afterward the only pre-existing unrelated status entry is `?? ~archived/output/`.

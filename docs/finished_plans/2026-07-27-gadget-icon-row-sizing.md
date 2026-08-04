# Gadget Icon Row Sizing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让五种榜单的 PDF 与 XLSX 省略次要装备数量标记、使用正文行等高图标，并保证 Striker、Sentry 的 7 个装备按 4+3 两行完整可见。

**Architecture:** 在 `report_theme.py` 集中定义 PDF/XLSX 的正文行高、图标边界和每行装备数；两个 renderer 只负责把共享语义转换成各自单位。XLSX 通过透明 token 和锚点边界测试保证第二行不越界，PDF 通过文本提取与图像尺寸测试保证数量文字消失且图标保持比例。

**Tech Stack:** Python 3.12、Pillow、openpyxl、ReportLab、pypdf、unittest、artifact-tool、Poppler

## Global Constraints

- PDF 装备图标边界为 6 mm × 6 mm，与卡片正文 6 mm 行高一致。
- XLSX 装备 token 为 20 px × 17 px，可见图标限制在 17 px × 17 px。
- XLSX 单行装备区为 17 pt；7 件装备按 4+3 两行显示，装备区为 34 pt。
- PDF 与 XLSX 每行最多显示 4 件装备，保持输入顺序。
- 数量保留在 `GadgetItem.quantity` 数据模型中，但不得出现在 PDF 或 XLSX 的视觉输出中。
- 图标必须保持原始长宽比并居中，不允许拉伸。
- 直接在用户指定的当前 `main` 执行，不创建分支或 worktree。

---

### Task 1: Shared gadget sizing semantics

**Files:**
- Modify: `src/r6_report/report_theme.py`
- Test: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Produces: `GADGETS_PER_LINE: int`
- Produces: `PDF_CARD_BODY_ROW_MM: int`
- Produces: `PDF_GADGET_ICON_BOX_MM: int`
- Produces: `XLSX_CARD_BODY_ROW_PT: int`
- Produces: `XLSX_GADGET_TOKEN_PX: Tuple[int, int]`
- Produces: `XLSX_GADGET_ICON_BOX_PX: int`

- [ ] **Step 1: Write the failing shared-theme test**

```python
def test_shared_gadget_dimensions_match_card_body_rows(self):
    self.assertEqual(theme.GADGETS_PER_LINE, 4)
    self.assertEqual(theme.PDF_GADGET_ICON_BOX_MM, theme.PDF_CARD_BODY_ROW_MM)
    self.assertEqual(theme.XLSX_GADGET_TOKEN_PX, (20, 17))
    self.assertEqual(theme.XLSX_GADGET_ICON_BOX_PX, 17)
    self.assertEqual(theme.XLSX_CARD_BODY_ROW_PT, 17)
```

- [ ] **Step 2: Run the test and verify it fails**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_shared_gadget_dimensions_match_card_body_rows -v
```

Expected: `ERROR` because the shared constants do not yet exist.

- [ ] **Step 3: Add the minimal shared constants**

```python
GADGETS_PER_LINE = 4
PDF_CARD_BODY_ROW_MM = 6
PDF_GADGET_ICON_BOX_MM = PDF_CARD_BODY_ROW_MM
XLSX_CARD_BODY_ROW_PT = 17
XLSX_GADGET_TOKEN_PX: Tuple[int, int] = (20, 17)
XLSX_GADGET_ICON_BOX_PX = 17
```

- [ ] **Step 4: Run the focused test and the theme-related tests**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_shared_gadget_dimensions_match_card_body_rows tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_missing_card_fields_use_shared_gray_fill -v
```

Expected: both tests `PASS`.

- [ ] **Step 5: Commit the shared semantics**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/report_theme.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: define shared gadget row dimensions"
```

### Task 2: XLSX quantity-free tokens and complete 7-icon layout

**Files:**
- Modify: `src/r6_report/leaderboards.py`
- Modify: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Consumes: Task 1 shared constants.
- Produces: `draw_gadget_token(source: Path, destination: Path) -> None`
- Produces: `_make_gadget_tokens(...) -> Mapping[str, Path]`

- [ ] **Step 1: Write failing token tests**

```python
def test_gadget_token_omits_quantity_and_matches_body_row(self):
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "source.png"
        output = root / "output.png"
        Image.new("RGBA", (32, 16), "black").save(source)

        lb.draw_gadget_token(source, output)

        with Image.open(output) as rendered:
            self.assertEqual(rendered.size, theme.XLSX_GADGET_TOKEN_PX)
            alpha_bounds = rendered.getchannel("A").getbbox()
            self.assertIsNotNone(alpha_bounds)
            self.assertLessEqual(alpha_bounds[2] - alpha_bounds[0], 17)
            self.assertLessEqual(alpha_bounds[3] - alpha_bounds[1], 17)
            white_pixels = [
                pixel for pixel in rendered.convert("RGBA").getdata()
                if pixel[:3] == (255, 255, 255) and pixel[3] > 0
            ]
            self.assertEqual(white_pixels, [])
```

- [ ] **Step 2: Strengthen the Striker/Sentry regression test**

For each 7-gadget card, assert:

```python
self.assertEqual(len(gadget_anchors), 7)
self.assertEqual(
    sorted(round(anchor.rowOff / 9525) for anchor in gadget_anchors),
    [0, 0, 0, 0, 17, 17, 17],
)
self.assertEqual(sheet.row_dimensions[gadget_row].height, 34)
row_height_px = round(34 * 96 / 72)
for image in gadget_images:
    top_px = round(image.anchor._from.rowOff / 9525)
    height_px = round(image.anchor.ext.cy / 9525)
    self.assertLessEqual(top_px + height_px, row_height_px)
```

This verifies all seven image bottoms remain inside the 34 pt visible row.

- [ ] **Step 3: Run both tests and verify they fail**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_gadget_token_omits_quantity_and_matches_body_row tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_wraps_seven_gadgets_inside_striker_and_sentry_cards -v
```

Expected: token signature/size assertions and exact 17 px offsets fail against the current 24 px × 20 px numbered token implementation.

- [ ] **Step 4: Implement quantity-free token generation**

Change token keys from `(name, quantity)` to `name`, remove the quantity parameter from `draw_gadget_token`, create a transparent `20 × 17` canvas, crop to the alpha bounds, scale within `17 × 17`, and center the visible icon. Remove the Pillow badge drawing imports and all badge drawing code.

- [ ] **Step 5: Implement a bounded 4+3 XLSX layout**

Use `theme.GADGETS_PER_LINE`, `theme.XLSX_CARD_BODY_ROW_PT`, and `theme.XLSX_GADGET_TOKEN_PX` for:

```python
gadget_lines = ceil(len(card.gadgets) / theme.GADGETS_PER_LINE)
sheet.row_dimensions[gadget_row].height = gadget_lines * 17
x_offset = (gadget_index % 4) * 20
y_offset = (gadget_index // 4) * 17
```

Anchor each token as 20 px × 17 px without extra vertical padding, so the second row occupies px 17–34 and stays within a 34 pt Excel row.

- [ ] **Step 6: Run focused and renderer tests**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_gadget_token_omits_quantity_and_matches_body_row tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_wraps_seven_gadgets_inside_striker_and_sentry_cards tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_gadget_token_has_no_red_highlight_border -v
```

Expected: all tests `PASS`, with 7 visible in-bounds anchors on both Striker and Sentry.

- [ ] **Step 7: Commit the XLSX fix**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/leaderboards.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "fix: keep seven gadget icons visible in xlsx"
```

### Task 3: PDF quantity-free labels and body-row-sized icons

**Files:**
- Modify: `src/r6_report/pdf_leaderboards.py`
- Modify: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Consumes: Task 1 shared constants.
- Produces: `_gadget_label(gadget: GadgetItem) -> str`

- [ ] **Step 1: Write failing PDF behavior tests**

```python
def test_pdf_gadget_label_omits_quantity(self):
    self.assertEqual(
        pdf_lb._gadget_label(tier.GadgetItem("烟雾弹", 3)),
        "烟雾弹",
    )

def test_pdf_gadget_image_fits_six_mm_box(self):
    with tempfile.TemporaryDirectory() as directory:
        icon = Path(directory) / "wide.png"
        Image.new("RGBA", (40, 20), "black").save(icon)
        box = theme.PDF_GADGET_ICON_BOX_MM * mm
        self.assertEqual(pdf_lb._fit_image_size(icon, box, box), (box, box / 2))
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_label_omits_quantity tests.test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_image_fits_six_mm_box -v
```

Expected: `ERROR` because `_gadget_label` does not exist and the renderer still uses a hard-coded 5 mm box.

- [ ] **Step 3: Implement the PDF label and 6 mm icon box**

Add:

```python
def _gadget_label(gadget: GadgetItem) -> str:
    return gadget.name
```

Use `theme.PDF_GADGET_ICON_BOX_MM * mm` for both width and height, preserve `_fit_image_size`, and remove the `×%d` quantity suffix. Keep the four-items-per-line grouping and existing two-line card structure.

- [ ] **Step 4: Run focused PDF and CLI tests**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_label_omits_quantity tests.test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_image_fits_six_mm_box tests.test_r6_leaderboards.LeaderboardCliTests.test_cli_writes_five_leaderboard_workbooks_and_pdfs -v
```

Expected: all tests `PASS`; extracted PDF text contains no `×2` or `×3`.

- [ ] **Step 5: Commit the PDF change**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/pdf_leaderboards.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: simplify pdf gadget rows"
```

### Task 4: Regenerate and visually verify all reports

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/*.pdf`
- Modify: `docs/*.xlsx`
- Modify: `README.assets/*.png`
- Move after completion: `docs/superpowers/plans/2026-07-27-gadget-icon-row-sizing.md` to `~archived/superpowers-plans/2026-07-27-gadget-icon-row-sizing.md`

**Interfaces:**
- Consumes: Tasks 1–3 renderers.
- Produces: five verified PDFs, five verified XLSX workbooks, and refreshed README previews.

- [ ] **Step 1: Update the changelog**

Under `Unreleased / Changed`, record that PDF/XLSX gadget quantity markers were removed and icon sizes now match the card body row. Under `Fixed`, record that Striker/Sentry seven-gadget XLSX cards now keep the 4+3 layout fully visible.

- [ ] **Step 2: Run the full test suite**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest discover -s tests -v
```

Expected: all tests `PASS`.

- [ ] **Step 3: Regenerate the ten report artifacts**

Run the repository’s unified report command against the checked-in `data/` snapshot, writing temporary outputs under `~temp/`, then copy the five PDF and five XLSX results to `docs/`.

- [ ] **Step 4: Verify XLSX structure and formulas**

Use artifact-tool to import each workbook, render all 15 sheets, and scan formulas for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, and `#N/A`. Inspect the Striker and Sentry rows in native Excel/PDF output and confirm all 7 icons appear as 4+3 with no quantity badges.

- [ ] **Step 5: Verify all PDF pages**

Use Poppler to render all 15 PDF pages. Inspect every page for clipping, overlap, quantity suffixes, and stretched icons; confirm the gadget icons are visually equal in height to the other card body rows.

- [ ] **Step 6: Refresh README previews**

Update the ten report preview PNGs from the newly rendered attack/defense pages and confirm README references remain valid.

- [ ] **Step 7: Archive the completed plan**

Move the checked implementation plan into `~archived/superpowers-plans/` without deleting it.

- [ ] **Step 8: Run final verification and commit**

Run the complete suite again, recheck artifact counts and `git diff --check`, then commit all relevant source, tests, changelog, plans, reports, and previews:

```powershell
git -c safe.directory=<PROJECT_ROOT> add CHANGELOG.md README.assets docs src tests '~archived/superpowers-plans'
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: align gadget icons with card rows"
```

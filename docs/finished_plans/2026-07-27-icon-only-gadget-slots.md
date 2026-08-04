# Icon-Only Gadget Slots Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove PDF gadget names and make PDF/XLSX gadget icons occupy four fixed, evenly spaced, centered slots per row.

**Architecture:** Keep `GadgetItem` data unchanged and modify only the two report renderers. PDF gadget tables will contain image flowables directly in four equal columns; XLSX images will anchor to four consecutive information columns with one centered token per column. Existing shared row-count and icon-size semantics remain authoritative.

**Tech Stack:** Python 3.12、ReportLab、openpyxl、Pillow、unittest、artifact-tool、Poppler、Microsoft Excel

## Global Constraints

- PDF 与 XLSX 的装备行只显示图标，不显示数量或装备名称。
- 每行固定四个槽位；一至四件从左到右占槽，五至七件在第二行从左开始。
- PDF 图标继续在 6 mm × 6 mm 边界框内保持原始长宽比。
- XLSX token 保持 20 px × 17 px，装备行保持单行 17 pt、两行 34 pt。
- 卡片其他字段、颜色、灰底、勾号、排序、分档、分页和打印设置保持不变。
- 生成前必须确认没有 `docs/~$*.xlsx` Excel 锁文件。
- 直接在用户指定的当前 `main` 执行，不创建分支或 worktree。

---

### Task 1: PDF icon-only equal slots

**Files:**
- Modify: `src/r6_report/pdf_leaderboards.py:259-302`
- Modify: `tests/test_r6_leaderboards.py:673-708`

**Interfaces:**
- Consumes: `theme.GADGETS_PER_LINE`、`theme.PDF_GADGET_ICON_BOX_MM`
- Produces: `_card_flowable(...)` whose gadget table cells contain ReportLab `Image` flowables only

- [ ] **Step 1: Change the PDF component test to require image-only slots**

Replace the label expectation with assertions on the real card flowable:

```python
gadget_table = card_table._cellvalues[1][0]
gadget_icon = gadget_table._cellvalues[0][0]
self.assertIsInstance(gadget_icon, pdf_lb.Image)
self.assertEqual(
    gadget_table._colWidths,
    [19 * mm, 19 * mm, 19 * mm, 19 * mm],
)
self.assertEqual(gadget_icon.hAlign, "CENTER")
self.assertAlmostEqual(gadget_icon.drawWidth, 6 * mm, delta=0.01)
self.assertAlmostEqual(gadget_icon.drawHeight, 3 * mm, delta=0.01)
```

This fails if a slot still contains the current nested `Table([icon, name])`.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_row_omits_name_and_uses_equal_slots -v
```

Expected: `FAIL` because the first gadget cell is a nested `Table`, not an `Image`.

- [ ] **Step 3: Implement direct image cells**

In `_card_flowable`:

```python
gadget_row.append(icon)
```

Delete creation of the nested table and gadget-name paragraph. Keep four 19 mm columns, then apply table style commands that center every cell horizontally and vertically with zero inner padding:

```python
gadgets.setStyle(
    TableStyle(
        [
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]
    )
)
```

- [ ] **Step 4: Run focused and CLI tests**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest test_r6_leaderboards.LeaderboardCliTests.test_pdf_gadget_row_omits_name_and_uses_equal_slots test_r6_leaderboards.LeaderboardCliTests.test_main_generates_five_workbooks_and_five_complete_pdfs -v
```

Expected: both tests `PASS`.

- [ ] **Step 5: Commit the PDF change**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/pdf_leaderboards.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: use icon-only pdf gadget slots"
```

### Task 2: XLSX four-column centered slots

**Files:**
- Modify: `src/r6_report/report_theme.py`
- Modify: `src/r6_report/leaderboards.py:692-714`
- Modify: `tests/test_r6_leaderboards.py:607-658`

**Interfaces:**
- Consumes: `theme.GADGETS_PER_LINE`、`theme.XLSX_GADGET_TOKEN_PX`
- Produces: `theme.XLSX_GADGET_COLUMN_OFFSET_PX: int`
- Produces: gadget anchors whose zero-based columns are `first_info - 1 + slot_index`

- [ ] **Step 1: Strengthen the seven-gadget anchor test**

For each Striker/Sentry card, preserve image order and assert:

```python
self.assertEqual(
    [
        (
            image.anchor._from.col,
            round(image.anchor._from.colOff / 9525),
            round(image.anchor._from.rowOff / 9525),
        )
        for image in gadget_images
    ],
    [
        (name_cell.column - 1, 16, 0),
        (name_cell.column, 16, 0),
        (name_cell.column + 1, 16, 0),
        (name_cell.column + 2, 16, 0),
        (name_cell.column - 1, 16, 17),
        (name_cell.column, 16, 17),
        (name_cell.column + 1, 16, 17),
    ],
)
```

Keep the existing 34 pt row-height and image-boundary assertions.

- [ ] **Step 2: Run the test and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
$env:PYTHONPATH='tests'
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest test_r6_leaderboards.LeaderboardWorkbookTests.test_wraps_seven_gadgets_inside_striker_and_sentry_cards -v
```

Expected: `FAIL` because all seven images currently share one anchor column and use compact `0/20/40/60px` offsets.

- [ ] **Step 3: Implement four centered column anchors**

Add:

```python
XLSX_GADGET_COLUMN_OFFSET_PX = 16
```

For each gadget, calculate `slot = gadget_index % 4` and `line = gadget_index // 4`, then call `_add_offset_image` with:

```python
column=first_info + slot
x_offset=theme.XLSX_GADGET_COLUMN_OFFSET_PX
y_offset=line * theme.XLSX_GADGET_TOKEN_PX[1]
```

Keep token width/height and row heights unchanged.

- [ ] **Step 4: Run focused workbook tests**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest test_r6_leaderboards.LeaderboardWorkbookTests.test_wraps_seven_gadgets_inside_striker_and_sentry_cards test_r6_leaderboards.LeaderboardWorkbookTests.test_gadget_token_omits_quantity_and_matches_body_row -v
```

Expected: both tests `PASS`.

- [ ] **Step 5: Commit the XLSX change**

```powershell
git -c safe.directory=<PROJECT_ROOT> add src/r6_report/report_theme.py src/r6_report/leaderboards.py tests/test_r6_leaderboards.py
git -c safe.directory=<PROJECT_ROOT> commit -m "fix: distribute xlsx gadget icons across slots"
```

### Task 3: Regenerate and verify report artifacts

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/*.pdf`
- Modify: `docs/*.xlsx`
- Modify: `docs/previews/*.png`
- Move after completion: `docs/superpowers/plans/2026-07-27-icon-only-gadget-slots.md` to `~archived/superpowers-plans/2026-07-27-icon-only-gadget-slots.md`

**Interfaces:**
- Consumes: Tasks 1–2 renderers.
- Produces: five PDFs, five XLSX workbooks, and ten refreshed README previews.

- [ ] **Step 1: Update the changelog**

Under `Unreleased / Changed`, state that PDF gadget names were removed and both formats now use four fixed centered slots. Under `Fixed`, state that four-item XLSX rows no longer cluster at the left.

- [ ] **Step 2: Verify no Excel lock files exist**

Run:

```powershell
Get-ChildItem -LiteralPath docs -Filter '~$*.xlsx'
```

Expected: no output. If a lock exists, stop before overwriting reports.

- [ ] **Step 3: Run the full test suite**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest discover -s tests -v
```

Expected: all tests `PASS`.

- [ ] **Step 4: Regenerate all reports**

Run `python -m r6_report.leaderboards` with `data/r6_operator_stats.xlsx`, `data/` icon caches, and a `~temp/` output directory. Copy the five PDFs and five XLSX workbooks to `docs/`.

- [ ] **Step 5: Verify XLSX and PDF semantics**

Programmatically inspect production XLSX anchors to confirm every four-item first line uses four consecutive columns with the same 16 px inner offset. Verify Striker/Sentry remain 4+3 and all image bottoms stay inside the 34 pt row. Inspect PDF card components and extracted text to confirm equipment names are absent from gadget rows.

- [ ] **Step 6: Render and inspect all artifacts**

Use artifact-tool to import and render all 15 XLSX sheets and scan formula errors. Use Microsoft Excel native export to check the visible four-slot spacing. Use Poppler to render all 15 PDF pages and inspect icon centering, clipping, patch tables, headers, and footers.

- [ ] **Step 7: Refresh README previews**

Copy the new PDF attack/defense page renders to the ten existing `docs/previews/*.png` paths.

- [ ] **Step 8: Archive the plan and run final verification**

Move this plan to `~archived/superpowers-plans/`, run the complete 88-test suite again, run artifact structure checks, and run `git diff --check`.

- [ ] **Step 9: Commit final artifacts**

```powershell
git -c safe.directory=<PROJECT_ROOT> add CHANGELOG.md docs tests '~archived/superpowers-plans'
git -c safe.directory=<PROJECT_ROOT> commit -m "feat: unify icon-only gadget slots"
```

# Exclude Gadget Balance From Operator Markers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Keep pure secondary-gadget balance changes in patch notes while excluding them from operator `+`, `-`, and `~` card markers.

**Architecture:** Add a conservative report-layer predicate in `leaderboards.py`, using the canonical gadget slots to recognize gadget names and explicit loadout phrases. Filter only the marker aggregation path; leave source parsing, stored patch data, and patch-note rendering unchanged.

**Tech Stack:** Python 3.9+, `unittest`, openpyxl, ReportLab, Git

## Global Constraints

- Pure secondary-gadget performance changes do not count as operator buffs or nerfs.
- Operator gadget additions, removals, and replacements still count as operator changes.
- Mixed operator ability and gadget-loadout records still count.
- Unknown or ambiguous records still count.
- Patch data and PDF/XLSX patch-note pages remain complete.
- PDF and XLSX outputs must remain mutually consistent.
- Work directly on the current local `main`, as explicitly requested.
- Commit every repository change and publish through the sanitized `output/github-export/` repository.

---

### Task 1: Add marker eligibility semantics with regression tests

**Files:**
- Modify: `tests/test_r6_leaderboards.py`
- Modify: `src/r6_report/leaderboards.py`

**Interfaces:**
- Consumes: `PatchChange.detail` and `GADGET_SLOT_NAMES`
- Produces: `counts_as_operator_adjustment(change: PatchChange) -> bool`
- Produces: `patch_markers(sources: ReportSources) -> Mapping[str, str]` with gadget-only changes excluded

- [x] **Step 1: Add failing eligibility and aggregation tests**

Add tests that construct `PatchChange` values for:

```python
("增强", "Mute", "防弹摄像头电磁脉冲波的爆炸范围提升。", False)
("增强", "Wamai", "新增机动护盾。", True)
("混合", "Wamai", "冲击手榴弹被替换为遥控炸药。", True)
("增强", "Wamai", "磁力销毁系统充能提高；新增机动护盾。", True)
("削弱", "Ace", "技能持续时间缩短。", True)
```

Also create one `ReportSources` fixture containing a pure gadget change and an ordinary operator change. Assert the gadget-only subject is absent from `patch_markers()` and the ordinary subject retains its marker.

- [x] **Step 2: Run the focused test and verify RED**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest tests.test_r6_leaderboards.LeaderboardClassificationTests -v
```

Expected: failure because `counts_as_operator_adjustment` does not exist or gadget-only changes are still aggregated.

- [x] **Step 3: Implement the conservative predicate**

In `leaderboards.py`:

- Import `PatchChange` and `GADGET_SLOT_NAMES`.
- Flatten all non-`None` gadget names from `GADGET_SLOT_NAMES`.
- Define explicit loadout regular expressions for `新增<装备>`, `移除<装备>`,
  `<装备>被替换为<装备>`, `<装备>替换为<装备>`, `改为配备<装备>`, and
  `配备<装备>`.
- Split details on `。`, `；`, or line breaks and inspect the first non-empty clause.
- Return `False` only when the first clause mentions a known gadget and no explicit loadout expression appears anywhere in the detail.
- Make `patch_markers()` skip changes for which the predicate returns `False`.

- [x] **Step 4: Run focused tests and verify GREEN**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest tests.test_r6_leaderboards.LeaderboardClassificationTests -v
```

Expected: all classification tests pass.

- [x] **Step 5: Commit marker filtering**

Run:

```powershell
git add src/r6_report/leaderboards.py tests/test_r6_leaderboards.py
git commit -m "fix: exclude gadget balance from operator markers"
```

### Task 2: Prove patch notes retain gadget changes

**Files:**
- Modify: `tests/test_r6_patch_notes.py`

**Interfaces:**
- Consumes: unchanged `PatchRecord.changes`
- Produces: regression proof that patch-note rendering retains gadget-only rows

- [x] **Step 1: Add a patch-note retention test**

Build a patch containing:

```python
PatchChange(
    direction="增强",
    subject="Mute",
    detail="防弹摄像头电磁脉冲波的爆炸范围提升。",
)
```

Render the patch sheet and assert its last data rows contain subject `Mute`, direction `增强`, and the full detail text.

- [x] **Step 2: Run the focused patch-note test**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest tests.test_r6_patch_notes -v
```

Expected: all patch-note tests pass without production changes.

- [x] **Step 3: Commit the regression test**

Run:

```powershell
git add tests/test_r6_patch_notes.py
git commit -m "test: retain gadget balance in patch notes"
```

### Task 3: Regenerate and verify PDF/XLSX reports

**Files:**
- Modify: `docs/*.pdf`
- Modify: `docs/*.xlsx`
- Modify: `docs/previews/*.png`

**Interfaces:**
- Consumes: current `data/` snapshot and updated marker filtering
- Produces: five verified PDFs, five verified XLSX workbooks, and refreshed previews

- [x] **Step 1: Run the full test suite**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest discover -s tests -v
```

Expected: all tests pass.

- [x] **Step 2: Generate all reports into a fresh ignored temporary directory**

Run the existing leaderboard CLI with `data/` and `data/r6_operator_stats.xlsx`,
writing five PDFs and five XLSX workbooks under a new `~temp/` subdirectory.

- [x] **Step 3: Verify generated report structure**

Check that:

- Each workbook has three sheets.
- Each PDF has three pages.
- Mute and the other Y11S2.1 bulletproof-camera carriers no longer receive a marker solely from that change.
- Wamai retains `+` from Y11S2.2.
- The bulletproof-camera change remains on every patch-note page.

- [x] **Step 4: Copy verified reports to `docs/` and refresh previews**

Replace the five public PDFs and five public XLSX workbooks with verified outputs.
Render PDF previews using the established Poppler workflow and replace all ten
`docs/previews/*.png` images.

- [x] **Step 5: Perform visual and workbook QA**

Render all project PDFs and Excel-exported workbooks. Inspect every page for
marker accuracy, unchanged layout, clipping, overlap, fixed gadget slots, and
patch-note retention. Run the artifact-tool formula scan on all five XLSX files.

- [x] **Step 6: Commit regenerated artifacts**

Run:

```powershell
git add docs
git commit -m "docs: regenerate operator marker reports"
```

### Task 4: Final verification, archive, and sanitized push

**Files:**
- Move locally: `docs/superpowers/plans/2026-07-28-exclude-gadget-balance-from-operator-markers.md`
- To: `~archived/superpowers-plans/2026-07-28-exclude-gadget-balance-from-operator-markers.md`
- Update ignored export: `output/github-export/`

**Interfaces:**
- Consumes: verified local `main`
- Produces: archived local plan and fast-forward remote `main`

- [x] **Step 1: Run fresh final tests**

Run the complete 94+ test suite and all report-structure checks again.

- [x] **Step 2: Mark this plan complete and archive it**

Mark every checkbox `[x]`, move the plan into `~archived/superpowers-plans/`,
and commit the move locally. Do not include the plan or archive in the public export.

- [x] **Step 3: Synchronize the sanitized export**

Fetch remote `main` in `output/github-export/`, ensure it is clean, then copy only
publishable committed files from local `main`. Exclude `output/`, `~temp/`,
`~archived/`, `~ref/`, `docs/superpowers/`, non-example `secrets/`,
`__pycache__/`, `*.pyc`, `*.pyo`, and `~$*.xlsx`.

- [x] **Step 4: Scan, test, and commit the public export**

Verify there are no credentials, private paths, or forbidden files. Run the full
test suite in the export, stage the reviewed diff, and commit:

```powershell
git commit -m "fix: exclude gadget balance from operator markers"
```

- [x] **Step 5: Fast-forward push and verify GitHub**

Fetch `origin/main`, require it to be an ancestor of export `main`, and push
without force. Confirm export and remote SHAs match; source, export, and remote
each have only `main`; the existing English Description and eight Topics remain
unchanged.

# Dual-Column Report Previews Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Show attack and defense previews side by side for each of the five leaderboard PDFs in README.

**Architecture:** Render PDF page 1 and page 2 independently with Poppler so each preview preserves its content-fitted page height. Use one Markdown table per leaderboard with attack on the left and defense on the right.

**Tech Stack:** Markdown, Poppler `pdftoppm`, PNG, Git.

## Global Constraints

- Render attack from PDF page 1 and defense from PDF page 2.
- Keep five leaderboard sections in their current order.
- Use Markdown tables, not composite images or fixed-width HTML.
- Do not use Microsoft Excel or LibreOffice.
- Commit locally and do not push.

---

### Task 1: Render ten preview images

**Files:**
- Create or update: `docs/previews/video-rating-attack.png`
- Create or update: `docs/previews/video-rating-defense.png`
- Create or update: `docs/previews/primary-rpm-attack.png`
- Create or update: `docs/previews/primary-rpm-defense.png`
- Create or update: `docs/previews/speed-attack.png`
- Create or update: `docs/previews/speed-defense.png`
- Create or update: `docs/previews/rare-weapons-attack.png`
- Create or update: `docs/previews/rare-weapons-defense.png`
- Create or update: `docs/previews/secondary-gadgets-attack.png`
- Create or update: `docs/previews/secondary-gadgets-defense.png`

**Interfaces:**
- Consumes: the first two pages of each `docs/*.pdf`.
- Produces: ten stable PNG paths used by README.

- [x] **Step 1: Render page 1 and page 2**

Use bundled `pdftoppm.exe` at 110 DPI with `-f 1 -l 2 -png`.

- [x] **Step 2: Normalize filenames**

Map page suffix `-1.png` to `-attack.png` and `-2.png` to `-defense.png` for each report prefix.

- [x] **Step 3: Verify image dimensions**

Require all ten PNG files to exist, have positive dimensions, and preserve a common width.

### Task 2: Replace README previews with dual columns

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: the ten preview paths from Task 1.
- Produces: five Markdown tables with left attack and right defense columns.

- [x] **Step 1: Update each leaderboard section**

Keep one heading and one PDF link per leaderboard, then add:

```markdown
| 进攻方 | 防守方 |
|:---:|:---:|
| ![进攻方预览](docs/previews/example-attack.png) | ![防守方预览](docs/previews/example-defense.png) |
```

- [x] **Step 2: Record the layout change**

Add one CHANGELOG entry describing the five dual-column attack/defense previews.

### Task 3: Visual and automated verification

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-dual-column-report-previews.md` to `~archived/superpowers-plans/2026-07-26-dual-column-report-previews.md`

**Interfaces:**
- Consumes: final README and ten PNG files.
- Produces: a verified, committed local tree.

- [x] **Step 1: Inspect all ten images**

Check titles, full tier bands, cards, sources, and page numbers; reject clipping, overlap, or unreadable text.

- [x] **Step 2: Audit README**

Require exactly five PDF links, ten unique preview paths, and five tables whose first column uses
`-attack.png` and second column uses `-defense.png`.

- [x] **Step 3: Run full tests**

```powershell
$env:PYTHONPATH='src'
python -m unittest discover -s tests -v
```

- [x] **Step 4: Archive the completed plan and commit**

Move this plan to `~archived/superpowers-plans/`, run `git diff --check`, stage all preview,
README, CHANGELOG, and plan-archive changes, then create one local commit without push.

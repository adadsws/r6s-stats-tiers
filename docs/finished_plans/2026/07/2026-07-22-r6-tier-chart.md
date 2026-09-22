# R6 Tier Chart Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a reusable Python script that reads `r6_operator_stats.xlsx` and generates a two-sheet horizontal tier-list workbook with five operator cards per row, official Huiji Wiki Badge images, compact weapon facts, and icon-based secondary gadgets.

**Architecture:** Keep source parsing and compact text conversion in pure functions, isolate Huiji gadget-image resolution behind an explicit filename map, and render the final workbook with openpyxl using fixed card geometry. The generated workbook contains no source-data sheets; each side is rendered independently from validated source rows.

**Tech Stack:** Python 3.9+, openpyxl 3.x, Pillow, standard-library unittest, system curl, Huiji Wiki read-only MediaWiki API.

## Global Constraints

- Default input is `r6_operator_stats.xlsx`; default output is `r6_operator_tier_chart.xlsx`.
- Output sheets are exactly `进攻方简图` and `防守方简图`.
- Visible Tier order is `S, A, B, C, D, F`; raw score `0` is normalized to the current lowest letter tier F.
- Use five operator cards per tier row and wrap within the same tier.
- Card text uses `主手半自` as the compact label, based on a weapon with `equipment == 1`, `firerate == 0`, and `projectile == 1`.
- Use official operator Badge assets from `assets/operator-icons/badge/`.
- Use official Huiji Wiki gadget icons for every secondary gadget; do not use third-party images.
- Preserve unrelated workspace changes and keep `AGENTS.md` ignored.
- Git commands currently fail because `.git` is not recognized as a repository; do not attempt commits until repository metadata is restored.

---

### Task 1: Source Workbook Parser And Compact Facts

**Files:**
- Create: `r6_tier_chart.py`
- Create: `tests/test_r6_tier_chart.py`

**Interfaces:**
- Produces: `OperatorCard`, `load_operator_cards(path: Path) -> Dict[str, List[OperatorCard]]`
- Produces: `extract_rpms(value: str) -> Tuple[int, ...]`
- Produces: `parse_gadgets(value: str) -> Tuple[GadgetItem, ...]`

- [ ] **Step 1: Write failing parser tests**

Create a temporary source workbook with the nine exact headers and rows covering multiple RPMs, missing secondary automatic weapons, primary semiautomatic presence, secondary-shotgun presence, and gadget quantities. Assert:

```python
self.assertEqual(card.primary_rpms, (900, 700))
self.assertEqual(card.secondary_rpms, (1270,))
self.assertTrue(card.has_semiautomatic)
self.assertTrue(card.has_secondary_shotgun)
self.assertEqual(card.gadgets, (GadgetItem("手雷", 2), GadgetItem("烟雾弹", 2)))
```

Also assert missing headers, duplicate operator names, unknown scores, and nonnumeric speed raise `TierChartError` with the offending sheet or operator name.

- [ ] **Step 2: Run parser tests and confirm failure**

Run:

```powershell
python -m unittest tests.test_r6_tier_chart.ParserTests -v
```

Expected: failure because `r6_tier_chart` does not exist.

- [ ] **Step 3: Implement data classes and parser**

Add:

```python
@dataclass(frozen=True)
class GadgetItem:
    name: str
    quantity: Optional[int]

@dataclass(frozen=True)
class OperatorCard:
    side: str
    name: str
    speed: int
    score: int
    tier: str
    primary_rpms: Tuple[int, ...]
    secondary_rpms: Tuple[int, ...]
    has_semiautomatic: bool
    has_secondary_shotgun: bool
    gadgets: Tuple[GadgetItem, ...]
    source_order: int
```

Read sheets by their exact names, index columns by header text, extract RPM from full-width parentheses with `r"（([0-9]+(?:\.[0-9]+)?)）"`, and treat `无`/`无自动枪械` as absent. Normalize `手雷` and `破片手榴弹` to `破片手榴弹`, `烟雾手榴弹` to `烟雾弹`, and `眩晕手榴弹` to `闪光弹`.

- [ ] **Step 4: Run parser tests**

Run the command from Step 2. Expected: all parser tests pass.

### Task 2: Official Gadget Icon Resolver And Cache

**Files:**
- Modify: `r6_tier_chart.py`
- Modify: `tests/test_r6_tier_chart.py`
- Create at runtime: `assets/gadget-icons/*.png`

**Interfaces:**
- Consumes: normalized `GadgetItem.name`
- Produces: `prepare_gadget_icons(items, directory, ...) -> Dict[str, Path]`
- Produces: `resolve_wiki_file_url(file_title, query_json=...) -> str`

- [ ] **Step 1: Write failing icon-map and retry tests**

Assert all normalized gadget names resolve to these exact Wiki files:

```python
GADGET_FILES = {
    "倒刺铁丝网": "文件:R6S gp Barbed wire.png",
    "冲击手榴弹": "文件:R6S gp Impact Grenade.png",
    "感应警报器": "文件:R6S gp Proximity Alarm.png",
    "破片手榴弹": "文件:R6S gp Frag Grenade.png",
    "机动护盾": "文件:R6S gp Deployable Shield.png",
    "烟雾弹": "文件:R6S gp Smoke Grenade.png",
    "爆破炸药": "文件:R6S gp Breach Charge.png",
    "电磁脉冲式冲击弹": "文件:R6S gp Impact emp Grenade.png",
    "闪光弹": "文件:R6S gp Stun Grenade.png",
    "硬突破炸药": "文件:R6S gp SecondaryBreacher.png",
    "观测工具阻拦器": "文件:R6S gp Observation Blocker.png",
    "遥控炸药": "文件:R6S gp Nitro Cell.png",
    "阔剑地雷": "文件:R6S gp Claymore.png",
    "防弹摄像头": "文件:R6S gp Bulletproof camera.png",
}
```

Mock MediaWiki `imageinfo` responses and curl downloads. Assert unknown gadget names, missing `imageinfo`, non-Huiji URLs, invalid image bytes, and retry exhaustion raise `TierChartError`.

- [ ] **Step 2: Run icon tests and confirm failure**

```powershell
python -m unittest tests.test_r6_tier_chart.GadgetIconTests -v
```

Expected: missing resolver and mapping failures.

- [ ] **Step 3: Implement resolver and cache**

Use `https://r6s.huijiwiki.com/api.php` with `action=query`, `prop=imageinfo`, `iiprop=url`, `format=json`, and `formatversion=2`. Accept only URLs beginning with `https://huiji-public.huijistatic.com/r6s/`. Download with system curl, validate using Pillow, and skip valid cached files.

- [ ] **Step 4: Run icon tests**

Run Step 2 command. Expected: all gadget icon tests pass.

### Task 3: Five-Card Horizontal Tier Layout

**Files:**
- Modify: `r6_tier_chart.py`
- Modify: `tests/test_r6_tier_chart.py`

**Interfaces:**
- Consumes: parsed `OperatorCard` groups and prepared icon paths
- Produces: `write_tier_workbook(path, cards, operator_icons, gadget_icons) -> None`

- [ ] **Step 1: Write failing workbook-layout tests**

Build sample data with six S-tier attackers so the sixth card must wrap. Verify:

- Sheet names are exactly `进攻方简图`, `防守方简图`.
- Tier labels appear in score-descending order.
- S tier uses two card rows when it has six operators.
- Each card contains compact text matching `2速 · 主 900/700 · 副 1270` and `主手半自 ✓ · 副喷 ✓`.
- Operator Badge count equals operator count; gadget image count equals parsed gadget-item count.
- Page orientation is landscape, fit-to-width is 1, gridlines are hidden, and print area covers the rendered tier bands.

- [ ] **Step 2: Run layout tests and confirm failure**

```powershell
python -m unittest tests.test_r6_tier_chart.WorkbookLayoutTests -v
```

Expected: missing workbook writer.

- [ ] **Step 3: Implement fixed card geometry**

Use one tier-label column followed by five equal card groups. Each card group uses two worksheet columns: a narrow Badge column and a wider text column. Set fixed row heights, merge the tier label vertically across wrapped rows, anchor a 48x48 Badge in the narrow cell, and place the two compact fact lines in the text cell. Add a dedicated short gadget row under each card row, anchor 16x16 gadget images left to right, and write quantities beside them.

Use tier colors already established by the source workbook:

```python
TIER_COLORS = {
    "S": "E74C3C", "A": "F39C12", "B": "F1C40F",
    "C": "2ECC71", "D": "3498DB", "F": "7F8C8D",
}
```

- [ ] **Step 4: Run layout tests**

Run Step 2 command. Expected: all layout tests pass.

### Task 4: CLI, Documentation, Skill Sync, And Real Workbook

**Files:**
- Modify: `r6_tier_chart.py`
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Modify: `.codex/skills/build-r6-operator-report/SKILL.md`
- Modify: `.codex/skills/build-r6-operator-report/references/source-notes.md`
- Create: `.codex/skills/build-r6-operator-report/scripts/r6_tier_chart.py`
- Create: `r6_operator_tier_chart.xlsx`

**Interfaces:**
- Produces CLI: `python r6_tier_chart.py [--input PATH] [--output PATH] [--icons-dir PATH] [--gadget-icons-dir PATH]`

- [ ] **Step 1: Write failing CLI test**

Invoke `main()` with temporary input/output/icon directories and injected icon preparation. Assert exit code 0, output exists, and stdout reports attacker count, defender count, and absolute output path. Assert missing input returns 1 with a Chinese error message on stderr.

- [ ] **Step 2: Implement CLI and run full test suite**

```powershell
python -m unittest discover -s tests -v
```

Expected: all existing and new tests pass.

- [ ] **Step 3: Update user documentation and project Skill**

Document the tier-chart command, two output sheets, five-card wrapping, compact labels, gadget-icon source, and output filename in README and CHANGELOG. Add the command to the Skill workflow and copy the verified script into the Skill `scripts/` directory. Validate with `quick_validate.py` in UTF-8 mode.

- [ ] **Step 4: Generate and audit the real workbook**

```powershell
python r6_tier_chart.py --input r6_operator_stats.xlsx --output r6_operator_tier_chart.xlsx
```

Reopen with openpyxl and verify two sheets, 77 operator cards, tier order, five-card wrapping, all compact strings, and image counts. Open read-only with Microsoft Excel and confirm both sheets render.

- [ ] **Step 5: Visual verification**

Export the two sheets to PDF with Microsoft Excel COM, render each PDF page to PNG, and inspect both images. Fix clipping, overlap, unreadable icons, excessive whitespace, or broken tier merges before the final run.

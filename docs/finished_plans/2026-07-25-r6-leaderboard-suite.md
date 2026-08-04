# R6 Five-Leaderboard Workbook Suite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the old combined chart outputs with five independent, consistently rendered leaderboard workbooks generated from `r6_operator_stats.xlsx`.

**Architecture:** Keep data acquisition and the nine-column source workbook in `r6_operator_stats.py`, while adding a new configuration-driven `r6_leaderboards.py` for category membership, cross-leaderboard sorting, repeated cards, highlighting, and workbook rendering. Reuse the existing Huiji icon, patch-note, rating, and Excel image helpers instead of duplicating network or source logic.

**Tech Stack:** Python 3.9+, `openpyxl 3.0`, Pillow, standard-library `unittest`, Windows BAT, Microsoft Excel COM for final compatibility and visual verification.

## Global Constraints

- Follow both `<USER_HOME>/.codex/AGENTS.md` and the project `AGENTS.md`.
- Commit every source or documentation change; never commit `AGENTS.md`, generated `.xlsx`, `output/`, `~temp/`, `.tools/`, or `~archived/`.
- Archive instead of deleting old outputs.
- Keep root-file growth small: exactly one new Python entry point and one new BAT launcher.
- New workbook files belong under `output/`.
- The five global dimensions are ordered `video`, `primary_rpm`, `speed`, `rare`, `gadget`.
- Each workbook groups by its own dimension first, then sorts inside each band by the remaining dimensions in global order, then Wiki source order.
- Rare-weapon and gadget workbooks allow repeated operators; the other three do not.
- Primary RPM bands stay fixed at `860 / 780 / 700`.
- Rare-weapon bands are `有副喷`, `有主狙`, `副自动`, `都无`.
- Rare-weapon cards do not receive red highlighting.
- Card rows are name/video tier/speed, secondary shotgun/primary sniper, secondary RPM/primary RPM, then all gadget icons.
- Visible `主手半自` becomes `主狙`; exact weapon alias `P10 RONI转换套件衍生型` becomes `P10 RONI`.

---

### Task 1: Normalize source-workbook labels and weapon names

**Files:**
- Modify: `tests/test_r6_operator_stats.py`
- Modify: `tests/test_r6_tier_chart.py`
- Modify: `r6_operator_stats.py`
- Modify: `r6_tier_chart.py`

**Interfaces:**
- Produces: `normalize_weapon_name(name: str) -> str`.
- Produces source header `主狙`.
- `load_operator_cards(path)` consumes `主狙` while preserving `OperatorCard.has_semiautomatic`.

- [ ] **Step 1: Add failing naming tests**

Add assertions equivalent to:

```python
def test_normalizes_roni_conversion_variant(self):
    self.assertEqual(
        stats.normalize_weapon_name("P10 RONI转换套件衍生型"),
        "P10 RONI",
    )
    self.assertEqual(stats.normalize_weapon_name("Commando 9"), "Commando 9")

def test_workbook_uses_primary_sniper_header(self):
    self.assertIn("主狙", stats.HEADERS)
    self.assertNotIn("主手半自", stats.HEADERS)
```

Update parser fixtures so the sixth source column is `主狙`, then assert formatted card text contains `主狙` and does not contain `主手半自`.

- [ ] **Step 2: Run focused tests and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_operator_stats tests.test_r6_tier_chart -v
```

Expected: failures for missing `normalize_weapon_name` and old `主手半自` header/text.

- [ ] **Step 3: Implement exact alias and visible label migration**

Add:

```python
WEAPON_NAME_ALIASES = {
    "P10 RONI转换套件衍生型": "P10 RONI",
}

def normalize_weapon_name(name: str) -> str:
    return WEAPON_NAME_ALIASES.get(name, name)
```

Apply it when reading `WeaponData.zh_model`, change the sixth source header to `主狙`, change `r6_tier_chart.REQUIRED_HEADERS`, parser lookup, and visible card text to `主狙`.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run the same two-module command. Expected: all tests pass.

- [ ] **Step 5: Commit Task 1**

```powershell
git add r6_operator_stats.py r6_tier_chart.py tests/test_r6_operator_stats.py tests/test_r6_tier_chart.py
git commit -m "feat: normalize RONI and primary sniper labels"
```

### Task 2: Define five leaderboard dimensions and cross-sort rules

**Files:**
- Create: `r6_leaderboards.py`
- Create: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Consumes: `r6_tier_chart.OperatorCard`, `GadgetItem`, `SOURCE_SHEETS`.
- Produces: `LeaderboardSpec`, `LEADERBOARD_SPECS`, `bands_for_card`, `best_dimension_rank`, `sort_cards_for_band`, and `group_cards`.

- [ ] **Step 1: Write failing classification tests**

Create fixtures covering:

```python
spray_sniper_secondary = make_card(
    "Triple",
    primary=(900,),
    secondary=(1270,),
    speed=3,
    has_semiautomatic=True,
    has_secondary_shotgun=True,
    gadgets=("破片手榴弹", "闪光弹"),
)
```

Assert:

```python
self.assertEqual(
    lb.bands_for_card(spray_sniper_secondary, "rare", "进攻方"),
    ("有副喷", "有主狙", "副自动"),
)
self.assertEqual(
    lb.bands_for_card(spray_sniper_secondary, "gadget", "进攻方"),
    ("手雷", "眩晕手榴弹"),
)
self.assertEqual(
    lb.bands_for_card(no_specials, "rare", "进攻方"),
    ("都无",),
)
self.assertEqual(
    lb.bands_for_card(no_target_gadgets, "gadget", "进攻方"),
    ("这些都无",),
)
```

Test primary RPM boundary values `860`, `859`, `780`, `779`, `700`, `699`, and missing primary RPM.

- [ ] **Step 2: Run classification tests and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_leaderboards.LeaderboardClassificationTests -v
```

Expected: import failure because `r6_leaderboards.py` does not exist.

- [ ] **Step 3: Implement immutable specifications and memberships**

Define:

```python
DIMENSION_ORDER = ("video", "primary_rpm", "speed", "rare", "gadget")
VIDEO_BANDS = ("S", "A", "B", "C", "D", "F")
PRIMARY_RPM_BANDS = ("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ")
SPEED_BANDS = ("3速", "2速", "1速")
RARE_BANDS = ("有副喷", "有主狙", "副自动", "都无")
ATTACK_GADGET_BANDS = ("手雷", "眩晕手榴弹", "硬突破炸药", "这些都无")
DEFENSE_GADGET_BANDS = ("遥控炸药", "机动护盾", "冲击手榴弹", "这些都无")
```

Use internal gadget mappings `手雷 -> 破片手榴弹` and `眩晕手榴弹 -> 闪光弹`. Rare and gadget functions return every matched band; their fallback appears only when no target band matches.

- [ ] **Step 4: Write failing cross-sort and repeat tests**

Assert that:

- video groups sort by primary RPM band, then speed, rare best band, gadget best band, then source order;
- primary RPM groups sort by video, speed, rare, gadget, then source order;
- speed groups sort by video, primary RPM, rare, gadget, then source order;
- rare groups repeat cards across every owned band;
- gadget groups repeat cards across every owned target gadget;
- non-repeat groups contain each source operator exactly once.

- [ ] **Step 5: Run cross-sort tests and verify RED**

Run the complete new test module. Expected: failures for missing sorting/grouping behavior.

- [ ] **Step 6: Implement ranking and grouping**

Use:

```python
def sort_cards_for_band(cards, current_dimension, side):
    dimensions = tuple(
        dimension
        for dimension in DIMENSION_ORDER
        if dimension != current_dimension
    )
    return sorted(
        cards,
        key=lambda card: tuple(
            best_dimension_rank(card, dimension, side)
            for dimension in dimensions
        ) + (card.source_order,),
    )
```

`group_cards` expands rare/gadget memberships, validates full coverage, rejects duplicates in non-repeat dimensions, and returns bands in specification order.

- [ ] **Step 7: Run all new logic tests and verify GREEN**

Run:

```powershell
python -m unittest tests.test_r6_leaderboards -v
```

Expected: all classification, sorting, coverage, and repeat tests pass.

- [ ] **Step 8: Commit Task 2**

```powershell
git add r6_leaderboards.py tests/test_r6_leaderboards.py
git commit -m "feat: add five-dimension leaderboard sorting"
```

### Task 3: Render the structured four-row operator card

**Files:**
- Modify: `r6_leaderboards.py`
- Modify: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Produces: `render_side_sheet`, `write_leaderboard_workbook`, and highlighted gadget tokens.
- Reuses: `r6_tier_chart._add_offset_image`, `_excel_image`, `prepare_gadget_icons`, `operator_key`.

- [ ] **Step 1: Write a failing workbook layout test**

Generate a temporary workbook with six cards in one band and assert:

- five cards occupy the first chunk and the sixth wraps inside the same merged left band;
- each card spans one Badge column plus four information columns;
- row 1 contains name, video Tier, speed in that order;
- row 2 contains `副喷` then `主狙`;
- row 3 contains `副` RPM then `主` RPM;
- row 4 contains all gadget images;
- the status row is merged across the complete print width;
- `补丁说明.freeze_panes is None`.

Assert red font/fill behavior:

```python
self.assertEqual(video_s_tier_cell.fill.fgColor.rgb[-6:], TIER_COLORS["S"])
self.assertEqual(primary_i_rpm_cell.font.color.rgb[-6:], "E74C3C")
self.assertEqual(speed_three_cell.font.color.rgb[-6:], "E74C3C")
self.assertNotEqual(rare_spray_cell.font.color.rgb[-6:], "E74C3C")
```

Inspect a highlighted gadget token corner pixel or border pixel to prove the red marker is present only for attacker frag and defender nitro.

- [ ] **Step 2: Run workbook tests and verify RED**

Run:

```powershell
python -m unittest tests.test_r6_leaderboards.LeaderboardWorkbookTests -v
```

Expected: failures because rendering functions do not exist.

- [ ] **Step 3: Implement four-row cards**

Use five columns per card:

```text
Badge | info-1 | info-2 | info-3 | info-4
```

For each four-row card chunk:

```text
row 1: Badge | name merged over info-1:2 | video Tier | speed
row 2: Badge | 副喷 merged over info-1:2 | 主狙 merged over info-3:4
row 3: Badge | 副 RPM merged over info-1:2 | 主 RPM merged over info-3:4
row 4: all gadget icons merged across Badge and all info columns
```

Merge the Badge column across rows 1-3. Use stable row heights `18 / 17 / 17 / 20`, five cards per chunk, and a left band column merged across every chunk in the band.

- [ ] **Step 4: Implement first-band highlighting**

Use the current specification only:

- video: colored Tier cell, including S red;
- primary RPM: primary RPM text red/bold for `Ⅰ`;
- speed: speed text red/bold for `3速`;
- rare: no red highlight;
- gadget: red border around frag/nitro token in every repeated occurrence.

- [ ] **Step 5: Run workbook tests and verify GREEN**

Run the focused workbook test, then the complete new module. Expected: all pass.

- [ ] **Step 6: Commit Task 3**

```powershell
git add r6_leaderboards.py tests/test_r6_leaderboards.py
git commit -m "feat: render four-row leaderboard cards"
```

### Task 4: Generate all five workbooks from one CLI

**Files:**
- Modify: `r6_leaderboards.py`
- Modify: `tests/test_r6_leaderboards.py`
- Create: `run_r6_leaderboards.bat`

**Interfaces:**
- CLI: `python r6_leaderboards.py [--input PATH] [--output-dir PATH] [--icons-dir PATH] [--gadget-icons-dir PATH]`.
- Default output directory: `output`.

- [ ] **Step 1: Write failing suite and CLI tests**

Assert exact filenames:

```python
EXPECTED_OUTPUTS = (
    "r6_video_tier_chart.xlsx",
    "r6_primary_rpm_chart.xlsx",
    "r6_speed_chart.xlsx",
    "r6_rare_weapon_chart.xlsx",
    "r6_gadget_chart.xlsx",
)
```

Each output must contain exactly its descriptive attacker sheet, descriptive defender sheet, and `补丁说明`. Assert CLI prints `39` attackers, `38` defenders, and five resolved output paths.

Test the BAT contains `cd /d "%~dp0"`, Python/py fallback, the root stats input, `--output-dir "%~dp0output"`, preserved exit code, and no deletion command.

- [ ] **Step 2: Run suite/CLI tests and verify RED**

Run the new module. Expected: failures for missing suite writer, CLI, and BAT.

- [ ] **Step 3: Implement suite writer and CLI**

`write_all_leaderboards` prepares gadget icons once, then calls `write_leaderboard_workbook` for each immutable spec. Create `output/` but do not archive or delete from inside the generator.

- [ ] **Step 4: Add the BAT launcher**

The BAT invokes:

```bat
python "%~dp0r6_leaderboards.py" --input "%~dp0r6_operator_stats.xlsx" --output-dir "%~dp0output"
```

with the existing `py -3` fallback and exit-code preservation pattern.

- [ ] **Step 5: Run new module and full suite**

Run:

```powershell
python -m unittest tests.test_r6_leaderboards -v
python -m unittest discover -s tests -v
```

Expected: all tests pass.

- [ ] **Step 6: Commit Task 4**

```powershell
git add r6_leaderboards.py run_r6_leaderboards.bat tests/test_r6_leaderboards.py
git commit -m "feat: generate five leaderboard workbooks"
```

### Task 5: Update documentation and project Skill

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Modify: `.codex/skills/build-r6-operator-report/SKILL.md`
- Modify: `.codex/skills/build-r6-operator-report/scripts/r6_operator_stats.py`
- Modify: `.codex/skills/build-r6-operator-report/scripts/r6_tier_chart.py`
- Create: `.codex/skills/build-r6-operator-report/scripts/r6_leaderboards.py`
- Create: `.codex/skills/build-r6-operator-report/scripts/run_r6_leaderboards.bat`

**Interfaces:**
- Documents the five outputs, global order, repeated categories, card rows, fixed primary thresholds, archive location, and CLI.

- [ ] **Step 1: Update README and CHANGELOG**

Replace the old “primary output” guidance with the five-workbook suite while preserving legacy script documentation as compatibility notes. State that output files are generated, ignored, and stored under `output/`.

- [ ] **Step 2: Update Skill workflow**

Require exact five output filenames, exact sheet sets, repeated rare/gadget coverage, red-marker rules, `主狙`, P10 RONI alias, data-status rows, patch notes, and Excel visual QA.

- [ ] **Step 3: Synchronize scripts**

Copy the completed root scripts and BAT into the project Skill. Verify root/Skill SHA-256 equality for every synchronized file.

- [ ] **Step 4: Validate Skill and tests**

Run:

```powershell
$env:PYTHONUTF8='1'
python <USER_HOME>\.codex\skills\.system\skill-creator\scripts\quick_validate.py .codex\skills\build-r6-operator-report
python -m unittest discover -s tests -v
```

Expected: `Skill is valid!` and all tests pass.

- [ ] **Step 5: Commit Task 5**

```powershell
git add README.md CHANGELOG.md .codex/skills/build-r6-operator-report
git commit -m "docs: document five leaderboard workflow"
```

### Task 6: Archive old outputs, regenerate data, and verify Excel

**Files:**
- Generated/ignored: `r6_operator_stats.xlsx`
- Archived/ignored: `~archived/2026-07-25-before-leaderboards/*.xlsx`
- Generated/ignored: `output/*.xlsx`

**Interfaces:**
- Produces final user-facing files without changing tracked source state.

- [ ] **Step 1: Verify source tree before generation**

Run:

```powershell
python -m unittest discover -s tests -v
git status --short
```

Expected: all tests pass and tracked working tree is clean.

- [ ] **Step 2: Regenerate `r6_operator_stats.xlsx`**

Run:

```powershell
python r6_operator_stats.py --output r6_operator_stats.xlsx
```

Reopen it and verify both side headers contain `主狙`, no cell contains `P10 RONI转换套件衍生型`, and the two expected RONI entries display `P10 RONI`.

- [ ] **Step 3: Archive old chart workbooks**

Resolve every source and destination path and verify they remain under the project root before moving:

```text
r6_operator_tier_chart.xlsx
r6_operator_rpm_chart.xlsx
r6_operator_rpm_chart_updated.xlsx
```

Move them into `~archived/2026-07-25-before-leaderboards/`. If Excel locks a file, copy it to the archive, leave the locked original untouched, and report the lock instead of terminating Excel.

- [ ] **Step 4: Generate all five outputs**

Run:

```powershell
python r6_leaderboards.py --input r6_operator_stats.xlsx --output-dir output
```

Expected: 39 attackers, 38 defenders, and five output paths.

- [ ] **Step 5: Audit workbook structure and memberships**

For every file, reopen with `openpyxl` and verify exact sheet names, freeze panes, print areas, status rows, patch sheet, image counts, and band labels. Confirm unique coverage for video/primary/speed and repeated coverage for rare/gadget.

- [ ] **Step 6: Open and render all five files with Microsoft Excel**

Use hidden Excel COM read-only mode. Export all 15 sheets to PDF, render to PNG with local PyMuPDF, and inspect every page. Check:

- cards show the required four-row field order;
- no text, Badge, Tier cell, speed, or gadget icon overlaps;
- five cards wrap correctly;
- repeated cards appear in every matching rare/gadget band;
- only permitted first-band elements are red;
- rare cards have no red feature highlight;
- patch notes are legible and unfrozen.

- [ ] **Step 7: Final verification**

Run:

```powershell
python -m unittest discover -s tests -v
$env:PYTHONUTF8='1'
python <USER_HOME>\.codex\skills\.system\skill-creator\scripts\quick_validate.py .codex\skills\build-r6-operator-report
git status --short
```

Expected: tests and Skill pass; only ignored generated/archived workbooks exist outside tracked status.

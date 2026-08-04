# R6 Data Pipeline And Workbook Refresh Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single-entry R6 report pipeline in which the project Skill supplies the latest verified Athieno ratings, the script snapshots current Huiji Wiki data/icons/patches under `data/`, and offline generators create the source workbook plus five Chinese-named leaderboard workbooks.

**Architecture:** Move runtime code into `src/r6_report`, separate network collection from workbook generation, and define validated JSON contracts for rating, Wiki, and patch snapshots. The collector uses Huiji's MediaWiki API and atomically replaces complete snapshots; workbook modules only read local `data/`. Shared source metadata and footer rendering keep every worksheet auditable, while leaderboard rendering uses one bounded gadget-strip image per card.

**Tech Stack:** Python 3.9+, standard library `argparse/json/html.parser/pathlib/subprocess/tempfile/unittest`, `openpyxl>=3.0,<4`, `Pillow>=9,<12`, Windows BAT, Huiji MediaWiki read-only API.

## Global Constraints

- Root contains exactly one BAT: `run_r6_report.bat`.
- Root contains no Python source files or XLSX files.
- Runtime inputs and snapshots live under ignored `data/`; generated leaderboards live under ignored `output/`.
- Old files are archived under ignored `~archived/`, never deleted.
- Automatic weapons remain `firerate > 0` and `projectile == 1`.
- Primary semiautomatic remains `equipment == 1`, `firerate == 0`, and `projectile == 1`.
- Secondary shotgun remains `equipment == 2` and `projectile > 1`.
- `P10 RONI转换套件衍生型` displays as `P10 RONI`.
- Video score mapping remains `S=100, A=85, B=70, C=55, D=40, F=20, boof=0`.
- Leaderboard grouping, repeated rare/gadget memberships, and secondary sort priorities remain unchanged.
- Every repository modification is committed before beginning the next task.

---

## File Structure

Create or move these runtime modules:

- `src/r6_report/__init__.py`: package marker and version.
- `src/r6_report/sources.py`: rating/Wiki/patch source dataclasses, JSON validation, interval validation.
- `src/r6_report/wiki_client.py`: MediaWiki query/parse requests and Tabx parsing.
- `src/r6_report/patch_catalog.py`: patch-index HTML parsing, patch-page wikitext parsing, direction classification.
- `src/r6_report/collector.py`: snapshot orchestration, temporary staging, archive-and-replace, CLI.
- `src/r6_report/operator_stats.py`: operator joining and offline source-workbook generation.
- `src/r6_report/workbook_sources.py`: shared three-row worksheet source footer.
- `src/r6_report/patch_notes.py`: date-grouped patch worksheet rendering.
- `src/r6_report/tier_chart.py`: card loading, icon handling, legacy compact chart.
- `src/r6_report/leaderboards.py`: five leaderboard grouping and rendering.
- `src/r6_report/tiers.py`: tier normalization and colours.

Create or update tests:

- `tests/_path_setup.py`
- `tests/test_sources.py`
- `tests/test_collector.py`
- `tests/test_patch_catalog.py`
- `tests/test_r6_operator_stats.py`
- `tests/test_r6_patch_notes.py`
- `tests/test_r6_tier_chart.py`
- `tests/test_r6_leaderboards.py`
- `tests/test_project_layout.py`

Update project workflow:

- `run_r6_report.bat`
- `.codex/skills/build-r6-operator-report/SKILL.md`
- `.codex/skills/build-r6-operator-report/references/rating-contract.md`
- `README.md`
- `CHANGELOG.md`

## Task 1: Move Runtime Code Into A Package

**Files:**
- Create: `src/r6_report/__init__.py`
- Move: `r6_operator_stats.py` to `src/r6_report/operator_stats.py`
- Move: `r6_leaderboards.py` to `src/r6_report/leaderboards.py`
- Move: `r6_patch_notes.py` to `src/r6_report/patch_notes.py`
- Move: `r6_tier_chart.py` to `src/r6_report/tier_chart.py`
- Move: `r6_tiers.py` to `src/r6_report/tiers.py`
- Create: `tests/_path_setup.py`
- Modify: all `tests/test_*.py`

**Interfaces:**
- Consumes: existing public functions and dataclasses without behavior changes.
- Produces: importable modules under `r6_report`; commands run with `python -m r6_report.<module>`.

- [ ] **Step 1: Add a failing package import test**

```python
import _path_setup

class PackageLayoutTests(unittest.TestCase):
    def test_runtime_modules_are_importable_from_r6_report(self):
        from r6_report import leaderboards, operator_stats, patch_notes, tier_chart, tiers
        self.assertTrue(callable(operator_stats.main))
        self.assertTrue(callable(leaderboards.main))
```

- [ ] **Step 2: Run the package test and verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_project_layout.PackageLayoutTests.test_runtime_modules_are_importable_from_r6_report -v
```

Expected: `ModuleNotFoundError: No module named 'r6_report'`.

- [ ] **Step 3: Move modules and make imports relative**

Create:

```python
"""Rainbow Six Siege operator report package."""

__version__ = "1.0.0"
```

Use `tests/_path_setup.py`:

```python
from pathlib import Path
import sys

SRC = Path(__file__).resolve().parents[1] / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
```

Change package imports to `.patch_notes`, `.tier_chart`, and `.tiers`. Import `_path_setup` before `r6_report` in every test.

- [ ] **Step 4: Run all existing tests**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest discover -s tests -v
```

Expected: all existing tests plus the package test pass.

- [ ] **Step 5: Commit**

```text
git add src tests
git commit -m "refactor: move R6 runtime into package"
```

## Task 2: Define And Validate Source Snapshot Contracts

**Files:**
- Create: `src/r6_report/sources.py`
- Create: `tests/test_sources.py`
- Modify: `data/athieno_y11s2.json` before moving it to ignored `data/athieno/latest.json`

**Interfaces:**
- Produces: `RatingSource`, `WikiManifest`, `PatchSource`, `PatchChange`, `PatchRecord`, `ReportSources`.
- Produces: `load_rating_document(path)`, `load_wiki_manifest(path)`, `load_patch_document(path)`, `load_report_sources(data_dir)`.
- Consumes later: collector, source workbook, patch sheet, leaderboards, source footers.

- [ ] **Step 1: Write failing validation tests**

```python
def test_loads_complete_rating_source_and_interval(self):
    document = rating_document(
        season="Y11S2",
        covered_patch="Y11S2",
        covered_through="2026-06-02",
        captured_at="2026-07-25T10:00:00+08:00",
    )
    source, tiers, scores = sources.parse_rating_document(document)
    self.assertEqual(source.covered_patch, "Y11S2")
    self.assertEqual(scores["solid-snake"], 100)

def test_rejects_missing_timezone_and_patch_outside_interval(self):
    with self.assertRaisesRegex(sources.SourceDataError, "timezone"):
        sources.parse_iso_datetime("2026-07-25T10:00:00")
    with self.assertRaisesRegex(sources.SourceDataError, "outside"):
        sources.validate_patch_interval(rating, wiki, (future_patch,))
```

- [ ] **Step 2: Verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_sources -v
```

Expected: import failure for `r6_report.sources`.

- [ ] **Step 3: Implement immutable models and strict loaders**

Use dataclasses with these signatures:

```python
@dataclass(frozen=True)
class RatingSource:
    creator: str
    title: str
    url: str
    video_id: str
    published: date
    season: str
    covered_patch: str
    covered_through: date
    coverage_basis: str
    final_frame: str
    captured_at: datetime

@dataclass(frozen=True)
class WikiManifest:
    season: str
    season_name: str
    patch: str
    fetched_at: datetime
    sources: Mapping[str, str]
    counts: Mapping[str, int]

@dataclass(frozen=True)
class ReportSources:
    rating: RatingSource
    wiki: WikiManifest
    patches: Tuple[PatchRecord, ...]
    patch_index_url: str
```

Require HTTPS URLs, aware ISO datetimes, unique rating names, known score-map keys, and patch dates satisfying `covered_through < released <= fetched_at.date()`.

- [ ] **Step 4: Verify GREEN**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_sources -v
```

Expected: all source-contract tests pass.

- [ ] **Step 5: Commit**

```text
git add src/r6_report/sources.py tests/test_sources.py
git commit -m "feat: validate report source snapshots"
```

## Task 3: Parse Dynamic Huiji Patch Data

**Files:**
- Create: `src/r6_report/patch_catalog.py`
- Create: `tests/test_patch_catalog.py`

**Interfaces:**
- Consumes: MediaWiki `action=parse&prop=text` output for `更新补丁总表`.
- Consumes: MediaWiki revision wikitext for each selected `YxSy.z更新补丁`.
- Produces: `parse_patch_index_html(html) -> Tuple[PatchIndexEntry, ...]`.
- Produces: `select_patch_interval(entries, lower, upper) -> Tuple[PatchIndexEntry, ...]`.
- Produces: `parse_patch_wikitext(entry, wikitext, operator_names) -> PatchRecord`.

- [ ] **Step 1: Add failing parser tests with real compact fixtures**

```python
INDEX_HTML = """
<table><tr><th>所属赛季</th><th>补丁版本</th><th>推送日期</th></tr>
<tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.2更新补丁">Y11S2.2更新补丁</a></td><td>2026-07-14</td></tr>
<tr><td>系统覆盖行动</td><td><a href="/wiki/Y11S2.1更新补丁">Y11S2.1更新补丁</a></td><td>2026-06-23</td></tr>
</table>
"""

def test_parses_and_orders_patch_index(self):
    entries = patch_catalog.parse_patch_index_html(INDEX_HTML)
    self.assertEqual([entry.patch for entry in entries], ["Y11S2.1", "Y11S2.2"])

def test_extracts_operator_change_and_official_source(self):
    record = patch_catalog.parse_patch_wikitext(
        entry,
        "{{Infobox patch|来源=[https://www.ubisoft.com/x Ubisoft]}}"
        "{{干员改动|WAMAI|最大总充能数提高至 7 个（原为 6 个）。}}",
        {"Wamai"},
    )
    self.assertEqual(record.changes[0].subject, "Wamai")
    self.assertEqual(record.changes[0].direction, "增强")
```

- [ ] **Step 2: Verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_patch_catalog -v
```

Expected: import failure for `r6_report.patch_catalog`.

- [ ] **Step 3: Implement parsers**

Use `html.parser.HTMLParser` for the three-column index. Extract the balanced `{{干员改动 ... }}` template by counting nested `{{`/`}}`, split only top-level pipes, normalize Wiki links/templates to display text, and map uppercase operator keys to current Wiki names.

Direction rules:

```python
CURRENT_OVERRIDES = {
    ("Y11S2.1", "Dokkaebi"): "削弱",
    ("Y11S2.1", "Solid Snake"): "混合",
    ("Y11S2.2", "Dokkaebi"): "混合",
    ("Y11S2.2", "Jäger"): "混合",
}
```

Classify other details as `混合` when both positive and negative cues occur, `削弱` for cooldown extension/range or quantity reduction/removal, and `增强` for cooldown reduction/range or quantity increase/recoil reduction. Unknown text uses `混合` rather than being omitted.

- [ ] **Step 4: Verify GREEN with interval boundary tests**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_patch_catalog -v
```

Expected: index, interval, nested-template, no-change, direction, and invalid-date tests all pass.

- [ ] **Step 5: Commit**

```text
git add src/r6_report/patch_catalog.py tests/test_patch_catalog.py
git commit -m "feat: parse Huiji patch catalog"
```

## Task 4: Collect Atomic Wiki, Icon, And Patch Snapshots

**Files:**
- Create: `src/r6_report/wiki_client.py`
- Create: `src/r6_report/collector.py`
- Create: `tests/test_collector.py`
- Modify: `src/r6_report/operator_stats.py`
- Modify: `src/r6_report/tier_chart.py`

**Interfaces:**
- Produces: `collect_snapshot(data_dir, archive_dir, temp_dir, now, client) -> WikiManifest`.
- Produces CLI: `python -m r6_report.collector --data-dir data`.
- Consumes: `data/athieno/latest.json` to choose the patch interval.
- Produces: validated JSON under `data/wiki/`, images under `data/icons/`, and patch JSON under `data/patches/`.

- [ ] **Step 1: Write failing collector tests**

```python
def test_stages_complete_snapshot_before_replacing_active_data(self):
    manifest = collector.collect_snapshot(
        data_dir,
        archive_dir,
        temp_dir,
        aware_now,
        client=FakeHuijiClient.complete(),
    )
    self.assertEqual(manifest.patch, "Y11S2.2")
    self.assertTrue((data_dir / "wiki" / "operator.json").is_file())
    self.assertTrue((data_dir / "icons" / "operator" / "badge" / "ace.png").is_file())

def test_failed_icon_validation_keeps_previous_snapshot(self):
    old = write_complete_snapshot(data_dir)
    with self.assertRaisesRegex(collector.CollectionError, "badge"):
        collector.collect_snapshot(
            data_dir=data_dir,
            archive_dir=archive_dir,
            temp_dir=temp_dir,
            now=aware_now,
            client=FakeHuijiClient.missing_badge(),
        )
    self.assertEqual((data_dir / "wiki" / "manifest.json").read_bytes(), old)
```

- [ ] **Step 2: Verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_collector -v
```

Expected: import failure for `r6_report.collector`.

- [ ] **Step 3: Extract reusable network helpers and implement collection**

`HuijiClient` exposes:

```python
def fetch_tabx(self, title: str, fields: Sequence[str]) -> List[Dict[str, object]]
def fetch_parsed_html(self, title: str) -> str
def fetch_wikitext(self, title: str) -> str
def download_image(self, url: str, destination: Path) -> None
```

Write JSON with `ensure_ascii=False`, `indent=2`, and trailing newline. Validate image MIME bytes with Pillow. Stage all files under `~temp/r6-report-<uuid>`, archive any previous active directories to `~archived/data-snapshots/<kind>/<timestamp>/`, then replace complete directories with `Path.replace`.

- [ ] **Step 4: Verify GREEN and retry behavior**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_collector tests.test_r6_operator_stats tests.test_r6_tier_chart -v
```

Expected: collection, retry, empty response, 403, invalid JSON, atomic rollback, and existing parsing tests pass.

- [ ] **Step 5: Commit**

```text
git add src/r6_report/wiki_client.py src/r6_report/collector.py src/r6_report/operator_stats.py src/r6_report/tier_chart.py tests
git commit -m "feat: collect atomic R6 data snapshots"
```

## Task 5: Render Dynamic Patch Groups And Source Footers

**Files:**
- Create: `src/r6_report/workbook_sources.py`
- Modify: `src/r6_report/patch_notes.py`
- Modify: `src/r6_report/operator_stats.py`
- Modify: `tests/test_r6_patch_notes.py`
- Modify: `tests/test_r6_operator_stats.py`

**Interfaces:**
- Produces: `append_source_footer(sheet, last_column, report_sources) -> Tuple[int, int]`.
- Produces: `add_patch_notes_sheet(workbook, scores, report_sources) -> Worksheet`.
- Consumes: `ReportSources` from Task 2.

- [ ] **Step 1: Write failing workbook tests**

```python
def test_footer_contains_three_source_rows_with_links_and_versions(self):
    start, end = workbook_sources.append_source_footer(sheet, 9, report_sources())
    self.assertEqual(sheet.cell(start, 1).value.split("：", 1)[0], "评分来源")
    self.assertEqual(sheet.cell(start + 1, 1).value.split("：", 1)[0], "游戏数据")
    self.assertEqual(sheet.cell(end, 1).value.split("：", 1)[0], "补丁区间")
    self.assertTrue(sheet.cell(start, 1).hyperlink.target.startswith("https://"))

def test_patch_sheet_groups_oldest_to_newest_and_has_footer(self):
    sheet = patch_notes.add_patch_notes_sheet(workbook, scores, report_sources())
    headings = [cell.value for cell in sheet["A"] if isinstance(cell.value, str)]
    self.assertLess(headings.index("Y11S2.1 · 2026-06-23"), headings.index("Y11S2.2 · 2026-07-14"))
    self.assertIn("评分来源", headings[-3])
```

- [ ] **Step 2: Verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_r6_patch_notes tests.test_r6_operator_stats -v
```

Expected: missing `workbook_sources` and old patch-sheet signature failures.

- [ ] **Step 3: Implement footer and grouped patch renderer**

Replace `DATA_STATUS_TEXT` with three merged rows. Use aware ISO strings, shallow gray fills, hyperlink cells, and exclude footer rows from table filters while including them in print areas. Patch groups use a merged blue date/version heading, one source-link row, then direction/detail rows. Render “无影响本报告字段的变更” when a patch has no changes.

- [ ] **Step 4: Make source-workbook generation offline**

Change `operator_stats.main` to:

```python
rows = load_operator_snapshot(args.data_dir)
report_sources = load_report_sources(args.data_dir)
ratings = load_ratings(args.ratings, operator_names)
write_workbook(output, rows, ratings, badge_dir, report_sources)
```

No workbook-writing function may call `subprocess`, `curl`, or a network helper.

- [ ] **Step 5: Verify GREEN**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_r6_patch_notes tests.test_r6_operator_stats -v
```

Expected: all footer, patch grouping, offline generation, table-range, and source-workbook tests pass.

- [ ] **Step 6: Commit**

```text
git add src/r6_report/workbook_sources.py src/r6_report/patch_notes.py src/r6_report/operator_stats.py tests
git commit -m "feat: add dynamic workbook sources and patches"
```

## Task 6: Fix Card Geometry, Gadget Overflow, Labels, And Filenames

**Files:**
- Modify: `src/r6_report/leaderboards.py`
- Modify: `src/r6_report/tier_chart.py`
- Modify: `tests/test_r6_leaderboards.py`
- Modify: `tests/test_r6_tier_chart.py`

**Interfaces:**
- Produces: `draw_gadget_strip(gadgets, icon_paths, destination, highlighted_name=None) -> Tuple[int, int]`.
- Produces exact output names: `视频评分榜.xlsx`, `主武器射速榜.xlsx`, `速度榜.xlsx`, `稀有枪械榜.xlsx`, `次要装备榜.xlsx`.

- [ ] **Step 1: Add failing regression tests**

```python
def test_seven_gadgets_render_as_one_bounded_strip(self):
    size = lb.draw_gadget_strip(seven_gadgets, icons, output)
    self.assertEqual(size, (168, 20))
    with Image.open(output) as image:
        self.assertEqual(image.size, (168, 20))

def test_cards_have_blank_gutters_and_badges_shift_toward_own_text(self):
    workbook = render_fixture()
    sheet = workbook["进攻方视频Tier榜"]
    self.assertIsNone(sheet["G3"].fill.fill_type)
    badge_anchor = first_badge_anchor(sheet)
    self.assertGreaterEqual(badge_anchor._from.colOff, pixels_to_EMU(7))

def test_rare_labels_and_chinese_output_names(self):
    self.assertEqual(lb.RARE_BANDS, ("副喷", "主狙", "副自", "都无"))
    self.assertEqual(lb.EXPECTED_OUTPUTS, ("视频评分榜.xlsx", "主武器射速榜.xlsx", "速度榜.xlsx", "稀有枪械榜.xlsx", "次要装备榜.xlsx"))
```

- [ ] **Step 2: Verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_r6_leaderboards tests.test_r6_tier_chart -v
```

Expected: missing strip function, old labels/names, and no gutter assertions fail.

- [ ] **Step 3: Implement one image strip per card**

Render each 24×20 token into one transparent strip, preserving order and quantity. Anchor the strip once with a local offset smaller than its starting column width. Keep per-token red borders inside the strip for the top gadget in the gadget leaderboard.

- [ ] **Step 4: Implement card gutters and badge alignment**

Use five card blocks separated by four narrow blank columns. Set badge columns to 7.5, info columns to widths that preserve current text, gutter columns to 1.5, and Badge x-offset to at least 7 px. Update merges, `last_column`, title merge, print area, and freeze pane calculations.

- [ ] **Step 5: Apply labels, names, and footers**

Use exact rare labels and Chinese filenames. Pass `ReportSources` to every leaderboard and append the three-row footer to both side sheets and `补丁说明`.

- [ ] **Step 6: Verify GREEN**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_r6_leaderboards tests.test_r6_tier_chart -v
```

Expected: all grouping, sorting, repetition, highlight, seven-gadget, geometry, source-footer, and output-name tests pass.

- [ ] **Step 7: Commit**

```text
git add src/r6_report/leaderboards.py src/r6_report/tier_chart.py tests
git commit -m "fix: contain gadget icons and space leaderboard cards"
```

## Task 7: Add The Single BAT And Update The Project Skill

**Files:**
- Create: `run_r6_report.bat`
- Archive: `run_r6_operator_stats.bat`, `run_r6_leaderboards.bat`
- Modify: `.codex/skills/build-r6-operator-report/SKILL.md`
- Create: `.codex/skills/build-r6-operator-report/references/rating-contract.md`
- Archive: `.codex/skills/build-r6-operator-report/scripts/`
- Archive: `.codex/skills/build-r6-operator-report/references/athieno_y11s2.json`
- Create: `tests/test_project_layout.py`

**Interfaces:**
- Produces one user entry point that validates ratings, collects snapshots, builds source workbook, then builds five leaderboards.
- Produces a Skill workflow that writes `data/athieno/latest.json`.

- [ ] **Step 1: Write failing layout and launcher tests**

```python
def test_root_has_exactly_one_bat(self):
    self.assertEqual([path.name for path in ROOT.glob("*.bat")], ["run_r6_report.bat"])

def test_launcher_stops_after_each_failed_stage(self):
    text = (ROOT / "run_r6_report.bat").read_text(encoding="utf-8")
    self.assertLess(text.index("r6_report.collector"), text.index("r6_report.operator_stats"))
    self.assertLess(text.index("r6_report.operator_stats"), text.index("r6_report.leaderboards"))
    self.assertGreaterEqual(text.count("if errorlevel 1"), 4)
```

- [ ] **Step 2: Verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_project_layout -v
```

Expected: two BAT files and missing single launcher.

- [ ] **Step 3: Implement `run_r6_report.bat`**

The BAT sets `PYTHONPATH=%~dp0src`, selects `python` or `py -3`, validates `data\athieno\latest.json`, and runs collector → operator stats → leaderboards. Capture and return each failing `%ERRORLEVEL%`; never run the next command after failure.

- [ ] **Step 4: Rewrite the Skill around latest-video acquisition**

The Skill must direct the Agent to find Athieno's official latest current-season tier video, inspect the final list frame, record complete source metadata including `covered_patch` and `covered_through`, validate full unique operator coverage, archive the previous rating snapshot, and atomically write `data/athieno/latest.json`. It then tells the user to run `run_r6_report.bat`.

- [ ] **Step 5: Verify GREEN and validate the Skill**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_project_layout -v
$env:PYTHONUTF8='1'
python <USER_HOME>\.codex\skills\.system\skill-creator\scripts\quick_validate.py .codex\skills\build-r6-operator-report
```

Expected: layout tests pass and Skill validator reports success.

- [ ] **Step 6: Commit**

```text
git add run_r6_report.bat .codex tests/test_project_layout.py
git commit -m "feat: add single R6 report workflow"
```

## Task 8: Archive Root Files And Update Documentation

**Files:**
- Move: root Python files already represented in `src/`
- Move: root `r6_operator_stats.xlsx` to `data/r6_operator_stats.xlsx`
- Move: root `assets/` to initial `data/icons/`
- Archive: obsolete BATs, Skill script copies, English-named output workbooks, and released `.tools/`
- Modify: `.gitignore`
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Produces the final minimal root and documented clone/run/test workflow.

- [ ] **Step 1: Verify every move target remains inside the project**

Resolve each source and destination with `Resolve-Path`/`[IO.Path]::GetFullPath`; require both paths to begin with `<PROJECT_ROOT>\`. Do not move locked `.tools` content until Windows releases it.

- [ ] **Step 2: Move runtime data and archive obsolete files**

Keep `.gitignore` entries:

```text
~archived/
~temp/
data/
output/
```

Move old launchers, duplicated Skill scripts, fixed rating references, English workbooks, and any released `.tools` content into descriptive subdirectories under `~archived/2026-07-25-root-cleanup/`.

- [ ] **Step 3: Update README and CHANGELOG**

README documents:

```text
1. 使用 build-r6-operator-report Skill 生成 data/athieno/latest.json
2. 双击 run_r6_report.bat
3. 在 output/ 查看五个中文工作簿
```

Also document Python 3.9+, `pip install -r requirements.txt`, module commands, snapshot directories, failure behavior, and the test command. CHANGELOG records the package move, data contracts, patch interval, source footers, gadget fix, card spacing, Chinese names, and single BAT.

- [ ] **Step 4: Run structural tests**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_project_layout -v
```

Expected: one BAT, no root Python/XLSX, required docs/directories present.

- [ ] **Step 5: Commit**

```text
git add .gitignore README.md CHANGELOG.md src tests run_r6_report.bat .codex
git commit -m "chore: finish minimal R6 project layout"
```

## Task 9: Full Verification And Real Data Refresh

**Files:**
- Generated ignored data: `data/`
- Generated ignored outputs: `output/`
- Temporary visual previews: `~temp/qa/`

**Interfaces:**
- Verifies the complete user workflow without adding generated artifacts to Git.

- [ ] **Step 1: Run the complete unit suite**

```powershell
$env:PYTHONPATH='src'
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest discover -s tests -v
```

Expected: all tests pass with no warnings or tracebacks.

- [ ] **Step 2: Run the real collector**

```powershell
$env:PYTHONPATH='src'
python -m r6_report.collector --data-dir data
```

Expected: current Wiki manifest identifies `Y11S2.2` or a newer published patch, all three Tabx snapshots validate, and complete current icons exist.

- [ ] **Step 3: Generate the source workbook and leaderboards**

```powershell
python -m r6_report.operator_stats --data-dir data --ratings data/athieno/latest.json --output data/r6_operator_stats.xlsx
python -m r6_report.leaderboards --data-dir data --input data/r6_operator_stats.xlsx --output-dir output
```

Expected: source workbook plus exactly five Chinese-named workbooks.

- [ ] **Step 4: Perform structural workbook verification**

Reopen every workbook with `openpyxl`, verify expected sheet names, source footer labels/links, patch date ordering, no formulas errors, image counts, seven-gadget strips, print areas, and unfrozen `补丁说明`.

- [ ] **Step 5: Render every sheet with `artifact-tool`**

Render all 18 worksheets from the source workbook and five leaderboards into ignored `~temp/qa/`. Confirm no clipped text, invalid icon anchors, gadget overlap, card-boundary overlap, missing source footers, or bad patch grouping.

- [ ] **Step 6: Confirm Git cleanliness**

```text
git status --short
```

Expected: clean working tree; ignored generated data/output/previews do not appear.

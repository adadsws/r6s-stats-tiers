# 榜单卡片枪械图标 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在五张榜单的副喷、主狙、副手自动枪和主手自动枪字段右侧显示全部对应枪械图标，同时保持现有固定行高和排序行为。

**Status:** 已于 2026-08-04 完成；105 项测试通过，并验证 75 个枪械图标及 5 组 XLSX/PDF/PNG 输出。

**Architecture:** `tier_chart.py` 负责把统计表四个武器字段解析为不可变 `WeaponItem`，并继续填充旧 RPM／布尔接口；采集器为实际使用的枪械准备经过裁边和深色轮廓化的本地图标。`leaderboards.py` 与 `pdf_leaderboards.py` 只消费同一组 `WeaponItem` 和图标映射，分别在固定高度区域中排列图片。

**Tech Stack:** Python 3.9+、标准库 `unittest`、Pillow、openpyxl、ReportLab、灰机 Wiki MediaWiki API。

## Global Constraints

- 副喷与主狙有多把符合条件枪械时全部显示图标；无枪时保持 `-`、灰底和零图标。
- 自动枪显示全部射速及全部对应图标；射速降序，相同射速保持 Wiki 来源顺序。
- XLSX `theme.XLSX_CARD_BODY_ROW_PT` 和 PDF `theme.PDF_CARD_BODY_ROW_MM` 不变。
- 图标只能来自已验证的 `data/icons/weapon/`；渲染阶段不得联网。
- 卡片四行结构、阶级排序、输出文件名和 WebUI `api_version = 1` 不变。
- `output/` 与 `~temp/` 不纳入 Git；其余本次生成内容按仓库规则提交。
- 本计划所有 PowerShell 测试步骤先执行 `$workspacePython='<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'`，使用已安装 Pillow、openpyxl、ReportLab、pypdf 的工作区 Python 3.12。

---

### Task 1: 保留具体武器数据并兼容旧排序接口

**Files:**
- Modify: `src/r6_report/tier_chart.py:95-230`
- Modify: `tests/test_r6_tier_chart.py:70-125`
- Modify: `tests/test_r6_leaderboards.py:20-55`

**Interfaces:**
- Produces: `WeaponItem(name: str, icon_key: str, firerate: Optional[int])`。
- Produces: `weapon_icon_key(name: str) -> str`、`parse_automatic_weapons(value: object) -> Tuple[WeaponItem, ...]`、`parse_named_weapons(value: object) -> Tuple[WeaponItem, ...]`。
- Preserves: `OperatorCard.primary_rpms`、`secondary_rpms`、`has_semiautomatic`、`has_secondary_shotgun`。

- [ ] **Step 1: 写具体武器解析的失败测试**

在 `ParserTests` 增加：

```python
def test_parses_all_named_weapons_and_matching_rates(self):
    automatic = chart.parse_automatic_weapons(
        "G8A1（850）\nAUG A2（720）\n552 Commando（690）"
    )
    named = chart.parse_named_weapons("Mk 14 EBR\nBOSG.12.2")

    self.assertEqual(
        [(item.name, item.firerate) for item in automatic],
        [("G8A1", 850), ("AUG A2", 720), ("552 Commando", 690)],
    )
    self.assertEqual(
        [(item.name, item.firerate) for item in named],
        [("Mk 14 EBR", None), ("BOSG.12.2", None)],
    )
    self.assertEqual(chart.parse_automatic_weapons("无自动枪械"), ())
    self.assertEqual(chart.parse_named_weapons("无"), ())
```

- [ ] **Step 2: 运行测试并确认缺少新接口而失败**

Run:

```powershell
$env:PYTHONPATH='src;tests'
$env:PYTHONUTF8='1'
& $workspacePython -m unittest tests.test_r6_tier_chart.ParserTests.test_parses_all_named_weapons_and_matching_rates -v
```

Expected: ERROR，`tier_chart` 没有 `parse_automatic_weapons`。

- [ ] **Step 3: 实现最小数据模型和解析器**

在 `GadgetItem` 前加入：

```python
@dataclass(frozen=True)
class WeaponItem:
    name: str
    icon_key: str
    firerate: Optional[int] = None


def weapon_icon_key(name: str) -> str:
    normalized = normalize_weapon_name(name)
    return operator_key(normalized)
```

把自动枪行按 `名称（整数）` 逐行解析；`无自动枪械` 返回空元组。具名分类按非空行解析；`无` 返回空元组。每个 `WeaponItem.icon_key` 使用 `weapon_icon_key()` 生成。`extract_rpms()` 改为：

```python
return tuple(
    item.firerate
    for item in parse_automatic_weapons(value)
    if item.firerate is not None
)
```

- [ ] **Step 4: 扩展 OperatorCard 并让工作簿加载器填充四组武器**

在 `OperatorCard.source_order` 后增加默认字段：

```python
primary_weapons: Tuple[WeaponItem, ...] = ()
secondary_weapons: Tuple[WeaponItem, ...] = ()
semiautomatic_weapons: Tuple[WeaponItem, ...] = ()
secondary_shotguns: Tuple[WeaponItem, ...] = ()
```

加载每行时先解析四组武器，再用具体组构造旧接口：

```python
primary_rpms=tuple(item.firerate for item in primary_weapons),
secondary_rpms=tuple(item.firerate for item in secondary_weapons),
has_semiautomatic=bool(semiautomatic_weapons),
has_secondary_shotgun=bool(secondary_shotguns),
```

测试构造器 `make_card()` 同步生成命名稳定的 `WeaponItem`，保证已有渲染测试可提供图标。

- [ ] **Step 5: 运行解析、分档和排序测试**

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
& $workspacePython -m unittest discover -s tests -p 'test_r6_tier_chart.py' -v
& $workspacePython -m unittest discover -s tests -p 'test_r6_leaderboards.py' -v
```

Expected: 两组测试全部 PASS，刚完成的自身指标排序测试继续通过。

- [ ] **Step 6: 提交具体武器模型**

```powershell
git add -- src/r6_report/tier_chart.py tests/test_r6_tier_chart.py tests/test_r6_leaderboards.py
git commit -m "feat: retain leaderboard weapon details"
```

### Task 2: 原子采集并严格加载枪械图标

**Files:**
- Modify: `src/r6_report/tier_chart.py:240-380`
- Modify: `src/r6_report/wiki_client.py`
- Modify: `src/r6_report/collector.py:35-210`
- Modify: `src/r6_report/leaderboards.py:850-900`
- Modify: `tests/test_r6_tier_chart.py:195-340`
- Modify: `tests/test_collector.py:55-180`
- Modify: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Produces: `prepare_weapon_icons(items, directory, *, query_json, run_command, which, sleep) -> Dict[str, Path]`。
- Produces: `load_weapon_icons(items, directory) -> Mapping[str, Path]`。
- Produces: `HuijiClient.prepare_weapon_icons(items, directory)`。

- [ ] **Step 1: 写图标准备的失败测试**

使用两个 `WeaponItem`，伪造 MediaWiki `imageinfo` 返回和 curl 下载，断言：文件名分别为 `<icon_key>.png`；透明边缘被裁去；可见 RGB 统一为深色；相同 `icon_key` 只下载一次；损坏下载重试四次后抛出 `TierChartError`。

核心断言：

```python
self.assertEqual(tuple(paths), ("r4-c", "g8a1"))
with Image.open(paths["r4-c"]) as icon:
    self.assertEqual(icon.getchannel("A").getbbox(), (0, 0, *icon.size))
    self.assertTrue(all(pixel[:3] == (32, 35, 39) for pixel in visible))
```

- [ ] **Step 2: 运行测试并确认接口缺失**

```powershell
$env:PYTHONPATH='src;tests'
& $workspacePython -m unittest tests.test_r6_tier_chart.GadgetIconTests.test_prepares_all_weapon_icons_as_cropped_dark_silhouettes -v
```

Expected: ERROR，缺少 `prepare_weapon_icons`。

- [ ] **Step 3: 实现图标准备和加载**

为每个唯一枪械请求 `文件:R6S wpn <规范化名称>.png` 的可信灰机地址，下载到临时 `.download`，验证后裁透明边缘并以 `theme.COLOURS["text"]` 替换 RGB，最后原子替换目标。增加必要的显式文件名别名映射；未知或缺图直接失败。

`load_weapon_icons()` 对每个 `icon_key` 检查文件存在、图片有效且路径位于 `data/icons/weapon/` 下，缺失时抛出 `LeaderboardError`。

- [ ] **Step 4: 把枪械图标加入收集快照**

`collector._weapon_items(rows)` 从四组 `OperatorRow` 武器中按 `icon_key` 去重；`collect_snapshot()` 调用 `client.prepare_weapon_icons(..., icon_stage / "weapon")`，验证全部路径，并与 `wiki/icons/patches` 一起启用。`FakeHuijiClient` 生成每把测试枪的 PNG；新增缺枪图时旧快照保持不变的测试。

- [ ] **Step 5: 运行采集与图标测试**

```powershell
$env:PYTHONPATH='src'
& $workspacePython -m unittest discover -s tests -p 'test_collector.py' -v
& $workspacePython -m unittest discover -s tests -p 'test_r6_tier_chart.py' -v
```

Expected: 全部 PASS。

- [ ] **Step 6: 提交采集实现**

```powershell
git add -- src/r6_report/tier_chart.py src/r6_report/wiki_client.py src/r6_report/collector.py src/r6_report/leaderboards.py tests/test_r6_tier_chart.py tests/test_collector.py tests/test_r6_leaderboards.py
git commit -m "feat: collect validated weapon icons"
```

### Task 3: 在 XLSX 固定行高内渲染全部枪械图标

**Files:**
- Modify: `src/r6_report/report_theme.py`
- Modify: `src/r6_report/leaderboards.py:340-800`
- Modify: `tests/test_r6_leaderboards.py:450-820`

**Interfaces:**
- Consumes: `Mapping[str, Path]`，键为 `WeaponItem.icon_key`。
- Produces: `_add_weapon_icons_to_xlsx(sheet, items, weapon_icons, first_column, row) -> None`。

- [ ] **Step 1: 写 XLSX 多图标与固定行高失败测试**

构造含三把主手自动枪、两把主狙、两把副喷的卡片，调用工作簿写入器后断言：

```python
self.assertEqual(sheet.row_dimensions[feature_row].height, theme.XLSX_CARD_BODY_ROW_PT)
self.assertEqual(sheet.row_dimensions[rpm_row].height, theme.XLSX_CARD_BODY_ROW_PT)
self.assertEqual(len(weapon_row_images(feature_row)), 4)
self.assertEqual(len(weapon_row_images(rpm_row)), 3)
```

并断言每个图片锚点的列只落在当前字段的两列范围内，缺失卡片两行没有枪图。

- [ ] **Step 2: 运行测试并确认当前只有 Badge／装备图**

```powershell
$env:PYTHONPATH='src;tests'
& $workspacePython -m unittest tests.test_r6_leaderboards.LeaderboardWorkbookTests.test_renders_all_weapon_icons_without_changing_body_row_heights -v
```

Expected: FAIL，武器行图片数量为 0。

- [ ] **Step 3: 实现 XLSX 行内图片**

在主题中定义固定图标高度、单字段最大宽度、右侧起始偏移和最小间距。对每组最多三把枪先读取裁边后的宽高，再以可用宽度和固定行高共同限制等比缩放；用 `_add_offset_image()` 锚定在该字段右侧。字段为空时不创建图片对象；行高赋值保持原常量。

把 `weapon_icons` 参数贯穿 `write_leaderboard_workbook()`、`write_all_leaderboards()` 与 CLI `main()`，并继续把相同映射传给 PDF writer。

- [ ] **Step 4: 运行完整榜单测试**

```powershell
$env:PYTHONPATH='src'
& $workspacePython -m unittest discover -s tests -p 'test_r6_leaderboards.py' -v
```

Expected: 全部 PASS；旧文本、灰底、行高和排序断言不变。

- [ ] **Step 5: 提交 XLSX 渲染**

```powershell
git add -- src/r6_report/report_theme.py src/r6_report/leaderboards.py tests/test_r6_leaderboards.py
git commit -m "feat: render weapon icons in xlsx cards"
```

### Task 4: 在 PDF 固定行高内渲染全部枪械图标

**Files:**
- Modify: `src/r6_report/pdf_leaderboards.py:145-285`
- Modify: `tests/test_r6_leaderboards.py:800-1100`

**Interfaces:**
- Consumes: 与 XLSX 相同的 `Mapping[str, Path]`。
- Produces: `_weapon_field_flowable(text, items, weapon_icons, width, height)` 固定尺寸内层表格。

- [ ] **Step 1: 写 PDF 固定高度与全部图标失败测试**

构造三把主手自动枪和两把主狙，调用 `_card_flowable()`，断言信息表的两条正文 `rowHeights` 仍为 `theme.PDF_CARD_BODY_ROW_MM * mm`，并递归统计内层 `Image` 数量等于四组武器总数。

- [ ] **Step 2: 运行测试并确认图片数量不足**

```powershell
$env:PYTHONPATH='src;tests'
& $workspacePython -m unittest tests.test_r6_leaderboards.LeaderboardCliTests.test_pdf_renders_all_weapon_icons_in_fixed_body_rows -v
```

Expected: FAIL，PDF 卡片只有 Badge 与装备图片，没有枪械图片。

- [ ] **Step 3: 实现 PDF 文字／图标内层布局**

每个字段创建一行两列的固定高度 `Table`：左列为现有文字，右列为包含所有枪图的横向固定尺寸表。枪图使用 `_cropped_image_source()`，按图标数量共享剩余宽度并等比缩放；内层表所有 padding 为 0，外层正文行高保持 6 mm。

把 `weapon_icons` 传入 `_card_flowable()` 和 `write_leaderboard_pdf()`；不得在 PDF 层重新判断武器资格。

- [ ] **Step 4: 运行榜单与完整测试**

```powershell
$env:PYTHONPATH='src'
& $workspacePython -m unittest discover -s tests -p 'test_r6_leaderboards.py' -v
& $workspacePython -m unittest discover -s tests -v
```

Expected: 完整测试全部 PASS。

- [ ] **Step 5: 提交 PDF 渲染**

```powershell
git add -- src/r6_report/pdf_leaderboards.py tests/test_r6_leaderboards.py
git commit -m "feat: render weapon icons in pdf cards"
```

### Task 5: 获取真实图标、记录变更并重新生成五榜

**Files:**
- Modify: `CHANGELOG.md`
- Create: `data/icons/weapon/*.png`
- Generate locally: `output/*.xlsx`、`output/*.pdf`、`output/图片版/*/*.png`
- Move after completion: `docs/superpowers/plans/2026-08-04-weapon-icons-in-card-rows.md` → `~archived/superpowers-plans/2026-08-04-weapon-icons-in-card-rows.md`

**Interfaces:**
- Consumes: 当前 `data/wiki/*.json` 与新 `prepare_weapon_icons()`。
- Produces: 完整枪械图标缓存和更新后的五榜。

- [ ] **Step 1: 用临时目录准备当前快照所需全部枪械图标**

从 `data/r6_operator_stats.xlsx` 加载 77 张卡片，汇总四组 `WeaponItem`，调用 `prepare_weapon_icons()` 写入 `~temp/weapon-icons-stage`。只有全部图片通过格式、可见像素和数量校验后，才把完整目录移动到不存在的 `data/icons/weapon/`；若目标已存在则先验证并复用，禁止覆盖未知内容。

- [ ] **Step 2: 更新 Changelog**

在 `Unreleased > Added` 增加：

```markdown
- 五张榜单卡片在副喷、主狙、副手自动枪和主手自动枪字段右侧显示全部对应枪械图标；多把枪与全部射速按相同顺序保留，XLSX/PDF 正文行高不变。
```

- [ ] **Step 3: 运行完整测试并生成五榜**

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
& $workspacePython -m unittest discover -s tests -v
& $workspacePython -m r6_report.leaderboards --data-dir data --input data/r6_operator_stats.xlsx --output-dir output
```

Expected: 测试全部 PASS；生成 5 XLSX、5 PDF、15 PNG。

- [ ] **Step 4: 验证真实多枪卡片和固定高度**

读取 `output/视频评分榜.xlsx`，定位 IQ、Dokkaebi、Zero 等卡片；断言 IQ 主手射速行为 `850/720/690` 且该行有三张枪图，Dokkaebi 主狙行有两张枪图，所有 feature/rpm 行高度等于 `theme.XLSX_CARD_BODY_ROW_PT`。渲染 PNG 后人工抽查图标不越界、不重叠相邻字段。

- [ ] **Step 5: 提交 Changelog、真实图标和本次缓存**

```powershell
git add -- CHANGELOG.md data/icons/weapon src/r6_report/__pycache__ tests/__pycache__
git commit -m "feat: add weapon icon assets to reports"
```

- [ ] **Step 6: 归档计划并提交**

使用 `apply_patch` 把本计划完整移动到 `~archived/superpowers-plans/2026-08-04-weapon-icons-in-card-rows.md`，增加“已完成并归档”状态，然后提交：

```powershell
git add -- docs/superpowers/plans/2026-08-04-weapon-icons-in-card-rows.md ~archived/superpowers-plans/2026-08-04-weapon-icons-in-card-rows.md
git commit -m "docs: archive weapon icon plan"
```

- [ ] **Step 7: 最终验证**

重新运行完整测试、生成物数量、`git diff --check`、`git status --short` 和分支提交列表。最终工作区只允许保留任务开始前已有的未跟踪 `~archived/output/`。

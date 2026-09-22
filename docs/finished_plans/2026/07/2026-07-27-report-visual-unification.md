# PDF 与 XLSX 榜单视觉统一 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 统一五种榜单 PDF/XLSX 的颜色、字体、字号与状态表达，修复 PDF 次要装备图标横向压缩，并让所有语义明确的缺失值使用灰色底。

**Architecture:** 新建无渲染器依赖的 `report_theme.py`，集中维护颜色、字体、字号层级、状态文本和缺失值判定。XLSX 与 PDF 渲染器分别把共享语义转换为 openpyxl 和 ReportLab 样式；PDF 额外使用纯函数计算等比例图标尺寸。

**Tech Stack:** Python 3、openpyxl、ReportLab、Pillow、pypdf、pdfplumber、`@oai/artifact-tool`、unittest

## Global Constraints

- 保持现有数据、分档、排序、卡片结构、PDF 三页结构和 XLSX 工作表结构不变。
- 中文字体使用 Microsoft YaHei；PDF 不可用时沿用现有中文字体回退。
- PDF 与 XLSX 共享同一组颜色、字号语义和状态文本规则。
- 增强使用绿色，削弱使用红色，混合使用黄色。
- 有特性显示 `副喷 ✓`、`主狙 ✓`；没有特性显示 `副喷 -`、`主狙 -`。
- 所有语义明确的缺失字段使用灰色底，不把普通正文中的连字符当作缺失值。
- PDF 次要装备图标保持原始长宽比，在固定边界框内居中。
- 不改变 `api_version`、输入数据格式、公开命令参数或输出文件名。
- 不增加新的联网依赖。

---

### Task 1: 共享报告主题

**Files:**
- Create: `src/r6_report/report_theme.py`
- Create: `tests/test_report_theme.py`

**Interfaces:**
- Produces: `FONT_FAMILY: str`
- Produces: `COLOURS: Mapping[str, str]`
- Produces: `XLSX_FONT_SIZES: Mapping[str, float]`
- Produces: `PDF_FONT_SIZES: Mapping[str, float]`
- Produces: `feature_text(label: str, present: bool) -> str`
- Produces: `rpm_text(prefix: str, values: tuple[int, ...]) -> str`
- Produces: `is_missing_field(text: str) -> bool`

- [ ] **Step 1: 写共享主题失败测试**

```python
from r6_report import report_theme as theme


def test_shared_status_and_missing_value_rules():
    assert theme.feature_text("副喷", True) == "副喷 ✓"
    assert theme.feature_text("主狙", False) == "主狙 -"
    assert theme.rpm_text("副", ()) == "副 -"
    assert theme.rpm_text("主", (690, 650)) == "主 690/650"
    assert theme.is_missing_field("副喷 -")
    assert theme.is_missing_field("主 -")
    assert not theme.is_missing_field("Y11S2.1 - Y11S2.2")


def test_patch_direction_colours_are_shared():
    assert set(theme.PATCH_DIRECTION_COLOURS) == {"增强", "削弱", "混合"}
    assert theme.MISSING_FILL == "D9D9D9"
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m unittest tests.test_report_theme -v`

Expected: FAIL，`r6_report.report_theme` 尚不存在。

- [ ] **Step 3: 实现共享常量与纯函数**

```python
FONT_FAMILY = "Microsoft YaHei"
MISSING_FILL = "D9D9D9"
PATCH_DIRECTION_COLOURS = {
    "增强": "548235",
    "削弱": "C00000",
    "混合": "BF9000",
}


def feature_text(label: str, present: bool) -> str:
    return f"{label} {'✓' if present else '-'}"


def rpm_text(prefix: str, values: tuple[int, ...]) -> str:
    payload = "/".join(str(value) for value in values) or "-"
    return f"{prefix} {payload}"


def is_missing_field(text: str) -> bool:
    return str(text).strip() == "-" or str(text).rstrip().endswith(" -")
```

同时定义标题、分档、姓名、正文、装备、来源、页码所需的颜色和 XLSX/PDF 字号映射。

- [ ] **Step 4: 运行共享主题测试**

Run: `python -m unittest tests.test_report_theme -v`

Expected: PASS。

- [ ] **Step 5: 提交共享主题**

```bash
git add src/r6_report/report_theme.py tests/test_report_theme.py
git commit -m "feat: add shared report theme"
```

### Task 2: XLSX 卡片与补丁样式

**Files:**
- Modify: `src/r6_report/leaderboards.py`
- Modify: `src/r6_report/patch_notes.py`
- Modify: `tests/test_r6_leaderboards.py`
- Modify: `tests/test_r6_patch_notes.py`

**Interfaces:**
- Consumes: Task 1 的主题常量、`feature_text`、`rpm_text`、`is_missing_field`
- Produces: 五种 XLSX 使用统一状态文字、缺失灰底、字体层级和补丁方向颜色

- [ ] **Step 1: 写 XLSX 失败测试**

在 `tests/test_r6_leaderboards.py` 的工作簿渲染测试中加入：

```python
self.assertEqual(sheet["C4"].value, "副喷 ✓")
self.assertEqual(sheet["E4"].value, "主狙 ✓")
self.assertEqual(missing_sheet["C4"].value, "副喷 -")
self.assertEqual(missing_sheet["C4"].fill.fgColor.rgb[-6:], theme.MISSING_FILL)
self.assertEqual(missing_sheet["C5"].value, "副 -")
self.assertEqual(missing_sheet["C5"].fill.fgColor.rgb[-6:], theme.MISSING_FILL)
```

在 `tests/test_r6_patch_notes.py` 中逐项断言方向单元格颜色等于 `theme.PATCH_DIRECTION_COLOURS`。

- [ ] **Step 2: 运行目标测试确认失败**

Run: `python -m unittest tests.test_r6_leaderboards tests.test_r6_patch_notes -v`

Expected: 至少缺失值灰底和共享颜色断言失败。

- [ ] **Step 3: 接入共享主题**

- 用 `feature_text` 生成副喷/主狙文本。
- 用 `rpm_text` 生成主手/副手射速文本。
- 对 `is_missing_field(cell.value)` 为真的卡片字段设置 `MISSING_FILL`。
- 所有字体改用共享字体、字号和颜色常量。
- `patch_notes.py` 删除本地补丁方向颜色表，读取共享常量。

- [ ] **Step 4: 运行 XLSX 目标测试**

Run: `python -m unittest tests.test_r6_leaderboards tests.test_r6_patch_notes -v`

Expected: PASS。

- [ ] **Step 5: 提交 XLSX 改造**

```bash
git add src/r6_report/leaderboards.py src/r6_report/patch_notes.py tests/test_r6_leaderboards.py tests/test_r6_patch_notes.py
git commit -m "feat: unify xlsx report styling"
```

### Task 3: PDF 卡片、补丁颜色与等比例装备图标

**Files:**
- Modify: `src/r6_report/pdf_leaderboards.py`
- Modify: `tests/test_r6_leaderboards.py`

**Interfaces:**
- Consumes: Task 1 的共享主题
- Produces: `_fit_image_size(path: Path, box_width: float, box_height: float) -> tuple[float, float]`
- Produces: PDF 状态文字、缺失灰底、补丁方向颜色和统一字体层级

- [ ] **Step 1: 写 PDF 失败测试**

```python
def test_pdf_gadget_image_preserves_aspect_ratio(self):
    width, height = pdf_lb._fit_image_size(icon_path, 20, 20)
    self.assertEqual((width, height), (20, 10))


def test_pdf_uses_checkmarks_and_patch_direction_colours(self):
    pdf_lb.write_leaderboard_pdf(...)
    text = "\n".join(page.extract_text() or "" for page in PdfReader(output).pages)
    self.assertIn("副喷 ✓", text)
    self.assertIn("主狙 ✓", text)
```

同时解压 PDF 内容流或渲染补丁页像素采样，确认三个方向颜色均写入。

- [ ] **Step 2: 运行 PDF 目标测试确认失败**

Run: `python -m unittest tests.test_r6_leaderboards -v`

Expected: PDF 仍包含“是”，且固定 5 mm × 5 mm 图标不能通过比例测试。

- [ ] **Step 3: 实现 PDF 共享样式**

- `_style` 从共享 PDF 字号和颜色映射取值。
- 副喷/主狙改用 `feature_text`。
- 主副手射速改用 `rpm_text`。
- 对缺失状态与缺失射速单元格增加 `MISSING_FILL` 背景。
- 补丁表按每行 `change.direction` 设置第一列背景和高对比文字。
- 使用 Pillow 读取图标原始尺寸，由 `_fit_image_size` 计算固定边界框内的等比例尺寸。

- [ ] **Step 4: 运行 PDF 目标测试**

Run: `python -m unittest tests.test_r6_leaderboards -v`

Expected: PASS。

- [ ] **Step 5: 提交 PDF 改造**

```bash
git add src/r6_report/pdf_leaderboards.py tests/test_r6_leaderboards.py
git commit -m "feat: unify pdf report styling"
```

### Task 4: 生成报告并完成视觉验证

**Files:**
- Modify: `docs/视频评分榜.xlsx`
- Modify: `docs/视频评分榜.pdf`
- Modify: `docs/主武器射速榜.xlsx`
- Modify: `docs/主武器射速榜.pdf`
- Modify: `docs/速度榜.xlsx`
- Modify: `docs/速度榜.pdf`
- Modify: `docs/稀有枪械榜.xlsx`
- Modify: `docs/稀有枪械榜.pdf`
- Modify: `docs/次要装备榜.xlsx`
- Modify: `docs/次要装备榜.pdf`
- Modify: `docs/previews/*.png`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Tasks 1-3 的生成器
- Produces: 五份最终 XLSX、五份最终 PDF、十张 README 预览图

- [ ] **Step 1: 运行完整测试基线**

Run: `python -m unittest discover -s tests -v`

Expected: PASS，0 failures。

- [ ] **Step 2: 重新生成五种报告**

Run: `python -m r6_report.leaderboards --data-dir data --output-dir docs`

Expected: 五份 XLSX 和五份同名 PDF 更新成功。

- [ ] **Step 3: 渲染全部 PDF 页面**

对每份 PDF 执行：

```powershell
pdftoppm -png -r 110 "docs/视频评分榜.pdf" "~temp/pdf-review/视频评分榜"
```

检查每份报告三页：图标不变形，增强/削弱/混合颜色正确，所有文本不裁切，页脚和页码清楚。

- [ ] **Step 4: 用 artifact-tool 检查并渲染全部 XLSX**

导入每份 XLSX，逐工作表执行：

```javascript
await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
});
await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
```

确认没有公式错误；进攻方、防守方和补丁说明三张工作表均完成视觉检查。

- [ ] **Step 5: 更新预览和变更日志**

从新版工作簿渲染图更新 `docs/previews/` 十张图片；在 `CHANGELOG.md` 记录共享主题、勾选符号、缺失灰底、补丁方向颜色和 PDF 图标等比例修复。

- [ ] **Step 6: 运行最终验证**

Run:

```powershell
python -m unittest discover -s tests -v
git diff --check
git status --short
```

Expected: 所有测试通过、无 whitespace 错误，状态只包含本任务文件。

- [ ] **Step 7: 归档计划并提交最终产物**

将本计划移入 `~archived/superpowers-plans/2026-07-27-report-visual-unification.md`，不得删除。

```bash
git add src tests docs CHANGELOG.md ~archived/superpowers-plans/2026-07-27-report-visual-unification.md
git commit -m "feat: unify pdf and xlsx report visuals"
```

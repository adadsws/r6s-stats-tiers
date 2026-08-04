# PDF 榜单报告 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让统一榜单命令同时生成 5 个 XLSX 和 5 个完整 PDF，并在 README 展示 5 张交错阵营预览及 PDF 链接。

**Architecture:** 新增 `r6_report.pdf_leaderboards`，以现有 `OperatorCard`、`LeaderboardSpec`、`group_cards`、补丁来源和图标为输入，使用 ReportLab 绘制横向 A3 多页 PDF。`leaderboards.write_all_leaderboards` 继续生成 XLSX，并紧接着调用 PDF writer；示例 PDF 从当前 `data/` 生成到 `docs/`，Poppler 负责预览与视觉检查。

**Tech Stack:** Python 3.9+、ReportLab、Pillow、pypdf、Poppler、unittest

## Global Constraints

- 不调用 Microsoft Excel 或 LibreOffice。
- 每个 PDF 依次包含进攻方、防守方、补丁说明。
- XLSX 现有内容与文件名保持不变。
- README 链接 `docs/*.pdf`，预览阵营严格交错。
- `output/`、`~temp/` 继续忽略；源码、测试、示例 PDF、PNG、文档和归档计划进入本地 Git。
- 不执行远程 push。

---

### Task 1: 用失败测试定义 PDF 契约

**Files:**
- Modify: `tests/test_r6_leaderboards.py`
- Modify: `requirements.txt`

- [x] 新增测试：调用真实 `write_all_leaderboards` 后得到 10 个路径，每个榜单同时存在同名 XLSX/PDF。
- [x] 用 `pypdf.PdfReader` 验证每个 PDF 至少 3 页，提取文本包含“进攻方”“防守方”“补丁说明”。
- [x] 运行目标测试并确认因 PDF API 尚不存在而失败。
- [x] 在 `requirements.txt` 加入 `reportlab>=3.6,<5` 与 `pypdf>=3,<7`。

### Task 2: 实现 ReportLab PDF renderer

**Files:**
- Create: `src/r6_report/pdf_leaderboards.py`
- Modify: `src/r6_report/leaderboards.py`

- [x] 实现 `write_leaderboard_pdf(path, spec, cards, operator_icon_dir, gadget_icons, report_sources)`。
- [x] 用横向 A3、系统中文字体（缺失时回退 `STSong-Light`）、每行 5 卡片绘制阵营页面；复用 `group_cards`、`band_order`、`band_color` 与补丁 marker。
- [x] 绘制补丁说明页、来源页脚和页码。
- [x] 修改 `write_all_leaderboards`，按每个 spec 依次生成 XLSX/PDF 并返回 10 个路径。
- [x] 运行目标测试直到通过，再运行 `tests/test_r6_leaderboards.py`。

### Task 3: 生成示例 PDF 与 README 预览

**Files:**
- Create: `docs/视频评分榜.pdf`
- Create: `docs/主武器射速榜.pdf`
- Create: `docs/速度榜.pdf`
- Create: `docs/稀有枪械榜.pdf`
- Create: `docs/次要装备榜.pdf`
- Create/replace: `docs/previews/*.png`
- Modify: `README.md`
- Modify: `CHANGELOG.md`

- [x] 运行统一榜单命令生成 10 个 output 文件。
- [x] 将 5 个 PDF 复制到 `docs/`。
- [x] 用 Poppler 按交错阵营页渲染 5 张 PNG，并逐张视觉检查。
- [x] README 增加 5 个 PDF 链接和 5 张 PNG；输出清单补充 PDF。
- [x] CHANGELOG 记录 PDF renderer、依赖和 README 预览。

### Task 4: 验证、归档与提交

**Files:**
- Move: 本计划到 `~archived/superpowers-plans/`

- [x] 运行全部单元测试。
- [x] 用 pypdf 检查 5 个 docs PDF 的页数、文本和文件大小。
- [x] 检查 README 的 10 个链接存在、Git 忽略规则未扩张、工作区无未跟踪非输出内容。
- [x] 归档计划并创建 `feat: add PDF leaderboard reports` 本地 commit。

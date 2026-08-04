# 全局 AGENTS.md 合规改造 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在不改写 Git 历史、不发布远程仓库、不丢失现有文件的前提下，让当前工作树、运行路径、文档和 Git 跟踪规则满足全局 `AGENTS.md`。

**Architecture:** 将原始输入、可重建中间产物、交付结果和旧文件分别收敛到 `inputs/`、`~temp/`、`~outputs/`、`~archive/`。运行时代码只改变默认路径和命令行路径契约，不改变报表内容、数据格式或 `api_version`；项目治理由精简的 `AGENTS.md`、`AGENT_CONTEXT.md`、按日期组织的 `CHANGELOG.md` 和自动化布局测试固化。

**Tech Stack:** Python 3.9+、`unittest`、PowerShell/BAT、Git、Markdown

## Global Constraints

- 当前分支 `codex/leaderboard-self-metric-sorting` 与 `main` 均不重命名；不改写历史、不 push、不创建 PR。
- 保留当前未跟踪的 `~archived/output/` 和被忽略的 `output/`、`~temp/` 内容；迁移时不得覆盖同名文件。
- 普通源码、测试、文档和原始输入由本地 Git 跟踪；`~temp/`、`~outputs/`、`.worktrees/`、`.venv/` 和工具缓存不跟踪。
- 不创建技能默认的额外规格、实施计划或元计划；本文件是本次任务唯一计划。
- README 只保留功能、部署和使用说明；AGENT_CONTEXT 只保留架构、结构、开发流程和已知问题；项目 AGENTS 只保留项目特有约束。
- 普通功能迭代不递增 `api_version`；本任务不引入依赖或联网调研。

## 方案与待确认决策

### A. 完整语义迁移（推荐）

执行本计划全部任务：`data/` 改为 `inputs/`，生成的 `r6_operator_stats.xlsx` 改放 `~temp/`，`output/` 改为 `~outputs/`，`~archived/` 改为 `~archive/`。优点是当前工作树逐项符合全局目录语义，并由测试防止回退；代价是本地 CLI 默认值和 BAT 可见路径发生一次不兼容变化。

### B. 仅治理文件

只新增治理文档、日期化 CHANGELOG、ignore 规则和缓存清理，不迁移业务路径。改动小，但 `data/`、`output/`、`~archived/` 仍与全局规范冲突，因此只能算部分合规。

### C. 兼容期迁移

代码接受新旧两套命令行参数并在旧目录存在时回退。对旧脚本最友好，但长期保留双重路径语义、增加测试面，而且旧目录仍可能重新出现。

**确认方式：** 回复 `v` 表示批准推荐方案 A 并授权继续；也可回复 `B` 或 `C`，届时先按所选范围修订本计划再执行。

---

### Task 1: 用失败测试锁定合规契约

**Files:**
- Modify: `tests/test_project_layout.py`

**Interfaces:**
- Consumes: 项目根目录和现有布局测试。
- Produces: `test_global_agents_layout_contract()`，覆盖规范目录、治理文档、ignore 规则和非跟踪缓存。

- [x] **Step 1: 添加布局契约测试**

在 `PackageLayoutTests` 中增加测试，断言：

```python
def test_global_agents_layout_contract(self):
    root = Path(__file__).resolve().parents[1]
    for required in ("AGENTS.md", "AGENT_CONTEXT.md", "README.md", "CHANGELOG.md"):
        self.assertTrue((root / required).is_file(), required)
    for required_dir in ("inputs", "docs/plans", "docs/finished_plans", "~archive"):
        self.assertTrue((root / required_dir).is_dir(), required_dir)
    for legacy in ("data", "output", "~archived", "docs/superpowers/plans"):
        self.assertFalse((root / legacy).exists(), legacy)
    ignore = (root / ".gitignore").read_text(encoding="utf-8")
    for rule in ("/~temp/", "/~outputs/", "/.worktrees/", "/.venv/", "__pycache__/"):
        self.assertIn(rule, ignore)
```

- [x] **Step 2: 运行测试并确认失败**

Run: `$env:PYTHONPATH='src'; $env:PYTHONUTF8='1'; python -m unittest tests.test_project_layout.PackageLayoutTests.test_global_agents_layout_contract -v`

Expected: FAIL，首先报告缺少 `AGENTS.md` 或 `AGENT_CONTEXT.md`。

### Task 2: 迁移规范目录并修正 Git 跟踪边界

**Files:**
- Move: `data/` → `inputs/`
- Move: `data/r6_operator_stats.xlsx` → `~temp/r6_operator_stats.xlsx`
- Move: `output/` 内容 → `~outputs/`
- Move: `~archived/output/` → `~outputs/legacy-output-2026-08-04/`
- Move: `~archived/` → `~archive/`
- Move: `~archived.md` → `~archive.md`
- Modify: `.gitignore`
- Modify: `.gitattributes` only if migration reveals a missing binary/text rule

**Interfaces:**
- Consumes: 当前 Git 已跟踪输入与归档、当前忽略输出、当前未跟踪旧输出。
- Produces: 唯一规范目录集合，且所有原文件仍可从新位置访问。

- [x] **Step 1: 校验所有迁移源和目标**

用 `Resolve-Path` 确认源都位于 `<PROJECT_ROOT>`；确认 `inputs/`、`~outputs/`、`~archive/` 不存在。若目标已存在，停止迁移并报告冲突，不合并或覆盖。

- [x] **Step 2: 先分离生成产物，再执行目录重命名**

使用 PowerShell `Move-Item -LiteralPath` 完成上述移动；`~outputs/` 直接接收当前 `output/` 内容，旧归档输出放入独立的 `legacy-output-2026-08-04/`。移动后逐项比较文件数量与总字节数，源目录应消失且目标统计应相等。

- [x] **Step 3: 写入精确且带说明的 ignore 规则**

`.gitignore` 改为：

```gitignore
# 可随时重建的缓存、日志和中间产物。
/~temp/
# 需要本地检查、复用或交付但不进入 Git 的生成结果。
/~outputs/
# 本地固定副本和 Python 虚拟环境。
/.worktrees/
/.venv/
# Python 在源码或测试相邻位置生成的解释器缓存。
__pycache__/
*.py[cod]
```

- [x] **Step 4: 停止跟踪现有 Python 缓存但保留工作树文件**

用 `git rm -r --cached --ignore-unmatch` 仅从 index 移除当前被跟踪的 `__pycache__/` 和 `*.pyc`；禁止删除工作树缓存。验证 `git ls-files '*__pycache__*' '*.pyc'` 无输出。

### Task 3: 迁移运行路径并保持报表行为不变

**Files:**
- Modify: `run_r6_report.bat`
- Modify: `src/r6_report/collector.py`
- Modify: `src/r6_report/operator_stats.py`
- Modify: `src/r6_report/leaderboards.py`
- Modify: `src/r6_report/tier_chart.py`
- Modify: `tests/test_collector.py`
- Modify: `tests/test_r6_operator_stats.py`
- Modify: `tests/test_r6_leaderboards.py`
- Modify: `tests/test_r6_tier_chart.py`
- Modify: `tests/test_sources.py`

**Interfaces:**
- Consumes: `inputs/athieno/latest.json`、`inputs/wiki/`、`inputs/icons/`、`inputs/patches/`。
- Produces: `~temp/r6_operator_stats.xlsx` 和 `~outputs/` 下五组 XLSX/PDF/PNG；CLI 使用 `--inputs-dir`、`--archive-dir`、`--temp-dir`、`--output-dir`。

- [x] **Step 1: 先更新 CLI/BAT 路径测试并确认失败**

将现有测试中生产路径期望改为 `inputs`、`~temp`、`~outputs`、`~archive`，并要求 BAT 包含：

```text
--inputs-dir "%~dp0inputs"
--archive-dir "%~dp0~archive\data-snapshots"
--output "%~dp0~temp\r6_operator_stats.xlsx"
--output-dir "%~dp0~outputs"
```

Run: `$env:PYTHONPATH='src'; $env:PYTHONUTF8='1'; python -m unittest tests.test_collector tests.test_r6_operator_stats tests.test_r6_leaderboards tests.test_r6_tier_chart tests.test_sources tests.test_project_layout -v`

Expected: FAIL，错误仅来自旧默认路径、旧参数名或尚未迁移的文件位置。

- [x] **Step 2: 实现最小路径迁移**

将四个 CLI 的输入目录默认值统一为 `Path("inputs")`；collector 的归档默认值改为 `Path("~archive") / "data-snapshots"`；leaderboards 与 tier_chart 的默认输出改到 `~outputs/`。BAT 的第二阶段输出 `~temp/r6_operator_stats.xlsx`，第三阶段显式读取该中间文件。函数内部表示业务数据的普通变量名无需机械改名。

- [x] **Step 3: 运行路径相关测试并确认通过**

Run: `$env:PYTHONPATH='src'; $env:PYTHONUTF8='1'; python -m unittest tests.test_collector tests.test_r6_operator_stats tests.test_r6_leaderboards tests.test_r6_tier_chart tests.test_sources tests.test_project_layout -v`

Expected: PASS。

### Task 4: 建立职责分离的项目治理文档

**Files:**
- Create: `AGENTS.md`
- Create: `AGENT_CONTEXT.md`
- Modify: `README.md`
- Rewrite compactly: `CHANGELOG.md`
- Modify: `skills/build-r6-operator-report/SKILL.md`
- Modify: `skills/build-r6-operator-report/agents/openai.yaml`
- Modify: `requirements.txt`
- Move: `docs/superpowers/plans/2026-07-26-project-intro-video-script.md` → `docs/plans/2026-07-26-project-intro-video-script.md`
- Move: `~archive/superpowers-plans/` → `docs/finished_plans/`
- Move: `docs/superpowers/specs/` → `~archive/superpowers-specs/`
- Remove empty directory: `docs/superpowers/`

**Interfaces:**
- Consumes: 已迁移的目录和当前用户文档。
- Produces: 互不重复的用户说明、开发上下文、项目代理规则和逐日变更记录。

- [x] **Step 1: 创建精简的项目专用 AGENTS.md**

只记录本项目特有规则：报表数据流顺序、`run_r6_report.bat` 的 CRLF/UTF-8 约束、联网仅限 collector 与 Athieno Skill、发布示例先生成到 `~outputs/` 再人工审查。不得复制全局文件管理和 Git 规则。

- [x] **Step 2: 创建 AGENT_CONTEXT.md**

包含四节：技术架构、项目结构、开发流程、已知问题。明确 collector → operator_stats → leaderboards 数据流；列出 `src/r6_report/`、`inputs/`、`~temp/`、`~outputs/`、`~archive/`、`tests/`；已知问题记录 Windows/`curl.exe` 依赖、联网采集波动和 Git safe-directory 的本机所有者差异。

- [x] **Step 3: 更新 README 和项目 Skill**

把所有当前操作路径改成规范路径，部署与离线重建命令读取 `inputs/` 和 `~temp/r6_operator_stats.xlsx`，生成清单改为 `~outputs/`。Skill 继续只更新 Athieno 文件，但目标改为 `inputs/athieno/latest.json`。

- [x] **Step 4: 将 CHANGELOG 压缩为按日期的一日一节**

使用 Git 提交日期作为历史分组，保留用户可见变化并删除重复内部过程记录。顶层日期至少包括 `2026-08-04`、`2026-07-28`、`2026-07-27`、`2026-07-26`、`2026-07-25`；本次迁移记录在 `2026-08-04`，不保留 `Unreleased`。

- [x] **Step 5: 归位旧计划并封存技能遗留规格**

未完成的视频文案计划迁入 `docs/plans/` 并只修正路径；既有完成计划迁入 `docs/finished_plans/`；历史规格移到 `~archive/superpowers-specs/`。这只改变存放位置，不执行视频文案任务，也不把规格内容复制到治理文档。基线验证发现 Pillow `<12` 与现有图像测试 API 不兼容，因此将依赖约束修正为 `Pillow>=12,<13`。

### Task 5: 全量验证、敏感信息检查和实现提交

**Files:**
- Verify: all changed and moved files

**Interfaces:**
- Consumes: Tasks 1–4 的完整工作树。
- Produces: 可审计的通过证据和一次只包含本任务相关变化的实现提交。

- [x] **Step 1: 运行完整测试套件**

Run: `$env:PYTHONPATH='src'; $env:PYTHONUTF8='1'; python -m unittest discover -s tests -v`

Expected: 全部 PASS。

- [x] **Step 2: 运行结构、编码和 Git 边界检查**

验证 BAT 为 UTF-8/CRLF；`git diff --check` 无错误；`git check-ignore -v ~temp ~outputs .venv .worktrees` 命中相邻说明规则；`git check-ignore inputs '~archive' docs src tests` 不命中；除 `~temp/`、`~outputs/`、`.venv/`、`.worktrees/` 和精确工具缓存外无被忽略文件。

- [x] **Step 3: 检查大文件、隐私和 secrets 状态**

重新统计目录与大于 100 MiB 文件，确认 `~archive.md` 记录约 110.16 MiB 安装包及重建方式。扫描当前 diff 和待提交文件中的令牌、密码、私钥、Cookie、邮箱、绝对用户目录；确认项目没有真实 `secrets/**`，因此本次不初始化 `git-crypt`。

- [x] **Step 4: 提交实现**

只暂存本计划列出的相关路径，检查 `git diff --cached --stat` 与 `git diff --cached --name-status`，然后提交：

```text
chore: align project with global agent rules
```

### Task 6: 完成本次计划归档

**Files:**
- Move: `docs/plans/2026-08-04-global-agents-compliance.md` → `docs/finished_plans/2026-08-04-global-agents-compliance.md`

**Interfaces:**
- Consumes: 已通过验证并提交的实现。
- Produces: `docs/plans/` 只保留仍未完成的视频文案计划，本次完成计划进入 `docs/finished_plans/`。

- [x] **Step 1: 移动并核对计划状态**

确认本计划所有复选框均已完成后移动文件；验证 `docs/finished_plans/2026-08-04-global-agents-compliance.md` 存在，源路径不存在。

- [x] **Step 2: 提交计划归档**

提交：

```text
docs: archive global agents compliance plan
```

- [x] **Step 3: 最终复验**

重新运行完整 `unittest`、`git diff --check`、路径契约测试和 `git status --short --branch`。Expected: 测试全部 PASS，除被忽略的本地结果与缓存外工作树干净。

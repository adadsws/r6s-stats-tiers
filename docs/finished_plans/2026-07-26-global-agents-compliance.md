# 全局 AGENTS.md 合规改造 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让当前仓库的默认分支、目录跟踪规则和远程发布说明符合更新后的全局 `AGENTS.md`。

**Architecture:** 保留源仓库中的本地私有历史，只将默认分支改为 `main`；远程发布继续通过 `output/github-export/` 中的脱敏导出仓库完成。项目文档只补充与合规直接相关的说明，不改动现有业务实现和用户尚未提交的数据快照。

**Tech Stack:** Git、Markdown、PowerShell

## Global Constraints

- `output/`、`~temp/` 必须被 `.gitignore` 忽略，且不纳入 Git。
- `data/`、`~archived/` 及其他非输出内容必须由本地 Git 跟踪。
- `data/` 仅可在脱敏后发布；`~archived/` 及其历史禁止 push。
- 远程发布使用 `output/github-export/` 中仅含 `main` 的脱敏导出仓库，只允许 fast-forward push。
- 不修改或提交与本任务无关的既有工作区改动。

---

### Task 1: 对齐本地默认分支

**Files:**
- Modify: Git branch reference `master` → `main`

**Interfaces:**
- Consumes: 当前本地分支 `master` 及其完整本地私有历史。
- Produces: 唯一本地分支 `main`，供后续脱敏导出流程引用。

- [x] **Step 1: 确认当前分支和分支数量**

Run: `git -c safe.directory=<PROJECT_ROOT> branch --format="%(refname:short)"`

Expected: 仅输出 `master`。

- [x] **Step 2: 重命名当前分支**

Run: `git -c safe.directory=<PROJECT_ROOT> branch -m main`

Expected: 命令成功且不改动工作区文件。

- [x] **Step 3: 验证重命名结果**

Run: `git -c safe.directory=<PROJECT_ROOT> branch --format="%(refname:short)"`

Expected: 仅输出 `main`。

### Task 2: 固化发布约束并验证

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Move after completion: `docs/superpowers/plans/2026-07-26-global-agents-compliance.md` → `~archived/superpowers-plans/2026-07-26-global-agents-compliance.md`

**Interfaces:**
- Consumes: 全局 `AGENTS.md` 的本地记录和远程脱敏发布规则。
- Produces: 项目级可执行发布说明及本次合规变更记录。

- [x] **Step 1: 更新 README**

在目录说明之后增加“本地 Git 与远程发布”章节，明确：

1. 源仓库 `main` 保存本地私有历史，不得直接 push。
2. `data/` 发布时使用脱敏示例替换真实输入。
3. `~archived/` 及历史不得进入远程导出。
4. 远程操作仅在 `output/github-export/` 的 `main` 上执行，并要求同步、检查、测试和 fast-forward push。

- [x] **Step 2: 更新 CHANGELOG**

在 `Unreleased / Changed` 中记录默认分支改名和发布边界说明。

- [x] **Step 3: 运行仓库合规检查**

Run:

```powershell
git -c safe.directory=<PROJECT_ROOT> check-ignore -v output ~temp
git -c safe.directory=<PROJECT_ROOT> check-ignore data '~archived'
git -c safe.directory=<PROJECT_ROOT> ls-files --others --ignored --exclude-standard
```

Expected: 前两项分别显示 `.gitignore` 的根目录规则；`data` 与 `~archived` 不匹配忽略规则；被忽略文件只位于 `output/` 或 `~temp/`。

- [x] **Step 4: 运行项目测试**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest discover -s tests -v
```

Expected: 所有测试通过。

- [x] **Step 5: 归档计划**

将本计划移动到 `~archived/superpowers-plans/`，确保 `docs/superpowers/plans/` 不保留已完成计划。

- [x] **Step 6: 创建本地 commit**

仅暂存 `README.md`、`CHANGELOG.md` 和归档后的计划，保留既有 `data/` 与新快照改动不暂存。

Run: `git commit -m "chore: align repository with updated agent rules"`

Expected: commit 成功，工作区只剩用户原有的数据快照改动。

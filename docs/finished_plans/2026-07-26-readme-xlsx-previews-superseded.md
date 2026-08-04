# README 榜单工作簿预览 Implementation Plan（已被取代）

> 本计划在执行期间被“生成 XLSX 时同步生成完整 PDF，README 链接 PDF”的新需求取代。已完成的工作簿结构检查保留供追溯；未继续使用临时 Excel 截图方案。

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 `docs/` 中 5 个榜单工作簿按进攻方、防守方交错渲染为 5 张清晰 PNG，并加入 README 预览。

**Architecture:** 使用工作区依赖加载器提供的 Node.js 与 `@oai/artifact-tool`，在被 Git 忽略的 `~temp/xlsx-preview/` 中运行单一可复用脚本。脚本只读取现有工作簿并将目标工作表首屏渲染到 `docs/previews/`；README 使用相对链接展示图片和原始 `.xlsx`。

**Tech Stack:** Node.js、`@oai/artifact-tool`、Markdown、PowerShell、Git

## Global Constraints

- 原始工作簿保留在 `docs/`，不修改内容。
- 5 张图片按“进攻方、 防守方、进攻方、 防守方、进攻方”交错。
- PNG 必须保留标题、阵营、左侧分档和首屏卡片，不得拉伸或裁掉关键内容。
- 临时脚本和依赖 junction 只放在 `~temp/`，不得纳入 Git。
- 新增工作簿、预览图、README、CHANGELOG 和归档计划必须进入本地 Git。
- 不执行远程 push。

---

### Task 1: 检查工作簿结构

**Files:**
- Create temporarily: `~temp/xlsx-preview/render-previews.mjs`
- Read: `docs/视频评分榜.xlsx`
- Read: `docs/主武器射速榜.xlsx`
- Read: `docs/速度榜.xlsx`
- Read: `docs/稀有枪械榜.xlsx`
- Read: `docs/次要装备榜.xlsx`

**Interfaces:**
- Consumes: 5 个现有 `.xlsx` 文件和加载器提供的 `@oai/artifact-tool`。
- Produces: 每个工作簿的工作表名称、已用区域和目标阵营工作表，供 Task 2 配置渲染范围。

- [ ] **Step 1: 创建临时依赖 junction**

Run:

```powershell
New-Item -ItemType Directory -Force -Path '~temp/xlsx-preview'
New-Item -ItemType Junction -Path '~temp/xlsx-preview/node_modules' -Target '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules'
```

Expected: `~temp/xlsx-preview/node_modules` 指向加载器返回的依赖目录。

- [ ] **Step 2: 创建单一检查与渲染脚本**

Create `~temp/xlsx-preview/render-previews.mjs` with:

```javascript
import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = process.cwd();
const outputDir = path.join(root, "docs", "previews");
const jobs = [
  { input: "视频评分榜.xlsx", side: "进攻方", output: "video-rating-attack.png" },
  { input: "主武器射速榜.xlsx", side: "防守方", output: "primary-rpm-defense.png" },
  { input: "速度榜.xlsx", side: "进攻方", output: "speed-attack.png" },
  { input: "稀有枪械榜.xlsx", side: "防守方", output: "rare-weapons-defense.png" },
  { input: "次要装备榜.xlsx", side: "进攻方", output: "secondary-gadgets-attack.png" },
];

for (const job of jobs) {
  const file = path.join(root, "docs", job.input);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
  const summary = await workbook.inspect({
    kind: "sheet,region",
    maxChars: 5000,
    tableMaxRows: 4,
    tableMaxCols: 8,
  });
  console.log(JSON.stringify({ input: job.input, side: job.side }));
  console.log(summary.ndjson);
}
```

- [ ] **Step 3: 运行结构检查**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' '~temp/xlsx-preview/render-previews.mjs'
```

Expected: 5 个工作簿均成功导入，输出各工作表名称和已用区域；每个目标阵营都能唯一匹配一个榜单工作表。

### Task 2: 渲染并视觉验证 5 张预览

**Files:**
- Modify temporarily: `~temp/xlsx-preview/render-previews.mjs`
- Create: `docs/previews/video-rating-attack.png`
- Create: `docs/previews/primary-rpm-defense.png`
- Create: `docs/previews/speed-attack.png`
- Create: `docs/previews/rare-weapons-defense.png`
- Create: `docs/previews/secondary-gadgets-attack.png`

**Interfaces:**
- Consumes: Task 1 识别的目标工作表和首屏范围。
- Produces: README 可直接引用的 5 张 PNG。

- [ ] **Step 1: 将脚本切换为渲染模式**

保留 Task 1 的 imports 和 `jobs`，为每个 job 补充检查后确定的 `sheetName` 与 `range`，并将循环体替换为：

```javascript
await fs.mkdir(outputDir, { recursive: true });
for (const job of jobs) {
  const file = path.join(root, "docs", job.input);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
  const preview = await workbook.render({
    sheetName: job.sheetName,
    range: job.range,
    scale: 1.5,
    format: "png",
  });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(outputDir, job.output), bytes);
  console.log(`${job.output}\t${bytes.length}`);
}
```

- [ ] **Step 2: 生成预览图**

Run:

```powershell
& '<USER_HOME>\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' '~temp/xlsx-preview/render-previews.mjs'
```

Expected: `docs/previews/` 下生成 5 个非空 PNG。

- [ ] **Step 3: 逐张视觉检查**

用图像查看工具依次打开 5 张 PNG，确认：

- 标题和阵营名称可见。
- 左侧分档完整。
- 首屏卡片未被裁剪。
- 图标、颜色和文字清晰。
- 图片没有大面积无用空白。

如果任一图片不符合，单独调整对应 `range` 或 `scale` 后重新运行脚本并复查该图片。

### Task 3: 更新 README 和变更记录

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Task 2 的 5 张 PNG 和现有 5 个 `.xlsx`。
- Produces: 可浏览、可下载的 README 预览章节和变更记录。

- [ ] **Step 1: 在 README 的“输出”章节后添加预览**

添加以下结构，实际链接均使用仓库相对路径：

```markdown
## 工作簿预览

### 视频评分榜 · 进攻方

[打开工作簿](docs/视频评分榜.xlsx)

![视频评分榜进攻方预览](docs/previews/video-rating-attack.png)
```

按设计顺序继续加入主武器射速榜防守方、速度榜进攻方、稀有枪械榜防守方和次要装备榜进攻方。

- [ ] **Step 2: 更新 CHANGELOG**

在 `Unreleased / Added` 中记录 README 的 5 张交错阵营工作簿预览及原文件链接。

- [ ] **Step 3: 验证 README 引用**

解析 README 中新增的 10 个相对路径并用 `Test-Path` 检查。

Expected: 5 个 `.xlsx` 和 5 个 PNG 全部存在。

### Task 4: 完成验证、归档和提交

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-readme-xlsx-previews.md` → `~archived/superpowers-plans/2026-07-26-readme-xlsx-previews.md`
- Commit: `README.md`
- Commit: `CHANGELOG.md`
- Commit: `docs/previews/*.png`
- Commit: `~archived/superpowers-plans/2026-07-26-readme-xlsx-previews.md`

**Interfaces:**
- Consumes: Task 1–3 的完整实现。
- Produces: 经测试、已归档计划且工作区干净的本地 Git commit。

- [ ] **Step 1: 运行项目测试**

Run:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest discover -s tests -v
```

Expected: 81 项测试全部通过。

- [ ] **Step 2: 运行 Git 与文件检查**

Run:

```powershell
git -c safe.directory=<PROJECT_ROOT> diff --check
git -c safe.directory=<PROJECT_ROOT> ls-files --others --ignored --exclude-standard
```

Expected: 无空白错误；所有被忽略文件仍只位于 `output/` 或 `~temp/`。

- [ ] **Step 3: 归档计划**

将本计划移动到 `~archived/superpowers-plans/` 并把全部复选框标为完成，确保 `docs/superpowers/plans/` 没有已完成计划。

- [ ] **Step 4: 创建实现 commit**

Run:

```powershell
git add README.md CHANGELOG.md docs/previews '~archived/superpowers-plans/2026-07-26-readme-xlsx-previews.md'
git commit -m "docs: add workbook previews to README"
```

Expected: commit 只包含 README、CHANGELOG、5 张 PNG 和归档计划；不包含 `~temp/`。

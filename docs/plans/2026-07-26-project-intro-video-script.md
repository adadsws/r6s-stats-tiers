# R6 干员中文榜单介绍视频文案 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 `docs/` 创建一份面向普通《彩虹六号》玩家、时长 2–3 分钟的日常风格项目介绍视频文案与逐镜头素材提示。

**Architecture:** 使用单一 Markdown 文档承载分镜表、连续旁白、素材清单和剪辑注意事项。分镜按“五榜轮播型”组织，并以仓库现有预览图、PDF、GitHub 页面和可后期录制的屏幕画面作为素材来源。

**Tech Stack:** Markdown、PowerShell、Git

## Global Constraints

- 受众是普通《彩虹六号》玩家，不要求技术背景。
- 成片时长按正常中文语速控制在 2–3 分钟。
- 风格节奏明快、表达日常，不使用夸张宣传语。
- 五个榜单均独立介绍，并展示进攻方与防守方。
- 只编写文案与素材提示，不生成、下载或嵌入新图片和视频。
- 不要求展示 Microsoft Excel。

---

### Task 1: 编写并验证介绍视频文案

**Files:**
- Create: `docs/项目介绍视频文案.md`
- Read: `README.md`
- Read: `~archive/superpowers-specs/2026-07-26-project-intro-video-script-design.md`
- Move after completion: `docs/plans/2026-07-26-project-intro-video-script.md` → `docs/finished_plans/2026-07-26-project-intro-video-script.md`

**Interfaces:**
- Consumes: README 中的五榜功能、十张双阵营预览图、五份 PDF 路径和已确认设计。
- Produces: 可直接交给配音与剪辑人员使用的 `docs/项目介绍视频文案.md`。

- [ ] **Step 1: 核对现有素材路径**

运行：

```powershell
@(
  'docs/previews/video-rating-attack.png',
  'docs/previews/video-rating-defense.png',
  'docs/previews/primary-rpm-attack.png',
  'docs/previews/primary-rpm-defense.png',
  'docs/previews/speed-attack.png',
  'docs/previews/speed-defense.png',
  'docs/previews/rare-weapons-attack.png',
  'docs/previews/rare-weapons-defense.png',
  'docs/previews/secondary-gadgets-attack.png',
  'docs/previews/secondary-gadgets-defense.png',
  'docs/视频评分榜.pdf',
  'docs/主武器射速榜.pdf',
  'docs/速度榜.pdf',
  'docs/稀有枪械榜.pdf',
  'docs/次要装备榜.pdf'
) | ForEach-Object {
  if (-not (Test-Path -LiteralPath $_)) { throw "缺少素材：$_" }
}
```

预期：命令成功结束，没有“缺少素材”错误。

- [ ] **Step 2: 编写分镜表**

创建 `docs/项目介绍视频文案.md`，分镜总长 2 分 20 秒，按以下时间段编排：

```text
0:00–0:15  开场：选干员时信息分散
0:15–0:30  项目总览：五榜集中查看
0:30–0:47  视频评分榜
0:47–1:04  主武器射速榜
1:04–1:21  速度榜
1:21–1:38  稀有枪械榜
1:38–1:55  次要装备榜
1:55–2:10  双阵营、补丁、PDF/XLSX 与更新
2:10–2:20  GitHub 收尾
```

每段必须填写时间、建议画面、旁白、屏幕字幕、素材提示、剪辑提示六项。

- [ ] **Step 3: 补充连续旁白、素材清单和剪辑注意事项**

连续旁白必须与分镜旁白逐字一致，方便直接录音。素材清单必须逐项列出十张预览图、五份 PDF、GitHub 页面录屏、报告滚动录屏和可选的授权对局素材。剪辑注意事项必须说明字幕长度、双栏画面可读性、画面停留时间、音乐音量、隐私检查及不展示 Microsoft Excel。

- [ ] **Step 4: 验证文档完整性**

运行：

```powershell
$path = 'docs/项目介绍视频文案.md'
$text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
@('视频评分榜', '主武器射速榜', '速度榜', '稀有枪械榜', '次要装备榜',
  '进攻方', '防守方', '完整连续旁白', '素材清单', '剪辑注意事项') |
  ForEach-Object {
    if (-not $text.Contains($_)) { throw "缺少内容：$_" }
  }
if ($text -match 'TBD|TODO|待补充') { throw '文档含未完成占位符' }
git diff --check -- $path
```

预期：命令成功结束，无缺失内容、占位符或空白错误。

- [ ] **Step 5: 提交最终文档并归档计划**

将本计划移动到 `docs/finished_plans/2026-07-26-project-intro-video-script.md`，然后运行：

```powershell
git add -- 'docs/项目介绍视频文案.md' `
  'docs/plans/2026-07-26-project-intro-video-script.md' `
  'docs/finished_plans/2026-07-26-project-intro-video-script.md'
git commit -m "docs: add project intro video script"
```

预期：提交成功；工作区干净；`docs/plans/` 不再保留已完成计划。

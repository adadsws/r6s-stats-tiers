# R6 项目根目录精简设计

## 目标

在保留既有评分换算、榜单规则和卡片信息的前提下，完成指定的数据来源、补丁区间和排版改进，将项目整理为标准的 `src` 包结构，并把根目录入口统一为一个 `run_r6_report.bat`。

完整数据流分为三个明确阶段：

1. 用户或 Agent 使用项目 Skill，查找并人工核验 Athieno 最新的 R6 干员 Tier 视频，将完整评分保存为 `data/athieno/latest.json`。
2. 用户运行唯一的 BAT。采集脚本从灰机 Wiki 获取最新干员、武器、配置、图标和补丁，验证后保存到 `data/wiki/`、`data/icons/` 与 `data/patches/`。
3. 生成脚本只读取 `data/athieno/latest.json` 以及采集器保存的 `data/wiki/`、`data/icons/`、`data/patches/`，生成 `data/r6_operator_stats.xlsx`，再生成五个榜单工作簿到 `output/`。

任一步失败时必须立即停止并保留原始退出码。不得使用内存中尚未落盘的数据，也不得在采集失败后把新评分与旧 Wiki 快照静默组合。

## 最终根目录

根目录只保留下列必要文件：

- `run_r6_report.bat`
- `README.md`
- `CHANGELOG.md`
- `requirements.txt`
- `.gitignore`

根目录保留下列必要目录：

- `src/`：项目 Python 源码。
- `tests/`：自动化测试。
- `data/`：评分、Wiki 快照、图标和基础工作簿，Git 忽略。
- `output/`：五个榜单工作簿，Git 忽略。
- `docs/`：设计、计划和其他项目文档。
- `.codex/`：项目 Skill。
- `~archived/`：不再使用的旧文件，Git 忽略。
- `~temp/`：临时文件，Git 忽略。
- `.git/`、`.agents/`：仓库和代理运行所需的隐藏目录。

现有 `.tools/` 只剩一个被后台 Excel 进程锁定的零字节临时文件。Windows 释放文件锁后，将整个 `.tools/` 归档到本次整理对应的 `~archived/` 子目录；整理过程不强制终止跨会话 Excel 进程，也不删除该文件。

## Python 包结构

现有根目录 Python 文件迁移为：

| 当前文件 | 目标文件 |
| --- | --- |
| `r6_operator_stats.py` | `src/r6_report/operator_stats.py` |
| `r6_leaderboards.py` | `src/r6_report/leaderboards.py` |
| `r6_patch_notes.py` | `src/r6_report/patch_notes.py` |
| `r6_tier_chart.py` | `src/r6_report/tier_chart.py` |
| `r6_tiers.py` | `src/r6_report/tiers.py` |

新增：

- `src/r6_report/__init__.py`
- `src/r6_report/collector.py`：负责 Wiki 表数据及图标的联网采集和本地快照。

包内导入改为相对导入，例如 `.patch_notes`、`.tier_chart` 和 `.tiers`。现有业务规则、公开函数和数据模型尽量保持不变；`operator_stats` 改为从经过验证的本地快照生成基础工作簿，不再在写工作簿时联网。

命令行入口改为模块调用：

```text
python -m r6_report.collector
python -m r6_report.operator_stats
python -m r6_report.leaderboards
python -m r6_report.tier_chart
```

`collector` 默认写入 `data/wiki/`、`data/icons/` 和 `data/patches/`。`operator_stats` 默认读取 `data/athieno/latest.json` 及采集快照，输出 `data/r6_operator_stats.xlsx`。`leaderboards` 默认读取该文件并继续输出到 `output/`。`tier_chart` 保留独立命令行兼容性，其默认输入改为 `data/r6_operator_stats.xlsx`，默认输出改为 `output/传统分级简图.xlsx`，但不再提供单独 BAT。

## Athieno Skill 工作流

项目 Skill 从“固定 Y11S2 报告 Skill”升级为“获取 Athieno 最新 R6 Tier 并构建报告”的项目工作流。用户或 Agent 调用 Skill 时必须：

1. 在 Athieno 官方频道中确认最新一条明确面向当前赛季的 R6 干员 Tier List 视频，不用搜索结果中的搬运、剪辑或其他主播视频。
2. 记录视频标题、URL、视频 ID、发布日期、适用赛季/补丁和用于抄录的最终榜单画面时间。
3. 人工查看最终榜单画面并逐档抄录，不能根据口播、字幕、缩略图或模型推测评级。
4. 保留视频出现的原始 Tier 名称，并使用当前评分映射生成数值；`boof` 继续按最低档处理。
5. 写入 `data/athieno/latest.json` 前校验名称唯一、每名干员只出现一次、Tier 均有合法分值、来源元数据完整。
6. 若已有评分快照，先归档到 `~archived/data-snapshots/athieno/<时间戳>/`，再原子替换 `latest.json`。

Skill 获取评分时只负责生成评分快照，不直接生成 Excel。若无法可靠读取最终榜单画面，必须停止并说明缺少哪些信息，不得补猜。

`data/athieno/latest.json` 延续现有 JSON 的 `source`、`score_map` 和 `tiers` 结构，并在 `source` 中新增：

- `video_id`
- `season`
- `covered_patch`
- `covered_through`
- `captured_at`

`covered_patch` 是视频明确适用或已覆盖的补丁，`covered_through` 是该覆盖范围对应的 ISO 日期。评分文件中的干员集合最终由生成脚本与最新 Wiki 干员集合交叉验证。新增或缺失干员会导致明确失败，由用户或 Agent 重新核对视频，而不是静默补分。

## Wiki 数据和图标快照

采集器使用现有灰机 Wiki 官方只读 API 和重试规则，保存以下 UTF-8 JSON：

- `data/wiki/operator.json`
- `data/wiki/weapon.json`
- `data/wiki/weapon_config.json`
- `data/wiki/manifest.json`

`manifest.json` 记录来源页面、API 标题、当前赛季、当前补丁、ISO 抓取时间、记录数和快照格式版本。前三个文件保存已通过 Tabx 结构验证的记录，生成器不再重复解析网页或请求网络。

图标保存到：

- `data/icons/operator/white/`
- `data/icons/operator/badge/`
- `data/icons/gadget/`

采集过程先写入 `~temp/`。只有三张表、全部当前干员 Badge、可取得的白色图案和所有已引用次要装备图标都通过格式及来源域名校验后，才原子更新 `data/`。更新前的有效快照归档到 `~archived/data-snapshots/wiki/<时间戳>/`。采集失败时保留上一个完整快照，但本次 BAT 立即失败，不继续生成工作簿。

## 补丁区间

采集器同时读取灰机 Wiki 更新补丁总表及对应补丁页面，把补丁索引、相关变更和来源保存到：

- `data/patches/patches.json`
- `data/patches/manifest.json`

补丁区间的下界取 `data/athieno/latest.json` 中视频已覆盖的 `covered_through`，上界取 `data/wiki/manifest.json` 的 `fetched_at`。纳入发布日期满足以下条件的每一个补丁：

```text
covered_through < 补丁发布日期 <= fetched_at
```

若只能确认视频发布日期而不能确认视频覆盖到哪个补丁，Skill 必须把 `covered_through` 设为视频发布日期，并在评分来源说明中标为“按发布日期推定”；不得由脚本猜测一个补丁编号。

`patches.json` 按发布日期从旧到新保存，每个补丁记录名称、发布日期、赛季、灰机 Wiki URL、官方 URL，以及影响干员、武器、速度、装备或本报告显示字段的变更。区间内没有相关字段变更的补丁也必须保留，并标记“无影响本报告字段的变更”，证明该时间段没有被跳过。

生成器验证补丁区间连续、日期合法、来源为 HTTPS，且每条变更只属于一个补丁。`补丁说明` 工作表按“日期 + 补丁”建立独立分组，从旧到新排列；每组先显示补丁标题、发布日期和来源链接，再显示该补丁的增强、削弱及混合变更。

## 工作表来源区

每一个工作表，包括 `补丁说明`，都在主体内容下方空一行后显示统一的三行来源区：

1. `评分来源`：Athieno 视频标题、适用赛季、已覆盖补丁、发布日期、人工核验时间和视频 URL。
2. `游戏数据`：灰机 Wiki 当前赛季、当前补丁、抓取时间，以及干员/武器数据来源 URL。
3. `补丁区间`：视频覆盖日期、Wiki 数据抓取时间、实际纳入的最早和最新补丁，以及更新补丁总表 URL。

来源时间使用带时区的 ISO 8601 格式。URL 使用可点击的 HTTPS 超链接。来源区采用统一的浅灰底和紧凑字号，不进入表格筛选范围，但必须进入打印区域。原先只写“除评分外其他信息已是最新”的单行状态说明由该来源区替代。

若来源赛季、补丁或时间缺失，生成器立即失败，不生成缺少版本依据的工作簿。

## 榜单卡片排版

每行仍显示五名干员，卡片字段和信息顺序保持不变。调整统一适用于五个榜单：

- 五张卡之间加入无填充、无边框的窄间隔列，卡片不再共用相邻边界。
- Badge 在自己的图标列内靠右放置，使它与本卡文字距离更近，与上一张卡距离更远。
- 信息列宽略微收紧，但最长干员名、Tier、速度和射速仍须完整显示。
- 每张卡的外边框只包围自己的 Badge、文字和次要装备，不跨越间隔列。

次要装备不再把每个 token 都以不断增大的偏移锚定在卡片首列。生成器为每张卡制作一张透明的“装备图标带”，在图标带内部依次绘制全部装备及数量；次要装备榜需要标红的单个装备仍在图标带内部保留红框。图标带使用一个合法的单元格锚点并限制在本卡宽度内。

当前最多装备的 Striker 和 Sentry 各有 7 项装备，必须作为回归用例。1 至 7 项装备均保持原始顺序、数量可读、不重叠、不越过卡片边界，也不覆盖下一张卡。

稀有枪械榜左侧分档名称固定为：

```text
副喷
主狙
副自
都无
```

分档判定和允许干员重复出现的规则不变。

## 单一 BAT 工作流

`run_r6_report.bat` 从自身所在目录启动，并按现有兼容方式优先使用 `python`，不可用时尝试 `py -3`。两者都不存在时给出明确错误并返回 `1`。

BAT 将 `PYTHONPATH` 指向项目的 `src`，首先检查 `data/athieno/latest.json` 是否存在且格式有效。评分缺失时提示用户或 Agent 先调用项目 Skill，并返回非零状态。

评分有效后运行采集器：

```text
python -m r6_report.collector --data-dir data
```

采集成功后生成基础表：

```text
python -m r6_report.operator_stats ^
  --data-dir data ^
  --ratings data\athieno\latest.json ^
  --output data\r6_operator_stats.xlsx
```

基础表成功后生成五个榜单：

```text
python -m r6_report.leaderboards ^
  --input data\r6_operator_stats.xlsx ^
  --output-dir output ^
  --icons-dir data\icons\operator\badge ^
  --gadget-icons-dir data\icons\gadget
```

每条命令结束后立即检查 `%ERRORLEVEL%`。失败时打印发生失败的阶段和退出码，直接 `exit /b`；全部成功时打印评分来源、Wiki 抓取时间、基础表路径和输出目录。BAT 最终保留 `pause`，方便双击运行时查看结果。

旧的 `run_r6_operator_stats.bat` 和 `run_r6_leaderboards.bat` 归档，不删除。

## 数据和输出

- 当前评分文件迁移为初始 `data/athieno/latest.json`，保留其已核验的视频来源元数据。
- 根目录现有 `r6_operator_stats.xlsx` 移到 `data/r6_operator_stats.xlsx`。
- 当前 `assets/operator-icons/` 和 `assets/gadget-icons/` 作为初始快照迁移到 `data/icons/`；后续由采集器维护。
- 五个榜单工作簿继续位于 `output/`，文件名改为：

```text
视频评分榜.xlsx
主武器射速榜.xlsx
速度榜.xlsx
稀有枪械榜.xlsx
次要装备榜.xlsx
```

- 旧英文文件名的工作簿归档，不与新文件并存于 `output/`。
- `data/` 与 `output/` 始终由 `.gitignore` 忽略。
- 新入口每次都先更新 Wiki 和图标快照，再重新生成基础表；不允许在联网步骤失败后继续使用已有基础表。

## 测试

测试代码改为从 `r6_report` 包导入。增加统一的测试路径引导，使标准库 `unittest` 在未安装项目包时也能从 `src/` 加载代码。

验证命令为：

```powershell
$env:PYTHONPATH='src'
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest discover -s tests -v
```

现有 56 项功能测试必须继续通过，并补充或更新启动器测试，验证：

- 根目录只有 `run_r6_report.bat` 一个 BAT。
- 评分文件缺失或无效时，BAT 不启动 Wiki 采集。
- BAT 先采集 Wiki 数据和图标，再生成基础表，最后生成榜单。
- 任一阶段失败时不会运行后续阶段。
- 任一阶段失败时返回该阶段的非零退出码。
- 默认输入、输出路径分别位于 `data/` 和 `output/`。
- 采集器只有在完整验证后才替换现有快照。
- 生成器运行时不发起网络请求。
- 最新评分与最新 Wiki 干员集合不一致时明确报错。
- 每张工作表底部都有完整的评分、Wiki 和补丁来源区。
- 补丁只覆盖视频数据与 Wiki 快照之间的完整时间区间，并按补丁日期分组。
- 1 至 7 项次要装备在卡片内正确显示，所有图片锚点均合法。
- 五张卡之间有独立间隔列，Badge 更靠近本卡文字。
- 稀有枪械榜左侧只显示 `副喷`、`主狙`、`副自` 和 `都无`。
- `output/` 只包含五个规定的中文文件名。

完成后使用已核验评分执行一次真实联网冒烟测试，确认 Wiki/补丁 JSON、全部图标、基础表及五个榜单都可以重新生成并由 `openpyxl` 打开。再通过工作区提供的 `artifact-tool` 渲染所有工作表，逐页确认来源区、补丁分组、图标带和卡片间距没有裁切、重叠或越界。

## 文档和 Skill

- 更新 `README.md`，只展示新的 BAT 用法、模块命令、目录结构和测试命令。
- 更新 `CHANGELOG.md`，记录根目录整理、包迁移、单入口工作流和路径变化。
- 更新 `.codex/skills/build-r6-operator-report/SKILL.md`，使其负责发现、人工核验、保存最新 Athieno Tier，并说明随后运行项目 BAT。
- Skill 自带的旧脚本和固定评分参考归档到项目 `~archived/`；项目源码是唯一报告生成实现，避免 Skill 内继续维护重复副本。
- Skill 的参考文档记录评分 JSON 契约、视频核验清单、`covered_patch`/`covered_through` 判定和失败条件。
- 修改后再次运行 Skill 校验器。

## 归档规则

本次不删除任何旧文件。旧 BAT、已迁移的根目录 Python 文件及其他不再使用的入口统一归档到 `~archived/2026-07-25-root-cleanup/` 的对应子目录。若归档目标已存在，则使用清晰的子目录避免覆盖。

`.gitignore` 必须继续包含：

```text
~archived/
~temp/
data/
output/
```

## 非目标

- 不修改干员数据抓取和三表关联规则。
- 不修改自动枪械、主手半自、副喷和次要装备的判定。
- 不修改视频 Tier 到分数的换算、榜单分级和排序。
- 不修改 Badge 和次要装备图标的图案或颜色。
- 不新增安装包、构建系统或额外根目录配置文件。

## 验收标准

1. 根目录只有一个 `.bat` 文件，名称为 `run_r6_report.bat`。
2. 根目录没有 Python 源文件和工作簿。
3. 用户或 Agent 能按 Skill 指引把 Athieno 最新 Tier 完整保存为 `data/athieno/latest.json`。
4. 双击 BAT 能按顺序更新 `data/wiki/`、`data/icons/`、`data/patches/`，生成 `data/r6_operator_stats.xlsx` 和 `output/` 下五个榜单。
5. 生成 Excel 的阶段只读取 `data/` 中已验证的评分来源组和 Wiki 采集来源组，不访问网络。
6. 任一阶段失败时立即停止并返回非零状态。
7. 每个工作表底部注明各数据来源、更新赛季、补丁和带时区时间。
8. `补丁说明` 完整覆盖最旧评分来源到最新 Wiki 快照之间的补丁，并按补丁时间分组。
9. 最多 7 项次要装备的卡片无重叠、越界或跨卡显示。
10. 卡片间距、Badge 位置、稀有枪械分档标签符合本规范。
11. `output/` 中五个工作簿全部使用规定的中文文件名。
12. 全部自动化测试、Skill 校验和联网冒烟测试通过。
13. 所有仓库修改均有 Git 提交记录，旧文件只归档、不删除。

# 本地归档说明

`~archived/` 保存项目历史文件、旧数据快照、一次性 QA 资产和系统修复材料。本地 Git 必须继续跟踪该目录，但远程脱敏导出不得包含目录内容及其历史；远程仓库仅保留本文件作为同名说明。

## 不推送原因

- 全局规则禁止推送 `~archived/`。
- 2026-07-26 审计基准为约 299.44 MiB、3240 个文件。
- 目录包含约 110.16 MiB 的 PowerShell 安装包和约 83.58 MiB 的 FFmpeg 可执行文件，会显著增大仓库，且前者超过常见远程 Git 单文件限制。
- 当前报表生成流程不依赖这些归档文件。

## 主要内容

| 本地目录 | 内容 | 审计大小 |
|---|---|---:|
| `2026-07-25-root-cleanup/` | 旧 QA 工具、视频核对资产、PowerShell 修复包与日志 | 约 286.73 MiB |
| `2026-07-25-project-rebuild/` | 重构前脚本、旧输出和重复资产 | 约 6.84 MiB |
| `2026-07-25-before-leaderboards/` | 榜单重做前的工作簿 | 约 3.30 MiB |
| `data-snapshots/` | 历史 Wiki、补丁和图标输入快照 | 约 2.47 MiB |
| `superpowers-plans/` | 已完成并归档的实施计划 | 约 0.09 MiB |
| `2026-07-26-agents-compliance/` | 旧规则文件和临时工具残留 | 小于 0.01 MiB |

## 获取与重建

- 完整历史归档只能从本地 Git 仓库或本地备份恢复；制作远程导出时不得复制对应提交历史。
- 当前输入可运行 `run_r6_report.bat` 重新采集到 `data/`，当前工作簿可由同一命令重新生成到 `output/`。
- PowerShell 安装包来自 PowerShell 官方发行渠道，FFmpeg 二进制来自 `imageio-ffmpeg` 依赖；需要时应从各自官方来源重新下载，不从远程仓库分发归档副本。
- 可用以下 PowerShell 命令重新统计归档容量：

```powershell
$files = Get-ChildItem -LiteralPath '~archived' -Force -Recurse -File
$files | Measure-Object -Property Length -Sum
```

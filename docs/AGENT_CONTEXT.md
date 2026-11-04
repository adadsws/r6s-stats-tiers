# 开发上下文

## 技术架构

项目是一个 Python 3.9+ 的离线报表生成器。`r6_report.collector` 联网采集并原子更新数据快照；`r6_report.operator_stats` 将评分、Wiki、补丁和图标合成为临时基础工作簿；`r6_report.leaderboards` 再生成五类 XLSX、PDF 与逐页 PNG。`src/rebuild_leaderboards.bat` 仅在需要重建榜单时串联三阶段，并在任一阶段失败时停止。

## 项目结构

- `src/r6_report/`：采集、数据校验、分类、主题和报表渲染实现。
- `data/`：Git 跟踪的 Athieno 评分、Wiki、补丁和图标数据快照；远程发布前必须检查公开边界。
- `tests/`：基于 `unittest` 的单元、集成和布局契约测试。
- `skills/build-r6-operator-report/`：只负责人工核对 Athieno 最新完整 Tier List。
- `docs/`：项目规则、开发规范、变更记录、用户可下载的固定报表示例、当前计划、已完成计划和经验记录。
- `~temp/`：可重建的中间工作簿、采集暂存和日志，不进入 Git。
- `~outputs/`：待检查或交付的生成结果，不进入 Git。
- `~archive/`：已确认封存的历史文件，本地 Git 保留但禁止直接远程发布。
- `README.md` / `README_EN.md`：中文和英文用户入口；报告内容仍为中文。

## 开发流程

使用 Codex bundled Python 或按 `requirements.txt` 创建的 Python 3.9+ 虚拟环境。修改行为前先写失败测试；验证命令为 `$env:PYTHONPATH='src'; $env:PYTHONUTF8='1'; python -m unittest discover -s tests -v`。报表渲染变化还需运行 `src/rebuild_leaderboards.bat`，人工核对 `~outputs/` 中的 XLSX、PDF 和 PNG，再决定是否更新 `docs/` 示例。

`tier_chart` 的 CLI 只从 `data/` 读取已验证的次要装备图标；采集所用下载辅助函数仅由 `collector` 通过 `wiki_client` 调用。

## 已知问题

- 一键入口面向 Windows，并依赖可调用的 Python 3.9+ 与 `curl.exe`。
- Wiki 与视频来源可能变化或临时不可用；collector 失败时会保留上一份有效数据快照，但不会继续生成报表。
- 本机仓库所有者与执行账户不一致时，Git 需要单次 `-c safe.directory=D:/CODE/PY/r6_num` 参数；不要为此修改项目配置。
- 图像测试依赖 Pillow 12 的像素迭代 API；旧环境需按当前 `requirements.txt` 重建。

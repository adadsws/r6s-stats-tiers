# 项目专用规则

- 报表流水线固定为 `collector` → `operator_stats` → `leaderboards`；前一阶段失败时不得继续执行后续阶段。
- 只有 `collector` 和 `skills/build-r6-operator-report/` 可以联网。工作簿、PDF 与 PNG 生成阶段必须只读取 `inputs/` 中已验证的快照。
- `run_r6_report.bat` 必须保持 UTF-8（无 BOM）、CRLF、从脚本所在目录启动，并完整传递非零退出码。
- 新生成的报表先写入 `~outputs/`。只有经过人工检查、用于 README 展示的固定示例才复制到 `docs/`，并与相关实现一同更新。
- 调整榜单布局或渲染时，同步验证 XLSX、PDF 和逐页 PNG 三种输出，避免只修复单一格式。

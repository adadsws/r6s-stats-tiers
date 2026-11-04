# 项目专用规则

- `docs/plans/` 同一时间最多保留一个当前 Plan（`README.md` 不计）；开始独立任务前先完成并归档当前 Plan。同一任务的规范、补充规范、需求及其补充、澄清、变更、样例、验收标准和验证记录统一更新到该 Plan，不另建并行任务文档。正式项目规则、源码说明和必要交付文档仍按各自职责维护。
- 报表流水线固定为 `collector` → `operator_stats` → `leaderboards`；前一阶段失败时不得继续执行后续阶段。
- 只有 `collector` 和 `skills/build-r6-operator-report/` 可以联网。工作簿、PDF 与 PNG 生成阶段必须只读取 `data/` 中已验证的快照。
- `src/rebuild_leaderboards.bat` 仅用于重新采集数据和重建榜单；必须保持 UTF-8（无 BOM）、CRLF，从脚本位置定位项目根目录，并完整传递非零退出码。
- 新生成的报表先写入 `~outputs/`。只有经过人工检查、用于 README 展示的固定示例才复制到 `docs/`，并与相关实现一同更新。
- 调整榜单布局或渲染时，同步验证 XLSX、PDF 和逐页 PNG 三种输出，避免只修复单一格式。
- 运行时代码不得把 Markdown 当作数据输入；布局与契约测试可以只读项目 Markdown 和 Skill 文档，以验证路径、命令及治理约束没有回退。

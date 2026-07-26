# Athieno Tier Skill 单一职责设计

## 目标

将项目 Skill 从 `.codex/skills/build-r6-operator-report/` 迁移到根目录
`skills/build-r6-operator-report/`，并将职责收缩为从 Athieno 官方最新完整
Tier List 视频人工核对数据、更新 `data/athieno/latest.json`。

## Skill 边界

Skill 只执行以下工作：

1. 查找 Athieno 官方 YouTube 频道最新的完整干员 Tier List。
2. 核对标题、发布日期、赛季、覆盖补丁和最终完整榜单画面时间。
3. 逐档人工核对所有干员，验证每名干员恰好出现一次。
4. 按固定映射生成完整来源信息、`score_map` 和七个 `tiers` 数组。
5. 仅在信息可确认且数据完整时覆盖 `data/athieno/latest.json`。

Skill 不采集灰机 Wiki，不运行项目流水线，不生成 XLSX、PDF 或预览图，也不负责
报表测试和视觉检查。上述工作仍由 README 的后续本地部署步骤说明。

## 文件与发现方式

- Skill 目录固定为 `skills/build-r6-operator-report/`。
- `SKILL.md` 保留现有名称 `build-r6-operator-report`，避免破坏 README 中的调用方式。
- `agents/openai.yaml` 同步改为只描述 Athieno Tier 更新。
- 旧目录 `.codex/skills/build-r6-operator-report/` 不再存在。
- 根目录 `skills/` 由 Git 跟踪，发布时随普通源码上传。

## 验证

- 项目布局测试检查新路径存在、旧路径不存在。
- 测试检查 Skill 包含 Athieno 视频核对和 `data/athieno/latest.json` 数据契约。
- 测试禁止 Skill 出现 Wiki 采集、报表生成、XLSX、PDF、Poppler、
  Microsoft Excel 或 LibreOffice 职责。
- 使用 Skill Creator 的 `quick_validate.py` 验证目录、名称和 YAML frontmatter。

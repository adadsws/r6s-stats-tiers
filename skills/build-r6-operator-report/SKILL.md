---
name: build-r6-operator-report
description: Use when finding Athieno's latest complete Rainbow Six Siege operator Tier List video, manually verifying its final ranking frame, and updating this project's data/athieno/latest.json rating source.
---

# 更新 Athieno 干员 Tier

只从 Athieno 官方最新完整 Tier List 视频获取评分数据，并更新
`data/athieno/latest.json`。不要执行项目中的其他采集、生成或验证任务。

## 查找视频

1. 只搜索 Athieno 官方 YouTube 频道。
2. 选择最新的完整干员 Tier List；排除 Shorts、片段、其他创作者和没有最终完整榜单画面的视频。
3. 记录视频标题、URL、视频 ID、发布日期、赛季、覆盖补丁和最终完整榜单画面时间。
4. 无法确认视频覆盖版本或最终完整榜单画面时停止，不修改旧文件。

## 核对 Tier

1. 以 `final_frame` 对应的最终完整榜单画面为准，逐档人工读取所有干员。
2. 不根据自动字幕、口头评价或片段画面猜测档位。
3. 使用项目现有 ASCII 干员键；核对每名当前干员恰好出现一次，不得遗漏或重复。
4. 固定使用以下映射：

```text
S=100
A=85
B=70
C=55
D=40
F=20
boof=0
```

## 写入数据

将结果写入 `data/athieno/latest.json`，保留现有 JSON 结构并完整包含：

- `source`：`creator`、`title`、`url`、`video_id`、`published`、`season`、
  `covered_patch`、`covered_through`、`coverage_basis`、`final_frame`、`captured_at`
- 完整 `score_map`
- `S`、`A`、`B`、`C`、`D`、`F`、`boof` 七个 `tiers` 数组

写入前验证 JSON 可解析、七档成员互斥、所有干员恰好出现一次、`score_map` 与档位映射一致。
验证失败时停止并保留旧文件。完成后只报告视频来源、最终画面时间、覆盖版本、干员数量和
`data/athieno/latest.json` 路径。

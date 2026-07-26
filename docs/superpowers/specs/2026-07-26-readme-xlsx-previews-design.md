# 榜单 PDF 与 README 预览设计

## 目标

在现有 5 个榜单 XLSX 之外同步生成 5 个同名 PDF，并在根目录 `README.md` 中为每个榜单并排展示进攻方与防守方预览及对应 PDF 链接。PDF 不依赖 Microsoft Excel 或 LibreOffice。

## 生成架构

- `leaderboards.py` 保持现有 XLSX 生成行为，并把同一份干员数据、榜单规格、排序结果和补丁来源交给独立 PDF 渲染模块。
- 新模块使用 `ReportLab` 在 A2 内存画布上排版 PDF，再按每页实际内容边界收缩高度，避免从 XLSX 反向解析或调用办公软件转换。
- 每个 PDF 固定为“进攻方榜单、防守方榜单、补丁说明”三张大页，页面保持 420 mm 宽并使用各自的内容适配高度，任何章节都不自动换页。
- `write_all_leaderboards` 返回 10 个输出路径：每个榜单先列 `.xlsx`，再列同名 `.pdf`；CLI 逐一打印绝对路径。
- `requirements.txt` 包含 `reportlab`、`pypdf` 和 `pdfplumber`，分别负责初始排版、页面边界调整与内容边界测量。

## PDF 布局

- 页面先使用 A2 排版以确保章节不换页，再保留顶部内容与安全底边并缩短 `MediaBox`/`CropBox`。
- 每页显示标题、阵营或章节、榜单维度和页码。
- 榜单卡片每行最多 5 名干员，包含干员图标、名称、视频 Tier 与补丁方向、速度、主狙/副喷状态、主副手自动枪械射速和全部次要装备图标。
- 分档使用与 XLSX 一致的颜色并以整行色带标识；同一分档超过一行时在色带下继续排列卡片。
- PDF 文本保持可搜索，图标按原始宽高比缩放。
- 补丁说明按补丁发布日期从旧到新排列，使用与 XLSX 相同的增强、削弱和混合颜色，并显示来源 URL。
- 页脚显示评分来源、Wiki 快照时间、补丁覆盖区间和页码。
- 中文使用 ReportLab 的 `STSong-Light` CID 字体，英文与数字使用 Helvetica；不依赖系统 Office 字体。

## 输出文件

统一入口完成后在 `output/` 生成：

```text
视频评分榜.xlsx
视频评分榜.pdf
主武器射速榜.xlsx
主武器射速榜.pdf
速度榜.xlsx
速度榜.pdf
稀有枪械榜.xlsx
稀有枪械榜.pdf
次要装备榜.xlsx
次要装备榜.pdf
```

仓库中的可发布示例保存在 `docs/`。5 个 XLSX 保持原样；本次新增 5 个同名 PDF，README 链接 PDF。

## 展示顺序

5 个 PDF 按以下顺序展示，每个榜单左列为进攻方、右列为防守方：

1. `视频评分榜.pdf`
2. `主武器射速榜.pdf`
3. `速度榜.pdf`
4. `稀有枪械榜.pdf`
5. `次要装备榜.pdf`

## README 文件布局

- 预览图保存到 `docs/previews/`：
  - `video-rating-attack.png`、`video-rating-defense.png`
  - `primary-rpm-attack.png`、`primary-rpm-defense.png`
  - `speed-attack.png`、`speed-defense.png`
  - `rare-weapons-attack.png`、`rare-weapons-defense.png`
  - `secondary-gadgets-attack.png`、`secondary-gadgets-defense.png`
- README 在现有“输出”说明之后增加“报告预览”章节。
- 每项包含榜单名称、指向 `docs/<榜单名>.pdf` 的链接，以及“进攻方｜防守方”Markdown 双列表格。
- 图片使用描述性中文替代文本；双栏由 Markdown 表格控制，不拼接图片、不使用 HTML 固定宽度。

## 预览渲染规则

- 使用 Poppler 将每个 PDF 的第 1 页进攻方和第 2 页防守方分别渲染为 PNG。
- 预览覆盖完整 PDF 页面，保留标题、分档、卡片和页脚。
- PNG 应清晰可读，不拉伸、不裁掉关键内容。

## 验证

- 单元测试确认 5 个 XLSX 和 5 个 PDF 同时生成；每个 PDF 恰好三页，按页包含进攻方、防守方和补丁说明，且各页高度按内容收缩并保留页码。
- 使用 `pypdf` 检查 PDF 可打开、页面非空且关键标题可提取。
- 使用 Poppler 渲染全部 10 张阵营预览并进行视觉检查，确认标题、分档、卡片、图标和页脚完整可见。
- 检查 README 中 5 个 `.pdf` 链接和 10 个 PNG 路径均存在，且每个双列表格左侧为进攻方、右侧为防守方。
- 运行全部单元测试。
- 确认 `output/` 与 `~temp/` 仍是仅有的 Git 忽略目录，新增 PDF 和预览图均进入本地 Git。

## Git 边界

- 用户提供的 5 个原始工作簿继续保持在独立 commit 中。
- 更新后的设计规格单独 commit。
- PDF 生成代码、测试、依赖、示例 PDF、预览图、README、CHANGELOG 和归档计划作为实现 commit。
- 不执行远程 push。

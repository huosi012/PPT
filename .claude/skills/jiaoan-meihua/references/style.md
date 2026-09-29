# 蓝金风格规范

## 配色（三层 + 点缀）

| 角色 | 色值 | 用途 |
|:---|:---|:---|
| 主色·藏蓝 | `1F3864` | 页首横幅、分区标题栏（01/02/03/04）；标签文字、小标题文字 |
| 强调·金 | `C9A227` | 横幅下细线、分区编号、表格最左侧金边（sz 24）、▍短竖条 |
| 深金（文字） | `A8841A` | 编号数字 1/2/3、(1)、①、环节序号 01–04 |
| 标签底·浅蓝灰 | `EEF3F9` | 大类标签、授课信息标签、表头行 |
| 卡片底 | `F7F9FC` | 小类标签、内容卡片、环节名称列 |
| 细线 | `D5DEEA` | 所有单元格边框（sz 4） |
| 正文 / 说明 / 占位 | `333333` / `666666` / `AAAAAA` | |
| 关键词标签 | 底 `F6EDD2` 字 `7A5C00` | “长对正、高平齐、宽相等”等（全篇只高亮一次，放在“教学重点”） |
| 课程思政 | 标题 `9B1C1C`，底 `FAF6F0`，顶边细红线 | 不用粗红竖条 |
| 设计意图列 | 底 `FBF8F1`，字 `5C5346`；表头 `F3EEE2` | 让“意图”读起来像注释 |

不要再引入：中蓝大色块（会比页首还抢眼）、橙色、正红。图标只放在分区标题、大类标签、表头；小类标签和卡片标题不放图标（用 ▍短竖条）。

## 字体与字号（微软雅黑）

| 元素 | 半磅值 sz | 说明 |
|:---|:---|:---|
| 页首大标题 | 48 | 白字加粗，字距 30 |
| 分区标题 | 26 | 白字加粗，编号金色 |
| 卡片/大类标题 | 24 | 藏蓝加粗 |
| 授课信息、标签 | 22 | |
| 第1–2页正文 | 20–21 | |
| 教学实施正文 | 19 | 窄栏左对齐（两端对齐会拉开字距） |
| 图注 | 16 | 灰 `666666`，居中 |

行距：一律“固定值”，约为字号 × 1.3 × 倍数（`docx_kit.para` 已封装）。段落 `snapToGrid=0`（原文档常开“对齐文档网格”，会把行距撑乱）。

## 版式组件

- 页首横幅：独立小表（主表之上），左格图标、右格三行文字；下接一行金色细条；横幅与主表间 8 磅空段。
- 分区标题栏：跨全宽单格，藏蓝底，`图标 + 金色 01 + 白色标题`，左缩进 160。
- 大类标签：浅蓝灰底、左金边、图标在上、四字分两行竖排；纵向合并的续格同底色、上边线 nil。
- 编号条目：`金色序号\t正文`，悬挂缩进 300（“(1)”类 420）。
- 标签（chip）：run 级底色 + 同色 bdr（否则首尾空格底色错位）。
- 表格最左侧：每行首格左边 = 金色 sz24。

## 图标（Tabler Icons，MIT）

`assets/icons/<名称>-<颜色>.png`，已有：白色（分区栏）`info-circle chalkboard presentation message-2 ruler-measure file-analytics users-group target star alert-triangle bulb books`；藏蓝 `1F3864`（标签/表头）`file-analytics users-group target star alert-triangle bulb books clipboard-check route chart-bar sparkles tool list-details book-2 chalkboard users flag-3`；等。
新图标：`node tools/render_icons.js specs.json`，specs 形如 `[{"name":"target","color":"1F3864","out":"assets/icons/target-1F3864.png","stroke":1.8,"px":160}]`。先确认图标名存在：`ls node_modules/@tabler/icons/icons/outline/<name>.svg`。

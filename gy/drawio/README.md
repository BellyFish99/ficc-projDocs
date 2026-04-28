# 国元 CEO Deck V7 — drawio 图表使用说明

为 `CEO_Deck_v7.pptx` （45 张幻灯片）准备的 10 张 CEO 级战略图表。
每张 `.drawio` 源文件都可在 drawio 中二次编辑。

## 文件清单 — 按 V7 幻灯片顺序

| # | 文件 | V7 幻灯片 | 用途 | 图表模式 |
|---|------|---------|------|---------|
| 01 | `01_strategic_framework_house.drawio` | **S6** — 四大核心能力 | 战略落地框架：愿景 → 目标 → 四支柱 → 底座 | House/Temple (McKinsey signature) |
| 06 | `06_aladdin_four_layer.drawio` | **S12** — 贝莱德 Aladdin | 四层架构：应用 / 分析 / 数据 / 技术 + 关键规模面板 | Layered architecture |
| 07 | `07_marquee_evolution.drawio` | **S13** — 高盛 Marquee | 1993 SecDB → 2014 Marquee → Today 演进 + 六大模块 | Timeline + modules |
| 08 | `08_calypso_six_solutions.drawio` | **S14** — Calypso Adenza | 六大解决方案围绕中央平台 | Hub-and-spoke |
| 09 | `09_murex_three_pillars.drawio` | **S15** — Murex MX.3 | 三大支柱：投资组合交易 / 合规风控 / 会计交易后 | Three-pillar temple |
| 10 | `10_vendor_comparison_heatmap.drawio` | **S16** — 对标总结 | 7 厂商 × 6 维度 能力热图，华锐 POMS 高亮 | Heat map / vendor matrix |
| 02 | `02_three_pillar_value_matrix.drawio` | **S22** — 价值体系 | 三支柱价值驱动树 + 能力 × 支柱映射矩阵 | Driver tree + coverage matrix |
| 04 | `04_system_landscape_context.drawio` | **S24 后插入** / **S36 替换** | 系统上下文：POMS 中心 + 国元现有系统 + 市场基础设施 | C4 Context / System landscape |
| 03 | `03_trade_lifecycle_swimlane.drawio` | **S32** — 连得通 | 交易生命周期前后对比泳道图 | Swimlane before/after |
| 05 | `05_phased_roadmap_gantt.drawio` | **S44** — 项目建议 | 三期 18 月甘特图 + 里程碑 ▲ + 阶段价值 ▼ | Gantt with milestones |

---

## 使用方法

### 方法 1 —— drawio 桌面版（推荐）
1. 下载 drawio Desktop: https://github.com/jgraph/drawio-desktop/releases
2. 打开 `.drawio` 文件
3. 根据实际数据修改数字、名称、颜色
4. 导出：**File → Export as → PNG**（**300 DPI** + 勾选 **Transparent Background**）
5. 将导出的 PNG 拖入 PPT 对应幻灯片

### 方法 2 —— drawio 网页版
1. 访问 https://app.diagrams.net/
2. **Open Existing Diagram** → 上传 `.drawio`
3. 编辑后 **File → Export as → PNG / SVG**

### 方法 3 —— PowerPoint 直接嵌入
drawio Desktop 支持 Office 插件，可直接把 drawio 嵌入 PPT，适合需要频繁迭代的图表。

---

## 颜色约定（整套 deck 保持一致）

| 用途 | 颜色 | Hex |
|------|------|-----|
| 主题 / 标题 / 顶层框 | 深海军蓝 | `#1a3a5c` |
| 二级强调 / 中层 | 中海军蓝 | `#2e75b6` |
| 三级 / 辅助 | 天蓝 | `#5b9bd5` |
| 背景浅蓝 | 浅蓝 | `#dae8fc` / `#e1ecf7` |
| 热图 · 强 | 绿 | `#c6efce` (填充) / `#548235` (边框) |
| 热图 · 中 | 浅蓝 | `#ddebf7` / `#2e75b6` |
| 热图 · 弱 | 灰 | `#e7e6e6` / `#999` |
| 红色警告 / 痛点 / 限制 | 红 | `#c0392b` / `#fce4d6` |
| 绿色收益 / 强项 | 绿 | `#27ae60` / `#c6efce` |

修改任何图表时保持这个色板，整套 deck 的视觉一致性就能维持。

---

## V7 推荐插图位置总览

```
S1  封面 (保留)
S2  内容目录 (需手动更新至 6 个 Part)
S3  Part 1 分节页
S4  国元自营增长曲线 (保留)
S5  自营发展核心要求 (保留)
S6  四大核心能力 ← 🎨 [插入 01 House Framework]
━━━━━━━━━━━━━━━━━━━━━━ 新 Section 02 ━━━━━━━━━━━━━━━━━━━━━━
S7  Part 2 分节页
S8  行业现状 (文字布局已足够)
S9  华泰大象平台 (文字布局已足够)
S10 平安领航 + FITS
S11 山西证券
S12 贝莱德 Aladdin ← 🎨 [插入 06 Aladdin 四层架构]
S13 高盛 Marquee ← 🎨 [插入 07 Marquee 演进]
S14 Calypso ← 🎨 [插入 08 Calypso 六大解决方案]
S15 Murex ← 🎨 [插入 09 Murex 三支柱]
S16 对标总结 ← 🎨 [插入 10 对标热图]
S17 五大启示 (文字布局已足够)
━━━━━━━━━━━━━━━━━━━━━━ 原 Section 03 ━━━━━━━━━━━━━━━━━━━━━━
S18 Part 3 分节页 (原 02 项目介绍)
S19 四大核心价值 (保留)
S20 压力测试场景 (保留)
S21 投资经理能效 (保留)
S22 年化价值预估 ← 🎨 [插入 02 Three-Pillar Value Matrix]
━━━━━━━━━━━━━━━━━━━━━━ 原 Section 04 ━━━━━━━━━━━━━━━━━━━━━━
S23 Part 4 分节页 (原 03 华锐 POMS)
S24 POMS 六大引擎 ← 🎨 [04 System Landscape 可插入之后]
S25-S31 各模块详情 (保留)
S32 连得通 ← 🎨 [插入 03 Trade Lifecycle Swimlane]
S33-34 配得优+算得快 / IFRS9 (保留)
━━━━━━━━━━━━━━━━━━━━━━ 原 Section 05 ━━━━━━━━━━━━━━━━━━━━━━
S35 Part 5 分节页 (原 04 华锐优势)
S36 全景图 ← 🎨 [也可用 04 System Landscape]
S37-S42 各模块 + 华锐定位
━━━━━━━━━━━━━━━━━━━━━━ 原 Section 06 ━━━━━━━━━━━━━━━━━━━━━━
S43 Part 6 分节页 (原 05 项目建议)
S44 项目路线图 ← 🎨 [插入 05 Phased Roadmap Gantt]
S45 下一步 (保留)
```

---

## 修改指南 —— 常见改动

### 改文字
双击任意方框 → 输入新文字 → Enter 确认

### 改颜色
选中方框 → 右侧 Format Panel → **Fill / Line** → 选颜色

### 改位置 / 大小
拖拽即可；对齐用 `Arrange → Align` 菜单

### 改连线
单击连线 → 拖拽端点到新目标；样式在 **Style** 面板修改

### 热图单元格改级别
单击单元格 → 右侧 Format Panel → Fill 选色：
- `#c6efce` 绿 = 强项 ●●●
- `#ddebf7` 浅蓝 = 中等 ●●
- `#e7e6e6` 灰 = 弱项 ●
- `#fce4d6` 红 = 限制 / 不适用 ✗

同时修改文字内容以对应（● 数量 + 描述）

### 增加里程碑 ▲（甘特图）
复制现有 ▲ 字符（Text 形式），放到正确日期位置，下方加注释

---

## 常见问题

**Q: 中文字体在导出 PNG 时变成方块？**
A: drawio 默认字体不含完整中文字符。导出前：
**Extras → Configure Diagram → fontFamily** 设为 `Microsoft YaHei, SimHei, sans-serif`
或在导出对话框勾选 **Embed fonts**

**Q: 导出的图分辨率不够？**
A: 导出时选 300 DPI（PPT 呈现）或 600 DPI（A3 打印）

**Q: 如何一次性全部导出？**
A: drawio Desktop → **File → Export All Files** 或使用 drawio CLI 批处理

---

## 当前数据假设（需根据国元实际情况调整）

- AUM：约千亿（2025H1）
- 年化直接价值：1-1.8 亿（基于千亿 AUM × 行业基准 1/10 保守估算）
- 单次极端避损：1-2 亿（基于 924 事件案例）
- Phase 1：6 个月见效；Phase 2：6-12 月；Phase 3：12-18 月
- 主要竞争对手：Aladdin / Calypso / Murex / Bloomberg / Hundsun / SimCorp

具体数字以国元实际数据和合同敲定为准。

---

## 文件大小参考

```
01_strategic_framework_house.drawio       ~10 KB
02_three_pillar_value_matrix.drawio       ~15 KB
03_trade_lifecycle_swimlane.drawio        ~13 KB
04_system_landscape_context.drawio        ~15 KB
05_phased_roadmap_gantt.drawio            ~17 KB
06_aladdin_four_layer.drawio              ~9 KB  (新)
07_marquee_evolution.drawio               ~11 KB (新)
08_calypso_six_solutions.drawio           ~12 KB (新)
09_murex_three_pillars.drawio             ~12 KB (新)
10_vendor_comparison_heatmap.drawio       ~20 KB (新)
```

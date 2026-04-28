# 解决方案架构文档 — 最佳实践、资源与学习指南

---

## 一、三大方法论框架（推荐组合使用）

我们的国元方案文档应融合三套方法论，各取所长：

### 1. McKinsey 金字塔原理 + SCQA（用于"讲故事"）

> 来源：Barbara Minto《金字塔原理》，McKinsey标准方法论

**核心逻辑：先说结论，再说为什么**

```
         ┌─────────────┐
         │   主控结论    │  ← CEO只看这一句就知道你要说什么
         └──────┬──────┘
        ┌───────┼───────┐
   ┌────┴───┐┌──┴───┐┌──┴────┐
   │支撑论点1││论点2 ││论点3  │  ← 2-4个关键论点（MECE）
   └────┬───┘└──┬───┘└──┬────┘
   ┌────┴───┐┌──┴───┐┌──┴────┐
   │数据证据 ││证据  ││证据   │  ← 底层用数据和案例支撑
   └────────┘└──────┘└───────┘
```

**SCQA框架（每章的故事线）：**
- **S**ituation（情境）：国元自营业务现状
- **C**omplication（矛盾）：5大能力断层阻碍稳定收益目标
- **Q**uestion（问题）：如何系统性解决？
- **A**nswer（答案）：华锐POMS一体化平台

**MECE原则**：所有分解必须"不重叠、不遗漏"
- 我们的5大挑战分解（看不清/算不快/管不住/连不通/比不过）就是MECE的

**应用到我们的文档：**
- 每章开头用SCQA讲故事
- 每页PPT的标题就是该页的结论（Action Title）
- 不要用描述性标题如"行业分析"，要用结论性标题如"头部券商已建成一体化平台，窗口期正在关闭"

### 2. TOGAF 架构方法（用于"技术深度"）

> 来源：The Open Group Architecture Framework

**TOGAF推荐的架构文档结构：**

| 文档层次 | 对应我们的章节 | 内容 |
|---------|--------------|------|
| Architecture Vision | 第0章执行摘要 + 第1章战略理解 | 业务驱动力、目标、范围、利益相关方 |
| Business Architecture | 第3章需求分析 | 业务流程、能力地图、用例 |
| Application Architecture | 第4章方案设计（功能模块） | 应用组件、交互、接口 |
| Data Architecture | 第4章（IBOR部分） | 数据模型、数据流、主数据管理 |
| Technology Architecture | 第4章（技术架构部分） | 技术选型、部署架构、非功能性 |
| Migration Planning | 第6章实施路径 | 分期规划、迁移策略、风险 |

**Architecture Vision文档标准结构（可直接参考）：**
1. Executive Summary
2. Business Drivers and Goals
3. Scope (In/Out of Scope)
4. Stakeholders and Concerns
5. Architecture Vision (To-Be Overview)
6. Business Capabilities Map
7. Constraints and Assumptions
8. Value Proposition (Benefit Analysis)
9. Initial Risk Assessment
10. Approval and Sign-off

### 3. MITRE 实用方案架构指南（用于"落地性"）

> 来源：MITRE Corporation - Guide for Creating Useful Solution Architectures

**MITRE的核心理念：Just Enough Architecture**
- 不是写得越厚越好，而是"刚好够用"
- 支持敏捷和瀑布两种开发模式
- 重点是"有用"而不是"完整"

**MITRE推荐的方案架构内容：**
- 业务上下文和驱动力
- 当前状态 vs 目标状态
- 解决方案概览（概念架构）
- 关键架构决策记录（ADR）
- 集成和接口设计
- 安全和合规考虑
- 运营和支持模型

---

## 二、我们应该学习的最佳文档模板

### GitHub上的实战模板

| 模板 | 特点 | 推荐用途 | 链接 |
|------|------|---------|------|
| **shekhargulati/software-architecture-document-template** | 13个核心章节，含约束/假设/非功能性需求/技术选型，简洁实用 | 技术架构章节的参考结构 | github.com/shekhargulati/software-architecture-document-template |
| **bwgartner/SA-template** | AsciiDoc格式，支持Enterprise/Reference Architecture变体，可输出HTML/PDF/EPUB | 文档工程化的参考 | github.com/bwgartner/SA-template |
| **bflorat/architecture-document-template** | 法国电信出品，含应用视图/开发视图/规模视图/基础设施视图 | 多视图架构文档的参考 | github.com/bflorat/architecture-document-template |
| **joelparkerhenderson/architecture-decision-record** | ADR（架构决策记录）模板和大量示例 | 记录关键技术决策 | github.com/joelparkerhenderson/architecture-decision-record |
| **unlight/solution-architecture** | 解决方案架构学习资源大全（文章/书/视频/课程） | 系统学习解决方案架构 | github.com/unlight/solution-architecture |

### 权威框架和指南

| 资源 | 特点 | 链接 |
|------|------|------|
| **MITRE Guide for Creating Useful Solution Architectures** | 最实用的方案架构"how-to"指南，政府级标准 | mitre.org/news-insights/publication/guide-creating-useful-solution-architectures |
| **Gartner Solution Architecture Document Template** | 行业标准模板（需Gartner订阅） | gartner.com/en/documents/4324899 |
| **TOGAF Architecture Vision Template** | 企业架构标准，10个核心章节 | pubs.opengroup.org/togaf-standard/ |
| **Microsoft Azure Well-Architected Framework** | 云架构最佳实践，含Solution Architect职责定义 | learn.microsoft.com/en-us/azure/well-architected/ |

### 持续学习资源

| 资源 | 类型 | 链接 |
|------|------|------|
| **awesome-software-architecture (mehdihadeli)** | GitHub资源大全，含专门网站 | github.com/mehdihadeli/awesome-software-architecture |
| **Databricks Financial Services Investment Mgmt Reference Architecture** | 金融行业参考架构 | databricks.com (Financial Services section) |
| **Barbara Minto《金字塔原理》** | 书籍 | 中文版可在各大书店购买 |
| **arc42 Architecture Documentation** | 实用架构文档模板 | arc42.org |

---

## 三、综合建议：我们的文档应该怎么写

### 黄金法则

1. **先讲结论，再讲过程**（金字塔原理）
   - 每章开头就是该章结论
   - CEO翻到任何一页都能3秒内知道这页在说什么

2. **每个方案必须回答一个问题**（McKinsey铁律）
   - No problem, no solution
   - 我们的"5大挑战→5大方案"映射就是这个逻辑

3. **多层次受众设计**（TOGAF + Codewave最佳实践）
   - CEO看：执行摘要 + 问题方案映射 + 价值 + 为什么是我们
   - CTO/CIO看：技术架构 + 技术选型 + 非功能性需求
   - 业务负责人看：功能模块详解 + 业务流程
   - 项目经理看：实施路径 + 风险 + 里程碑

4. **活文档，不是死文档**（MITRE + Codewave）
   - 用版本号管理（v0.1 骨架 → v0.5 初稿 → v1.0 正式）
   - 每次评审后更新

5. **架构决策要记录理由**（ADR最佳实践）
   - 不只记"选了什么"，更要记"为什么选、放弃了什么"
   - 例如：为什么选事件驱动而不是定时轮询？

### 文档格式建议

**对CEO的PPT版本**：20-30页
- 每页一个Action Title（结论性标题）
- 大量图表，少量文字
- 重点章节：执行摘要、问题树、方案映射、竞争差异、实施路径

**对技术团队的Word版本**：50-80页
- 完整技术架构
- 接口设计
- 非功能性需求详细指标
- 架构决策记录（ADR）

---

## 四、对我们solution_skeleton_v2.md的改进建议

基于以上最佳实践研究，建议对现有骨架做以下增强：

| 改进点 | 方法论来源 | 具体做法 |
|--------|-----------|---------|
| 每章加SCQA故事线 | McKinsey金字塔 | 在每章开头加S-C-Q-A四行引言 |
| 加入"明确的范围边界" | TOGAF Architecture Vision | 加In-Scope / Out-of-Scope章节 |
| 加入架构决策记录 | ADR最佳实践 | 附录加5-10个关键ADR |
| 加入利益相关方分析 | TOGAF + MITRE | 明确CEO/CTO/业务/PM各关注什么 |
| 加入当前状态架构图 | Codewave HLD指南 | 第3章加As-Is架构图 vs To-Be架构图 |
| Action Title替代描述性标题 | McKinsey presentation | 所有标题改为结论句 |
| 加入约束和假设 | shekhargulati模板 | 第3章加技术/预算/时间约束 |
| 加入back-of-envelope计算 | shekhargulati模板 | 第4章加容量/性能粗算 |

# 🎯 MCK Strategic 10 Tests — AI Strategy Coach

```
    ╔══════════════════════════════════╗
    ║                                  ║
    ║   M C K   S T R A T E G I C     ║
    ║       1 0   T E S T S           ║
    ║                                  ║
    ║  "Have You Tested Your Strategy?"║
    ║                                  ║
    ╚══════════════════════════════════╝
```

**不是又一个无聊的战略问卷。是一个会拍桌子挑战你的 AI 战略教练。**

基于 MCK 经典方法论 *"Have You Tested Your Strategy Lately?"* (2011)，这个 Skill 把十大战略检验维度变成了一场有温度、有锐度的**教练式对话**——你说出你的想法，AI 顾问帮你拷问它。

---

## 🤔 这东西解决什么问题？

你有一个战略想法——可能是一次创业、一个投资决策、一个职业选择、甚至一段关系的经营方向。

你需要的不是一个问卷让你自己打分（那不是诊断，那是自欺欺人）。

你需要的是**一个懂行的人，听你说完之后，问你那些你不想面对的问题**。

这就是这个 Skill 做的事。

---

## ✨ 它跟普通的"战略评估工具"有什么不同？

| 普通工具 | 这个 Skill |
|---------|-----------|
| 机械地问10个固定问题 | 基于你的具体内容**灵活追问** |
| 让你自己选 ABCD | **AI 自行判断**每个维度的得分 |
| 问题跟你的情况无关 | 每个问题都**紧扣你的具体想法** |
| 像填表 | 像被一个**资深战略顾问 coaching** |
| 冷冰冰的报告 | 有性格、有情绪、有 ASCII 头像的角色 |

---

## 🎭 10 个场景，10 个角色

你的想法是什么，就会匹配什么样的顾问。不是换了个名字，是**完全不同的性格、语言风格和提问角度**。

```
┌──────────────────────────────────────────────────────┐
│                                                      │
│  🏢 企业战略     Prof. Sterling    冷静、精确、拷问式  │
│  🚀 创业战略     Alex Storm       犀利、快节奏、中英混 │
│  📈 投资战略     Dr. Warren Li    数据导向、风控思维    │
│  💼 职业发展     Coach Maya       温暖但直接、心理学    │
│  🧭 人生规划     Sage 道元        禅意、隐喻、直击本质  │
│  💑 婚姻与关系   Dr. Helen        温柔、敏锐、非评判    │
│  🎓 教育战略     Prof. Edison     实证主义、前瞻性      │
│  ⚡ 数字化转型   CTO Nova         技术语言说战略        │
│  🏥 健康管理     Dr. Vita         系统思维看健康        │
│  🌍 社会影响力   Ambassador Kai   理想+现实主义        │
│                                                      │
└──────────────────────────────────────────────────────┘
```

每个角色都有自己的 ASCII 头像和情绪系统——他们会在肯定你时微笑，在发现弱点时皱眉，在你回避问题时拍桌子。

```
    ┌─────────┐          ┌─────────┐          ┌─────────┐
    │ ◉    ◉ │          │ ●    ● │          │ ♥    ♥ │
    │    ▽    │          │    ◇    │          │    ◡    │
    │  ╰───╯  │          │  ╰═══╯  │          │  ╰═══╯  │
    │ ┌─────┐ │          │  /HOOD\ │          │ ╭─────╮ │
    └─┤SUIT ├─┘          └─┤ DEV  ├─┘          └─┤HEART├─┘
      │& TIE │            │SHIRT │              │HEALER│
      └─────┘              └─────┘              └─────┘
    Prof. Sterling       Alex Storm           Dr. Helen
```

---

## 📊 MCK 战略十问维度

每个维度 1-4 分，AI 基于对话内容自行评分。

| # | 维度 | 核心问题 |
|---|------|---------|
| 1 | 市场竞胜力 | 你的战略能让你持续超越竞争者吗？ |
| 2 | 优势来源 | 你的竞争优势是真实的还是假设的？ |
| 3 | 精准聚焦 | "在哪里竞争"的定义够精准吗？ |
| 4 | 趋势前瞻 | 你是在追赶趋势还是引领趋势？ |
| 5 | 独到洞见 | 你有别人没有的信息或认知吗？ |
| 6 | 不确定性管理 | 你系统性地管理了不确定性吗？ |
| 7 | 承诺-灵活平衡 | 大胆押注和保持灵活之间平衡了吗？ |
| 8 | 去偏见 | 决策过程被认知偏见污染了吗？ |
| 9 | 执行决心 | 你有真正的决心去执行吗？ |
| 10 | 行动计划 | 战略已转化为具体行动了吗？ |

---

## 🚀 怎么用？

### 方式一：在 WorkBuddy 中使用（推荐）

1. 将 `SKILL.md` 和 `references/` 目录放入：
   ```
   ~/.workbuddy/skills/mck-10-tests/
   ```

2. 在对话中说出你的想法：
   ```
   "帮我用战略十问评估一下——我想做一个面向中小企业的 AI 财务助手"
   ```

3. AI 会自动加载 Skill，匹配角色，开始诊断

### 方式二：直接使用 SKILL.md 作为 System Prompt

把 `SKILL.md` 的内容作为任何 LLM 的 system prompt，然后开始对话。

---

## 📁 文件结构

```
mck-10-tests-skill/
├── SKILL.md                    # 核心 Skill 定义（角色、流程、行为规范）
├── references/
│   └── methodology.md          # 十问详细方法论（评分标准、理论依据）
├── scripts/
│   └── render_report_card.py   # 报告生成器 v5（PNG长图 + Word + PDF，楷体·雷达图·紧凑排版）
├── CHANGELOG.md                # 版本变更记录
├── LICENSE                     # Apache License 2.0
└── README.md                   # 你正在读的这个
```

---

## 💬 一次真实的诊断长什么样？

> **你**：我们公司想从 B 端原料供应商转型做 C 端消费品牌。
>
> **Prof. Sterling**：*调整袖扣* 经典的 B2B → B2C 延伸命题。让我先确认——你们的核心产品和技术壁垒是什么？这决定了优势迁移路径。
>
> **你**：我们有完整的日本工艺标准，核心客户是肯德基...
>
> **Prof. Sterling**：有意思。也就是说你们的技术已经在 C 端被验证了——只不过品牌溢价被客户拿走了，你们只赚原料的薄利。这恰恰是 B 端最痛苦的地方：价值创造者和价值获取者是分离的...
>
> *（10-15轮自然对话后）*
>
> **输出**：完整的十维评分报告，包含得分、证据、建议和行动清单

---

## 🧠 设计哲学

这个 Skill 的设计遵循几个核心信念：

1. **好的战略诊断不是填表，是对话。** 最有价值的洞见来自被追问到不舒服的时刻。

2. **AI 应该有观点。** 不是让用户自己打分然后汇总——而是 AI 基于对话内容做出判断，用户可以质疑和补充。

3. **角色不是装饰，是功能。** 不同场景的提问角度完全不同——投资顾问关心 downside，创业教练关心 PMF，婚姻治疗师关心互动模式。

4. **情绪价值是战略工具。** ASCII 头像和角色台词不是花哨——它们让用户在高强度的思考中保持投入感。

---

## 📖 方法论来源

- Bradley, C., Hirt, M., & Smit, S. (2011). *"Have you tested your strategy lately?"* MCK Quarterly.
- Bradley, C., Hirt, M., & Smit, S. (2018). *Strategy Beyond the Hockey Stick*. Wiley.
- Rumelt, R. (2011). *Good Strategy Bad Strategy*. Crown Business.

---

## 📝 Changelog

See [CHANGELOG.md](./CHANGELOG.md) for full version history.

**Latest: v2.1.0** (2026-04-09)
- Triple output: PNG long-image + Word (.docx) + PDF — all from the same data
- Word document preserves identical layout: KaiTi + Arial, navy headings, radar chart, score bars
- `generate_all(data, base_path)` one-call convenience function

---

## 📜 License

Apache License 2.0. See [LICENSE](./LICENSE) for details.

```
Copyright 2026 Kaku Li

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0
```

---

*Built by [Kaku Li](https://github.com/likaku) — because every strategy deserves to be stress-tested before it burns cash.*

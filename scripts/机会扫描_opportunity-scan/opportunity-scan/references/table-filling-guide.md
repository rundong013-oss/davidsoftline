# 公司跟踪表填写规范

每家覆盖公司有一张标准化的基本面和业绩跟踪表（Excel）。表格的行结构（科目）由 David 预先制作好，Claude 按季度填入数据和定性内容。

**核心原则：不写 row number——不同公司模板行号不同，一律按 section 名称和科目名称定位。**

### 定量数据源优先级

**开始填写前，先检查 `library/{TICKER}/` 下是否存在 model 文件**（如 `model_{TICKER}.xlsx`、`model-*.xlsx` 等）。

- **如果 model 存在 → 默认使用 Model**：所有定量数据（财务、segment、渠道、产品拆分、门店、增速）优先从 model 一次读取，无需调 API 或解析 filing。仅当 model 中某项数据缺失时，再 fallback 到 Filing / FMP。
- **如果 model 不存在 → 使用 Filing + FMP**：从 10-Q/10-K 原文和 FMP API 获取。
- **如不确定**：询问 David 确认。

无论哪种模式，定性内容（Section 4 指引+定性拆解、Section 5 景气周期）始终从 Transcript 获取，不变。

---

## 第 0 步：定位目标列

填写任何数据之前，先确定本季度数据应填入哪一列：

1. 读取 Excel，扫描 row 2–4 的日期字段，找到与目标季度匹配的列。
2. 如果目标季度的列已存在（row 2 日期匹配），直接在该列填写/更新。
3. 如果不存在，在最后一个数据列之后新增一列，填入以下 header：
   - **Row 1**：`季度`
   - **Row 2**：日历季度末日期（如 `2026-06-30`）
   - **Row 3**：财季起始日（如 `2026-05-04`），来自 Filing
   - **Row 4**：财季截止日（如 `2026-08-02`），来自 Filing
   - **Row 7**：财报货币（如 `USD`）
4. 注意跳过模板列（含填写说明文字的列）和分隔列（row 1 为 `S` 的列）。

---

## Section 0: 基础信息（表头区域）

| 科目 | 说明 | Source |
|------|------|--------|
| 财季起止日期 | 该季度的真实起止日（如 2026-05-04 至 2026-08-02），填入 row 3–4 | Filing（10-Q / 10-K 封面） |
| 绝对涨幅 | 该财季期间公司股价涨幅 | FMP historical prices；fallback yfinance |
| 超额涨幅 | 绝对涨幅 − 行业指数涨幅。行业指数来自 `output/monitor/apparel_updating file.xlsx` → "行业指数" sheet：运动服饰公司用"运动服饰等权指数"，奢侈品公司用"奢侈品等权指数"。指数为周度，日期取最接近财季起止日的数据点计算区间涨跌。如指数不覆盖该财季（如历史数据起点晚于财季），暂空 | 计算 |

---

## Section 1: Set-up

包含：收入 YoY / 净利润 YoY / OPM / 净利率 / Forward PE / Forward PS。

- **Source**：`output/monitor/Valuation Screening` 中对应公司的最新数据。
- **更新流程**：先把 Valuation Screening 中股价和一致预期更新到最新，再引用结果填入。
- **格式**：纯数字。
- **如 Valuation Screening 未及时更新**：收入 YoY / 净利润 YoY / OPM / 净利率可从 Section 2 的财务数据计算填入；Forward PE / Forward PS 暂空，标注"待更新"。

---

## Section 2: 核心财务指标

包含：单季收入 / 毛利率 / SGA 费用率 / OPM / 净利率 / 净利润 + 库存。

- **Source**：FMP `income-statement`（quarterly）+ `balance-sheet-statement`（quarterly）。
- **Fallback**：对应季度 Transcript 中的数字。
- **格式**：纯数字。
- **注意**：只更新季度数据，年度数据已有不动。
- **库存 QoQ**：依赖上一季度同一 row 的库存金额。如上季库存金额缺失，先补上一季数据再计算 QoQ。

---

## Section 3: 业绩速评

| 科目 | 说明 | Source |
|------|------|--------|
| 收入 beat/miss | actual vs estimated 差额或百分比 | FMP `earnings-calendar` 端点 |
| 利润 beat/miss | actual vs estimated 差额或百分比 | FMP `earnings-calendar` 端点 |
| 其他速评 | 暂空 | — |

- **格式**：数字（差额或百分比）。
- **如无 FMP API 访问**：从 Transcript 中提取定性描述（如"came in below expectations"、"beat our internal plan"），写入速评行；精确 beat/miss 数字待后续补充。

---

## Section 4: 指引及定性拆解

**这是最核心、最耗时的部分。**

- **Source**：**Transcript（唯一来源）**，不可编造。
- **总体原则**：不只写数字，关键是定性描述和 rationale。定性格子翻译管理层原话即可，保留 rationale，不要只写结论。

### 4.1 指引（Guidance）

| 科目 | 内容 |
|------|------|
| 收入 | 管理层给出的下季 / 下财年 / 长期收入指引 + 背后原因 |
| 毛利率与成本 | 毛利率指引 + 驱动因素（关税、mix、促销等） |
| SGA 费用率 | 费用率指引或目标 |
| OPM / 净利率 | 利润率指引 |

**格式规则**：
- 如果同时有下季度、下财年、长期目标，分段写在同一格子内
- 标明时间范围：`Q3 指引：`、`全年：`、`长期目标：`
- 直接翻译管理层原话即可，保留定量 + 定性

### 4.2 增长 Driver 定性拆解

#### 4.2.0 Summary & 战略

当前公司 under 什么战略，战略的重点措施精简总结。

#### 4.2.0 行业

Transcript 中关于宏观、行业、细分赛道的评论和展望，对公司业绩的影响。

#### 4.2.0 竞争格局

关于竞争、竞争对手的评论。

#### 4.2.1 产品（非常重要，最认真写）

**按公司模板中的产品线拆分逐一填写。**

每条产品线的内容包括：
1. **产品表现**：哪些产品线好 / 不好，具体数据
2. **变化**：过去这个季度发生了什么变化（新品上市、停产、战略调整）
3. **因果关系**：产品是否是业绩的主要原因（如管理层说收入放缓是因产品问题，要强调）

注意：
- 产品战略写在产品总行，各子品类写具体表现
- 这是整张表最重要的 section，需要逐一覆盖模板中列出的所有产品线

#### 4.2.1 人事

重大管理层任免。

#### 4.2.2 品牌定位/热度

和营销事件区分。内容包括：
- 品牌定位变化、重新定位描述
- 品牌热度指标（搜索量、awareness、sentiment 等）
- 目标客群变化

#### 4.2.2 品牌营销（事件）

- 具体营销活动（活动名称、签约代言人、赞助合作、联名）
- 取得的效果（参与人数、曝光量、销售转化等）

#### 4.2.3 渠道

**直营（DTC）**：

| 维度 | 内容 |
|------|------|
| 门店数 | 各地区变化、开关店计划、重点拓展方向、开店 guidance |
| 门店绩效 | 面积、坪效、客流、同店增长、可比增长、升级/翻新计划 |
| 线上渠道 | 线上相关变化/计划/成就、渗透率、会员数变化 |

**经销（Wholesale）**：

- 经销渠道表现、动作（收缩 / 扩展）
- 表现好或差的原因
- 新渠道拓展（进入哪些新零售商）

#### 4.2.3 区域

- 按地区顺序组织管理层对各地区的评论
- 突出：业绩解释、展望、目标
- 关注门店扩张计划、渗透率、市场份额

#### 4.2.4 其他信息

特殊事项（关税、一次性费用等）。

---

## Section 5: 景气周期

| 科目 | 说明 | Source |
|------|------|--------|
| 涨价/折扣 | 关于折扣率、促销、全价销售的评论；有定量数据不可遗漏 | Transcript |
| 渠道库存 | 重要的库存情况评论（总量已在 Section 2，此处放定性） | Transcript |

---

## Section 6: 分地区拆分（定量）

| 科目 | Source | Fallback |
|------|--------|----------|
| Total Revenue / YoY / FX-Neutral | FMP | — |
| 分地区收入 | Filing（10-Q） | Transcript |
| 分地区份额 / 增速 | 计算填入 | — |
| 分地区 Comparable Growth | Transcript | Filing |
| 门店 Summary（分地区门店数 + 增速） | Filing（10-Q） | Transcript |

- **格式**：纯数字。

---

## Section 7: 分渠道拆分（定量）

| 科目 | Source | Fallback |
|------|--------|----------|
| 渠道收入（Store / Direct / Other） | Filing（10-Q） | Transcript |
| 渠道份额 / 增速 | 计算 | — |
| Comparable Sales / LFL Growth | Transcript | Filing |
| 坪效（Sales per Avg Sq Ft） | Filing | — |

- **格式**：纯数字。

---

## Section 7.1: 分产品拆分（定量）

| 科目 | Source | Fallback |
|------|--------|----------|
| 产品线收入 | Filing（10-Q） | Transcript |
| 品类份额 / 增速 | 计算 | — |

- **格式**：纯数字。

---

## 数据源优先级总结

| Section | 首选 Source | Fallback |
|---------|-----------|----------|
| 基础信息（涨幅） | FMP historical prices | yfinance |
| Set-up（估值） | Valuation Screening | FMP quote + estimates |
| 财务指标 | FMP income-statement / balance-sheet（quarterly） | Transcript |
| Beat/Miss | FMP earnings-calendar | Transcript 定性描述 |
| 指引 + 定性拆解 | **Transcript（唯一）** | 不可编造 |
| 分地区 / 渠道 / 产品收入（定量） | Filing（10-Q） | Transcript |
| Comparable Growth / FX-Neutral 增速 | Transcript | Filing |
| 门店数据 | Filing（10-Q） | Transcript |

> **说明**：分地区/渠道/产品的精确收入数字，Filing（10-Q）精确到千元，优于 Transcript（通常只给百分比）。但 Comparable Growth 和 FX-Neutral 增速在 Transcript 中由管理层直接给出，比从 Filing 反推更可靠。

---

## 关键原则

1. **定量格子只填数字**，不可自己编造，必须有 source。
2. **定性格子翻译管理层原话即可**，保留 rationale，不要只写结论。
3. **不写 row number**——不同公司模板行号不同，按 section 名称和科目名称定位。
4. **指引要标时间范围**——`Q3 指引：`、`全年：`、`长期目标：` 分段。
5. **产品 section 最重要**——按模板中的产品线子分类逐一填写。
6. **只更新季度数据**——年度数据已有不动。
7. **先定位列，再填数据**——按"第 0 步"确定目标列，不要覆盖模板列或分隔列。
8. **前置依赖要处理**——库存 QoQ 需要上季数据，Forward PE/PS 需要 Valuation Screening；缺失时标注"待更新"而非留空不说明。
9. **单元格格式**——目标列的所有单元格默认设置 `wrap_text=True`（单元格内自动换行）+ `vertical='center'`（上下居中）+ `horizontal='center'`（左右居中），列宽设为 80。**所有格子一律居中**（包括定性长文本）。数值类型格式规则：
   - **增速、利润率、份额** → `number_format='0.0%'`，值用小数（如 0.05 = 5.0%）
   - **收入、利润等绝对金额** → 纯数字，不加百分比格式
   - **Forward PE / PS** → 纯数字，保留两位小数
   - 这是默认行为，每次填写都必须执行，不可遗漏。
10. **每次填写前重新读模板行号**——不可缓存旧行号。每次填写前扫描 column B–E 的标签，动态确定每个科目的行号再填入。模板可能被 David 调整过（增删行），行号会漂移。
11. **Segment 数据用 Filing，不用 FMP geographic API**——FMP 的 `revenue-geographic-segmentation` 返回的是 country-level 数据，与公司报表中的 operating segment 口径不同（如 LULU 的"China Mainland"segment 不含 HK/TW/Macau，但 FMP 的"CHINA"含义不明且数值不匹配）。分地区/渠道/产品收入一律以 Filing（10-Q/10-K）为准。填入后做 sanity check：各 segment 加总必须等于 Total Revenue。
12. **长文本分段换行**——定性格子中，不同主题之间用换行分隔。`【跑步】`/`【足球】` 等品类标记前、`Q2指引：`/`全年：`/`长期目标：` 等时间标记前、`【北美】`/`【EMEA】` 等地区标记前，都加 `\n` 换行。不要把所有内容挤在一行。
13. **每个格子只放该格子定义的内容，严禁混放**——产品格子只写产品（哪些品类好/差、新品、战略调整），不放收入数字、门店数、地区增速；行业格子只写管理层对行业/宏观的评价，不放公司自己的增速；门店格子只写门店数量变化和开店计划，不放产品或财务数据；区域格子只写地区评论。填写前先想清楚"这条信息属于哪个格子"，不确定就对照规范中每个格子的定义。绝不能偷懒把 transcript 大段 copy 进格子——必须先理解、分类、精简，再填入。
14. **判断值是否存在时用 `is not None`，不用 `if val`**——0、0.0、0.0% 都是合法数据（如 FX-Neutral 增速 0%、FX impact 0%）。`if val:` 会把 0 当成 False 跳过，导致漏填。所有填入逻辑必须用 `if val is not None:` 判断。

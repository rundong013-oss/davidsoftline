# 数据说明与脚本用法

skill 自带从《运动服饰机会复盘 Masterfile.xlsx》整理出的数据，位于 `data/`。数字以这些文件为准；`case-library.md` 里的叙述性数字来自原始文档，二者不一致时以数据文件为准。

## 目录

1. 数据文件与字段
2. 覆盖范围与已知问题
3. 脚本：匹配、扫描、更新数据
4. Masterfile 里还有什么（需要时直接读原文件）

---

## 1. 数据文件与字段

| 文件 | 内容 | 主要字段 |
|---|---|---|
| `monthly_panel.csv` | 32 家公司月度面板，2015-03 至 2025-02（多数公司指标到 2024-12） | `company, month, mkt_cap（亿美元）, fwd_rev（亿美元）, fwd_npm, fwd_ps, fwd_pe` |
| `quarterly_panel.csv` | 季度热图面板，3Q15–4Q24；含月度面板没有的公司（COLM、UAA、SHOO 等） | `company, quarter, mkt_cap, fwd_ps, fwd_npm` |
| `cases.json` | 32 个案例：类型标签、分组说明、前后指标、收益分解、人事/产品/品牌/渠道/行业Beta/评论 | 见下 |
| `windows.csv` | 时间轴上高亮的 40 个机会窗口（含 2020 年疫情纯反弹窗口） | `company, start, end, mkt_cap_start, mkt_cap_end, return` |
| `beta_periods.csv` | 时间轴上标注的行业 Beta 起点 | `month, description` |
| `declines.csv` | 13 段大幅下跌 | `company, window, return` |

`cases.json` 每条记录：

```
company, window（如 "22Q3–24Q4"）, size_bucket（>300% / 100%-300%）,
return_multiple（窗口总涨幅，11.7 表示 +1170%）, cagr,
types（自身成长 / 困境反转 / 行业Beta / 假反转 / 未分类，可多个）,
groups（Masterfile 中该案例所在分组的说明文字，即它属于哪个子范式）,
before / after：{mkt_cap, fwd_rev, fwd_npm, fwd_pe, fwd_ps}（亿美元；Forward 口径）,
decomposition：{mkt_cap_x, rev_x, npm_x, ps_x, pe_x}（after ÷ before）,
人事, 产品, 品牌, 渠道, 行业Beta, 评论（定性文字）,
panel_window：{start, end, return}（对应的月度高亮窗口）, note（手工修正说明）
```

公司代码统一为：Nike、Adidas、PUM、LULU、DECK、ON、CROX、SKX、ANF、GAP、VSCO、WWW、KTB、LEVI、RL、AS（Amer）、GII（吉登）、HBI（哈尼斯）、UAA、COLM、SHOO、ZGN、BZUN、GOLDWIN、安踏、李宁、特步、361度、波司登、比音勒芬、迅销、亚瑟士。

## 2. 覆盖范围与已知问题

- **Forward 指标是当时的一致预期**，不是事后实际值。这正是判断 set-up 时市场看到的东西。
- **Adidas 为欧元口径**，其他为美元（中国公司已折算）。跨公司比较规模时注意。
- 部分公司月度面板只有市值（亚瑟士、COLM、UAA、SHOO、GII），估值与利润率要看季度面板或留空。
- 部分公司历史较短：ON 从 2021-10、VSCO 从 2021-08、AS 从 2024-02、KTB 和 LEVI 从 2019 年起。PS 分位需要至少 24 个月数据，不足时为空。
- 已知异常：比音勒芬 Forward 收入在 2024 年有口径跳变（12 个月变化 +600% 以上），波司登月度数据稀疏。遇到极端值先核对，不要直接用。
- Masterfile 里同一案例在不同 sheet 的数字偶有不一致（例如李宁 16Q1–21Q4 的"市值 After"一处写 67、一处写 285；按月度面板为 8.2 → 285.9 亿美元，约 34 倍）。窗口涨幅优先看 `panel_window.return`。
- 手工修正都写在 `scripts/extract_masterfile.py` 顶部的 `WINDOW_ALIASES` 和 `CASE_OVERRIDES` 里，并在案例的 `note` 字段注明。

## 3. 脚本

所有命令在 skill 根目录下运行，依赖 pandas、numpy（提取还需要 openpyxl）。

### 3.1 匹配历史起点：`scripts/match_setup.py`

在历史案例的**起点月**计算同样的指纹，按加权标准化距离排序，返回最像的几个起点，并附上它们之后 12/24 个月的市值、净利率、PS 路径，以及兑现前后的指标。

指纹（均为 Forward 口径）：净利率、PS、PE、PS 在自身过去最多 10 年的分位、收入距过去峰值的回撤、收入预期 12 个月变化、净利率 12 个月变化、收入规模。

```bash
# 公司在面板里（用于回测或看历史某个时点）
python scripts/match_setup.py --company PUM --month 2024-09 --exclude-self

# 公司不在面板里，或要用最新数据：手动输入指纹，缺的项不填
python scripts/match_setup.py --npm 0.018 --ps 0.35 --ps-pct 0.05 --rev-vs-peak -0.10 \
    --rev-g12 -0.04 --npm-chg12 -0.012 --rev 45 --pe 19

# 只在某类机会里找
python scripts/match_setup.py ... --type 困境反转

# 面板扫描：列出某月所有公司的指纹，并标出"净利率≤3%""PS≤10%分位""收入距峰值≤-8%"等
python scripts/match_setup.py --screen [--month 2024-12]
```

距离 d > 0.8 的结果会标"不够相似"（回测中历史案例起点的首个类比距离中位数约 0.5、90 分位约 0.7；全部公司-月份的 90 分位约 0.8）；最近的都超过阈值时，脚本会提示没有足够相似的历史起点。

默认排除 2020 年 3 月后的疫情纯反弹窗口（宏观驱动、指纹无区分度），需要时加 `--include-rebound`。

**怎么读结果**：距离 d 越小越像；"参与比较的指纹"少于 5 项时，匹配可信度低。回测（每个案例起点排除自身后找最像的历史起点）：困境反转类的首个类比有 5/9 仍是困境反转，成长类之间（LULU ↔ DECK、ON → LULU、安踏 → LULU）也能互相找到；行业 Beta 类主要由宏观驱动，指纹匹配区分度弱，要用 `beta_periods.csv` 和定性判断。结论：**定量匹配用来缩小范围和提供赔率参照，类型判断以定性为准。**

**从 FMP 等数据源取指纹**：
- `npm` = 下一财年一致预期净利润 ÷ 一致预期收入
- `ps` = 当前市值 ÷ 下一财年一致预期收入；`pe` = 当前市值 ÷ 下一财年一致预期净利润
- `rev` = 下一财年一致预期收入（亿美元）
- `ps_pct`、`rev_vs_peak`、`rev_g12`、`npm_chg12` 需要历史的一致预期序列；拿不到就不填，脚本会只用可比的指纹

### 3.2 写入 Obsidian：`scripts/render_obsidian.py`

```bash
python scripts/render_obsidian.py --vault <vault 路径> [--dir 机会扫描] NKE.json ANF.json
python scripts/render_obsidian.py --vault <vault 路径> --summary-only
```

JSON 结构和笔记格式见 `output-template.md`。

### 3.3 追加最新数据

把新月份按 `monthly_panel.csv` 的列格式追加到一个副本，然后：

```bash
python scripts/match_setup.py --panel my_panel.csv --company NKE --month 2026-09
```

### 3.4 Masterfile 更新后重新生成数据

```bash
python scripts/extract_masterfile.py <Masterfile.xlsx 路径>
```

脚本只读 Masterfile，按标签定位各 sheet 的数据；机会窗口从"时间轴（月度）(output)"里的高亮填充识别。

## 4. Masterfile 里还有什么

需要更细的内容时，直接用 openpyxl/pandas 读原文件：

| Sheet | 用途 |
|---|---|
| `时间轴（月度） (output)` | 月度面板原表；第 2–3 行标注行业 Beta 时期；公司行的高亮 = 机会窗口；下方还有行业涨幅中位数、各指标 MoM 变化与平均值 |
| `按机会大小整理` | 按涨幅分档（>300%、100–300%、下跌）的案例及前后指标 |
| `自身成长…`、`困境反转…`、`行业beta` | 每个案例的人事/产品/品牌/渠道/行业 Beta/评论，及前后指标与变化倍数 |
| `机会类型的统计` | 24 个翻倍以上机会（不含仅 2020M3 纯反弹）的类型计数 |
| `公司市值-PS-净利率高峰低谷热图` | 季度市值、Forward PS、Forward 净利率热图 |
| `LULU高点研究`、`LULU mini case-高点研究` | LULU 高点研究：季度实际业绩（收入、分地区增速、同店、毛利率、OPM、净利率、SG&A、库存）+ 按时间标注的市场观点、行业、产品（主要矛盾）、渠道 |
| `12-14x PE Forward研究` | 与月度时间轴同结构，用于研究 PE 区间 |

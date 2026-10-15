# 机会扫描（opportunity-scan）打包

## 目录

| 路径 | 内容 |
|---|---|
| `opportunity-scan.skill` | 可直接安装的 skill 包（与下方文件夹内容相同） |
| `opportunity-scan/` | skill 源文件，可直接放进 Claude Code 的 skills 目录或手动修改 |
| `opportunity-scan/SKILL.md` | 主流程：基本面整理 → 门槛判断 → 类型与类比 → 机会大小 → 写入 Obsidian |
| `opportunity-scan/references/` | 输出模板、机会类型、案例库、参照数字、数据说明 |
| `opportunity-scan/data/` | 从 Masterfile 整理出的月度/季度面板、案例库、机会窗口、Beta 时期、下跌窗口 |
| `opportunity-scan/scripts/` | `match_setup.py`（历史起点匹配）、`render_obsidian.py`（写笔记与 summary）、`extract_masterfile.py`（Masterfile 更新后重建 data） |
| `tools/fmp_snapshot.py` | 从 FMP 拉取财务、一致预期与市值的脚本（需设置环境变量 FMP_API_KEY） |
| `sources/` | 原始材料：运动服饰经验 PDF、机会复盘 Masterfile |

## 安装与使用

1. 把 `opportunity-scan/` 文件夹放到 Claude Code 的 skills 目录（如 `~/.claude/skills/`），或在 Claude 设置中上传 `opportunity-scan.skill`。旧的 `apparel-setup-analogs` 如已安装，请删除。
2. 在工作目录的 `CLAUDE.md` 里写明 Obsidian vault 路径、filings 和 transcript 的存放位置。
3. 对 Claude Code 说"机会扫描 NKE ANF DECK"之类即可触发。
4. Masterfile 更新后：`python opportunity-scan/scripts/extract_masterfile.py sources/运动服饰机会复盘Masterfile.xlsx`

脚本依赖：pandas、numpy、openpyxl、requests。

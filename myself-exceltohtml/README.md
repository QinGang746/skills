---
name: lottery-analysis
description: |
  彩票数据处理双能力技能 —— (功能一) 文本(JSON/CSV/TSV)转 Excel，支持列拆分/合并表头/农历日期/波色查找；(功能二) 彩票分析 Excel 转交互式 HTML 仪表盘。
  触发场景：
  (1) 用户说"把这个文本转成 Excel"、"txt 转 xlsx"、"json 转 excel"、"把开奖数据导出为表格"
  (2) 用户说"把这个彩票分析 Excel 做成 HTML"、"帮我可视化这些开奖数据"
  (3) 用户提供包含单双/大小/生肖/频率分析的 Excel 文件
  (4) 用户输入"帮助"/"【帮助】"时，输出功能菜单
  关键词：txt转excel、json转excel、文本转表格、生成xlsx、农历、波色、彩票分析、开奖记录、Excel转HTML、数据可视化、生肖分析
allowed-tools:
  - Read
  - Write
  - Bash
  - Edit
---

# lottery-analysis — 彩票数据处理（TXT→Excel / Excel→HTML）

> **核心心智模型**：本技能含**两个独立功能**，围绕彩票开奖数据的两个不同处理阶段：
>
> ```
> 功能一  文本(JSON/CSV/TSV)  ──txt2excel.py──▶  Excel 开奖记录表(.xlsx)
> 功能二  彩票分析 Excel(2 sheet) ──excel2html_formula.py──▶  交互式 HTML 仪表盘(.html)
> ```
>
> **⚠️ 两功能不自动串联**：功能一产出的是**原始开奖记录表**（id/date/平码/特码…），功能二需要的是**含"分析结果+历史记录"两 sheet 的分析 Excel**（`公式统计.xlsx` 格式）。二者数据结构不同，中间的分析加工由外部完成，本技能不负责衔接。

## 功能总览

| 功能 | 脚本 | 输入 | 输出 |
|------|------|------|------|
| 一：TXT→Excel | `02-script/txt2excel.py` | `03-data/txt-source/*.txt` (JSON/CSV/TSV) | `04-output/*.xlsx` |
| 一（增量）：合并→Excel→HTML 一键串联 | `02-script/update_excel.py` | `03-data/txt-source/*.txt`（自动增量：仅 2026 年有新增时才重新生成） | `04-output/开奖记录-全年份合并.xlsx` + `*report-*.html` |
| 二：Excel→HTML 仪表盘（公式统计） | `02-script/excel2html_formula.py` | `03-data/*.xlsx`（公式统计结构，2 sheet 固定偏移） | `04-output/*.html` |
| 二：Excel→HTML 仪表盘（开奖记录 v系列） | `02-script/excel2html_lottery.py` | 开奖记录明细 xlsx（含「开奖明细」sheet，支持两层结构：v5 标准版 44 列 / txt2excel 产出 29 列） | `04-output/*-report-时间戳.html` |
| 二（简版）：Excel→HTML 平表 | `02-script/excel2html_flat.py` | 任意 `.xlsx` | `.html` |
| **三：myself-pyhtml 独立工具** | `../myself-pyhtml/run.bat` | 接口拉取 + txt-source | `myself-pyhtml/output/*.xlsx` + `*-report-*.html` |

**目录约定**：模板在 `01-template/`（波色查找表 `01-template/波色号码对照表.xlsx`），脚本在 `02-script/`，数据源在 `03-data/`（txt 放 `03-data/txt-source/`），所有产物输出到 `04-output/`。

---

## 【帮助】响应约定

**MUST**：当用户输入包含"帮助"（如"帮助"、"【帮助】"、"help"）时，**只输出下方功能菜单，不执行任何转换、不读取任何文件**。无例外。

菜单内容（原样输出给用户）：

```
📊 lottery-analysis 技能 —— 两大功能

【功能一】文本 → Excel（txt2excel.py）
  把 JSON/CSV/TSV 文本转成带格式的 Excel，支持列拆分、合并表头、农历日期、波色查找。
  用法：
    python 02-script/txt2excel.py <输入txt> [输出xlsx] [选项]
  彩票开奖完整示例：
    python 02-script/txt2excel.py "03-data/txt-source/临时测试-2026.txt" "04-output/临时测试-2026.xlsx" \
      --format json \
      --split "num,shengxiao,wuxing" \
      --split-prefixes "平码,平肖,平五行" \
      --split-special "特码,特肖,特五行" \
      --lunar-date \
      --bose-table "01-template/波色号码对照表.xlsx"
  常用选项：
    --format json|csv|tsv     输入格式（默认按扩展名推断）
    --encoding utf-8|gbk       输入编码（默认 utf-8）
    --split <列名>             拆分逗号分隔列，如 num,shengxiao,wuxing
    --split-prefixes <前缀>    平码子列前缀，如 平码,平肖,平五行
    --split-special <特列名>   特列标签，如 特码,特肖,特五行
    --split-normal-count N     平码数量（默认 6）
    --lunar-date               date 列旁新增农历日期列
    --bose-table <路径>        特码列旁新增波色列（红/蓝/绿波，带底色）
    --timestamp                文件名追加 -YYYYMMDD-HHmm
    --sheet-name <名>          Sheet 名称（默认 Sheet1）
  现成数据（03-data/txt-source/）：
    临时测试-2020.txt ~ 临时测试-2026.txt（7 个年度开奖 JSON）

【功能二】彩票分析 Excel → HTML 仪表盘（excel2html_formula.py）
  把含"分析结果 + 历史记录"两 sheet 的分析 Excel 生成交互式 HTML（5 个 Tab 仪表盘）。
  用法（自动取 03-data 下首个 .xlsx，输出到 04-output）：
    python 02-script/excel2html_formula.py
  简版（任意 Excel 转平表 HTML）：
    python 02-script/excel2html_flat.py <输入.xlsx> [-o 输出.html] [-t 标题]

【目录】
  01-template/  波色号码对照表.xlsx   02-script/  脚本   03-data/  数据源（txt-source/ 放文本）   04-output/  产物
```

---

## 角色声明

> **角色**：彩票数据处理工程师（数据转换 + 前端可视化）
> **默认立场**：自动推断数据格式，不确定时询问；优先保证数据准确性，保留原始值而非猜测
> **角色外**：不做数据分析/预测、不给投注建议、不修改源数据文件——只做格式转换与可视化

## 禁令（MUST NOT）

**MUST NOT** 在未 Read 源文件内容前直接生成产物。无例外。
**MUST NOT** 修改或丢弃源数据中的任何字段，也不修改用户原始 Excel 文件。无例外。
**MUST NOT** 在 HTML 中引入外部 CDN 依赖（如 Chart.js），HTML 需离线独立运行。无例外。
**MUST NOT** 对分析数据做出博彩建议或"下一期预测"。无例外。

============================================================

# 功能一：文本 → Excel

> **核心能力**：列拆分 + 合并表头 + 农历日期转换 + 波色查找。始终使用 `02-script/txt2excel.py`，不手写 openpyxl。

## 工作流程

> - 场景 A：单文件转换 → Step 1 → Step 2
> - 场景 B：多文件合并转换（如 txt-source 下多个 txt）→ **推荐用 `update_excel.py` 一键完成**（自动增量判断 + 合并 + Excel + HTML）
> - 场景 C：格式不确定 → 先 Read 文件，询问用户

### Step 1: 读取源文件

先 Read 源文件确认 JSON 结构、编码和字段名。

多文件合并时（按 date 倒序合并为临时文件）：
```python
import json, os
all_rows = []
folder = "03-data/txt-source"
for f in sorted(os.listdir(folder)):
    if not f.endswith(".txt"):
        continue
    with open(os.path.join(folder, f), encoding="utf-8") as fh:
        all_rows.extend(json.load(fh))
all_rows.sort(key=lambda r: r.get("date", ""), reverse=True)
with open("04-output/merged.json", "w", encoding="utf-8") as fh:
    json.dump(all_rows, fh, ensure_ascii=False)
```

### Step 2: 执行转换

```bash
python 02-script/txt2excel.py <输入文件路径> [输出xlsx路径] [选项]
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--encoding` | 输入编码 (utf-8/gbk) | utf-8 |
| `--format` | 输入格式 (json/csv/tsv) | 自动推断 |
| `--sheet-name` | Sheet 名称 | Sheet1 |
| `--timestamp` | 文件名追加 `-YYYYMMDD-HHmm` | 关闭 |
| `--lunar-date` | date 列旁新增农历日期列 `农历YYYY-MM-DD` | 关闭 |
| `--bose-table <路径>` | 特码列旁新增波色列（红/蓝/绿波），按农历年查表 | 关闭 |
| `--split <列名>` | 拆分逗号分隔列 | — |
| `--split-prefixes` | 拆分子列前缀 | — |
| `--split-special` | 特列名 | — |
| `--split-normal-count` | 平码数量 | 6 |

**完整示例**（彩票开奖数据，输出到 04-output）：

```bash
python 02-script/txt2excel.py "03-data/txt-source/临时测试-2026.txt" "04-output/临时测试-2026.xlsx" \
  --format json \
  --split "num,shengxiao,wuxing" \
  --split-prefixes "平码,平肖,平五行" \
  --split-special "特码,特肖,特五行" \
  --timestamp \
  --lunar-date \
  --bose-table "01-template/波色号码对照表.xlsx"
```

### Step 3: 验证

检查生成的 `.xlsx` 文件是否存在且大小 > 0。

**用了 `--lunar-date` 时必须抽查农历**：农历位运算 + 闰月易错，且"春节初一对 ≠ 全年对"。抽查方法——用连续多天 + 已知锚点交叉验证，不能只测一个日期：
- 春节正月初一（如 2026-02-17→农历01-01、2024-02-10→农历01-01）
- 中秋八月十五（如 2025-10-06→农历08-15、2024-09-17→农历08-15）
- 闰月边界（如 2025 闰六月：2025-07-25→闰六月初一）
- 打印一段连续日期确认农历**逐日单调递增**、月末→次月初一无跳变
- 已知修复：`solar_to_lunar` 月大小位用 `(16-m)`（非 `m+3`），闰月按 `leap_month` 显式插入分支（详见反模式表）

### 输出结构

```
| id | year | qishu | date | 农历date | 平码(合并6列) | 特码 | 波色 | 平肖(合并6列) | 特肖 | 平五行(合并6列) | 特五行 | nyear | nqi |
```
- 第 1 行：合并表头（平码/平肖/平五行各占 6 列合并单元格）
- 波色列：根据农历年从 `01-template/波色号码对照表.xlsx` 查特码对应波色（红波/蓝波/绿波），含单元格背景色

### 一键增量更新（推荐）— `update_excel.py`

适用于日常更新场景：`03-data/txt-source/` 下仅 2026 年数据每日更新，历史年份（2020-2025）已固定。

```
python 02-script/update_excel.py
```

**执行逻辑**：
1. 检查 `04-output/开奖记录-全年份合并.xlsx` 是否存在
2. **首次**（不存在）：全量合并 7 个年份 txt → 生成 Excel → 生成 HTML
3. **增量**（存在）：读取 Excel 中最新日期 → 从 2026.txt 筛选晚于该日期的记录 → **有新增**才重新合并生成；**无新增**则跳过，只检查 HTML 是否存在
4. 自动串联 `txt2excel.py` + `excel2html_lottery.py`，输出 xlsx 和 html 到 `04-output/`

### 输出格式
```
转换完成：[输入文件] → [输出文件]
行数：[N]  列数：[M]
```

### 功能一反模式

| 错误做法 | 问题 | 正确做法 |
|---------|------|---------|
| 未 Read 文件就开始转换 | 格式推断错误 | 先 Read 确认结构再调脚本 |
| 手写 openpyxl 代码 | 不一致、效率低 | 始终用 `02-script/txt2excel.py` |
| 转换前手动编辑数据 | 破坏原始数据 | 只转换不修改 |
| 农历换算自行实现 | 位偏移和基准日易错 | 用脚本内置已验证算法 |
| 月大小位索引写成 `(m+3)` | 12 个月大小首尾颠倒，年内日期少一天（春节/年份边界仍对，隐蔽） | 第 m 月大小位在 bit `(16-m)`；`lunar_year_days` 与 `solar_to_lunar` 均用此 |
| 闰月靠 `for lm in range(1,13)` 内 `lm-=1` 回退 | range 迭代器下一轮覆盖 lm，自减失效，闰月年（如 2025 闰六月）算错 | 每个常规月后按 `leap_month` 显式插入闰月分支 |
| 改完农历只测单个日期（如春节） | 春节初一永远偏移 0，与月长/闰月无关，测不出年内 bug | 连续多天 + 春节/中秋/闰月边界多锚点交叉验证单调性 |
| 把阳历平/闰年与农历闰月混为一谈 | 误改对的阳历逻辑 | 阳历平闰年 `solar_day_of_year` 已判 4/100/400 且正确；出错的是农历闰月，两码事 |

### 扩展：特码衍生列 + 统计分析 sheet（按需）

用户常在开奖记录表基础上追加"特码属性衍生列"，并单独出一张统计 sheet。以下为已确认口径（六合彩 1-49，特码为准）：

**明细表衍生列口径**：
| 列 | 规则 |
|----|------|
| 头数 | 特码十位 `n//10`，**只有 0-4**（49 的十位最大是 4，无 5-9 头） |
| 尾数 | 特码个位 `n%10`，0-9 |
| 单双 | 奇=单 偶=双 |
| 大小 | ≥25 大，≤24 小 |
| 大小单双组合 | 大小+单双 → 大单/大双/小单/小双 |
| 合数 | 十位+个位；合数单双取其奇偶 |
| 质合 | 质数集 `{1,2,3,5,7,11,13,17,19,23,29,31,37,41,43,47}`（含1计质）为质，余为合 |
| 半波 | 波色色系+大小 → 红大/红小/蓝大…；色单双 → 波色色系+单双 |
| 尾大尾小 | 尾数 ≥5 尾大，≤4 尾小；头单双 → 头数奇偶 |
| 家禽野兽 | 家禽 牛马羊鸡狗猪 / 野兽 鼠虎兔龙蛇猴 |
| 天肖地肖 | 天肖 兔马猴猪牛龙 / 地肖 蛇鼠鸡狗羊虎 |
| 前肖后肖 | 前肖 鼠牛虎兔龙蛇 / 后肖 马羊猴鸡狗猪 |

**统计分析 sheet 口径**（六块：频率 / 遗漏冷热 / 连开重号 / 热温冷+遗漏榜 / 按年+星期 / 生肖×波色）：
- 表格式参考 `03-data/公式统计.xlsx` 的「【工具】统计分析」——每维度一张"类型/次数/…"表。
- **⚠️ 只按实际数据统计频率，不写"理论占比/偏差"**：号码↔生肖/波色/五行的映射**每年更换**，任何固定理论值（如生肖 4/49）都不成立。频率表只出「类型 / 次数 / 实际占比」三列。
- **实际占比用百分比**（如 `49.80%`，保留两位小数），不用 0.498 这种小数。
- **每个分析块 title 下加一行计算说明**（灰色斜体小字，小白可懂，如"计算：特码 ≥25 记为'大'…"）；一级标题（一/二/三）下也加总说明。
- **号头统计只列 0-4**（不要补 5-9 空行）。
- 遗漏值：当前遗漏=距今多少期未出（0=最新期开出），另给出现次数 + 历史最大遗漏 + **【当前最新期数】【上一次最近期数】**（各展开为 year/qishu/date 三子列，用合并表头分组）。当前最新期数=全表最新一期（固定，所有行相同）；上一次最近期数=该项最近一次开出那期（即 `yqd[gap]`，从未出现则留空）；两者间隔恰=当前遗漏。三个遗漏块（号码/生肖/属性）都加这两列。
- **号码补零两位**（`1`→`01`），号码块标题写"（01-49）"。
- **遗漏期数列加单位**：year→`2026年`、qishu→`186期`，**date 不加**（保持 `2026-07-05`）。
- 连开/重号：特码与上期重号、本期特码在上期平码中、平码与上期重叠个数分布。
- 数据须按 date 倒序（最新在前）后再算遗漏。
- **四、热温冷分层 + 遗漏排行榜**：热温冷以「当前遗漏」为指标、以全部号码**平均遗漏为界**（<均值=热带红底、=均值=温带黄底、>均值=冷带蓝底）；遗漏排行榜按当前遗漏**降序**列号码/生肖/属性三个维度（越靠前越久未出）。
- **五、按年统计 + 星期效应**：按年=各特码属性（大小/单双/波色/家禽野兽/质合）逐年占比对比；星期效应=date 转周几后统计各属性占比。占比均=当年(该周几)该取值次数÷当年(该周几)总期数。
- **六、生肖 × 波色 交叉表**：12 生肖 × 红蓝绿波 次数矩阵，行末合计=该生肖总次数，末行给三波色列合计。
- **⚠️ 全部只用特码维度**（特码/特肖/波色/五行等），用户明确不需要平码平肖的分析。

**「分年统计」sheet 口径**（独立第 3 个 sheet，与「统计分析」并列）：
- 一个 sheet 内**按年从上到下分区**（2020→2026，每年一段，深蓝分区条 `════ YYYY 年（共 N 期，起~止）════`）。
- 每年段含四类：①频率统计 ②遗漏值/冷热 ③连开/重号 ④热温冷分层+遗漏排行榜——口径与「统计分析」对应块相同，但**全部只用该年数据**。
- **⚠️ 分年遗漏在该年内独立计算**：当前遗漏=从该年最后一期往回数、历史最大遗漏=该年内最长连续未出、上一次最近期数=该年内最近开出那期（本年未出现则留空）。不跨年——跨年遗漏由「统计分析」sheet 体现。
- 各年占比分母=该年期数（非总期数）；热温冷均值也按该年重算。

============================================================

# 功能二：彩票分析 Excel → HTML 仪表盘

> **⚠️ 两个生成器，按输入结构选**：
> - `excel2html_formula.py` —— 仅适配 **`公式统计.xlsx`** 那种结构（2 sheet、固定列偏移的分析结果表）。下文 Step1-3 讲的是它。
> - `excel2html_lottery.py` —— 适配 **开奖记录明细 xlsx**（`开奖记录-全年份合并-v5.xlsx` 这类，含「开奖明细」44 列 sheet）。见本区末尾「excel2html_lottery.py」小节。
> 两者数据结构完全不同、不可混用；跑错会报 `IndexError`/读到空。

> **核心心智模型**：彩票分析 Excel 是扁平、无交互的表格数据，用户需要快速洞察偏差趋势。
> 本功能将 Excel 数据提取为结构化数据，注入预设计的 HTML 模板，生成带标签页导航的交互式仪表盘。

```
Excel (.xlsx)  →  Python 数据提取 (pandas + openpyxl)  →  HTML 模板注入  →  {excel名}-时间戳.html
    2 sheets         结构化 JSON + 单元格样式                5 Tab 仪表盘          浏览器打开
```

## 工作流程

> **一句话执行路径**：提取数据 + 样式 → 生成 JSON → 注入模板 → 输出 HTML

### Step 1: 读取 Excel 并提取数据

对 Excel 两个 sheet 执行数据提取，**同时用 openpyxl 提取单元格样式**（字体色、背景色、粗体）：

```python
import pandas as pd
import openpyxl

# Sheet 1: 分析结果 (index 0)
df1 = pd.read_excel(excel_path, sheet_name=0, header=None)
# Sheet 2: 历史开奖记录 (index 1)
df2 = pd.read_excel(excel_path, sheet_name=1, header=None)
# 用 openpyxl 读取样式（字体颜色、背景色、粗体）
wb = openpyxl.load_workbook(excel_path, data_only=False)
ws1, ws2 = wb.worksheets[0], wb.worksheets[1]
```

**样式提取规则**：
- RGB 类型颜色直接转换 `#RRGGBB`
- Indexed 颜色查表映射
- Theme 类型忽略（默认色）
- 跳过白色 `#FFFFFF` 背景、`#262626` 近黑字体（均为默认值）

**Sheet 1 数据结构**（当前脚本适配 `公式统计.xlsx`，39 行 × 53 列，无标准表头；区块 title 在 r2、表头在 r3、数据从 r4）：
- cols 3-7：【01】开奖属性（类型/次数/实际占比/理论占比/偏差），数据 r4-12：野兽/家禽/大/小/单/双/红波/蓝波/绿波（合并表，overview 的波/大小/单双/兽源统计均来自此块）
- cols 9-14：【02】单双统计（类型/次数/实际占比/理论占比/偏差/备注），数据 r4-13：大单/小单/大双/小双/红单红双/蓝单蓝双/绿单绿双
- cols 16-21：【03】号头统计（号头/次数/实际占比/理论占比/偏差/范围），数据 r4-8
- cols 23-28：【04】号尾统计（号尾/次数/实际占比/理论占比/偏差/范围），数据 r4-13
- cols 30-49：【05】数字出现频率，4 个块（每块 6 列含分隔列），block 起点 `[30, 36, 42, 48]`，数据 r4-16
- rows 20-31 cols 3-7：【06】生肖统计（生肖/最近次数/最新期数/偏差值/是否注意）
- rows 20-31 cols 9-14：【07】生肖频率（生肖/次数/实际占比/理论占比/偏差/范围）
- 总期数：号头/号尾汇总 `r16 c17`（当前 277），读不到时兜底 "277"

> **⚠️ 结构变体警告**：此 Excel 系列存在"整体偏移"的历史变体。旧版 `分析结果.xlsx`（30 行 × 50 列）各区块相对当前版本**左移 3 列、上移 2 行**：单双 c0-4(r2起) / 大小双 c6-11 / 号头 c13-18 / 号尾 c20-25 / 频率 blocks `[27,33,39,45]` / 生肖 r18起 c0、c6 / 总期数 r14c14。**换新 Excel 时务必先 dump 网格核对区块起点，不能假设偏移不变**（见反模式表"直接沿用旧列偏移"）。

**Sheet 2 数据结构**（`公式统计.xlsx` 51 行 × 159 列；旧版 36 行 × ~151 列）：
- 10 个月份块，每块 15 列，从 col 2 开始（2, 17, 32, 47, 62, 77, 92, 107, 122, 137）——**块起点两版本相同**
- 每块内：周期(0)/日序(1)/日期(2)/期数(3)/生肖(4)/颜色(5)/大小(6)/单双(7)/色波(8)/野兽(9)/号头(10)/号尾(11)/颜色单双(12)/组合(13)
- 数据行：`公式统计.xlsx` 从 row 5 开始（row0 标题 / row1 "注意"说明行 / row2 表头 / 部分月首行空）；旧版从 row 4 开始。提取时用 `range(4,36)` 扫描 + 守卫（跳过表头/"期"串行、空行由 notna 跳过）兼容两版
- 渲染时"星期"列改名为"日序"并加 tooltip "当月第几日"

**换入新 Excel 的排查步骤**（当 `03-data/` 换了文件或页面出现大面积 0/`-`）：
1. 先 dump 网格到 UTF-8 文件核对结构（终端直接 print 中文会乱码）：
   ```python
   import openpyxl, io
   wb = openpyxl.load_workbook(path, data_only=True)
   out = io.open('_dump.txt','w',encoding='utf-8')
   for si,ws in enumerate(wb.worksheets):
       out.write('== SHEET %d %s (%dx%d) ==\n'%(si,ws.title,ws.max_row,ws.max_column))
       for r in range(1, min(ws.max_row,40)+1):
           cells=['c%d=%r'%(c-1,ws.cell(r,c).value) for c in range(1,ws.max_column+1) if ws.cell(r,c).value is not None]
           if cells: out.write('r%d: %s\n'%(r-1,' | '.join(cells)))
   out.close()
   ```
2. 用 `_dump.txt` 定位每个区块的 title 行、表头行、数据起始行、列起点，对照上面的结构表改提取常量（`iloc[行,列]`、`range(...)`、`block_starts`）。
3. 生成后**用 Python 反解 HTML 里 `var DATA = {...};` 校验**关键字段（total_count / hot / freq 数量 / history 记录数），不要只靠肉眼看截图。
4. 完成后删除 `_dump.txt` 临时文件。

### Step 2: 生成 HTML

生成 HTML 时遵循以下样式规范：

1. **浅色主题**：背景 `#f0f2f5`，卡片 `#fff`，圆角 `12px`，box-shadow 柔和
2. **标签页导航**：首页概览 / 综合分析 / 历史记录 / 号码走势 / 其它走势 五个 Tab，`position: sticky; top: 0` 固定顶部，`active` 状态蓝色 `#1a73e8`
3. **首页概览**：
   - 顶部统计卡片（总期数 / 最热号码 / 最冷号码 / 最高偏差生肖），每张卡悬停显示 tooltip 解释
   - **热门/冷门数字**：左右并排两个卡片式盒子（`hotcold-box`，红/蓝左边框），每个号码一张独立小卡片（`hc-card`）
   - **红蓝绿波分布**：2×2 网格（`wave-grid`），每格内含彩色小卡片（`wc-card`，带浅色背景）：
     - 红蓝绿波（红/蓝/绿三张小卡片）
     - 大小（橙/紫两张小卡片）
     - 单双（蓝/金两张小卡片）
     - 家畜/野兽（绿/红两张小卡片，绿底家畜、红底野兽）
   - **生肖偏差排行 + 生肖出现频率**（双列并排 `analysis-grid`）：
     - 生肖名字按野兽/家畜分类着色：野兽（鼠虎兔龙蛇猴）浅红底 `#fde8e8`，家畜（牛马羊鸡狗猪）浅绿底 `#e8f5e9`
     - 偏差排行标题右侧有"排序"按钮（`sort-toggle`），点击切换升序/降序，段描述含阈值提示 `忽略<25≤提醒<30≤可以追`
     - 偏差排行标签按阈值自动显示：≥30"可以追"绿底、25~30"提醒"黄底、<25"忽略"灰底
4. **综合分析页**：双列网格布局（`grid-template-columns: 1fr 1fr`，移动端单列），每个板块包含三段信息：
   - 标题（部分需 strip 掉 Excel 中嵌入的条件表达式）
   - **段描述**：灰色小字说明该表含义
   - **参考值**：黄色 insight-box（`background: #fef9e7; border-left: 3px solid #f39c12`），AI 编写的总结建议
    - 板块：单双分析 / 大小双统计 / 开头统计 / 尾数统计 / 数字频率 / 生肖统计 / 生肖频率
5. **号码走势页**：Canvas 网格表格，X=号码1-49 / Y=期数（上=新→下=旧）：
    - 格子大小 24×24px，总宽约 1230px 自适应无横向滚动
    - Y 轴标签格式 `275期`→`1期` 全部展示（9px 灰色），X 轴 1-49 全部展示
    - 命中格子按波色浅底（红/蓝/绿波）或默认浅蓝 `rgba(26,115,232,0.15)` + 号码，未命中留白
    - **整页滚动**（不套内层 `overflow:auto` 容器；`overflow-x:auto` 会连带纵向局部滚动、削弱 sticky）
    - **X 轴表头用独立 sticky canvas**（`t1h`，`position:sticky;top:var(--navh)`）绘制，与主 canvas 共用同一 `padLeft=55` 与 `cellW`，做到列像素级对齐；禁止用 HTML flex 盒另画表头（坐标系不一致必然错位）
6. **其它走势页**：4 张属性色块图并列展示（flex 布局），**与号码走势统一为"整页滚动 + sticky canvas 表头"模式**：
    - 大小走势：X=大(浅橙 `#F8C471`) / 小(浅紫 `#C39BD3`)，Y=期数
    - 单双走势：X=单(浅蓝 `#85C1E9`) / 双(浅金 `#F7DC6F`)，Y=期数
    - 波色走势：X=红波(浅红 `#F1A2AB`) / 蓝波(浅蓝 `#81BBF8`) / 绿波(浅绿 `#C1E77E`)，Y=期数
    - 家禽野兽走势：X=家畜(浅绿 `#82C99A`) / 野兽(浅红 `#F1948A`)，Y=期数
    - **属性色块统一用浅色系**（避免深色刺眼、且便于叠深色文字），Canvas 宽度 180px/张，24px/行
    - **色块内嵌文案**：每个命中单元格在填色后于中央绘制对应文字（大/小、单/双、红波/蓝波/绿波、家畜/野兽）；文字色用 `textColorFor()` 按填充色亮度（`0.299R+0.587G+0.114B`，阈值 150）自动取深 `#333` / 白 `#fff`
    - **X 轴分类标签置于顶部并 sticky 固定**：每张图拆为 header canvas（`aNh`，`position:sticky;top:var(--navh)`，画彩色标题+分类子标签）+ body canvas（`aN`，画色块），header 用 `drawAttrHeader` 绘制、与 body 共用 `pad.left=30` 及列宽保证对齐；滚动时标签始终可见
    - **单元格黑色边框**：色块先填充，再用 `strokeStyle='#000';lineWidth=1` 叠画完整网格（所有行线+列线+外框），Y 轴标签 `275期`→`1期`（9px）
    - 外层 flex 容器必须 `align-items:flex-start`（禁用 `stretch`），否则高 canvas 卡片会被拉伸/溢出
    - 每次切换 Tab 强制重绘 Canvas（不缓存）
    - **走势图 sticky 表头的 `top` 必须等于 `.tab-nav` 实际高度**：用 CSS 变量 `--navh`（当前 51px，略大于实测 50.4px 留缝）统一控制，否则表头被导航栏遮挡
7. **百分比格式化**：占比列 ×100 显示 `XX%`；偏差列 ±符号，正值绿色 `#27ae60`、负值红色 `#e74c3c`
   - 注意：偏差列统一在 col 4（`c === 4`），非最后一列
8. **红蓝绿波高亮**：表"类型"列中 `红波`→浅红底、`蓝波`→浅蓝底、`绿波`→浅绿底（仅匹配含"波"后缀的完整词）
9. **生肖"是否注意"列**：三色区分 —— `忽略`灰底 `#e8e8e8`、`提醒`黄底 `#fff3cd`、`可以追`绿底 `#c8e6c9`，正常表格边框不丢失
10. **表格列排序**：综合分析页所有数值列（次数/占比/偏差/最新期数）表头可点击排序：
    - 表头显示 `⇅`（灰色 14px），点击升序 `▲`（蓝色）、再点降序 `▼`（蓝色）、再点恢复原始
    - `buildTable` 通过 `sortCols = [1,2,3,4]` 标记可排序列，cell 加 `data-sv` 属性存储排序值
    - 生肖统计和生肖频率的独立表格也支持相应列排序
11. **表格**：`border-collapse`，表头 `#f8f9fa`，hover 行高亮，保留 Excel 原始字体色/背景色（inline style）
12. **历史记录**：月份切换按钮（圆角 pill）+ 横向滚动表格，保留 Excel 原始样式
13. **段描述 + 参考值**：综合分析页每个板块均需包含三段信息：
    - 标题下方灰色小字段落描述：说明该表数据含义（如"统计大小、单双、波色等基础属性各出现了多少次"）
    - 黄色 `insight-box`（`background: #fef9e7; border-left: 3px solid #f39c12; font-size: 12px;`）：以"参考值："开头，AI 编写的趋势总结
    - 适用于全部 7 个板块（01~07），历史记录和首页概览不需要 insight-box
14. **列标题 tooltip**：以下表头带 `title` 属性悬停解释 + 虚线下划线（`border-bottom: 1px dashed`）：
    - 占比 → "实际占比=该类型出现次数÷总期数，理论占比=按数学概率计算的期望值"
    - 偏差 → "偏差=实际占比-理论占比，正值偏多、负值偏少"
    - 色波 → "号码对应的颜色波形：红波/蓝波/绿波"
    - 野兽 → "号码对应的兽类属性：野兽/家禽"
    - 颜色单双 → "色波与单双的组合，如蓝单=蓝波+单数"
    - 组合 → "大小与单双的组合，如大双=大+双数"
    - 是否注意 → "忽略=暂时观望；提醒=需留意趋势；可以追=建议关注"
15. **【06】生肖统计标题**：移除 Excel 源数据中嵌入的 `（忽略<25<=提醒<30<=可以追）` 条件表达式，改为段描述自然语言说明

### Step 3: 验证

生成完成后执行：
```bash
$file = Get-ChildItem -LiteralPath "04-output" -Filter "*.html" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$file.Length
```
确认文件 > 10KB，包含 `<script>` 标签内的 JS 数据数组。

### 输出格式

```
已生成 {excel文件名}-YYYYMMDD-HHMMSS.html

标签页导航（5 个 Tab，sticky 顶部）：
- 首页概览：统计卡片（带 tooltip）+ 热门冷门数字（独立卡片）+ 红蓝绿波分布（2×2 网格含小卡片）+ 生肖偏差排行/频率（双列含排序）
- 综合分析：单双/大小双/开头/尾数/数字频率/生肖（双列网格 + 段描述 + 黄色参考值 + 列排序 + 列标题 tooltip）
- 历史记录：[N] 个月份，按月切换，表头有 tooltip，保留 Excel 原始字体/背景色
- 号码走势：Canvas 网格表格，X=1-49（顶部 sticky canvas 表头）/ Y=期数（上=新→下=旧），24×24px 格子，命中按波色浅底，80vh 滚动
- 其它走势：并列 4 张 Canvas 色块图（大小/单双/波色/家禽野兽），X=属性（顶部 sticky 分类标签）/ Y=期数，24px/行，黑色单元格边框
```

### 功能二反模式

| 错误做法 | 问题 | 正确做法 |
|---------|------|---------|
| 凭记忆硬编码数据而不先读 Excel | 数据与实际文件不符 | 先 `python -c` 读 Excel，提取真实数据 |
| 在 HTML 中引入 CDN 链接（Chart.js 等） | 离线环境无法打开 | 纯 HTML/CSS/原生 JS，所有可视化手写 |
| 直接用 `pd.read_excel` 读全表当表格渲染 | 多 sheet 非标准表头无法正确解析 | 按列偏移量逐区段提取，跳过表头行 |
| 频率 block_starts 用 [27,32,37,42,47] | 总列数 50，47+4=51 越界 | 用 `[27, 33, 39, 45]`（块间有 1 列分隔）|
| 偏差列用 `headers.length - 1` | 6 列表最后一列是备注/范围不是偏差 | 固定 `c === 4`（所有分析表偏差都在 col 4）|
| 给 `<td>` 加 `class="alert-tag"` | `inline-block` 覆盖 `table-cell` 导致边框消失 | 用 inline style 直接设背景/文字色 |
| 将条件表达式 `（忽略<25<=提醒<30<=可以追）` 直接渲染为标题 | 用户无法理解代码式条件 | JS 端 strip 掉，用段描述自然语言说明阈值 |
| 保留旧版段描述文案不更新 | 描述与 AI 参考值脱节 | 每个板块必须同时提供段描述（说明数据含义）和参考值（AI 总结建议） |
| Canvas 走势图的 X 轴表头用 HTML flex 盒另画 | flex 盒模型受 padding/滚动条/亚像素舍入影响，与 canvas 绝对像素坐标不一致，越往右偏移越大 | 表头也用 canvas，与主图共用同一 `pad.left`/`cellW`，做到像素级对齐 |
| 高 canvas 卡片外层 flex 用默认 `align-items:stretch` | 卡片被拉伸到视口高度，内部高 canvas 戳出父容器溢出 | 外层 flex 设 `align-items:flex-start`，卡片随 canvas 内容撑开 |
| 走势图 X 轴分类标签画在 canvas 底部 | 长表滚动后标签滚出视野，无法对照列含义 | 标签放顶部独立 sticky canvas（`top:var(--navh)`），滚动时固定可见 |
| sticky 表头 `top` 写死小于 `.tab-nav` 高度（如 `top:36px`） | 表头顶部被导航栏遮挡一截 | `top` 用 `--navh` 变量对齐 tab-nav 实际高度（约 50px，取 51px 留缝）|
| 属性色块用高饱和深色（`#e67e22`/`#2e7d32` 等）| 大面积深色刺眼，且叠白字/黑字都易糊 | 统一浅色系，文字色按亮度自适应（`textColorFor`）|
| 色块单元格只填色不写文案 | 需悬停/看表头才知含义，可读性差 | 填色后在格子中央绘制对应文字（红波/大/家畜等）|
| `os.listdir` 取数据源不过滤 `~$` 文件 | Excel 打开时产生的临时锁文件被当数据源，触发 `PermissionError` | 扫描 `.xlsx` 时排除 `startswith('~$')` 的锁文件（波色模板已隔离在 `01-template/`，不在数据目录） |
| 换新 Excel 时直接沿用旧列偏移 | 同系列文件存在"整体偏移"变体（`分析结果`→`公式统计` 右移3列/下移2行），套旧偏移会读到全空、页面显示 0/`-` 但仍生成文件 | 先 dump 网格核对每个区块起点（见"新 Excel 排查步骤"），再改提取常量 |
| 把源数据的 `#N/A`/`#VALUE!` 当提取 bug | 这些是 Excel 公式未算出的值（如生肖偏差列），非代码问题 | try/except 优雅跳过，如实告知用户是源数据缺失；不要为凑数据而猜测填值 |
| txt2excel 产出的 29 列 xlsx 直接喂给 `excel2html_lottery.py` 不改列映射 | v5 版 `load_rows` 读 bose col29/tx col19/wx col26，29 列版对应在 col13/col20/col27，不修正则波色全空、生肖/五行串位 | 先 dump 表头确认列结构，按映射表修正 `load_rows` 中的列号；参见 excel2html_lottery 小节的列映射表 |
| 历史记录表「颜色」列名称 | 用户误解为颜色码而非特码号码 | 列头用「特码」，单元格展示号码（如 `01`），带波色底色 |
| 历史记录「波色」列只有文字无背景色 | 纯文字 `红波/蓝波/绿波` 缺乏视觉区分 | 每个单元格加 `boseCls` CSS 类：红波`#F1A2AB`/蓝波`#81BBF8`/绿波`#C1E77E` |

============================================================

### excel2html_lottery.py —— 开奖记录明细 → 仪表盘

用于 `开奖记录-全年份合并-v5.xlsx` 这类**开奖记录明细**（区别于公式统计）。也支持功能一 `txt2excel.py` 产出的 29 列 Excel（通过 `--split/--lunar-date/--bose-table` 生成）。

- 用法：`python 02-script/excel2html_lottery.py [xlsx路径]`（默认取 `04-output/开奖记录-全年份合并-v5.xlsx`）。
- **数据来源**：只读「开奖明细」sheet，在 Python 里**重算全部统计**（复用功能一「扩展」小节的口径），**不解析「统计分析/分年统计」sheet 的单元格**——按行列位置解析脆弱，重算更稳、逻辑已验证。
- **Excel 列结构**：该脚本有两种列映射，取决于输入 xlsx 来源：

| 字段 | v5 标准版（44 列） | txt2excel 产出（29 列） | 说明 |
|------|-------------------|----------------------|------|
| `id` | col 1 | col 1 | 同 |
| `year` | col 2 | col 2 | 同 |
| `qishu` | col 3 | col 3 | 同 |
| `date` | col 4 | col 4 | 同 |
| `tema` | col 12 | col 12 | 特码，同 |
| `plain` | col 6-11 | col 6-11 | 平码，同 |
| `bose` | col 29 | **col 13** | ⚠️ 不同！29 列版在特码后的 col 13 |
| `tx` (生肖) | col 19 | **col 20** | ⚠️ 不同！29 列版在 col 20 |
| `wx` (五行) | col 26 | **col 27** | ⚠️ 不同！29 列版在 col 27 |

> **⚠️ 换入 txt2excel 产出的 xlsx 时，必须修正 `load_rows` 中的 col 19→20、col 26→27、col 29→13**，否则 bose/tx/wx 全读错（波色列为空、生肖/五行串位）。

- **验证列映射**：换入新 xlsx 后先 dump 表头确认结构：
  ```python
  import openpyxl
  wb = openpyxl.load_workbook(path, data_only=True)
  ws = wb['开奖明细']
  for c in range(1, ws.max_column+1):
      print(f'col {c}: {ws.cell(1, c).value}')
  wb.close()
  ```
- 输出 `04-output/{名}-report-{时间戳}.html`，重名加 `-vN`（不删旧产物）。
- **生肖映射表**：`load_zodiac_map()` 从 `01-template/波色号码对照表.xlsx` 提取各年份的号码→生肖映射 `{2026:{1:'猴'...},2025:{1:'兔'...},...}`，存入 `DATA.zodiac_map`。JS 侧用 `DATA.zodiac_map[year][num]` 取当年生肖。表格结构为每两年一组（奇数行=肖、偶数行=号码），cols 2-50 对应号码 1-49。生辰标签色：野兽（鼠虎兔龙蛇猴）红底 `#F1948A` / 家禽（牛马羊鸡狗猪）绿底 `#82C99A`。
- **8 个 Tab（顺序）**：首页概览 / 频率统计 / 遗漏·冷热 / **分年统计** / 走势图 / 走势图(分年) / 历史记录 / 连开·重号。（分年统计紧跟遗漏·冷热之后）
- **频率统计 & 分年统计 区块顺序**（`freqs` 列表，参见 `build_stats` 中定义）：1大小 / 2单双 / 3大小单双组合 / 4波色 / 5质合 / 6家禽野兽 / 7五行 / 8号头(0-4) / 9生肖 / 10号尾(0-9)。左右两列 `grid2` 布局，左侧锚点菜单对应各区块 id。
- **左侧 sticky 锚点导航**（`withSideNav`/`attachSideNav`）：频率统计、遗漏·冷热、分年统计三个 tab 内容区左侧各有固定竖菜单，点击平滑滚到对应区块（区块须带 `id`，菜单项 `data-target` 指向）；点击项高亮。窄屏（<820px）自动转横排。⚠️ 加了 150px 左菜单后右侧变窄，表格易换行——须给 `td` 加 `white-space:nowrap`、给 `.side-body .card` 加 `overflow-x:auto`（表格超宽时横向滚动而非挤压换行）。
- **历史记录：顶部年份按钮 + 左侧 sticky 月份菜单**（`histYearMap`/`fillHistYear`/`fillHistory`）：顶部一排年份 pill（新→旧，默认最新年），左侧竖菜单列出**当前选中年**的月份（新→旧，默认当年最新月）；切年时左侧月份列表联动刷新。表格月内按日期正序（旧→新）。
- **首页概览底部 4 个生肖横条报表**：生肖偏差排行（总览）、生肖出现频率（总览）、生肖偏差排行（最新年·当年）、生肖出现频率（最新年·当年）。样式：横条形图，生肖名标签**野兽 `#F1948A`（浅红、白字）/ 家禽 `#82C99A`（浅绿、白字）**（野兽=鼠虎兔龙蛇猴，家禽=牛马羊鸡狗猪）；偏差排行=按「当前遗漏期数」降序（越久未出偏差越大，红色条+`+N`值），带阈值标签 `忽略<25≤提醒<30≤可以追`（灰/黄`#fff3cd`/绿`#c8e6c9`底）；频率条内显次数、右侧显实际占比%。数据取 `DATA.all.rank_zod` 与 `DATA.years[末项]`（当年）。
- **遗漏/冷热页布局**：号码遗漏 与 号码热温冷分层 左右各 50%（`grid2` 并排）；生肖遗漏 与 属性遗漏并排；**属性遗漏拆成 3 张表**（大小/单双/波色 一张、号头一张、号尾一张）。分年统计页同理（热温冷 与 遗漏榜 各 50%）。

- **分年统计热温冷/遗漏榜 生肖列**：号码后紧跟生肖列，生肖从 `DATA.zodiac_map` **按当年年份**取对应映射（每年号码↔生肖不同），用 `zxTag(n)` 生成野兽红底/家禽绿底标签。`hotFn` 的 `ci` 索引相应后移（当前遗漏从 ci=1→ci=2，热温冷 lvlCol 从 3→4，遗漏排行榜 gapCol/mxCol 各+1）。
- **表格排序**：`tableHtml` 传 `sortAll:true` 让**所有列**（含文本列）可点击排序；表头默认显示灰色 `⇅` 图标提示可排，点击升序 `▲`/降序 `▼`（蓝色），切到别列时原列图标复位为 `⇅`。数值列按数值排、文本列按字典序。`tableHtml` 支持的 option：`sortable`（可排序列索引）/ `sortAll`（全列可排）/ `boseCol`（波色底色列）/ `lvlCol`（热温冷底色列）/ `pctCols`（百分比格式列）/ `typeBgCol`（类型列底色）/ `hotFn`（标红函数）/ **`htmlCols`**（该列值为原始 HTML，不转义——用于嵌入 `<b>`/`<span>` 等标签）。
- **频率表「类型」列底色**（`typeBgCol`）：生肖/家禽野兽用 `#82C99A`(家禽)/`#F1948A`(野兽) 白字加粗；大小/单双/波色/质合/号头号尾/五行等用各自柔和浅色（见 `typeBgStyle`）。
- **生肖名带野/家后缀**（`zodLabel`）：除首页概览的横条报表外，所有含生肖的表格（频率·生肖、生肖遗漏、生肖遗漏排行榜、分年·生肖频率）显示为 `鼠(野)`/`牛(家)`（野兽=鼠虎兔龙蛇猴、家禽=牛马羊鸡狗猪）。`typeBgStyle` 会先剥离 `(野)/(家)` 后缀再判底色。首页概览 4 个横条报表保持纯生肖名。
- **需要关注的值加粗标红**（`.hot-val` = `#e74c3c` 粗体，`hotFn`/`makeMissHot` 生成）：遗漏冷热与分年统计里——当前遗漏 `≥30`、热温冷"冷"号行的遗漏、当前遗漏 Top3、历史最大遗漏 Top3、频率表实际占比 `>1.3×均值`，均标红。
- **分年统计二级 tab 从新到旧**（2026→2020，默认选最新年）；按钮倒序渲染但 `fillYear(idx)` 仍指向 `DATA.years` 真实下标。
- **走势图 / 走势图(分年)**（`renderTrend`/`renderTrendYear`+`TREND_DEFS`）：**3 个子级 tab**（属性走势 / 号码走势 / 遗漏报表），子 tab 状态在面板内独立保持。

  **① 属性走势**：4 个属性色块图并排（大小/单双/波色/家禽野兽），每列左带日期+期数，命中格填色+文字。色块：大`#F8C471`/小`#C39BD3`、单`#85C1E9`/双`#F7DC6F`、红`#F1A2AB`/蓝`#81BBF8`/绿`#C1E77E`、家禽`#82C99A`/野兽`#F1948A`。

  **② 号码走势**：双 Canvas 架构（表头 sticky + 内容滚动）。表头含日期/期数/1-49 列头，带 `#999` 边框，`position:sticky;top:0`。内容每格 21×21px，上=新→下=旧（最近 200 期），全部格子有可见边框（`#999` 1px），未命中浅灰底 `#fcfcfc`，命中格按波色着色+写号码文字。日期列 72px、期数列 58px，内容居中。数据取 `stats.trend`（含 `tema` 序列）和 `bose` 序列。

  **③ 遗漏报表**：可排序表格（`tableHtml` + `attachSort`），9 列=号码/生肖/当前遗漏/遗漏可视化/出现次数/历史最大/最近·年/最近·期/最近·日期。遗漏可视化=横条（绿<15<黄<30<红），`gap-cell` 容器左对齐。生肖从 `DATA.zodiac_map` 按最新年份取映射，野兽红底/家禽绿底。当前遗漏 ≥30 标红。分年版年份 pill 倒序切换。

- 数据取 `stats.trend`（labels 期数 + dates 阳历日期 + dx/ds/bose/qs/tema 序列，新→旧）。
- **历史记录**（`renderHistory`/`fillHistory`+`build_history`）：完整开奖明细表，14 列=周期(id)/日序(月内序号)/日期/期数/生肖/**特码**(带波色底)/大小/单双/**波色**(带波色底+文案 红波/蓝波/绿波)/野兽/号头/号尾/颜色单双/组合；按月份 pill 切换（默认最新月），月内按日期正序（旧→新）。`build_history` 从明细重算，`load_rows` 需取 `id`(第1列)。
- 复用 excel2html_formula.py 的浅色 Tab 视觉：`--navh` sticky 导航、卡片、`tableHtml` 可排序表格。纯离线无 CDN。`.tab-page` **宽度占 96%**（`width:96%;margin:0 auto`），内部表格/网格自适应。
- **配色（柔和浅色，避免刺眼）**：仅**家禽 `#82C99A` / 野兽 `#F1948A`** 用较饱和色（用户指定）；其余全部浅色系——波色表格/卡片底色 红`#fbe0e3`/蓝`#dcebfb`/绿`#e6f4d5`，热温冷 热`#fce4d6`/温`#fff2cc`/冷`#dce9f7`，首页分布卡片 大`#fdebd0`/小`#e8ddf3`/单`#d6eaf8`/双`#fdf3d0`。覆盖频率统计/遗漏冷热/分年统计/连开重号各页。
- 验证：**① 先用 node 校验生成 HTML 的 JS 语法**（改过模板 JS 后必做，比只看浏览器更早发现问题）：`node -e "const fs=require('fs');const h=fs.readFileSync('产物.html','utf8');const m=h.match(/<script>([\s\S]*)<\/script>/);try{new Function(m[1]);console.log('JS OK')}catch(e){console.log('JS ERROR:',e.message)}"`——⚠️ 曾因编辑时残留重复 `function xxx(){` 行导致 JS 语法错、整页白屏（浏览器只报 `Unexpected end of input`），node 校验 + 括号平衡扫描能快速定位。② 反解 HTML 里 `var DATA` 核对 total/热号/年份数/**zodiac_map 年数**/**trend.tema 长度 = total**。③ 起本地 http 服务（`python -m http.server` 于 `04-output`；file:// 被 playwright 拦，URL 需 encodeURI 中文名）用浏览器 `browser_evaluate` 驱动确认 Tab 切换/顺序、排序、分年切换、生肖后缀、走势图子tab切换/号码走势canvas、历史记录年月联动、左侧锚点导航、波色/热温冷底色、标红。收尾删 `.playwright-mcp/` 临时目录、`pkill` 关掉 http 服务。

============================================================

# 功能三：myself-pyhtml — 独立一键生成工具

> 位于 `myself-exceltohtml` 同级的 `myself-pyhtml/` 目录，独立运行，无需 AI。

## 概述

将功能一+功能二封装为独立工具，供**非技术人员双击使用**。含接口数据拉取 + 环境检测 + 一键生成。

```
myself-pyhtml/
├── run.bat                  唯一入口（双击）
├── 看前必读.txt
├── 01-data/
│   ├── 01-template/         波色号码对照表.xlsx（已扩展到2031年）
│   └── 02-txt-source/       7个年份txt（2020-2026）
├── 02-script/
│   ├── config.json           接口配置（url/headers/参数）
│   ├── requirements.txt      依赖清单
│   ├── env_install.bat       环境一键安装（纯bat，首次用）
│   ├── env_setup.py          环境检测 + 缺依赖自动安装
│   ├── fetch_data.py         从接口拉取当年数据
│   └── generate.py           合并txt → xlsx → html
└── output/                   产物目录
```

**命名规范**：`env_` 前缀 = 环境脚本，无前缀 = 业务脚本。

## 工作流

```
run.bat 双击
  → env_setup.py     检查 Python/pip/依赖，缺失自动安装
    ├─ 通过 → 自动继续
    └─ 未通过 → 按回车退出
  → fetch_data.py    调用接口拉取当年数据，按id去重合并
    ├─ 成功 → 写入 临时测试-{year}.txt
    └─ 失败 → 跳过，用现有数据
  → generate.py      增量判断 → txt2excel → excel2html_lottery
    → output/*.xlsx + output/*-report-*.html
```

## fetch_data.py 接口数据拉取

**调用方式**：POST JSON，配置集中在 `02-script/config.json`：

```json
{
  "api": {
    "url": "https://kj.9bkj.com:1888/kj",
    "method": "POST",
    "timeout": 300,
    "headers": { "Content-Type": "application/json", "Referer": "..." },
    "body": { "g": "am" }
  }
}
```

| 配置项 | 说明 |
|--------|------|
| `url` | 接口地址（后端数据服务器） |
| `headers.Referer` | 前端网站地址（跨域请求用，与 url 不同属正常） |
| `body` | 请求参数，年份 `y` 由 `datetime.now().year` 自动填充 |
| `timeout` | 超时秒数，接口响应较慢（可能数分钟），默认 300s |

**接口变更时只需编辑 config.json，不改代码。**

**去重逻辑**：按 `id` 字段去重，新记录追加，已有的跳过。

**跨年自动适配**：
- `fetch_data.py` 用 `datetime.now().year` 动态生成接口参数 `y` 和文件名 `临时测试-{year}.txt`
- 到 2027 年自动调 `y=2027`、自动创建 `临时测试-2027.txt`
- `generate.py` glob 所有 `*.txt`，新文件自动纳入合并

**时效判断**：通常每晚 23:00 后调用，接口返回当天全部数据。`generate.py` 通过 xlsx 中 max date 与 txt 数据对比判断是否需要重新生成。

## env_setup.py 环境检测

| 检测项 | 通过 | 失败 |
|--------|------|------|
| Python >=3.7 | 继续 | 提示下载地址 → 按回车退出 |
| pip | 继续 | 自动 `ensurepip` 安装 |
| openpyxl | 继续 | 自动 `pip install` |
| pandas | 继续 | 自动 `pip install` |
| requests | 继续 | 自动 `pip install` |

全部通过 → `sys.exit(0)` → run.bat 自动继续执行后续任务。失败 → `sys.exit(1)` → run.bat 暂停等待确认。

## 与功能一/功能二的关系

- `generate.py` 通过 subprocess 调用 `myself-exceltohtml/02-script/` 下的 `txt2excel.py` + `excel2html_lottery.py`
- `myself-pyhtml` 依赖 `myself-exceltohtml` 存在（同目录），不独立复制脚本代码
- Excel 列映射在 `excel2html_lottery.py` 的 `load_rows` 中已适配 29 列格式

============================================================

## 边界说明

本技能覆盖两个功能，各自边界如下：

**功能一（TXT→Excel）仅覆盖**：结构化文本文件到 `.xlsx` 的格式转换。

| 超出范围 | 处理方式 |
|--------------|---------|
| 非结构化文本 | 停止，告知用户先结构化 |
| 数据分析、图表绘制 | 停止，指向对应功能/技能 |
| Excel 转文本 | 停止，告知超出范围 |

**功能二（Excel→HTML）仅覆盖**：将指定格式的彩票分析 Excel 转为 HTML 仪表盘。

| 超出范围 | 处理方式 |
|--------------|---------|
| Excel 结构不同（不同列偏移/新增 sheet） | 按"换入新 Excel 排查步骤"dump 网格、核对区块起点后调整提取常量；已知 `分析结果`/`公式统计`（formula.py）与 `v5版/txt2excel版`（lottery.py）两类变体，后者需修正 `load_rows` 列映射（见 excel2html_lottery 小节列映射表） |
| 用户想修改 Excel 源数据 | 拒绝，告知只读操作 |
| 用户要求做数据分析或预测 | 拒绝，这是可视化工具，不做预测 |

# easydata-agent-testtool 使用指南

## 这是什么

测试 **AI Agent 能不能正确理解用户意图并调用对工具** 的自动化工具。
输入 MCP 工具定义 (JSON Schema) + 你填写的真实数据 → 输出自然语言测试用例 Excel + 可执行的 YAML + 分析报告。

## 快速上手 (完整流程)

### 第一步: 唤起 Skill
在聊天框里直接说 (任选一种):
```
用 easydata-agent-testtool 帮我生成测试用例
```

### 第二步: AI 询问时选择步骤
AI 会先展示一张表格让你选择从哪个步骤开始:

| 步骤 | 做什么 | 什么时候用 |
|------|--------|-----------|
| Step1 | 提取 MCP Schema 必填参数 → 生成 `2-data/{name}.md` 模板 | 第一次用, 需要知道每个工具有哪些参数 |
| Step2 | 生成自然语言用例 Excel | 你已填好 `2-data/{name}.md` 中的真实数据 |
| Step3 | ★ 执行测试 + 生成分析报告 | 已有 TC 用例 Excel, 要用双方案(NL+CLI)执行并生成报告 |
| Step4 | 转 YAML 格式 | Excel 用例已完成, 要导入评估系统执行 |
| Step5 | Excel → TC 模板 | 需要把 Step2 的 Excel 转成 TC 模板格式 |
| Step6 | TC 模板 → YAML | 已有 TC 模板格式的 Excel (非 Step2 产物), 要转成 YAML |

**工作流 A**: Step1 → **你手动填写真实数据** → Step2 → Step4
**工作流 B**: Step2 → Step3 (执行测试 + 报告)
**工作流 C**: Step2 → Step5 (Excel 转 TC 模板)
**工作流 D**: Step6 独立使用 (已有 TC 模板 Excel, 直接转 YAML)

**给不同用户的建议**:

| 你的情况 | 建议 |
|---------|------|
| **第一次用, 什么都没准备** | 选 Step1, 拿到模板后填入真实数据, 再走 Step2 + Step4 |
| **已经填好 2-data/{name}.md 真实数据** | 选 "从 Step2 开始" |
| **只想要 Excel 手工看** | 选 "Step2", 不做后续 |
| **已有 TC 用例 Excel, 想执行测试看结果** | 选 "Step3", 双方案执行 + 生成分析报告 |
| **已经有别人给的 Excel, 想转 YAML** | 选 "Step4", 把 Excel 转 YAML |
| **想要 TC 模板格式的 Excel** | 选 "Step5", 把 Step2 的 Excel 转格式 |
| **已有 TC 模板格式的 Excel (13列)** | 选 "Step6", 直接按产品模块拆分转 YAML |

### 第三步: 跟着 AI 一步步走
AI 会在每个步骤执行前展示具体做什么, 你只需要:
1. 看表格确认 AI 要做什么
2. 回复 "可以" 或 "开始"
3. Step1 后: 打开 `2-data/{name}.md`, 把 `待填写` 改成真实数据
4. AI 执行后会告诉你产出文件在哪里

## 完整流程示例
假设你要为 `控制台 (easyconsole)` 生成测试用例:
```
你:   帮我对 easyconsole 生成 agent 测试用例
AI:   本工具提供 6 个独立步骤... 请问从哪个开始?
你:   Step1, schema 在 "easydata-agent-testtool/1-resource/mcpTool/easyconsole"
AI:   [Step1] 开始提取 config.json 的参数...
      | tool | required_params |
      | list_schedule_calendars | product |
      已生成 2-data/easyconsole.md，含 1 个工具、1 个必填参数（标记为待填写）
      请手动填写真实数据后回复"继续 Step2"
你:   (打开 2-data/easyconsole.md, 把 product 的待填写改成 mammut)
      继续 Step2
AI:   [Step2] generate.py path1 --schema-dir "1-resource/mcpTool/easyconsole" --product easyconsole
      生成 N 条用例 -> 3-output/easyconsole-20260605-xxxx.xlsx
      确认继续 Step4 (转 YAML)?
你:   继续
AI:   [Step4] generate.py path2 --excel 3-output/easyconsole-20260605-xxxx.xlsx
      生成 N 条 YAML -> 3-output/easyconsole-20260605-xxxx.yml
      全部完成!
```

### Step3 独立使用示例
假设你已有 TC 用例 Excel (`easytaskops-tag-20260609-142907.xlsx`)，想执行测试并生成报告:
```
你:   用 easydata-agent-testtool 帮我对这个 Excel 执行测试，生成报告
AI:   本工具提供 6 个独立步骤... 请问从哪个开始?
你:   Step3
AI:   [Step3] 读取 easytaskops-tag-20260609-142907.xlsx
      发现 3 条测试用例，开始双方案执行...
      
      --- 方案一: 自然语言 ---
      TC-0001: "帮我查下 autoapi_test 项目里有哪些标签" → PASS
      TC-0002: "翻到第3页，每页10个" → PASS
      TC-0003: "有没有叫 noexist_tag_xyz 的标签" → PASS
      
      --- 方案二: CLI命令 ---
      TC-0001: list_tag 默认分页 → PASS
      TC-0002: pageNum=3, pageSize=10 → PASS
      TC-0003: 空结果查询 → PASS
      
      6/6 PASS (100%)
      报告已生成: 4-report/testReport-easytaskops-tag-20260609-142907.xlsx
      (4 Sheet: 测试汇总/测试结论/自然语言测试/CLI命令测试)
```

## 各步骤详解

### Step1: 提取 MCP Schema 必填参数 → 生成参数模板
**做什么**: 读取 JSON 文件, 告诉你每个工具有哪些必填参数, 参数类型是什么, **并自动生成 `2-data/{serviceName}.md` 模板文件**。
**产出**:
1. 参数清单分析 (表格形式)
2. `2-data/{serviceName}.md` — 必填参数模板, 各参数标记为`待填写`
**你需要做的**: 打开生成的 MD 文件, 把 `待填写` 改成真实数据

**示例输出**:
```
从 config.json 提取 1 个工具:
  list_schedule_calendars: required=[product]
已生成 2-data/easyconsole.md, 含 1 个工具、1 个必填参数（均标记为待填写）
请手动填写真实数据后回复"继续 Step2"
```

### Step2: 生成自然语言用例 Excel
**做什么**: 读取你填好的 `2-data/{name}.md` + JSON Schema, 自动生成带**真实数据**的自然语言用例
**前置条件**: `2-data/{name}.md` 中不能有 `待填写` (必须先完成填写)
**命令**:
```bash
python easydata-agent-testtool/scripts/generate.py path1 \
  --schema-dir "easydata-agent-testtool/1-resource/mcpTool/easyconsole" \
  --product easyconsole
```
**产出**: `3-output/{name}-{YYYYMMDD}-{HHMMSS}.xlsx` (11 列, 5~6 个 Sheet；多工具产品含「6-场景组合」)

**自动生成的 Prompt 示例**:
| 类别 | Prompt |
|------|--------|
| 正向 | "帮我查一下 mammut 项目的调度日历列表" |
| 缺失参数 | "帮我查一下调度日历列表，但不要提供项目" |
| 异常边界 | "帮我查一下，用不存在的 product" |
| 安全注入 | "帮我操作功能，参数里包含 ' OR '1'='1 这样的内容" |

### Step3: ★ 执行测试 + 生成分析报告
**做什么**: 读取 TC 用例 Excel，以 **两种方案并行** 执行测试并生成 4-Sheet 结构化分析报告
**前置条件**: 已有 TC 用例 Excel（Step2 或 Step5 产物），`easydata-cli` 已配置

**双方案**:
- **方案一 (自然语言)**: AI 理解 Prompt → Schema匹配 → 参数提取 → 执行 CLI → 记录过程
- **方案二 (CLI命令)**: 直接构造 JSON 参数 → 执行 CLI → Schema校验

**命令**:
```bash
# AI 执行测试后，调用脚本生成报告
python easydata-agent-testtool/scripts/step3_gen_report.py \
  --source "原始用例.xlsx" \
  --results "测试结果.json" \
  --outdir "easydata-agent-testtool/4-report/"
```
**产出**: `4-report/testReport-{name}.xlsx` (4 Sheet) + `4-report/reportHtml/testReport-{name}.html` (可视化 HTML)

**报告特色**:
- 🟠 人工复核列橙色标题 — 区分自动判定和人工确认
- 🟡 空白复核位黄色底色 — 引导填写
- 📊 概要表 + 双方案对比 — 一目了然
- 🔢 自动版本管理 (`-v1`, `-v2`)

**后处理①  添加分组标题行**: 报告默认平铺所有用例。如需按源 Sheet（1-正向/2-异常/3-安全/4-回归/5-敏感/6-场景组合）分组展示：
```bash
python easydata-agent-testtool/scripts/add_section_headers_by_prefix.py \
  --report "4-report/testReport-{name}.xlsx"
```
按用例标题里的 `[源Sheet]` 前缀分组，在"自然语言测试"和"CLI命令测试"中各插入深蓝底白字标题行（带图标+用例数）。用安全的"快照→重写"法，不会损坏数据。

**后处理②  重建 HTML（补单场景栏）**: 自动产出的 HTML 若只有【多工具链场景结果】栏、缺【单场景用例】栏，或 Excel 改动后需重新出 HTML：
```bash
python easydata-agent-testtool/scripts/gen_html_from_report.py \
  --report "4-report/testReport-{name}.xlsx"
```
直接读报告 Excel 重建，自动分【单场景用例(按层级)】+【多工具链场景】两栏，默认覆盖同名 .html。

### Step4: Excel → YAML 转换
**做什么**: 把 Step2 的 Excel 转成 `easydata-mode` YAML, 可直接导入 Agent 评估系统跑
**命令**:
```bash
python easydata-agent-testtool/scripts/generate.py path2 \
  --excel "3-output/{name}-{YYYYMMDD}-{HHMMSS}.xlsx"
```
**产出**: `3-output/{name}-{YYYYMMDD}-{HHMMSS}.yml`

### Step5: Excel → TC 模板 Excel
**做什么**: 把 Step2 的 Excel (5~6 Sheet) 按 `1-resource/tcCaseTemplate/template.xlsx` 格式转换
**两种模式**: `--mode nl`(默认,执行步骤填自然语言 Prompt) / `--mode cli`(执行步骤填 CLI 命令)
**命令**:
```bash
# 自然语言版 → 5-tcCase/TCCase-{name}.xlsx
python easydata-agent-testtool/scripts/generate.py path4 \
  --excel "3-output/{name}.xlsx"

# CLI 命令版 → 5-tcCase/tcCaseCLI/TCCase-{name}-CLI.xlsx
python easydata-agent-testtool/scripts/generate.py path4 \
  --excel "3-output/{name}.xlsx" --mode cli
```
**产出**:
- `nl`: `5-tcCase/TCCase-{name}.xlsx`
- `cli`: `5-tcCase/tcCaseCLI/TCCase-{name}-CLI.xlsx`（仅"执行步骤"列换成 CLI 命令，其余列同自然语言版）

**防覆盖**: 同名产物已存在时自动追加 `-v1`/`-v2`…，不会冲掉旧文件（如需指定路径用 `--output`）。

### Step6: TC 模板 Excel → YAML (按模块拆分)
**做什么**: 读取 TC 模板格式的 Excel (13列: Suite一~五 + 测试用例 + 用例分级 + 前置条件 + 执行步骤 + 预期结果 + 备注 + 用例Tag), AI 分析后按 **Suite二 (产品模块)** 拆分为多个 YAML 文件
**适用场景**: 已有 TC 模板 Excel (非 Step2 产物), 需要转 YAML 导入评估系统

**AI 分析功能**:
- 🔍 **跨列工具名识别**: 从 Suite四/三/二 中自动识别 snake_case 格式的 MCP 工具名
- 🏷️ **分类推断**: 自动识别 boundary (异常边界) / stability (安全回归) / normal (正向功能)
- 📝 **Promp 生成**: 优先执行步骤, 为空则用标题

**命令**:
```bash
python easydata-agent-testtool/scripts/convert_tc_to_yaml.py \
  --excel "path/to/TC模板.xlsx" \
  --outdir "easydata-agent-testtool/3-output/"
```
**产出**: `3-output/{Excel名}-{Suite二}.yml` (每个产品模块一个文件)

## 文件存放位置
```
easydata-agent-testtool/
├── SKILL.md                             # Skill 定义 (给 AI 看的)
├── README.md                            # 本文件 (给你看的)
├── 1-resource/                         # Step1/Step5 输入素材
│   ├── mcpTool/                         # MCP Schema JSON (按产品分目录)
│   │   ├── easyconsole/config.json
│   │   ├── easydqc/dqc.json
│   │   └── easytaskops/{tag,task,log,...}.json
│   └── tcCaseTemplate/                  # Step5 TC 用例模板
│       └── template.xlsx
├── 2-data/                              # 参数模板 + 真实数据 (按产品维护)
│   └── easyconsole.md                   # Step1 生成, 你手动填写真实值
├── scripts/
│   ├── generate.py                      # Step1 + Step4 + Step5 引擎
│   ├── step2_write_excel.py             # Step2: AI 用例 → 格式化 Excel
│   ├── step3_gen_report.py              # Step3: 测试结果 → 4-Sheet 报告 (自动出 HTML)
│   ├── gen_html_report.py               # Step3+: 测试结果 JSON → 可视化 HTML (step3 自动调)
│   ├── add_section_headers_by_prefix.py # Step3 后处理①: 按[源Sheet]前缀加分组标题行 (推荐)
│   ├── gen_html_from_report.py          # Step3 后处理②: 读报告Excel重建HTML (单场景+链场景两栏)
│   ├── add_section_headers.py           # (旧)按标题匹配源Excel — 标题带前缀失效, 勿用
│   └── convert_tc_to_yaml.py            # Step6: TC 模板 → 按模块拆分 YAML
├── 3-output/                            # 产物 (Step2/Step4/Step6)
│   ├── easyconsole-20260605-xxxx.xlsx   # Step2 产物
│   ├── easyconsole-20260605-xxxx.yml    # Step4 产物
│   └── {Excel名}-{Suite二}.yml          # Step6 产物 (按模块拆分)
├── 5-tcCase/                            # Step5 产物
│   ├── TCCase-{name}.xlsx               #   自然语言版 (--mode nl)
│   └── tcCaseCLI/                       #   CLI 命令版 (--mode cli)
│       └── TCCase-{name}-CLI.xlsx
└── 4-report/                            # Step3 产物: 测试报告
    ├── testReport-{name}.xlsx           # 4-Sheet 分析报告
    └── reportHtml/                      # HTML 报告子目录
        └── testReport-{name}.html       # 可视化 HTML 报告
```

## 常见问题

### Q: 第一次用, 需要做什么准备?
1. 确认 `python` 能用, `openpyxl` 已安装 (`pip install openpyxl`)
2. 记住你的 Schema 文件路径 (如 `1-resource/mcpTool/easyconsole/`)
3. 准备好你的真实数据 (product 名、集群名等)

### Q: Step1 生成的模板怎么填?
打开 `2-data/{name}.md`, 找到每个参数下面的 `待填写`, 替换成你项目中真实的值。比如:
```markdown
- **product** (string/resource_id): 产品账号
  - 真实值: `mammut`     ← 把"待填写"改成这个
```

### Q: 怎么加入更多真实数据?
在 `2-data/{name}.md` 中补充更多参数值。非必填参数也可以写进去, Step2 会用到。

### Q: 生成的 Excel 可以自己改吗?
可以。Step2 的 Excel 就是给你审查用的。你可以在 Excel 里手动修改 prompt, 改完后再走 Step4 转 YAML。

### Q: Step4 和 Step6 有什么区别?
- **Step4**: 处理 Step2 生成的 6-Sheet Excel (含场景组合)，转 easydata-mode YAML
- **Step6**: 处理 TC 模板格式 Excel (13 列, 含 Suite 层级), AI 跨列分析工具名, 按产品模块拆分输出

### Q: Step3 和其他步骤有什么不同?
- **Step3** 是执行步骤，不是生成步骤 — 它会**真实调用 easydata-cli** 执行 API，并收集结果
- 前置条件: `easydata-cli` 必须已配置 (`config check` 返回 `configured: true`)
- 产物: 4-Sheet 分析报告 Excel，含人工复核列

### Q: Step3 两种方案有什么区别?
- **自然语言方案**: 模拟真实用户输入 → 验证 AI Agent 能否正确理解意图并调用对工具
- **CLI命令方案**: 直接传参 → 验证接口 Schema 和参数处理是否正确
- 两者结果应一致，不一致时需人工排查

### Q: Step3 报告可以自己改吗?
可以。报告 Excel 中的"人工复核-结论"和"复核备注"列就是给你填写的。你也可以直接修改其他列的文案。

### Q: Step6 能自动识别工具名吗?
能。AI 会从 Suite四/三/二 列中识别 snake_case 格式的工具名 (如 `list_yarn_queues`, `get_datasource_list`), 过滤中文分类名 (如 `1.基础资源查询`)。匹配不到工具名时只保留 `answer_correctness` 评分器。

### Q: 其他子产品能用吗?
能。把 `--schema-dir` 换成对应的 Schema 目录, `--product` 换成对应名字。

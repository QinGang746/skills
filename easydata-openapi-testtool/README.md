# easydata-openapi-testtool — OpenAPI 接口测试用例生成器

从 **OpenAPI 接口文档** 到 **接口测试用例 + 执行报告** 的 3 阶段 AI 驱动管线（step1 文档→分析 / step2 设计用例 / step3 执行+报告）。每一步独立可执行，各自产出一份文件，便于人工核对与断点续跑。

```
【step1】落 1-input/：①转纯文本(脚本) → ②接口清单 → ③接口详情(入参/出参)
【step2】落 2-output/：用例 Excel(按接口拆 sheet（内按分类加标题行）,每条带curl)
【step3】落 3-report/：执行 curl → 执行报告(Excel+HTML)
```

## 设计原则

1. **提取、设计全部由 AI 完成**：接口识别、入参出参提取、curl 拼接、用例设计都是 AI 读文档后现场判断；脚本只做机械活——docx 转文本、把用例渲染成 Excel、执行 curl 出报告。
2. **接口文档不固定**：不写死任何接口路径、字段名、章节号、版式假设，适配任意 OpenAPI 文档。
3. **每条用例配可运行 curl**：URL 前缀统一（默认 `http://10.45.132.131:10098/openapi/easytaskops`），鉴权参数 `authType=TEST`，复制即可执行。curl **并入「执行步骤」列的最后一步**（不单独成列），使 Excel 12 列对齐平台导入模板 `scripts/template.xlsx`。
4. **每步落文件、防覆盖**：用例 Excel 直接落 2-output/（按接口拆 sheet（内按分类加标题行））；重新生成时旧产物保留，新产物落 `-v{N}` 副本。

## 安装依赖

```bash
pip install openpyxl
# docx 转文本用标准库；执行 curl 用系统自带 curl 命令
```

## 使用方式

### 一句话触发

> "用 easydata-openapi-testtool 帮我对 v1.23.0.docx 生成接口测试用例"

AI 会按 3 阶段管线推进，并在接口较多时先与你确认处理范围。

### 分步说明

| 阶段·子步 | 谁来做 | 输入 | 产物 |
|------|--------|------|------|
| step1·①（仅 docx） | 脚本 step1_extract_docx | `接口文档.docx` | `1-input/序1-转纯文本-{文件名}-{时间戳}.txt` |
| step1·② 提取接口URL | AI | 文档文本 | `1-input/序2-接口清单-{文件名}-{时间戳}.md` |
| step1·③ 提取入参出参 | AI | 文本+接口清单 | `1-input/序3-接口详情-{文件名}-{模块}-{时间戳}.md` |
| step2 设计用例(带curl)+渲染 | AI 设计 + step2_gen_excel | 接口详情 | `2-output/测试用例-{文件名}-{模块}-{时间戳}.xlsx`（按接口拆 sheet（内按分类加标题行）） |
| step3 执行 curl + 报告 | step3_run_curl + step3_gen_html_report | step2 用例 Excel | `3-report/执行报告-…xlsx + reportHtml/执行报告-…html` |

### 脚本命令

```bash
# step1·①：docx 转文本
python scripts/step1_extract_docx.py --docx "path/to/接口文档.docx"

# step2：渲染用例 Excel（AI 先把 testcase_data.py 写到 scripts/，每条用例含第6个 curl 元素，渲染时并入「执行步骤」）
#   --doc/--module 决定文件名，--ts 省略时自动沿用 1-input 批次时间戳；按接口拆 sheet（内按分类加标题行），直接落 2-output/
python scripts/step2_gen_excel.py --project "easytaskops大盘统计接口" \
  --data "scripts/testcase_data.py" --doc v1.23.0 --module 大盘统计 --cleanup

# step3：执行 curl 出 Excel 报告，再转 HTML（两脚本）
python scripts/step3_run_curl.py \
  --excel "2-output/测试用例-v1.23.0-大盘统计-20260618-103307.xlsx" \
  --doc v1.23.0 --module 大盘统计 --timeout 20
# 可选 --dry-run 只解析不执行
python scripts/step3_gen_html_report.py \
  --excel "3-report/执行报告-v1.23.0-大盘统计-20260618-103307.xlsx"
```

> **批次时间戳**：全管线产物沿用 1-input 序1 文件名里的同一个 `{年月日-时分秒}`，光看时间戳即可认出同一批；同名按 `-v1/-v2`（最早不带后缀）递增。step2/step3 脚本 `--ts` 省略时自动从 1-input 提取。

## 目录结构

```
easydata-openapi-testtool/
├── SKILL.md            # 技能定义（完整流程与规则）
├── README.md           # 本文件
├── scripts/
│   ├── step1_extract_docx.py     # docx → 纯文本（保留表格）
│   ├── step2_gen_excel.py        # testcase_data.py → Excel（按接口拆 sheet（内按分类加标题行））
│   ├── step3_run_curl.py         # 执行用例 curl → Excel 执行报告
│   ├── step3_gen_html_report.py  # Excel 执行报告 → HTML
│   ├── generate_excel.py         # Excel 渲染引擎（按接口多 sheet,内按分类加标题行）
│   ├── generate_xmind.py         # 工具函数依赖
│   └── _common.py                # step2/step3 共用工具（命名/防覆盖/读用例）
├── 1-input/            # step1：纯文本 + 接口清单 + 接口详情（序1~序3，统一批次时间戳）
├── 2-output/           # step2：测试用例-…xlsx（直接落此，按接口拆 sheet（内按分类加标题行））
└── 3-report/           # step3：执行报告-…xlsx + reportHtml/ 下的 …html
```

## 用例覆盖度框架（step2）

逐接口按下列维度推导用例（适用即写、不适用即跳过，不凑数）：

- **正向**：仅必填参数 / 必填+可选全参
- **参数边界**：数值/长度/范围的最小、最大、临界值
- **参数校验**：缺必填 → 非法请求；类型错误；枚举非法值
- **业务异常**：不存在的资源、状态不支持的操作（按文档业务响应码验证）
- **权限**：无权限场景（按文档权限响应码验证）
- **安全**：注入、越权（同类合并）
- **出参校验**：字段完整性、分页一致性
- **写操作敏感**：重复创建、幂等、删除/终止类标注谨慎执行

## Excel 列结构（用例 Excel，对齐平台导入模板 `template.xlsx`，共 12 列）

| Suite一 | Suite二 | Suite三 | Suite四 | Suite五 | 用例标题 | 用例分级 | 用例Tag | 前置条件 | 执行步骤 | 预期结果 | 备注 |
|--------|--------|--------|--------|--------|---------|---------|--------|---------|---------|---------|------|
| 项目名 | 接口模块 | 用例类型 | 预留 | 预留 | 用例名 | P0~P3 | 预留 | 多条合并 | 多条合并（**curl 为最后一步**） | 多条合并 | 预留 |

> **平台导入约束**（来自 `template.xlsx` 的「模板使用说明」）：列数/列名须与模板一致（12 列）；「用例分级」只能是 `P0/P1/P2/P3`，否则报错；**合并单元格**导入时会被拆开并给每格填同值——本工具的「分类标题行」是合并行，若目标平台不接受合并单元格，导入前需先删掉标题行（分类信息仍在「用例Suite三」列，不丢失）。

## 执行报告（step3 产物）

- **判定**：HTTP 2xx 且 `code:0` → PASS；校验/异常/安全/边界/权限类拿到错误响应 → PASS(预期)；**负向类却返回 code:0（预期/实际矛盾）→ 需人工复核**；网络错误 → ERROR；正向类无法自动确认 → 需人工复核。
- **判定分析(AI)** 列：每条用例附一句话说明「为什么是这个判定」，PASS 也能看到依据，无需回翻响应体。
- 报告 Excel/HTML 均**按接口拆 sheet/分 section，内按用例类型分组**（与用例 Excel 布局一致）。判定为自动初判，最终结论须人工结合「预期」列复核。

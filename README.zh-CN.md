<p align="center">
  <img src="assets/logo.png" alt="Excel-Ops 标志" width="180">
</p>

# Excel-Ops

中文 | [English](README.md)

### 我讨厌 VLOOKUP。这就是原因。

Excel-Ops 是一套由对话驱动、供 AI Agent 使用的工作流。把文件交给 Agent，用日常语言描述想要的结果；Agent 规划操作、选择合适的本地或云端连接器、隔离不确定项、验证结果，并交付一套可检查的文件，而不是把猜测包装成答案。它并非又一个表格网页应用，也不是 AI 公式生成器。

VLOOKUP 令人头疼的地方从来不只是公式本身。真实表格中的名称不统一、记录重复、表头变动，日期也可能互相矛盾。这个设计源于一项真实且反复发生的报告工作：每周人工核对来自多个来源的杂乱记录；它并非为了演示而虚构的场景。Excel-Ops 把这些问题当作一整套业务流程来处理，而不是指望一个函数就能解决：

**导入 → 提取 → 标准化 → 匹配 → 人工复核 → 写入 → 验证 → 交付**

它与 [PPT-Ops](https://github.com/luochen211/ppt-ops) 是互补项目：PPT-Ops 把原始材料变成可审查的演示文稿交付物；Excel-Ops 把图片和杂乱的表格输入变成可审查的数据交付物。

## 在 Agent 中使用

在 Codex 或其他遵循 Agent Skills 标准的 Agent 中调用仓库技能：

> 使用 `$excel-agent`，从这些截图中提取记录；只把高置信度记录追加到工作簿，其余不确定内容全部放进复核表。

`$excel-agent` 是唯一面向用户的入口。Python CLI 仍作为 Agent 和自动化测试使用的内部确定性能力。

初期连接器的顺序是：本地 XLSX/CSV、本地云同步目录、Google Sheets、Dropbox API、飞书表格和 WPS 表格。云端凭据只应保存在用户环境中，不得写入项目或交付包。

第一个垂直切片面向一种常见企业任务：导入 XLSX 或 CSV 记录（也可使用不依赖特定服务商的图片提取 JSON），识别布局，按声明的数据结构验证，然后只把通过检查的记录写入 Excel 工作簿。不确定的记录进入单独的复核工作表，不会被静默猜测。XLSX 数据页根据内容选择；布局无法识别时，运行会停止，不会生成部分结果。

## MVP 当前实际能做什么

输入从截图、扫描件、收据或表单提取、且不依赖特定服务商的 JSON 后，Excel-Ops 会为每条记录保留源图片名称和提取置信度，验证必填字段，并将结果分流：通过检查的记录进入 `Accepted`；低置信度或不完整的记录进入 `Review`，不会被静默猜测；`Audit` 工作表记录数量和处理元数据。模糊匹配不能自行作出最终决定。

仓库里的代码和示例均未使用真实生产工作流中的文件：没有客户文件、地址、工资记录或凭据。设计受到真实、重复发生的报告流程启发，但仓库中的材料全部是合成的。

## 合成场景示例

假设一家设施服务团队每周收到格式各异的巡检表格，偶尔还会收到手机拍摄的截图。每位客户要求使用不同的工作簿模板交付结果。

Excel-Ops 应当：

1. 识别输入布局；
2. 提取每条观察记录并进行标准化；
3. 使用严格规则将记录匹配到声明的目标；
4. 将不确定的匹配放入人工复核队列；
5. 把通过检查的记录写入对应的工作簿布局；
6. 交付前核对记录数量、日期和必填字段。

仓库中的名称、地点和标识符均为虚构。这个示例展示的是工作流，不包含任何机构的数据或私有配置。

## 快速开始

```bash
python -m pip install -e .
excel-ops examples/extracted-records.json output.xlsx
# 同一命令也接受 .xlsx、.xlsm 和 .csv 输入。
```

如果已知日期顺序或小数分隔符，可传入 `--locale en-US`（或其他明确的地区提示）。交付文件包含 `Type Inference` 工作表，JSON 结果提供字段级置信度和歧义计数。数字形状的标识符和前导零会保留为文本；歧义日期进入复核，不会被静默转换。

对话约定见 [Agent 工作流](docs/agent-workflow.zh-CN.md)；本地和云端表格的行为见[连接器合同](docs/connectors.zh-CN.md)。

字段级批量确认、本次运行/项目作用域和冲突重确认见[批量歧义确认与 Recipe](docs/ambiguity-recipes.zh-CN.md)。

在追加、合并或写回前比较工作簿结构，见 [Schema Drift 检查](docs/schema-drift.zh-CN.md)。

## 输入格式

```json
{
  "source": "synthetic-patrol-sheet.png",
  "records": [
    {
      "location": "100 Example Avenue",
      "event_date": "2026-09-08",
      "identifier": "DEMO 123",
      "category": "synthetic example",
      "confidence": 0.97
    }
  ]
}
```

图片转 JSON 的适配器与工作簿写入刻意分开。未来的适配器可以使用本地 OCR 或视觉模型，但必须输出同样可供复核的数据格式。

## 产品原则

1. **不偷偷猜。** 不确定性本身就是需要输出的结果，而非实现细节。
2. **保留来源。** 每条写入的记录都能追溯到输入来源。
3. **提取与接纳分开。** 模型可以提出数据；验证规则决定数据去哪里。
4. **验证交付物。** 写入后检查输出数量和必填字段。
5. **让重复工作可复现。** 规则应写进配置，而非只存在某个人的记忆中。

## 路线图

包含验收标准的详细路线图见 [Roadmap issue #5](https://github.com/Schlaflied/excel-ops/issues/5)：

1. **Phase 1 — 本地 XLSX：**可靠的本地导入、标准化、匹配、复核、模板写回、验证与 recipe。
2. **Phase 2 — 云端连接器：**将同一工作流扩展到同步目录、Google Sheets、Dropbox、飞书和 WPS；之后再根据需求评估 Microsoft Graph。
3. **Phase 3 — 从提示词到分析：**安全的工作簿合并、多 Tab 交付分组、经过验证的汇总与透视表，以及可重复执行的自动化。

## 当前能力

v0.3.0 已发布多来源导入、类型与地区格式推断、Schema Drift 检查、严格匹配和离线人工复核流程。`main` 还包含业务周期解析、周期感知日期刷新和安全文件命名；这三项尚未进入更新的 Release。

[查看完整能力清单、实现状态、安全边界以及对应的 Issue/PR](docs/capabilities.zh-CN.md)。

这些基础能力尚未完成 Phase 1 的完整交付闭环。下一项垂直切片工作是安全写回现有模板（[#3](https://github.com/Schlaflied/excel-ops/issues/3)），随后是独立验证交付文件（[#4](https://github.com/Schlaflied/excel-ops/issues/4)）。

## 许可证

MIT

# Excel-Ops

中文 | [English](README.md)

### 从图片和杂乱表格，到一份经过检查、可以复跑的数据交付

Excel-Ops 是一套给 AI Agent 使用的对话式工作流，不是另一个表格网页应用。用户只需要把文件交给 Agent，并用大白话说明想要的结果；Agent 负责制定操作计划、调用本地或云端连接器、隔离不确定项、验证结果并打包交付。

Excel-Ops 不是一个“帮你写公式”的 AI。它把表格工作视为一条完整的交付流程：

**导入 → 提取 → 标准化 → 匹配 → 人工复核 → 写入 → 验证 → 交付**

它与 [PPT-Ops](https://github.com/luochen211/ppt-ops) 是一对互补项目：PPT-Ops 把资料变成可以审查和交付的演示文稿；Excel-Ops 把图片和混乱的表格输入变成可以审查和交付的数据文件。

## 在 Agent 中使用

在 Codex 或其他兼容 Agent Skills 标准的 Agent 中调用：

> 使用 `$excel-agent`，把这些截图中的记录提取出来；只有高置信度记录可以写入工作簿，其余内容放进人工复核表。

`$excel-agent` 是唯一面向用户的入口。Python CLI 仍然保留，但它只是 Agent 和自动化测试调用的内部确定性能力，不要求用户自己选择 parser、connector 或命令。

连接器优先级暂定为：本地 XLSX/CSV、本地云同步目录、Google Sheets、Dropbox API、飞书表格、WPS 云表格。云端凭据只存在于用户环境中，不得写进项目或交付包。

## 为什么要做 Excel-Ops

用户通常不是真的想学 `VLOOKUP`。他们只是想把两个来源里的同一批记录正确地对应起来，同时确保：

- 数据没有丢失或重复；
- 不确定的匹配没有被系统偷偷决定；
- 每个结果都能追溯到原始来源；
- 下周或下个月可以用同一套规则重新执行；
- 最终 Excel 在交付前已经经过明确的检查。

## 第一版范围

第一个垂直切片处理一种常见企业任务：从截图、扫描件或结构化提取结果中获得多行记录，按照明确的数据结构进行验证，再把通过检查的数据写入 Excel。

- 接收由任意 OCR 或视觉模型产生的统一 JSON；
- 为每条记录保留源图片名称和提取置信度；
- 写入前检查必填字段、日期和数据类型；
- 高置信度且完整的记录进入 `Accepted`；
- 低置信度、缺字段或格式异常的记录进入 `Review`；
- `Audit` 工作表记录处理数量与交付信息；
- 模糊匹配永远不能作为不可见的最终决定。

## 合成业务示景

假设一家设施服务团队每周收到多种格式的巡检记录：有些来自不同结构的 Excel，有些来自工作人员用手机拍摄的表格截图。不同客户又要求使用各自的 Excel 模板交付。

Excel-Ops 应当完成：

1. 识别输入表格或图片中的数据布局；
2. 提取记录并统一字段、日期与标识符格式；
3. 使用配置中声明的规则匹配目标文件；
4. 把不确定、冲突或重复的匹配交给人审核；
5. 把确认后的数据写入对应模板；
6. 独立检查记录数量、日期范围、必填字段和输出结构；
7. 只把验证通过的文件放进最终交付目录。

仓库中的名称、地点和编号均为虚构内容，不包含任何公司的客户数据、员工数据、凭据或私有配置。

## 快速开始

```bash
python -m pip install -e .
excel-ops examples/extracted-records.json output.xlsx
```

完整对话与交付规则见[Agent 工作流](docs/agent-workflow.md)，本地和云端表格的统一行为见[连接器合同](docs/connectors.md)。

输入示例：

```json
{
  "source": "synthetic-inspection-sheet.png",
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

图片识别与 Excel 写入被刻意拆成两个模块。未来无论接入本地 OCR 还是视觉模型，都必须输出相同、可审查的数据合同，不能直接对工作簿进行不可追踪的修改。

## 产品原则

1. **不偷偷猜。** 不确定性本身就是需要交付的结果。
2. **保留来源。** 每条写入记录都能找到对应的输入来源。
3. **提取不等于确认。** 模型可以提出数据，验证规则决定数据去哪里。
4. **交付必须检查。** 写完文件后重新读取并核对记录与结构。
5. **重复工作必须可复现。** 业务规则进入配置，而不是只存在某个人的记忆里。

## 路线图

包含验收标准的完整规划见 [Roadmap issue #5](https://github.com/Schlaflied/excel-ops/issues/5)：

1. **Phase 1 — Local XLSX：**先完成可靠的本地导入、标准化、匹配、人工复核、模板写回、验证和 recipe。
2. **Phase 2 — Cloud Connectors：**将同一工作流接到本地同步目录、Google Sheets、Dropbox、飞书、WPS；再按真实需求评估 Microsoft Graph。
3. **Phase 3 — Prompt-to-Analysis：**支持安全的工作簿合并、多 Tab 交付分组、经过验证的汇总/透视表和周期性自动化。

### 当前 Phase 1 实现 PR

| 能力 | Issue | Pull request | 状态 |
|---|---:|---:|---|
| 多来源、布局感知导入 | [#1](https://github.com/Schlaflied/excel-ops/issues/1) | [#29](https://github.com/Schlaflied/excel-ops/pull/29) | Open |
| 类型、单位和地区格式推断 | [#24](https://github.com/Schlaflied/excel-ops/issues/24) | [#28](https://github.com/Schlaflied/excel-ops/pull/28) | Open |
| Schema Drift 检测与映射 | [#10](https://github.com/Schlaflied/excel-ops/issues/10) | [#30](https://github.com/Schlaflied/excel-ops/pull/30) | Open；stacked on #28 |
| 严格匹配与去重 | [#2](https://github.com/Schlaflied/excel-ops/issues/2) | [#27](https://github.com/Schlaflied/excel-ops/pull/27) | Open |
| 离线人工复核闭环 | [#18](https://github.com/Schlaflied/excel-ops/issues/18) | [#31](https://github.com/Schlaflied/excel-ops/pull/31) | Open |

这些仍是待审核分支，不代表功能已经发布。下一段纵向闭环是安全写回现有模板（[#3](https://github.com/Schlaflied/excel-ops/issues/3)），然后独立重新打开并验证交付文件（[#4](https://github.com/Schlaflied/excel-ops/issues/4)）。

## License

MIT

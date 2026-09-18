中文 | [English](delivery-pipeline.md)

# 端到端交付流程

`run_delivery(...)` 是可由 Agent 调用的唯一集成入口，把一次本地表格任务从输入文件跑到经过验证的交付文件。它不重新实现匹配、歧义确认、模板写回或交付验证逻辑，只把这些已分别合并并单测过的模块串起来，并返回一个可 JSON 序列化的结果。

```text
导入（xlsx / csv / 图片提取 JSON）
  -> 数据合同校验          （候选接受 / 拒绝）
  -> 严格匹配与去重        （matching.preallocate_records）
  -> 批量歧义确认          （ambiguity + 项目 Recipe）
  -> 可检查的计划 / dry run（在任何文件被触碰之前）
  -> 写入 staging 模板副本 （template_writer.write_template）
  -> 独立验证              （delivery_verification.verify_and_deliver）
  -> 可选公式检查          （formula_verification.FormulaVerifier）
  -> 只有验证通过才进入交付
```

## 保证

- **原始文件不被修改。** 输入和企业模板只读，所有写入都发生在 staging 副本上。
- **保存不等于交付。** 只有 `verify_and_deliver` 重新打开落盘文件、且全部阻断性检查通过，目标才会被报告为已交付。验证失败时不发布任何交付文件，机器可读报告说明原因。
- **每条输入记录恰好落入一个状态**：`written`、`accepted`（已接受但未写入，例如 dry run）、`review`、`rejected`、`skipped_existing`，数量与输入总数可对账。
- **未解决的问题不会进入 Accepted。** pending / unknown / conflict 状态的歧义、模糊候选、冲突和重复 record ID 全部进入复核，不会被自动写入。
- **不交付写了一半的行。** 如果写入器跳过了已映射单元格（受保护公式、非锚点合并单元格），运行以 `incomplete_write` 失败关闭，而不是发布半行数据。
- **每个交付单元格都可追溯**到源文件、工作表与行号或图片区域，以及稳定 record ID。

## 声明一个目标

```python
from excel_ops import (
    Destination, DeliveryTarget, PeriodExpectation, TemplateMapping, run_delivery,
)

target = DeliveryTarget(
    destination=Destination("北区仓库", ("北区仓库", "North Warehouse")),
    template_path="templates/北区模板.xlsx",
    mapping=TemplateMapping(
        sheet="巡检",
        header_row=4,
        data_start_row=5,
        field_columns={
            "record_id": "A",
            "location": "B",
            "event_date": "C",
            "identifier": "D",
            "category": "E",
            "source": "F",
        },
    ),
    required_fields=("record_id", "location", "event_date", "identifier", "category", "source"),
    period_expectations=(PeriodExpectation("巡检", "B2", "2026-09-07 至 2026-09-13"),),
)

result = run_delivery(
    ["sources/第37周.xlsx", "sources/extraction.json"],
    [target],
    staging_dir="staging",
    delivery_dir="delivery",
)
print(result.delivered, result.counts, result.delivery_paths)
```

只能映射 `excel_ops.delivery.CONTRACT_FIELDS` 中列出的字段：`record_id`、`location`、`event_date`、`identifier`、`category`、`confidence`、`source`、`source_file`、`source_sheet`、`source_row`、`source_region`。映射其他字段会在计划阶段成为阻断项，而不是写出一个空单元格。

## 先做 dry run

`plan_delivery(...)`（等价于 `run_delivery(..., dry_run=True)`）返回同一个结果对象，其中填好计划且不创建任何文件：输入数量、目标模板、字段 mapping、staging 与交付路径、每个目标的预期写入与复核数量、未解决歧义和阻断项。

## 复用人工决定

歧义匹配会被合并成一个确认批次。决定可以保存为项目 Recipe 并在后续运行中复用：

```python
from excel_ops import decide, save_project_recipe

pending = [item for item in plan.confirmation_batch.items if item.status == "pending"]
save_project_recipe(
    [decide(pending[0].ambiguity, "北区仓库", scope="project")],
    "recipes/project-recipe.json",
)
run_delivery(..., recipe_path="recipes/project-recipe.json")
```

## 命令行

```bash
excel-ops deliver delivery-plan.json --dry-run
excel-ops deliver delivery-plan.json --recipe recipes/project-recipe.json --result run.json
```

配置文件声明输入与目标，路径相对于配置文件所在目录：

```json
{
  "inputs": ["sources/第37周.xlsx", "sources/extraction.json"],
  "staging_dir": "staging",
  "delivery_dir": "delivery",
  "targets": [
    {
      "key": "北区仓库",
      "aliases": ["北区仓库", "North Warehouse"],
      "template": "templates/北区模板.xlsx",
      "sheet": "巡检",
      "header_row": 4,
      "data_start_row": 5,
      "field_columns": {
        "record_id": "A",
        "location": "B",
        "event_date": "C",
        "identifier": "D",
        "category": "E",
        "source": "F"
      },
      "required_fields": ["record_id", "location", "event_date", "identifier", "category", "source"],
      "period_expectations": [
        { "sheet": "巡检", "cell": "B2", "expected": "2026-09-07 至 2026-09-13" }
      ]
    }
  ]
}
```

`--dry-run` 成功时始终以状态 0 退出，即使 `delivered` 为 `false`（dry-run 本就不会交付，这不是失败）。非 dry-run 的真实运行仅在出现失败，或运行结束后仍未交付时，才以非零状态退出。

## 失败码

| 代码 | 含义 |
|---|---|
| `unknown_layout` | 无法识别输入表头，未写入任何内容。 |
| `unreadable_input` | 输入文件无法读取或解析。 |
| `unreadable_recipe` | 项目 Recipe 缺少格式标记或已损坏。 |
| `plan_blocked` | 计划存在阻断项，例如映射了数据合同之外的字段、record ID 未映射、模板缺失。 |
| `template_write_failed` | 声明的 mapping 与模板不匹配，例如工作表不存在。 |
| `incomplete_write` | 写入器跳过了已映射单元格；不交付写了一半的行。 |
| `verification_failed` | 重新打开的文件未通过阻断性检查，验证报告列出全部 finding。 |

## 本流程不覆盖的范围

- 跨运行指纹幂等仍属于 [#19](https://github.com/Schlaflied/excel-ops/issues/19)。重跑同一份已确认输入不会向已包含这些 record ID 的目标追加重复记录，但流程尚不能识别已交付记录的值发生变化、交付文件被改名或移动，以及运行中断后的部分状态。
- 来源与验证 Manifest（[#20](https://github.com/Schlaflied/excel-ops/issues/20)）、币种与精度（[#23](https://github.com/Schlaflied/excel-ops/issues/23)）、多格式导出（[#22](https://github.com/Schlaflied/excel-ops/issues/22)）。
- 云端连接器、多 Tab 编排、跨来源事实判断、地区法规计算。
- openpyxl 不做公式计算。声明的公式检查是静态的，运行会显式报告 `recalculation_not_verified` 警告，而不是声称工作簿已重算。

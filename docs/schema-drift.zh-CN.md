中文 | [English](schema-drift.md)

# Schema Drift 检查

在写入任何数据之前，将一个基线工作簿与一个或多个新工作簿进行比较：

```bash
excel-ops schema-drift baseline.xlsx week-1.xlsx week-2.xlsx \
  --key-field "Employee ID" --output drift-report.json
```

JSON 报告会分别列出新增和删除字段、已确认的重命名、顺序与类型变化、必填字段变化、小型枚举变化，以及每个 key 对应行数的变化。可能的字段重命名只会进入 `mapping_suggestions`，并保持 `status: review`；在人工确认前，它会阻止 append 和 join，不会被静默接受。

人工确认候选字段到基线字段的映射后，可以将它保存供后续运行复用：

```bash
excel-ops schema-drift baseline.xlsx incoming.xlsx --output drift-report.json \
  --mapping mappings.json --confirm "Staff Number=Employee ID"
```

只有映射文件中的规则可以被自动接受。即使之前确认的映射仍然适用，新出现或被删除的字段仍会触发复核。

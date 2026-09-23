中文 | [English](multi-format-export.md)

# 多格式交付导出契约

本文定义 [Issue #22](https://github.com/Schlaflied/excel-ops/issues/22) 已实现的多格式交付边界。CSV 与 PDF adapter 只消费已经验证的 XLSX，不重复执行匹配或工作簿业务逻辑。

## 用户请求

交付请求可以选择一种或多种格式：

```yaml
delivery:
  formats: [xlsx, pdf]
  csv:
    mode: one-file-per-sheet
  pdf:
    sheets: all
```

首个垂直切片支持 `xlsx`、`csv`、`pdf`。默认值必须保持未设置；用户没有选择时，Agent 应推荐或询问，不能静默替用户选择格式。

## 能力契约

每个格式都生成独立的输出工件，并返回请求格式、实际格式、来源工作簿、包含的 Sheet、输出路径、内容摘要、已知信息损失、验证状态和结构化 findings。扩展名必须由实际 writer 决定，不能只修改用户提供的后缀。

- **XLSX**：继续保留多 Sheet、公式、样式、图片、条件格式和可编辑结构，并沿用现有写回、读回验证、命名、幂等和 Manifest 契约。
- **CSV**：只表示单张扁平表，不能保留公式、样式、合并单元格、图片、图表或其他 Sheet。多 Sheet 时必须明确指定 Sheet 或选择 `one-file-per-sheet`，不得默认取活动 Sheet。
- **PDF**：是固定版面审阅和归档工件，不是可编辑工作簿。必须明确导出全部 Sheet 或指定 Sheet。渲染后，验证器会解析每一页，检查页数与页面尺寸、文本可读性、工作簿缩放和重复表头配置；单 Sheet PDF 还会确认配置的表头出现在每一页，并检查打印边界文本是否因截断而缺失。任何 finding 都会令工件保持 `unverified`，并写入结构化证据。

## 流程边界

格式选择属于 delivery plan，在写入前解析。格式适配器只消费已经验证的交付结果，不重复实现来源匹配、歧义决策、模板映射或业务计算：

```text
plan → dry run → 用户确认 → 写入并验证 XLSX → 导出 CSV/PDF → 记录各格式验证状态 → Manifest
```

单个格式导出失败时，结果必须分别报告成功、阻断和需要复核的工件，不能把整个交付伪装成成功。

## Manifest

Manifest 为每个导出工件记录格式、来源工作簿、包含的 Sheet、输出摘要、验证状态和信息损失警告；不得写入单元格值、凭据或 token。

## 实现切片

1. 在 delivery plan 中校验并规范化格式选择；
2. 增加 XLSX 直通工件和共享导出结果/Manifest 结构；
3. 增加显式单 Sheet 与逐 Sheet CSV 导出及测试；
4. PDF 导出在能力检查后执行，并记录页面级验证 findings；
5. CLI/MCP 暴露、示例及端到端测试已通过现有、受确认保护的 `run_delivery` 工具接入。

`prepare_delivery` 接受上面的 `delivery` 对象；`plan_delivery` 返回 `export_plan`；用户明确批准后，`run_delivery` 先写入并验证 XLSX，再生成所选格式，并在交付 Manifest 中写入 `exports` 数组。每项包含实际格式、来源工作簿、Sheet 范围、内容摘要、验证状态、能力损失警告和 renderer 摘要。见 [`examples/multi-format-delivery.json`](../examples/multi-format-delivery.json)。

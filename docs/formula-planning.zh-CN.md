中文 | [English](formula-planning.md)

# 公式规划契约

Issue [#25](https://github.com/Schlaflied/excel-ops/issues/25) 首先建立类型化规划边界。Agent 负责解释用户的自然语言业务目标；Python 核心只接受明确的单元格、范围、Sheet、Excel 版本和输出模式决策。这样可以避免把自由文本直接拼进公式语法。

首个切片支持查找、条件汇总和日期偏移计划。Excel 2021 与 Microsoft 365 使用 `XLOOKUP`；Excel 2016 与 2019 使用区分横向/纵向的 `INDEX/MATCH` 兼容方案。由于当前契约不携带旧版数组公式的录入元数据，旧版 Excel 只允许返回单行或单列。`SUMIFS` 与 `EDATE` 适用于全部已声明目标版本。

公式模式返回公式文本，并明确标记仍需独立复算。静态值模式不会让 openpyxl 假装计算公式：调用方必须提供来自独立计算路径的值。

## 安全写入工作簿

`plan_formula_application(...)` 针对明确的 Sheet 和有边界的目标范围生成不含单元格值的 dry run ledger。`apply_formula_plan(...)` 必须显式传入 `confirmed=True`，始终写入新工作簿并保留源文件；已有输出会被拒绝，默认跳过受保护公式和合并区域非锚点，批量填充时转换相对引用，并在发布文件前重新打开暂存副本，逐项核对计划写入。只有显式设置 `overwrite_formulas=True` 才能替换公式。静态值暂时只允许写入单个目标单元格，直到后续契约可以为每一行携带独立计算值。

## 完整性与独立复算

`verify_formula_recalculation(...)` 首先运行现有 #11 静态检查，覆盖断裂引用、缺失 Sheet、无效范围、错误 token、声明区域的填充缺口和循环引用；随后优先在 Windows 使用 Microsoft Excel，不可用时使用 LibreOffice，对暂存副本执行复算，并把计算引擎生成的缓存值与明确提供的独立期望值进行对账，可配置数值容差。只有所有检查通过才发布复算副本。

没有可用计算引擎时返回 `recalculation_not_verified`；未提供期望值时返回 `recalculation_expectations_missing`。两者都保持 `unverified`，不会冒充 `verified`。公式错误或期望值不一致会返回 `failed` 并丢弃暂存输出。

## Delivery 与 Agent 集成

交付目标现在可以声明 `formulas`。每条规则包含序列化的类型化计划、明确的 Sheet 和目标范围、公式模式所需的独立期望值、覆盖策略与计算引擎。`run_delivery(...)` 只在暂存副本上应用规则；任何跳过单元格或未通过独立复算的结果都会阻断交付，之后仍需通过原有的持久化工作簿验证。

规则身份会进入 #19 的 mapping 指纹，因此业务规则、Excel 版本、输出模式、目标范围、覆盖策略或期望范围发生变化时，旧的 no-op 基线都会失效。交付 Manifest 记录业务规则、兼容策略、写入范围、计算引擎和验证状态，但不复制期望单元格值。同一批字段会经由 `excel_ops.prepare_delivery` 进入计划文件，因此现有 MCP 流程可以直接规划并执行公式交付，无需 Agent 另外串联脚本。

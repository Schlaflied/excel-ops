中文 | [English](formula-planning.md)

# 公式规划契约

Issue [#25](https://github.com/Schlaflied/excel-ops/issues/25) 首先建立类型化规划边界。Agent 负责解释用户的自然语言业务目标；Python 核心只接受明确的单元格、范围、Sheet、Excel 版本和输出模式决策。这样可以避免把自由文本直接拼进公式语法。

首个切片支持查找、条件汇总和日期偏移计划。Excel 2021 与 Microsoft 365 使用 `XLOOKUP`；Excel 2016 与 2019 使用区分横向/纵向的 `INDEX/MATCH` 兼容方案。由于当前契约不携带旧版数组公式的录入元数据，旧版 Excel 只允许返回单行或单列。`SUMIFS` 与 `EDATE` 适用于全部已声明目标版本。

公式模式返回公式文本，并明确标记仍需独立复算。静态值模式不会让 openpyxl 假装计算公式：调用方必须提供来自独立计算路径的值。工作簿应用、已有公式保护、静态结果验证，以及与 #11 验证器和 #19 幂等记录的集成，保留为 #25 的后续切片。

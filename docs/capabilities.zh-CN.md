中文 | [English](capabilities.md)

# 当前能力与实现状态

本页记录 Excel Ops **当前确实可以做什么、代码位于哪里、有什么安全边界，以及能力是否已经发布**。未来规划仍以 [Roadmap #5](https://github.com/Schlaflied/excel-ops/issues/5) 为准。

## 状态口径

| 状态 | 含义 |
|---|---|
| 已发布 | 已合入 `main`，并包含在对应 GitHub Release 中。 |
| main 已合并 | 已合入 `main` 且通过仓库测试，但尚未包含在新的 Release 中。 |
| 规划中 | 只有 Issue、接口设想或 Roadmap，不应被描述成现有能力。 |

PR 合并、自动测试、实际文件验证和业务人员批准是不同证据。某个模块“已合并”不代表完整业务闭环已经成立。

## 已发布：v0.3.0

[v0.3.0](https://github.com/Schlaflied/excel-ops/releases/tag/v0.3.0) 发布了 Phase 1 的第一组本地表格基础能力。

### 1. 多来源、布局感知导入

- **能做什么**：读取 `.xlsx`、`.xlsm`、`.csv` 和 provider-neutral 图片提取 JSON；根据内容识别数据 sheet、表头和数据区域。
- **输出与证据**：统一为 `ExtractedRecord`，保留文件、sheet、行号或图片区域来源。
- **安全边界**：未知布局会停止，不生成静默的部分输出；`.xlsm` 可以读取，但不会执行宏。
- **实现**：[Issue #1](https://github.com/Schlaflied/excel-ops/issues/1) / [PR #29](https://github.com/Schlaflied/excel-ops/pull/29)

### 2. 数据类型、单位与地区格式推断

- **能做什么**：识别日期、Excel serial date、数字、百分比、货币和标识符；接受明确的 locale 提示。
- **输出与证据**：提供字段级类型、置信度、单位、locale、歧义数和低置信度计数。
- **安全边界**：员工号、邮编、电话和前导零优先保留为文本；歧义日期进入复核，不静默转换。
- **实现**：[Issue #24](https://github.com/Schlaflied/excel-ops/issues/24) / [PR #28](https://github.com/Schlaflied/excel-ops/pull/28)

### 3. Schema Drift 检测与确认映射

- **能做什么**：比较基线工作簿与候选工作簿，识别字段新增、删除、重排、类型、必填性、枚举和记录粒度变化。
- **输出与证据**：生成机器可读 drift report，并报告对 append、join 和 delivery grouping 的影响。
- **安全边界**：启发式 rename 只能成为复核建议；只有人工确认的映射可以保存并复用。
- **实现**：[Issue #10](https://github.com/Schlaflied/excel-ops/issues/10) / [PR #30](https://github.com/Schlaflied/excel-ops/pull/30) / [详细说明](schema-drift.md)

### 4. 严格匹配、预分配与去重

- **能做什么**：按配置精确匹配、忽略大小写匹配和规范化匹配的顺序，将一条记录预分配给最多一个目标。
- **输出与证据**：返回稳定 record ID、使用的规则、置信信息、候选目标和异常原因。
- **安全边界**：模糊匹配永远只能进入 review；冲突、重复和未匹配记录不会自动写回。
- **实现**：[Issue #2](https://github.com/Schlaflied/excel-ops/issues/2) / [PR #27](https://github.com/Schlaflied/excel-ops/pull/27)

### 5. 离线人工复核包

- **能做什么**：生成普通 Excel 可以打开和填写的 Review Pack；支持 `accept`、`correct`、`reject` 和 `cannot determine`。
- **输出与证据**：复核结果按稳定 record ID 回导，并保留幂等的审核历史。
- **安全边界**：`correct` 必须提供修正值；`cannot determine` 不会进入已接受结果；隐藏 manifest 检测表头、ID、候选值和重复 ID 篡改。
- **实现**：[Issue #18](https://github.com/Schlaflied/excel-ops/issues/18) / [PR #31](https://github.com/Schlaflied/excel-ops/pull/31)

## main 已合并，尚未发布

以下能力已进入 `main`，但晚于 v0.3.0，因此不能写成“v0.3.0 已发布”。

### 6. 确定性业务周期解析

- **能做什么**：解析周、月、季度、自定义日期、财年和财季，并支持明确的 `as_of`、时区、周起始日和财年起始月。
- **输出与证据**：`resolve_period(...)` 返回统一的 `PeriodResult`，供刷新与命名模块复用。
- **安全边界**：“最近一周”等依赖隐含当前时间的歧义表达会报错，不会自行猜测。
- **实现**：[Issue #16](https://github.com/Schlaflied/excel-ops/issues/16) / [PR #33](https://github.com/Schlaflied/excel-ops/pull/33)

### 7. 周期感知的日期刷新

- **能做什么**：按声明式日期槽位更新允许修改的周期日期，支持 dry run、零记录期间更新和写后逐槽验证。
- **输出与证据**：刷新计划与实际写入结果分开，使用 `PeriodResult` 作为日期权威来源。
- **安全边界**：历史日期、公式日期和未声明区域不被自动改写；重复运行保持幂等。
- **实现**：[Issue #9](https://github.com/Schlaflied/excel-ops/issues/9) / [PR #34](https://github.com/Schlaflied/excel-ops/pull/34)

### 8. 安全、周期感知的输出文件名

- **能做什么**：从业务周期生成如 `payroll-2026-09-12.xlsx` 的安全文件名，处理扩展名、非法字符、目标目录和重名版本。
- **输出与证据**：同内容目标可以 no-op；不同内容使用 `-rN` 后缀，并在落盘后核验实际路径。
- **安全边界**：不读取系统当天日期；命名日期必须来自上游已经解析的业务周期。
- **实现**：[Issue #17](https://github.com/Schlaflied/excel-ops/issues/17) / [PR #35](https://github.com/Schlaflied/excel-ops/pull/35)

### 9. 批量歧义确认与项目 Recipe

- **能做什么**：把重复出现的字段级问题合并为一个确认批次，并按 `this-run` 或 `project` 作用域复用已保存的决定。
- **输出与证据**：每个条目携带状态、选定值、决定来源和影响数量；项目级决定持久化为带格式版本的 Recipe 文件。
- **安全边界**：`unknown` 和冲突的决定会阻断交付，而不是被套用；候选发生变化时，已保存的决定失效而不会被静默复用。
- **实现**：[Issue #26](https://github.com/Schlaflied/excel-ops/issues/26) / [PR #41](https://github.com/Schlaflied/excel-ops/pull/41) / [详细指南](ambiguity-recipes.zh-CN.md)

### 10. 安全写回现有模板

- **能做什么**：复制企业模板，只修改声明的 `TemplateMapping` 授权的单元格。
- **输出与证据**：返回实际写入路径，并生成记录每个原值/新值与每个跳过项及原因的 JSON 变更日志。
- **安全边界**：源模板永不被修改，已存在的输出永不被覆盖，公式和非锚点合并单元格被跳过而不是替换；格式来自声明的样式行。
- **实现**：[Issue #3](https://github.com/Schlaflied/excel-ops/issues/3) / [PR #40](https://github.com/Schlaflied/excel-ops/pull/40)

### 11. 独立验证交付文件

- **能做什么**：重新打开写入器的实际输出、再打开即将发布的副本，核对工作表、表头、必填字段、record ID、写入值和声明的周期槽位。
- **输出与证据**：生成机器可读验证报告，每个问题都带 finding 代码、说明、建议和位置。
- **安全边界**：失败关闭——未知严重级别一律阻断交付，staging 与交付路径必须不同，验证失败时不发布任何文件。
- **实现**：[Issue #4](https://github.com/Schlaflied/excel-ops/issues/4) / [PR #42](https://github.com/Schlaflied/excel-ops/pull/42)

### 12. 静态公式完整性检查

- **能做什么**：扫描公式文本中的错误值、外部引用、断裂引用、缺失工作表、非法区域、不支持的动态数组和循环引用；校验声明的公式区域并独立对账声明的汇总值。
- **输出与证据**：返回可直接作为交付验证器使用的 findings。
- **安全边界**：openpyxl 不做公式计算，因此重算能力以 `recalculation_not_verified` 警告显式说明，缓存值不被当作证据。
- **实现**：[Issue #11](https://github.com/Schlaflied/excel-ops/issues/11) / [PR #43](https://github.com/Schlaflied/excel-ops/pull/43)

### 13. 集成端到端交付运行

- **能做什么**：`run_delivery(...)` 把导入、数据合同、严格匹配、歧义确认与 Recipe、可检查计划/dry run、staging 模板写回、独立验证和可选公式检查串成一次可由 Agent 调用的运行，并通过 `excel-ops deliver` 暴露。
- **输出与证据**：一个可 JSON 序列化的结果，包含每条记录的终态、稳定 record ID、每个交付单元格的来源、各目标的实际交付路径、验证 findings 和失败码；由包含两种输入布局、一个图片提取 JSON 和两个模板的合成 fixture 端到端覆盖。
- **安全边界**：保存永不被当作交付——只有重新打开并通过验证的目标才会发布；未解决的歧义、冲突和重复记录不会进入 Accepted；被跳过的已映射单元格会让运行失败关闭，而不是交付半行数据；输入与模板保持不变。
- **实现**：[Issue #46](https://github.com/Schlaflied/excel-ops/issues/46) / [详细指南](delivery-pipeline.zh-CN.md)

## 当前可以组合到什么程度

当前模块已经覆盖 Phase 1 本地闭环的全部环节：

```text
ingest → normalize/type inference → schema check → strict match → ambiguity/Recipe
       → plan/dry run → staged template write → independent verification → delivery
                                  period resolution → date refresh → safe naming
```

该闭环已用合成 fixture 端到端跑通，并以重新打开的交付文件为判据。但它仍未进入任何已发布 Release，且只覆盖已声明的本地工作簿模板。

## 尚未实现或尚未形成完整闭环

- [Issue #19](https://github.com/Schlaflied/excel-ops/issues/19)：完整跨运行幂等。集成运行不会向目标追加它已包含的记录，但尚不能对已交付记录的值变化、被改名或移动的交付文件以及中断的运行做指纹识别；
- [Issue #20](https://github.com/Schlaflied/excel-ops/issues/20) 来源与验证 Manifest、[Issue #23](https://github.com/Schlaflied/excel-ops/issues/23) 币种与精度、[Issue #22](https://github.com/Schlaflied/excel-ops/issues/22) 多格式导出；
- 公式重算证据，需要 Excel 或 LibreOffice，openpyxl 无法提供；
- 本地云同步目录、Google Sheets、Dropbox、飞书、WPS 等连接器；
- Prompt 驱动的 append、join、多 Tab delivery grouping、汇总和透视表；
- 跨来源事实判断与地区法规计算。

## 维护规则

- 合并功能 PR 时更新对应条目，但状态先写“main 已合并”。
- 只有发布 GitHub Release 后，才把能力移动到“已发布”并注明版本。
- 每项能力必须同时写清用途、边界和实现链接，不能只罗列模块名。
- Roadmap 的未来计划留在 Issue #5；本页不复制不断变化的完整待办列表。

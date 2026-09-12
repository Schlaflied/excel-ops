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

## 当前可以组合到什么程度

当前模块已经覆盖：

```text
ingest → normalize/type inference → schema check → strict match → human review
                                  period resolution → date refresh → safe naming
```

这些能力有明确的数据结构和测试，但完整的 `write → verify → deliver` 纵向闭环仍未完成。因此，当前不能声称 Excel Ops 已经可以把任意企业模板端到端无人值守交付。

## 尚未实现或尚未形成完整闭环

- [Issue #3](https://github.com/Schlaflied/excel-ops/issues/3)：安全写回现有企业工作簿模板；
- [Issue #4](https://github.com/Schlaflied/excel-ops/issues/4)：重新打开并独立验证实际交付文件；
- 公式完整性、逐项对账、来源 manifest 和完整 recipe；
- 本地云同步目录、Google Sheets、Dropbox、飞书、WPS 等连接器；
- Prompt 驱动的 append、join、多 Tab delivery grouping、汇总和透视表；
- 完整合成数据端到端 fixture：`ingest → normalize → match → review → write → verify → deliver`。

## 维护规则

- 合并功能 PR 时更新对应条目，但状态先写“main 已合并”。
- 只有发布 GitHub Release 后，才把能力移动到“已发布”并注明版本。
- 每项能力必须同时写清用途、边界和实现链接，不能只罗列模块名。
- Roadmap 的未来计划留在 Issue #5；本页不复制不断变化的完整待办列表。

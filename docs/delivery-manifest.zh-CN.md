中文 | [English](delivery-manifest.md)

# 交付 Manifest：一次交付的证据

一次周期性交付跑完之后，自然会问一个问题：*这次运行到底交付了什么？*
`excel_ops.delivery_manifest` 用每个交付工作簿各一份机器可读文件和一份人类可读文件来回答这个问题，两者完全由已完成的 `run_delivery(...)` 调用，以及它真正写到磁盘上的文件生成。

```text
run_delivery(...) 结束，工作簿已真实落盘
  -> 为每个已交付目标各构建一个 DeliveryManifest
       sources        每个输入：内容哈希 + 记录数 + 状态分布
       tabs           每个工作表：run 认为写入的数量 vs. 重新打开文件数出的真实行数
       template       实际使用模板的内容哈希
       recipe         实际生效的 Recipe 决定的内容摘要版本号
       verification   #4 的验证结论，压缩为代码/严重度/计数
       output         交付文件名 + 内容哈希
  -> 写出 <output>.manifest.json（机器可读）
  -> 写出 <output>.manifest.txt （人类可读，同一份事实的文字版）
```

Manifest **不计算任何新结论**。它报告的每一个数字，要么直接来自刚结束的
`DeliveryRun`，要么是写入之后从落盘工作簿里重新读出来的——从不凭空假设，也不单纯沿用内存里的计数器。

## 为什么行数是"读回来的"，而不是直接信任

run 本身已经在追踪它认为每个目标写入了多少条记录（`written`）。Manifest 会额外**重新打开交付文件**，按照映射的 record-ID 列数出真实行数。当两者不一致时——交付后被人手改过、行丢了，或别的原因——`reconciled` 为
`False`，不一致会被列在 `discrepancies` 里（例如
`row_count_mismatch:North` 或 `written_count_mismatch`），而不是被悄悄抹平。这正是 Manifest 是"可证伪的证据"而不是"重复的日志行"的关键。

## Manifest 里有什么

| 字段 | 含义 |
|---|---|
| `output`、`output_name`、`output_hash` | 交付文件的路径、文件名与内容哈希——在写入完成、且验证已发布该文件*之后*才计算哈希。 |
| `destination_key` | 这份 Manifest 对应哪个声明的目标。 |
| `template`、`template_version` | 实际使用的模板文件及其内容哈希。 |
| `recipe_path`、`recipe_version` | 使用的项目 Recipe（如有）及实际生效决定的内容摘要版本——参见[歧义 Recipe](ambiguity-recipes.zh-CN.md)。用新的时间戳或审阅人重新保存同一个答案不会改变版本号，只有决定本身变化才会。 |
| `sources[]` | 每个有贡献的输入文件各一条：文件名、路径、内容哈希、从该文件摄取的记录数、写入*本次*输出的记录数，以及运行/状态分布。 |
| `tabs[]`（JSON 中按工作表名为键） | 每个工作表：`written`（run 自己的计数）vs. `rows`（从真实文件重新数出），以及是哪些来源为它提供了数据。 |
| `accepted`、`review`、`rejected`、`written`、`rows` | run 的终态计数，收窄到对单个输出有意义的部分——见下文。 |
| `run_counts`、`target_counts` | 上面两个汇总字段所依据的完整明细。 |
| `verification` | `status`（`passed`/`failed`）、`findings`（数量）、`codes`（仅 finding 代码）、`severities`（按严重度计数），以及完整验证报告的路径。 |
| `period_start`、`period_end`、`period_display_text` | 传给 `run_delivery(..., period=...)` 的已解析报告周期；未声明时为 `null`——从不猜测。 |
| `run_fingerprint` | 启用了幂等时的 #19 整次运行指纹，仅作交叉引用。 |
| `reconciled`、`discrepancies` | 上述数字是否全部有文件为证；不一致时列出具体代码。 |

`accepted` 与 `rejected` 是运行级别的：被拒绝的记录从未进入匹配阶段，所以不属于任何一个交付目标。`review` 按*候选资格*归属——被搁置复核的记录会把自己仍可能落入的每个目标都计入，但并不代表已经交付。`written` 与逐 tab 的 `rows` 是唯一既属于单个输出、又与真实文件重新核对过的计数。

## Manifest 里绝不会出现什么

没有任何单元格值、以及由单元格值派生出的任何内容会进入 Manifest：

- 不含地点、标识符、类别或任何其他字段值。
- 不含 record ID——它是从源值派生出来的，出于同样的理由被排除。
- 不含验证 finding 的**消息文本**。一条 finding 消息可能会引用失败的单元格本身（例如 `written_value_mismatch: expected 'X', got 'Y'`）；只有 finding 的**代码**与**严重度**会被带入 Manifest。完整消息仍留在 `delivery_verification` 已经写出的验证报告里，由 `report_path` 引用。

留下来的只有：文件名、文件内容哈希、工作表名、声明的字段名，以及整数计数。

## 使用方式

Manifest 自动生成，无需额外开启。

```python
from excel_ops import run_delivery

result = run_delivery(inputs, targets, staging_dir="staging", delivery_dir="delivery")

for manifest in result.manifests:
    print(manifest.output_name, manifest.reconciled, manifest.verification.status)
```

默认情况下 `run_delivery(...)` 还会在每个交付文件旁边写出
`<output>.manifest.json` 与 `<output>.manifest.txt`。传入
`write_manifest=False`（命令行为 `excel-ops deliver --no-manifest`）可以只把 Manifest 保留在 `result.manifests` 中——证据仍然会被计算和返回，只是不落盘为同名文件。

```bash
excel-ops deliver delivery-plan.json --no-manifest
```

没有交付任何文件的运行——dry run、no-op、验证失败的目标——对应目标不会生成 Manifest：Manifest 描述的是真实落盘的文件，从不描述一个计划或一次失败的尝试。

读回一份 Manifest：

```python
from excel_ops.delivery_manifest import load_delivery_manifest, format_manifest

payload = load_delivery_manifest("delivery/North-Warehouse.xlsx.manifest.json")
```

`load_delivery_manifest` 会拒绝任何未标记为
`excel-ops-delivery-manifest-v1` 的内容，这与 `idempotency.py` 和 Recipe 加载器对各自格式采用的做法一致。

## 与 #4 验证、#19 幂等指纹不是一回事

这三个系统在流程里紧挨在一起，很容易混为一谈。它们回答的问题不同，谁也不能替代谁：

| | 回答的问题 | 存放位置 | 以什么为键 |
|---|---|---|---|
| **#4 验证**（`delivery_verification`） | "重新打开的文件是否真的和应该写入的内容一致？"在文件被算作已交付*之前*运行——检查失败会完全阻断交付。 | `<output>.verification.json` | 这次交付本身。 |
| **#19 幂等**（`idempotency.py`） | "这个完全相同的任务是否已经成功完成过，本次运行可以跳过？"在任何摄取、匹配或写入发生*之前*运行。 | `.excel-ops/idempotency.json` | 一个不透明的任务指纹摘要——不含路径、名称或单元格值。 |
| **#20 交付 Manifest**（本模块） | "这次已完成的交付实际包含了什么？"在文件交付*之后*构建，作为对它的记录。 | `<output>.manifest.json` / `.manifest.txt` | 它所紧挨的交付文件。 |

Manifest 的 `verification` 区块是在*转述* #4 的结论——不会重新验证任何东西。它的 `run_fingerprint` 字段是启用幂等时对 #19 记录的交叉引用——Manifest 从不读写 `.excel-ops/idempotency.json`
本身，两份文件也不可互换：指纹文件是一个不可读的组件哈希摘要，Manifest 是关于一次交付的、人类与机器都可读的证据。丢失或删除一份 Manifest 不会影响后续运行是否被判定为 no-op，反过来也一样。

`delivery_manifest.py` 之所以这样命名，是为了避免第四种混淆：
`scripts/update-refresh-manifest.mjs` 追踪的、与之无关的
`refresh-manifest.json`（见[更新流程](refresh.zh-CN.md)）是仓库工具层面对*受追踪源文件*的清单，与表格交付毫无关系。

## 未实现的部分

- **不是跨运行的审计轨迹。** 一份 Manifest 只描述一次交付。把多份 Manifest 汇总成项目历史审计包是后续的审计包路线图项（issue #12），不属于本项。
- **不是第二次验证。** `verification` 只是转述 #4 的结果；从不独立重新检查文件。
- **没有签名。** Manifest 是证据，不是防篡改的签名证明——没有任何机制阻止旁边的
  `.manifest.json` 本身事后被编辑。它防的是*漂移*：重新打开真实交付文件，发现它与 run 自己的计数器不一致。

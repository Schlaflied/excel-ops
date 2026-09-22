中文 | [English](idempotency.md)

# 幂等执行：运行指纹与 no-op 短路

周期性任务会被重复运行：可能是排程又触发了一次，可能是有人不确定上次跑完没有，也可能是同一份文件被重新下载成了另一个名字。`excel_ops.idempotency` 在任何工作开始之前先回答一个问题：*这个任务是不是已经成功完成过了？*

```text
计算运行指纹                （只用内容 hash，不用文件名）
  -> 读取持久化的运行记录     (.excel-ops/idempotency.json)
  -> 连接器 revision 冲突检查 （扩展点）
  -> 判定：no_op | changed | retry
       no_op   -> 立即返回，什么都没有重做
       changed -> 跑完整流程
       retry   -> 跑完整流程（上一次没有成功收尾）
```

这与 `run_delivery(...)` 里已有的**按记录去重**不是同一个机制。后者用稳定 record ID 与目标工作簿中已有的 record ID 比较，保证同一行不会被追加两次——但它要在完成摄取、匹配和计划之后才知道这件事。指纹机制位于它**之上**，直接拦住整次运行。两者同时生效，互不替代。

## 指纹范围

| 组件 | 覆盖内容 |
|---|---|
| `inputs` | 每个输入文件**内容**的 SHA-256 |
| `templates` | 每个模板的内容 hash，加上声明的 Template Profile 版本标记 |
| `mapping` | 声明的工作表、列映射、必填字段、record ID 字段与别名，即“mapping 版本” |
| `recipe` | 可复用的项目 Recipe 决定，直接取 `ambiguity.py` 已解析出的 `RecipeDecision`，不另行解析文件 |
| `confirmations` | 本次运行实际应用的人工确认记录 |
| `period` | 声明的报告周期，包含工作簿内的周期期望 |
| `connector` | 输出目标的连接器身份 |
| `output` | 目标位置上已存在内容的 hash |
| `remote_revision` | 目标的远端 revision，用于冲突检查 |

每个组件各自存为 SHA-256 摘要，指纹则是这组摘要的摘要，因此判定结果能说出**是哪个组件**变了（`changed_components`），而不需要存下关于它的任何内容。

## 保证

- **看内容，不看名字。** 输入与模板按内容 hash 识别。被改名、被移动或被重新下载但字节相同的文件产生相同指纹，不会重复处理；名字没变而内容被编辑过的文件则会触发新执行。文件名、路径、大小和修改时间从不进入指纹。
- **列出顺序不构成身份。** 每个文件的摘要会排序，因此同一批文件按不同顺序声明仍是同一个任务。
- **失败的运行永远不会被当成完成基线。** 只有显式记为 `succeeded` 的记录才能支撑 `no_op`。失败、被计划阻断、验证未通过，或在写入与收尾之间被中断的运行会被记为 `failed` / `started`，之后同样指纹的运行返回 `retry` 并继续执行。`retry` 与 `changed` 是两个不同判定，报告里理由不会含混。
- **新的人工决定一定重新执行。** 修改、新增或撤回一个已确认的复核决定都会改变 `confirmations` 组件。重新保存**同一个**答案不会：决定的 `source` 与 `decided_at` 属于来源信息而非规则内容，已被排除。
- **同一周期在更晚一天解析出来仍是同一周期。** `PeriodResult` 还带 `as_of_date`，它每天都在动；把它算进去会让周报任务仅仅因为时钟就失去幂等。指纹只取解析出的窗口与展示文本。
- **改动已交付文件会重新执行。** 输出内容 hash 在范围内，因此事后被手工改过或被删除的交付件算作变化。
- **不持久化任何原始内容。** 每个组件都是摘要。运行记录里没有路径、没有目标名称、没有单元格值、没有云端标识符、没有凭据。自由格式的 `detail` 同样被清洗：嵌套结构降级为类型名，键名看起来像机密的（`token`、`password`、`authorization` 等）值替换为摘要。凭据仍然可以通过 `ConnectorTarget.credential_material` **成为任务身份的一部分**——它被 hash 进连接器组件，从不写出。
- **损坏或格式不符的运行状态文件永远不会授权 no-op。** 加载器拒绝一切非 `excel-ops-run-record-v1` 的内容；`run_delivery(...)` 随后以 `unreadable_run_state` 为理由继续执行并重写记录。

## 运行记录存放位置

```text
<delivery_dir>/.excel-ops/idempotency.json
```

这沿用 `refresh.mjs` 的 `.refresh/state.json` 惯例：一个与所描述工作同级的隐藏状态目录，里面一个带版本的小 JSON 文件，以及一个拒绝其他格式的 `format` 标记。文件按任务分键，因此多个周期性任务可以共用一个交付目录而不会互相覆盖记录。任务键默认为输出目标的不透明摘要；传 `task_key=` 可以自己命名。

```json
{
  "format": "excel-ops-run-record-v1",
  "runs": {
    "weekly-north": {
      "fingerprint": "…64 位十六进制…",
      "status": "succeeded",
      "components": {"inputs": "…", "templates": "…", "…": "…"},
      "recorded_at": "2026-09-21T09:30:00+00:00",
      "delivered": true,
      "attempt": 1,
      "detail": {"counts": {"written": 2}, "failure_codes": [], "targets": 1}
    }
  }
}
```

每次运行会写**两次**记录：第一次写在首个写入动作之前，状态为 `started`，好让被中断的运行能被发现并重试；第二次写在收尾时，带最终状态。最终记录保存的是交付**之后**重新计算的指纹，因此它已经包含本次运行产出的输出——否则输出内容 hash 在下一次运行时永远不可能匹配。

## 用法

整次运行的幂等是可选开启的。不传 `idempotency=` 的调用方行为与之前完全一致，按记录去重照旧生效。

```python
from excel_ops import IdempotencyOptions, run_delivery

options = IdempotencyOptions(task_key="weekly-north", template_profile_version="profile-v1")

first = run_delivery(inputs, targets, staging_dir=..., delivery_dir=..., idempotency=options)
second = run_delivery(inputs, targets, staging_dir=..., delivery_dir=..., idempotency=options)

assert second.no_op is True
assert second.run_decision.decision == "no_op"
print(second.run_decision.explain())
# no_op: unchanged_since_successful_run (fingerprint 5af5cf177346)
```

命令行：

```bash
excel-ops deliver delivery-plan.json --run-state --task-key weekly-north
excel-ops deliver delivery-plan.json --run-state runs/state.json --task-key weekly-north
```

no-op 的退出码是 `0`：声明的交付件已经就位，什么都没变，也没有任何失败。`no_op` 作为独立字段上报，而不是报成 `delivered: true`，因为这次运行没有交付任何东西——交付的是上一次。

**dry run 永不短路。** `plan_delivery(...)` 仍然返回完整、可检查的计划，只是同时在 `run_decision` 里报出判定。

## 连接器冲突接口

Issue 的最后一条验收标准是**云端目标 revision 变化会触发冲突检查**。本仓库目前还没有云端连接器——那是 Phase 2——所以这里交付的是云端连接器将来接入的接口，并用 stub 做了测试覆盖。它是一个向前兼容的扩展点，**不是可用的云端集成**。

```python
class RevisionSource(Protocol):
    def current_revision(self) -> str | None: ...
```

`check_connector_conflict(previous_record, connector, revision_source=...)` 比较目标当前 revision 与上次运行记录中的 revision，不同则返回 `ConnectorConflict(code="remote_revision_changed", …)`。`evaluate_run(...)` 会先调用它；冲突永远不可能落到 `no_op`：无论指纹其余部分如何，判定都会变成 `changed` 并附带该冲突。一旦记录过 revision，当前 revision 未知（`None`）也绝不被当作“没有变化”。

Phase 1 的目标是本地文件，`ConnectorTarget.local([output_path])` 把它描述为 `local:<输出路径>`，并归一化为 POSIX 分隔符，使在 Windows 上生成指纹的任务在别处仍然匹配。将来的云端连接器提供自己的稳定目标 id 与自己的 `RevisionSource`，指纹格式无需改变。

## 尚未实现

- **没有云端连接器。** 冲突检查只有接口与基于 stub 的测试。这里的任何代码都不与云端表格通信，也不决定**如何**化解冲突——它只报告冲突存在。
- **没有更细粒度。** 判定针对整次运行。一旦有变化，流程会完整跑一遍，实际追加什么由既有的按记录去重决定；本模块不做“只交付变化的记录”。
- **未与工作目录扫描打通。** `scan_workdir(...)`（#15）负责判断**哪些**文件是本期输入；本模块把输入清单当参数接收，自己从不扫描或分类文件。
- **已记录的成功被信任。** 除内容 hash 之外，本模块不会重新验证之前交付过的文件。被改过又被改回到字节相同的交付件，与从未被碰过的交付件无法区分——这也是正确的。
- **没有锁。** 运行记录是状态，不是互斥量。同一瞬间启动的两次运行可能都看到“没有成功过”；并发控制不在范围内。

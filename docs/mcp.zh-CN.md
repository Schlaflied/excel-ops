中文 | [English](mcp.md)

# Agent MCP 服务

Excel-Ops 为能够发现和调用工具的 Agent 提供本地 stdio MCP 服务。它只是现有 `excel-ops` Python CLI 的协议适配层，不重新实现工作簿解析、匹配、写回、验证、幂等或 Manifest 生成。

## 安装与启动

需要 Node.js 20 或更高版本，以及 Python 3.11 或更高版本。

```bash
python -m pip install -e .
npm ci
node mcp/server.mjs
```

通常应由 MCP Host 启动服务，而不是手动运行。通用本地配置如下：

```json
{
  "mcpServers": {
    "excel-ops": {
      "command": "node",
      "args": ["/absolute/path/to/excel-ops/mcp/server.mjs"],
      "env": {
        "EXCEL_OPS_CLI": "excel-ops"
      }
    }
  }
}
```

桥接层默认执行 `python -m excel_ops.cli`。`EXCEL_OPS_PYTHON` 可以指向其他 Python 可执行文件；`EXCEL_OPS_CLI` 可以直接指向已安装的 `excel-ops` 可执行文件并跳过模块调用。`EXCEL_OPS_TIMEOUT_MS` 可以修改默认十分钟的子进程超时。这些变量都不能包含 shell 参数；服务从不调用 shell。

## 工具

| 工具 | 是否写入 | 用途 |
|---|---:|---|
| `excel_ops.scan_workdir` | 否 | 扫描明确授权的目录，分类文件并报告重复文件或无法确定的新旧版本。 |
| `excel_ops.prepare_delivery` | 仅计划文件 | 把 Agent 明确选择的输入、模板和映射写入 `planPath` 显式指定的计划文件；缺少决定时返回 review 而不猜测。 |
| `excel_ops.plan_delivery` | 否 | 执行 `excel-ops deliver --dry-run`，返回拟议写入、复核项和阻断项。 |
| `excel_ops.run_delivery` | 是 | 执行已获批准的交付，验证 XLSX，生成所选 XLSX/CSV/PDF 工件，并返回 Manifest 证据。 |
| `excel_ops.scan_workbook_health` | 否 | 扫描单个 XLSX/XLSM，返回 findings 与绑定源文件的修复 actions。 |
| `excel_ops.plan_workbook_health_repair` | 否 | 校验选定 action ID 和 baseline，不创建输出工作簿。 |
| `excel_ops.run_workbook_health_repair` | 是 | 应用明确批准的 allowlist，重新打开并复扫候选文件，只发布对账通过的输出。 |

### 工作簿健康流程

先调用 `scan_workbook_health`。它和 CLI 使用同一份带版本的 JSON 合同，
包含 `before`、`actions`、`changes`、`remaining`、`new` 和 `failures`。
审核 action ID 并补齐必要 baseline 后调用 `plan_workbook_health_repair`；
它只是 dry run，不会创建目标工作簿。

只有用户明确批准写入后才能调用 `run_workbook_health_repair`，并传入
`confirmed: true`。MCP 层只校验输入并调用 Python，不复制扫描、修复或
验证逻辑。Python 写入候选副本、重新打开并复扫；只有授权 finding 全部
消失、未授权 finding 保持且没有新增 finding 时才发布输出。

```bash
excel-ops health-scan workbook.xlsx
excel-ops health-repair workbook.xlsx repaired.xlsx --request examples/health-repair-request.json --dry-run
excel-ops health-repair workbook.xlsx repaired.xlsx --request examples/health-repair-request.json --confirm
```

`mode` 区分 `scan`、`dry_run` 与 `verified_repair`。对账失败会返回
`verified: false` 和具名失败项，不发布输出，并以非零 CLI 状态结束。

推荐的 Agent 顺序是 `scan_workdir → prepare_delivery → plan_delivery → 用户明确确认 → run_delivery`。准备工具接收结构化意图而不是自然语言：由 Host Agent 选择输入并提供目标模板、sheet 与字段映射。缺少业务决定时返回 `needs_review`，不会写出可执行计划。所有路径必须留在明确授权的根目录中；默认拒绝覆盖已有计划，替换时必须提供当前 SHA-256 摘要。

目标还可以包含由用户业务目标生成的类型化 `formulas`。公式模式必须提供明确的独立期望值；dry run 会展示规则和目标范围，获批后的正式运行只修改暂存工作簿，使用 Excel 或 LibreOffice 复算，并把公式证据写入交付 Manifest。公式规则同时参与幂等指纹。

`run_delivery` 强制要求 `confirmed: true`。调用方只有在展示 dry-run 计划并获得用户明确授权后才能设置它。MCP annotation 可以帮助 Host 展示风险区别，服务端还会独立验证确认字段。

每个工具都返回 `excel-ops-agent-v1` 信封：

```json
{
  "contractVersion": "excel-ops-agent-v1",
  "operation": "plan_delivery",
  "ok": true,
  "status": "planned",
  "durationMs": 42,
  "result": {}
}
```

CLI 无法启动、超时或被外部信号终止时归类为 `environment_error`；CLI 输出不是合法 JSON 时归类为 `unexpected_error`；CLI 返回合法结果但退出码非零时归类为 `business_logic_blocked`。业务阻断仍保留 CLI 的原始结构化结果，让 Agent 可以直接解释复核或验证失败，无需解析终端文字。

## 边界

- stdout 只承载 MCP JSON-RPC；诊断信息进入 stderr。
- 源文件保持不变；交付沿用直接 CLI 调用时相同的 staging、写后回读和 Manifest 规则。
- `prepare_delivery` 只写显式指定的计划文件，不写工作簿，也不能绕过 dry run。
- 除非另行实现的 connector 明确调用远端平台，否则路径和工作簿都留在本地。
- 本服务不实现 Google Sheets、Dropbox、飞书、WPS 或其他云端连接器。

使用 `npm run test:mcp` 运行协议与桥接测试；也可以用 MCP Inspector 启动 `node mcp/server.mjs`，交互式检查工具。

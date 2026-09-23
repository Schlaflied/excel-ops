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
| `excel_ops.prepare_delivery` | 仅计划文件 | 把 Agent 明确选择的输入、模板和映射写成经过校验且不含凭据的 `delivery-plan.json`；缺少决定时返回 review 而不猜测。 |
| `excel_ops.plan_delivery` | 否 | 执行 `excel-ops deliver --dry-run`，返回拟议写入、复核项和阻断项。 |
| `excel_ops.run_delivery` | 是 | 执行已获批准的交付，重新读取并验证输出，再返回交付证据。 |

推荐的 Agent 顺序是 `scan_workdir → prepare_delivery → plan_delivery → 用户明确确认 → run_delivery`。准备工具接收结构化意图而不是自然语言：由 Host Agent 选择输入并提供目标模板、sheet 与字段映射。缺少业务决定时返回 `needs_review`，不会写出可执行计划。所有路径必须留在明确授权的根目录中；默认拒绝覆盖已有计划，替换时必须提供当前 SHA-256 摘要。

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

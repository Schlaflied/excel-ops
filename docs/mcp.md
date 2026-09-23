[中文](mcp.zh-CN.md) | English

# Agent MCP server

Excel-Ops includes a local stdio MCP server for agents that can discover and call tools. It is a thin protocol adapter over the existing `excel-ops` Python CLI. It does not reimplement workbook parsing, matching, write-back, verification, idempotency, or Manifest generation.

## Install and start

Node.js 20 or later and Python 3.11 or later are required.

```bash
python -m pip install -e .
npm ci
node mcp/server.mjs
```

The server is normally launched by an MCP host rather than by hand. A generic local configuration is:

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

By default the bridge runs `python -m excel_ops.cli`. `EXCEL_OPS_PYTHON` may point to another Python executable, while `EXCEL_OPS_CLI` may point to the installed `excel-ops` executable and bypass module invocation. `EXCEL_OPS_TIMEOUT_MS` changes the default ten-minute child-process timeout. These variables may not contain shell arguments; the server never invokes a shell.

## Tools

| Tool | Mutation | Purpose |
|---|---:|---|
| `excel_ops.scan_workdir` | No | Classify files under explicitly authorized roots and report duplicates or uncertain versions. |
| `excel_ops.prepare_delivery` | Plan file only | Write the caller-declared plan file from explicit Agent selections at `planPath`; return review items instead of guessing missing mappings. |
| `excel_ops.plan_delivery` | No | Run `excel-ops deliver --dry-run` and return the proposed writes, review items, and blockers. |
| `excel_ops.run_delivery` | Yes | Run an approved delivery, verify XLSX, generate selected XLSX/CSV/PDF artifacts, and return Manifest evidence. |

The intended Agent sequence is `scan_workdir → prepare_delivery → plan_delivery → explicit user approval → run_delivery`. Preparation accepts structured intent rather than natural language: the Host Agent selects inputs and supplies the target template, sheet, and field mapping. Missing business decisions return `needs_review` and no executable plan is written. Paths are confined to explicit authorized roots, existing plans are not overwritten by default, and a replacement requires the current SHA-256 digest.

`run_delivery` requires `confirmed: true`. The caller must set it only after presenting the dry-run plan and obtaining explicit user approval. MCP annotations help a host display the distinction, but the required confirmation field is also validated by the server.

Every tool returns the `excel-ops-agent-v1` envelope:

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

CLI startup, timeout, and external-signal termination failures are `environment_error`; malformed CLI output is `unexpected_error`; and a valid nonzero CLI result is `business_logic_blocked`. The original structured CLI result is retained for business blockers so the agent can explain review or verification failures without parsing terminal text.

## Boundaries

- stdout belongs exclusively to MCP JSON-RPC; diagnostics go to stderr.
- Source files remain unchanged. Delivery follows the same staging, reread verification, and Manifest rules as direct CLI use.
- `prepare_delivery` writes only the declared plan file. It never writes a workbook or bypasses the dry run.
- Paths and workbooks stay local unless a separately implemented connector explicitly uses a remote platform.
- This server does not implement Google Sheets, Dropbox, Feishu, WPS, or any other cloud connector.

Run the protocol and bridge tests with `npm run test:mcp`. The MCP Inspector can also launch `node mcp/server.mjs` for interactive tool inspection.

import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import test from "node:test";

import { runExcelOps, scanArgs } from "../mcp/bridge.mjs";

const examples = fileURLToPath(new URL("../examples", import.meta.url));

test("the MCP bridge invokes the installed Python CLI and parses its JSON", async () => {
  const response = await runExcelOps(
    "scan_workdir",
    scanArgs({ directory: examples, stabilityWindow: 0 }),
  );

  assert.equal(response.ok, true, response.error?.detail ?? response.error?.message);
  assert.equal(response.status, "completed");
  assert.ok(Array.isArray(response.result.files));
  assert.equal(response.result.scope.roots.length, 1);
});

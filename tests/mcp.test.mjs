import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import { fileURLToPath } from "node:url";
import test from "node:test";

import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

import {
  classifyCliResult,
  deliveryArgs,
  runExcelOps,
  scanArgs,
} from "../mcp/bridge.mjs";
import { createExcelOpsServer } from "../mcp/server.mjs";

const serverPath = fileURLToPath(new URL("../mcp/server.mjs", import.meta.url));

function fakeProcess(run) {
  return () => {
    const child = new EventEmitter();
    child.stdout = new PassThrough();
    child.stderr = new PassThrough();
    child.kill = () => true;
    queueMicrotask(() => run(child));
    return child;
  };
}

test("bridge preserves successful, dry-run, and blocked CLI meanings", async () => {
  const delivered = classifyCliResult(
    "run_delivery",
    0,
    JSON.stringify({ delivered: true, dry_run: false, failures: [] }),
    "",
    4,
  );
  assert.equal(delivered.ok, true);
  assert.equal(delivered.status, "delivered");

  const planned = classifyCliResult(
    "plan_delivery",
    0,
    JSON.stringify({ delivered: false, dry_run: true, plan: { writable: true } }),
    "",
    5,
  );
  assert.equal(planned.ok, true);
  assert.equal(planned.status, "planned");

  const blocked = classifyCliResult(
    "run_delivery",
    1,
    JSON.stringify({
      delivered: false,
      dry_run: false,
      failures: [{ code: "verification_failed" }],
    }),
    "verification blocked delivery",
    6,
  );
  assert.equal(blocked.ok, false);
  assert.equal(blocked.status, "blocked");
  assert.equal(blocked.error.category, "business_logic_blocked");
  assert.equal(blocked.result.failures[0].code, "verification_failed");
});

test("bridge turns malformed output and missing executables into structured errors", async () => {
  const malformed = classifyCliResult("scan_workdir", 0, "not-json", "", 1);
  assert.equal(malformed.ok, false);
  assert.equal(malformed.error.code, "invalid_cli_json");

  const invalidArguments = classifyCliResult(
    "run_delivery",
    2,
    "",
    "usage: excel-ops deliver ...",
    1,
  );
  assert.equal(invalidArguments.error.category, "input_error");
  assert.equal(invalidArguments.error.code, "cli_usage_error");

  const missing = await runExcelOps("scan_workdir", ["scan-workdir", "."], {
    command: `missing-excel-ops-${process.pid}`,
    timeoutMs: 5_000,
  });
  assert.equal(missing.ok, false);
  assert.equal(missing.error.category, "environment_error");
  assert.equal(missing.error.code, "cli_unavailable");
});

test("bridge decodes UTF-8 safely when a character crosses stream chunks", async () => {
  const bytes = Buffer.from(JSON.stringify({ delivered: true, note: "北区" }), "utf8");
  const split = bytes.indexOf(Buffer.from("北", "utf8")) + 1;
  const response = await runExcelOps("run_delivery", [], {
    command: "fake",
    timeoutMs: 5_000,
    spawnImpl: fakeProcess((child) => {
      child.stdout.write(bytes.subarray(0, split));
      child.stdout.write(bytes.subarray(split));
      child.stdout.end();
      child.stderr.end();
      child.emit("close", 0, null);
    }),
  });

  assert.equal(response.ok, true);
  assert.equal(response.result.note, "北区");
});

test("bridge reports signal termination as an environment failure", async () => {
  const response = await runExcelOps("run_delivery", [], {
    command: "fake",
    timeoutMs: 5_000,
    spawnImpl: fakeProcess((child) => {
      child.stdout.end(JSON.stringify({ delivered: false }));
      child.stderr.end();
      child.emit("close", null, "SIGTERM");
    }),
  });

  assert.equal(response.ok, false);
  assert.equal(response.status, "failed");
  assert.equal(response.error.category, "environment_error");
  assert.equal(response.error.code, "cli_terminated");
  assert.equal(response.error.detail, "SIGTERM");
});

test("argument builders preserve CLI boundaries without using a shell", () => {
  assert.deepEqual(
    deliveryArgs("plans/weekly.json", {
      dryRun: true,
      recipe: "recipes/client.json",
      runState: "",
      taskKey: "weekly",
      templateProfileVersion: "v3",
      writeManifest: false,
    }),
    [
      "deliver",
      "plans/weekly.json",
      "--dry-run",
      "--recipe",
      "recipes/client.json",
      "--run-state",
      "--task-key",
      "weekly",
      "--template-profile-version",
      "v3",
      "--no-manifest",
    ],
  );

  assert.deepEqual(
    scanArgs({
      directory: "work",
      alsoAllow: ["templates"],
      recursive: false,
      periodStart: "2026-09-01",
      periodEnd: "2026-09-30",
      overrides: ["old.xlsx=prior_delivery:exclude"],
    }),
    [
      "scan-workdir",
      "work",
      "--json",
      "--also-allow",
      "templates",
      "--no-recursive",
      "--period-start",
      "2026-09-01",
      "--period-end",
      "2026-09-30",
      "--override",
      "old.xlsx=prior_delivery:exclude",
    ],
  );
});

test("MCP lists three bounded tools and calls the dry-run path end to end", async (t) => {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const calls = [];
  const server = createExcelOpsServer({
    invoke: async (operation, args) => {
      calls.push({ operation, args });
      return {
        contractVersion: "excel-ops-agent-v1",
        operation,
        ok: true,
        status: operation === "plan_delivery" ? "planned" : "completed",
        durationMs: 1,
        result: { dry_run: operation === "plan_delivery" },
      };
    },
  });
  const client = new Client({ name: "excel-ops-test", version: "1.0.0" });

  t.after(async () => {
    await client.close();
    await server.close();
  });
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const listed = await client.listTools();
  assert.deepEqual(
    listed.tools.map((tool) => tool.name).sort(),
    [
      "excel_ops.plan_delivery",
      "excel_ops.run_delivery",
      "excel_ops.scan_workdir",
    ],
  );
  const writeTool = listed.tools.find((tool) => tool.name === "excel_ops.run_delivery");
  assert.equal(writeTool.annotations.readOnlyHint, false);
  assert.equal(writeTool.annotations.destructiveHint, true);
  assert.ok(writeTool.inputSchema.required.includes("confirmed"));

  const response = await client.callTool({
    name: "excel_ops.plan_delivery",
    arguments: { config: "plans/weekly.json" },
  });
  assert.equal(response.isError, false);
  assert.equal(response.structuredContent.status, "planned");
  assert.deepEqual(calls[0], {
    operation: "plan_delivery",
    args: ["deliver", "plans/weekly.json", "--dry-run"],
  });
});

test("MCP rejects an unconfirmed write before invoking the CLI", async (t) => {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  let calls = 0;
  const server = createExcelOpsServer({
    invoke: async () => {
      calls += 1;
      throw new Error("should not run");
    },
  });
  const client = new Client({ name: "excel-ops-test", version: "1.0.0" });

  t.after(async () => {
    await client.close();
    await server.close();
  });
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const response = await client.callTool({
    name: "excel_ops.run_delivery",
    arguments: { config: "plans/weekly.json" },
  });
  assert.equal(response.isError, true);
  assert.equal(calls, 0);
});

test("the packaged stdio entrypoint keeps stdout clean and lists its tools", async (t) => {
  const client = new Client({ name: "excel-ops-stdio-test", version: "1.0.0" });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [serverPath],
    stderr: "pipe",
  });
  t.after(async () => client.close());

  await client.connect(transport);
  const listed = await client.listTools();
  assert.equal(listed.tools.length, 3);
  assert.ok(listed.tools.some((tool) => tool.name === "excel_ops.plan_delivery"));
});

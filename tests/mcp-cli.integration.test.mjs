import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

import {
  deliveryArgs,
  prepareArgs,
  runExcelOps,
  scanArgs,
} from "../mcp/bridge.mjs";

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

test("the MCP bridge asks Python for missing plan decisions", async () => {
  const response = await runExcelOps("prepare_delivery", prepareArgs(), {
    stdin: JSON.stringify({
      directory: examples,
      planPath: "delivery-plan.json",
      inputs: [],
      targets: [],
    }),
  });

  assert.equal(response.ok, true, response.error?.detail ?? response.error?.message);
  assert.equal(response.status, "needs_review");
  assert.equal(response.result.prepared, false);
  assert.ok(response.result.review.some((item) => item.field === "inputs"));
});

test("the Python preparation CLI returns structured invalid-request errors", async () => {
  const response = await runExcelOps("prepare_delivery", prepareArgs(), {
    stdin: JSON.stringify({
      directory: examples,
      planPath: "delivery-plan.json",
      alsoAllow: null,
      inputs: [],
      targets: [],
    }),
  });

  assert.equal(response.status, "blocked");
  assert.equal(response.error.category, "business_logic_blocked");
  assert.equal(response.result.error.code, "invalid_request");
});

test("structured MCP intent prepares a plan that the dry-run consumes", async (t) => {
  const working = await mkdtemp(path.join(tmpdir(), "excel-ops-prepare-"));
  t.after(async () => rm(working, { recursive: true, force: true }));
  await writeFile(
    path.join(working, "source.csv"),
    "record_id,location,event_date,identifier,category,source\nR-1,North,2026-09-22,I-1,Demo,synthetic\n",
    "utf8",
  );
  const template = path.join(working, "template.xlsx");
  const python = process.env.EXCEL_OPS_PYTHON ?? "python";
  const fixture = spawnSync(
    python,
    [
      "-c",
      "from openpyxl import Workbook; import sys; w=Workbook(); s=w.active; s.title='Data'; s.append(['record_id','location','event_date','identifier','category','source']); w.save(sys.argv[1])",
      template,
    ],
    { encoding: "utf8", windowsHide: true },
  );
  assert.equal(fixture.status, 0, fixture.stderr);

  const scanned = await runExcelOps(
    "scan_workdir",
    scanArgs({ directory: working, stabilityWindow: 0 }),
  );
  assert.equal(scanned.status, "completed", scanned.error?.message);
  assert.ok(scanned.result.files.some((item) => item.path.endsWith("source.csv")));

  const prepared = await runExcelOps("prepare_delivery", prepareArgs(), {
    stdin: JSON.stringify({
      directory: working,
      planPath: "delivery-plan.json",
      inputs: ["source.csv"],
      targets: [
        {
          key: "North",
          template: "template.xlsx",
          sheet: "Data",
          fieldColumns: {
            record_id: "A",
            location: "B",
            event_date: "C",
            identifier: "D",
            category: "E",
            source: "F",
          },
        },
      ],
    }),
  });
  assert.equal(prepared.status, "prepared", prepared.error?.message);

  const planned = await runExcelOps(
    "plan_delivery",
    deliveryArgs(prepared.result.plan_path, { dryRun: true }),
  );
  assert.equal(planned.status, "planned", planned.error?.detail ?? planned.error?.message);
  assert.equal(planned.result.dry_run, true);
  assert.equal(planned.result.failures.length, 0);
});

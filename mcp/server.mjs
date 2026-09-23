#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/server";
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { pathToFileURL } from "node:url";
import * as z from "zod/v4";

import { deliveryArgs, prepareArgs, runExcelOps, scanArgs } from "./bridge.mjs";

const ResultSchema = z.object({
  contractVersion: z.literal("excel-ops-agent-v1"),
  operation: z.enum(["scan_workdir", "prepare_delivery", "plan_delivery", "run_delivery"]),
  ok: z.boolean(),
  status: z.enum([
    "completed",
    "prepared",
    "needs_review",
    "planned",
    "delivered",
    "no_op",
    "blocked",
    "failed",
  ]),
  durationMs: z.number().nonnegative(),
  result: z.record(z.string(), z.unknown()).optional(),
  error: z
    .object({
      category: z.enum([
        "input_error",
        "environment_error",
        "business_logic_blocked",
        "unexpected_error",
      ]),
      code: z.string(),
      message: z.string(),
      detail: z.string().optional(),
    })
    .optional(),
});

const DeliveryOptions = {
  config: z.string().min(1).describe("Path to the Excel-Ops delivery-plan JSON file."),
  recipe: z.string().min(1).optional().describe("Optional project Recipe JSON path."),
  runState: z
    .string()
    .optional()
    .describe("Optional idempotency state path. Use an empty string for the default location."),
  taskKey: z.string().min(1).optional().describe("Stable key for a recurring delivery task."),
  templateProfileVersion: z.string().min(1).optional(),
};

function toolResult(payload) {
  return {
    content: [{ type: "text", text: JSON.stringify(payload) }],
    structuredContent: payload,
    isError: !payload.ok,
  };
}

export function createExcelOpsServer(options = {}) {
  const invoke = options.invoke ?? runExcelOps;
  const server = new McpServer(
    { name: "excel-ops", version: "0.1.0" },
    {
      instructions:
        "Plan first, preserve source files, never guess ambiguous business facts, and only call run_delivery after the user has approved the write.",
    },
  );

  server.registerTool(
    "excel_ops.prepare_delivery",
    {
      title: "Prepare an Excel delivery plan",
      description:
        "Writes a validated, credential-free delivery-plan JSON from explicit Agent selections. It never writes a workbook and returns review items instead of guessing missing business mappings.",
      inputSchema: z.object({
        directory: z.string().min(1),
        alsoAllow: z.array(z.string().min(1)).optional(),
        planPath: z.string().min(1),
        inputs: z.array(z.string().min(1)).default([]),
        stagingDir: z.string().min(1).default("staging"),
        deliveryDir: z.string().min(1).default("delivery"),
        confidenceThreshold: z.number().min(0).max(1).default(0.85),
        delivery: z
          .object({
            formats: z.array(z.enum(["xlsx", "csv", "pdf"])).min(1),
            csv: z
              .object({
                mode: z.enum(["single-sheet", "one-file-per-sheet"]),
                sheet: z.string().min(1).optional(),
              })
              .optional(),
            pdf: z
              .object({ sheets: z.union([z.literal("all"), z.array(z.string().min(1)).min(1)]) })
              .optional(),
          }),
        recipe: z.string().min(1).optional(),
        targets: z
          .array(
            z.object({
              key: z.string().optional(),
              aliases: z.array(z.string()).optional(),
              template: z.string().optional(),
              sheet: z.string().optional(),
              headerRow: z.number().int().positive().optional(),
              dataStartRow: z.number().int().positive().optional(),
              fieldColumns: z.record(z.string(), z.union([z.string(), z.number().int()])).optional(),
              requiredFields: z.array(z.string()).optional(),
              recordIdField: z.string().optional(),
              templateType: z.string().optional(),
              styleSourceRow: z.number().int().positive().optional(),
              maxRows: z.number().int().nonnegative().optional(),
              periodExpectations: z.array(z.record(z.string(), z.unknown())).optional(),
              deliveryName: z.string().optional(),
              formatPolicy: z.record(z.string(), z.unknown()).optional(),
            }),
          )
          .default([]),
        replace: z.boolean().default(false),
        expectedDigest: z.string().regex(/^[0-9a-f]{64}$/).optional(),
      }),
      outputSchema: ResultSchema,
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (input) =>
      toolResult(
        await invoke("prepare_delivery", prepareArgs(), {
          stdin: JSON.stringify(input),
        }),
      ),
  );

  server.registerTool(
    "excel_ops.scan_workdir",
    {
      title: "Scan an Excel working directory",
      description:
        "Read-only scan of explicitly authorized directories. Classifies current inputs, templates, prior deliveries, review returns, duplicates, and uncertain files.",
      inputSchema: z.object({
        directory: z.string().min(1),
        alsoAllow: z.array(z.string().min(1)).optional(),
        recursive: z.boolean().default(true),
        periodStart: z.iso.date().optional(),
        periodEnd: z.iso.date().optional(),
        stabilityWindow: z.number().nonnegative().optional(),
        recipe: z.string().min(1).optional(),
        overrides: z.array(z.string().min(1)).optional(),
      }),
      outputSchema: ResultSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (input) => toolResult(await invoke("scan_workdir", scanArgs(input))),
  );

  server.registerTool(
    "excel_ops.plan_delivery",
    {
      title: "Plan an Excel delivery",
      description:
        "Runs the complete Excel-Ops delivery path in dry-run mode and returns the proposed writes, review items, blockers, and verification plan without creating delivery files.",
      inputSchema: z.object(DeliveryOptions),
      outputSchema: ResultSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (input) =>
      toolResult(
        await invoke(
          "plan_delivery",
          deliveryArgs(input.config, { ...input, dryRun: true }),
        ),
      ),
  );

  server.registerTool(
    "excel_ops.run_delivery",
    {
      title: "Run an approved Excel delivery",
      description:
        "Writes and independently verifies delivery copies, then produces the plan's XLSX/CSV/PDF artifacts and Manifest evidence. Call only after showing the dry-run plan and obtaining explicit user approval. Source files remain unchanged.",
      inputSchema: z.object({
        ...DeliveryOptions,
        confirmed: z
          .literal(true)
          .describe("Must be true only after the user explicitly approves this delivery write."),
        writeManifest: z.boolean().default(true),
      }),
      outputSchema: ResultSchema,
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (input) =>
      toolResult(
        await invoke("run_delivery", deliveryArgs(input.config, input)),
      ),
  );

  return server;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void serveStdio(createExcelOpsServer);
}

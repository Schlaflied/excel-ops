import { spawn } from "node:child_process";

export const CONTRACT_VERSION = "excel-ops-agent-v1";
const DEFAULT_TIMEOUT_MS = 10 * 60 * 1000;
const MAX_OUTPUT_BYTES = 10 * 1024 * 1024;

function configuredTimeout() {
  const value = Number.parseInt(process.env.EXCEL_OPS_TIMEOUT_MS ?? "", 10);
  return Number.isFinite(value) && value > 0 ? value : DEFAULT_TIMEOUT_MS;
}

export function classifyCliResult(operation, exitCode, stdout, stderr, durationMs) {
  let result;
  try {
    result = JSON.parse(stdout.trim());
  } catch {
    const inputError = exitCode === 2;
    return {
      contractVersion: CONTRACT_VERSION,
      operation,
      ok: false,
      status: "failed",
      durationMs,
      error: {
        category: inputError ? "input_error" : "unexpected_error",
        code: inputError ? "cli_usage_error" : "invalid_cli_json",
        message: inputError
          ? "Excel-Ops CLI rejected the supplied arguments."
          : "Excel-Ops CLI did not return one valid JSON document on stdout.",
        detail: stderr.trim() || undefined,
      },
    };
  }

  if (exitCode !== 0) {
    return {
      contractVersion: CONTRACT_VERSION,
      operation,
      ok: false,
      status: "blocked",
      durationMs,
      result,
      error: {
        category: "business_logic_blocked",
        code: `cli_exit_${exitCode}`,
        message: "Excel-Ops completed without a deliverable result.",
        detail: stderr.trim() || undefined,
      },
    };
  }

  let status = "completed";
  if (result?.dry_run === true) status = "planned";
  else if (result?.no_op === true) status = "no_op";
  else if (result?.delivered === true) status = "delivered";

  return {
    contractVersion: CONTRACT_VERSION,
    operation,
    ok: true,
    status,
    durationMs,
    result,
  };
}

export function runExcelOps(operation, args, options = {}) {
  const explicitCli = process.env.EXCEL_OPS_CLI;
  const command =
    options.command ?? explicitCli ?? process.env.EXCEL_OPS_PYTHON ?? "python";
  const defaultPrefix = options.command || explicitCli ? [] : ["-m", "excel_ops.cli"];
  const commandArgs = [...(options.argsPrefix ?? defaultPrefix), ...args];
  const timeoutMs = options.timeoutMs ?? configuredTimeout();
  const spawnProcess = options.spawnImpl ?? spawn;
  const startedAt = performance.now();

  return new Promise((resolve) => {
    let stdout = "";
    let stderr = "";
    let settled = false;
    let timedOut = false;
    let child;
    let timer;

    const finish = (payload) => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      resolve(payload);
    };

    const fail = (code, message, detail) =>
      finish({
        contractVersion: CONTRACT_VERSION,
        operation,
        ok: false,
        status: "failed",
        durationMs: Math.round(performance.now() - startedAt),
        error: {
          category: "environment_error",
          code,
          message,
          detail: detail || undefined,
        },
      });

    const append = (current, chunk) => {
      const next = current + chunk;
      if (Buffer.byteLength(next, "utf8") > MAX_OUTPUT_BYTES) {
        child.kill();
        fail(
          "cli_output_too_large",
          "Excel-Ops CLI exceeded the 10 MiB MCP output limit.",
        );
        return current;
      }
      return next;
    };

    try {
      child = spawnProcess(command, commandArgs, {
        cwd: options.cwd,
        env: options.env ?? process.env,
        shell: false,
        windowsHide: true,
      });
    } catch (error) {
      fail(
        "cli_unavailable",
        `Could not start the Excel-Ops CLI executable: ${command}`,
        error.message,
      );
      return;
    }

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout = append(stdout, chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr = append(stderr, chunk);
    });
    child.on("error", (error) => {
      fail(
        "cli_unavailable",
        `Could not start the Excel-Ops CLI executable: ${command}`,
        error.message,
      );
    });
    child.on("close", (exitCode, signal) => {
      if (timedOut || settled) return;
      if (exitCode === null) {
        fail(
          "cli_terminated",
          `Excel-Ops CLI terminated by ${signal ?? "an external signal"}.`,
          signal ?? undefined,
        );
        return;
      }
      finish(
        classifyCliResult(
          operation,
          exitCode,
          stdout,
          stderr,
          Math.round(performance.now() - startedAt),
        ),
      );
    });

    timer = setTimeout(() => {
      timedOut = true;
      child.kill();
      fail("cli_timeout", `Excel-Ops CLI exceeded the ${timeoutMs} ms timeout.`);
    }, timeoutMs);
  });
}

export function deliveryArgs(config, options = {}) {
  const args = ["deliver", config];
  if (options.dryRun) args.push("--dry-run");
  if (options.recipe) args.push("--recipe", options.recipe);
  if (options.runState !== undefined) {
    args.push("--run-state");
    if (options.runState) args.push(options.runState);
  }
  if (options.taskKey) args.push("--task-key", options.taskKey);
  if (options.templateProfileVersion) {
    args.push("--template-profile-version", options.templateProfileVersion);
  }
  if (options.writeManifest === false) args.push("--no-manifest");
  return args;
}

export function scanArgs(input) {
  const args = ["scan-workdir", input.directory, "--json"];
  for (const directory of input.alsoAllow ?? []) args.push("--also-allow", directory);
  if (input.recursive === false) args.push("--no-recursive");
  if (input.periodStart) args.push("--period-start", input.periodStart);
  if (input.periodEnd) args.push("--period-end", input.periodEnd);
  if (input.stabilityWindow !== undefined) {
    args.push("--stability-window", String(input.stabilityWindow));
  }
  if (input.recipe) args.push("--recipe", input.recipe);
  for (const override of input.overrides ?? []) args.push("--override", override);
  return args;
}

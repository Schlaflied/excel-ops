import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";
import {
  agentCheck, apply, inspect, isSystemPath, rollback, systemHash, validateManifest,
} from "../refresh.mjs";

const hash = (value) => createHash("sha256").update(value).digest("hex");
const response = (manifest) => async () => ({ ok: true, json: async () => manifest });
const makeManifest = (files) => ({
  schema: 1, repository: "Schlaflied/excel-ops", version: "0.3.0", files,
});
const RAW = "https://raw.githubusercontent.com/Schlaflied/excel-ops/main/";

// Serves the published manifest plus the raw bytes of the listed system files.
function remote(manifest, contents = {}, counter = {}) {
  counter.calls = 0;
  const fetcher = async (url) => {
    counter.calls += 1;
    const name = url.slice(RAW.length).split("/").map(decodeURIComponent).join("/");
    if (name === "refresh-manifest.json") return { ok: true, json: async () => manifest };
    if (!(name in contents)) return { ok: false };
    return { ok: true, arrayBuffer: async () => Buffer.from(contents[name]) };
  };
  return fetcher;
}

const exists = async (...parts) => {
  try {
    await readFile(path.join(...parts));
    return true;
  } catch {
    return false;
  }
};
const readState = async (root) =>
  JSON.parse(await readFile(path.join(root, ".refresh", "state.json"), "utf8"));

async function fixture(run) {
  const root = await mkdtemp(path.join(os.tmpdir(), "excel-ops-refresh-"));
  const base = makeManifest({ "README.md": hash("base") });
  try {
    await writeFile(path.join(root, "README.md"), "base");
    await writeFile(path.join(root, "refresh-manifest.json"), JSON.stringify(base));
    await writeFile(path.join(root, "private-recipe.json"), "secret value");
    await run(root, base);
    assert.equal(await readFile(path.join(root, "private-recipe.json"), "utf8"), "secret value");
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

test("explicit system paths reject traversal and user data", () => {
  for (const name of ["../secret", "docs/../secret", "user-data/recipe.json",
    "docs\\secret", "C:/secret", "/secret", "docs//file", "docs/file:stream"]) {
    assert.equal(isSystemPath(name), false, name);
  }
  assert.equal(isSystemPath("docs/refresh.md"), true);
  assert.throws(() => validateManifest(makeManifest({ "user-data/secret": hash("x") })));
});

test("text hashes are stable across checkout line endings", () => {
  assert.equal(systemHash("README.md", Buffer.from("a\r\nb\r\n")),
    systemHash("README.md", Buffer.from("a\nb\n")));
});

test("same version still detects a changed remote system file", async () => {
  await fixture(async (root, base) => {
    const remote = makeManifest({ "README.md": hash("new"), "docs/refresh.md": hash("doc") });
    const result = await inspect({ root, fetcher: response(remote) });
    assert.equal(result.status, "update-available");
    assert.deepEqual(result.systemFiles.updated, ["README.md"]);
    assert.deepEqual(result.systemFiles.added, ["docs/refresh.md"]);
    assert.equal(result.remoteVersion, base.version);
    assert.equal(JSON.stringify(result).includes("secret value"), false);
  });
});

test("matching manifests report up-to-date", async () => {
  await fixture(async (root, base) => {
    const result = await inspect({ root, fetcher: response(base) });
    assert.equal(result.status, "up-to-date");
    assert.deepEqual(result.systemFiles.localModified, []);
  });
});

test("preview reports local conflicts without changing system or user files", async () => {
  await fixture(async (root) => {
    await writeFile(path.join(root, "README.md"), "locally changed");
    const before = await readFile(path.join(root, "README.md"));
    const result = await inspect({
      mode: "preview", root,
      fetcher: response(makeManifest({ "README.md": hash("remote change") })),
    });
    assert.equal(result.status, "local-system-changed");
    assert.deepEqual(result.systemFiles.conflicts, ["README.md"]);
    assert.equal(result.protectedUserPaths.length, 5);
    assert.deepEqual(await readFile(path.join(root, "README.md")), before);
  });
});

test("offline and invalid manifests fail open with distinct statuses", async () => {
  await fixture(async (root, base) => {
    const offline = await inspect({ root, fetcher: async () => { throw new Error("network down"); } });
    assert.equal(offline.status, "offline");
    assert.equal(offline.systemFiles.localModified.length, 0);
    const invalidRemote = await inspect({ root, fetcher: response({ ...base, files: { "../bad": hash("x") } }) });
    assert.equal(invalidRemote.status, "remote-manifest-invalid");
    await writeFile(path.join(root, "refresh-manifest.json"), "bad json");
    const invalidLocal = await inspect({ root, fetcher: response(base) });
    assert.equal(invalidLocal.status, "local-manifest-invalid");
  });
});

test("remote removals and unexpected local files at added paths are visible", async () => {
  await fixture(async (root) => {
    await mkdir(path.join(root, "docs"));
    await writeFile(path.join(root, "docs", "refresh.md"), "local draft");
    const published = makeManifest({ "docs/refresh.md": hash("new") });
    const result = await inspect({ root, fetcher: response(published) });
    assert.deepEqual(result.systemFiles.removed, ["README.md"]);
    assert.deepEqual(result.systemFiles.added, ["docs/refresh.md"]);
    assert.deepEqual(result.systemFiles.conflicts, ["docs/refresh.md"]);
  });
});

const update = makeManifest({ "README.md": hash("new"), "docs/refresh.md": hash("doc") });
const updateContents = { "README.md": "new", "docs/refresh.md": "doc" };

test("apply without --confirm explains the flag and changes nothing", async () => {
  await fixture(async (root, base) => {
    const result = await apply({ root, fetcher: remote(update, updateContents) });
    assert.equal(result.status, "confirmation-required");
    assert.match(result.message, /--confirm/);
    assert.equal(result.runId, null);
    assert.deepEqual(result.applied, { added: [], updated: [], removed: [] });
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    assert.equal(JSON.parse(await readFile(path.join(root, "refresh-manifest.json"), "utf8")).files["README.md"],
      base.files["README.md"]);
    assert.equal(await exists(root, ".refresh", "state.json"), false);
  });
});

test("apply --confirm backs up, writes, and verifies only manifest system files", async () => {
  await fixture(async (root) => {
    const result = await apply({ root, fetcher: remote(update, updateContents), confirm: true });
    assert.equal(result.status, "applied", JSON.stringify(result.diagnostics));
    assert.deepEqual(result.applied, { added: ["docs/refresh.md"], updated: ["README.md"], removed: [] });
    assert.equal(result.integrity.ok, true);
    assert.deepEqual(result.integrity.failures, []);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "new");
    assert.equal(await readFile(path.join(root, "docs", "refresh.md"), "utf8"), "doc");
    const manifest = JSON.parse(await readFile(path.join(root, "refresh-manifest.json"), "utf8"));
    assert.deepEqual(Object.keys(manifest.files).sort(), ["README.md", "docs/refresh.md"]);
    // A recoverable backup of every changed file, plus a state record rollback can consume.
    const backup = path.join(root, ...result.backupPath.split("/"), "files");
    assert.equal(await readFile(path.join(backup, "README.md"), "utf8"), "base");
    assert.equal(await readFile(path.join(backup, "refresh-manifest.json"), "utf8"),
      JSON.stringify(makeManifest({ "README.md": hash("base") })));
    const state = await readState(root);
    assert.equal(state.status, "applied");
    assert.equal(state.runId, result.runId);
    assert.deepEqual(state.entries.map((entry) => entry.action).sort(), ["added", "updated"]);
    assert.equal(JSON.stringify(result).includes("secret value"), false);
  });
});

test("apply refuses a conflicting local system change instead of picking a side", async () => {
  await fixture(async (root) => {
    await writeFile(path.join(root, "README.md"), "local edit");
    const result = await apply({ root, fetcher: remote(update, updateContents), confirm: true });
    assert.equal(result.status, "conflict-blocked");
    assert.deepEqual(result.conflicts, ["README.md"]);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "local edit");
    assert.equal(await exists(root, ".refresh", "state.json"), false);
    assert.equal(await exists(root, "docs", "refresh.md"), false);
  });
});

test("apply aborts before writing when a download fails verification", async () => {
  await fixture(async (root) => {
    const tampered = await apply({
      root, confirm: true,
      fetcher: remote(update, { "README.md": "new", "docs/refresh.md": "tampered" }),
    });
    assert.equal(tampered.status, "download-verification-failed");
    assert.deepEqual(tampered.diagnostics, [{ path: "docs/refresh.md", reason: "hash-mismatch" }]);
    const missing = await apply({ root, confirm: true, fetcher: remote(update, {}) });
    assert.equal(missing.status, "download-failed");
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    assert.equal(await exists(root, ".refresh", "state.json"), false);
  });
});

test("a failed apply leaves a state a rollback can clean up", async () => {
  await fixture(async (root) => {
    // A plain file occupies the parent of an incoming system path, so one write must fail.
    await writeFile(path.join(root, "docs"), "user note");
    const published = makeManifest({
      "AGENTS.md": hash("agents"), "README.md": hash("new"), "docs/refresh.md": hash("doc"),
    });
    const result = await apply({
      root, confirm: true,
      fetcher: remote(published, { "AGENTS.md": "agents", "README.md": "new", "docs/refresh.md": "doc" }),
    });
    assert.equal(result.status, "apply-failed");
    assert.equal(result.rollbackAvailable, true);
    assert.ok(result.diagnostics.length);
    assert.equal((await readState(root)).status, "failed");

    const undone = await rollback({ root });
    assert.equal(undone.status, "rolled-back-partial");
    assert.equal(await exists(root, "AGENTS.md"), false);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    assert.equal(await readFile(path.join(root, "docs"), "utf8"), "user note");
    assert.equal(undone.manifestRestored, true);
    assert.equal(JSON.parse(await readFile(path.join(root, "refresh-manifest.json"), "utf8")).files["README.md"],
      hash("base"));
  });
});

test("rollback undoes only the most recent recorded apply", async () => {
  await fixture(async (root) => {
    await apply({ root, fetcher: remote(update, updateContents), confirm: true });
    const first = await rollback({ root });
    assert.equal(first.status, "rolled-back");
    assert.deepEqual(first.restored, ["README.md"]);
    assert.deepEqual(first.removed, ["docs/refresh.md"]);
    assert.equal(first.integrity.ok, true);
    assert.deepEqual(first.integrity.localModified, []);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    assert.equal(await exists(root, "docs", "refresh.md"), false);

    const again = await rollback({ root });
    assert.equal(again.status, "already-rolled-back");
    assert.deepEqual(again.restored, []);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
  });
});

test("rollback keeps files changed after the apply and needs a recorded run", async () => {
  await fixture(async (root) => {
    const empty = await rollback({ root });
    assert.equal(empty.status, "no-apply-recorded");
    await apply({ root, fetcher: remote(update, updateContents), confirm: true });
    await writeFile(path.join(root, "README.md"), "edited after the update");
    await writeFile(path.join(root, "recipes-added-later.json"), "later user value");
    const result = await rollback({ root });
    assert.equal(result.status, "rolled-back-partial");
    assert.deepEqual(result.skipped, [{ path: "README.md", reason: "changed-after-apply" }]);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "edited after the update");
    assert.equal(await readFile(path.join(root, "recipes-added-later.json"), "utf8"), "later user value");
  });
});

test("apply fetches and trusts exactly one remote manifest", async () => {
  await fixture(async (root) => {
    // A second manifest read would see a different, unstaged file and could write it into the
    // final manifest, or reject and escape the JSON contract entirely.
    let manifestCalls = 0;
    const drifting = async (url) => {
      const name = url.slice(RAW.length).split("/").map(decodeURIComponent).join("/");
      if (name === "refresh-manifest.json") {
        manifestCalls += 1;
        if (manifestCalls > 1) throw new Error("network down");
        return { ok: true, json: async () => update };
      }
      if (!(name in updateContents)) return { ok: false };
      return { ok: true, arrayBuffer: async () => Buffer.from(updateContents[name]) };
    };
    const result = await apply({ root, fetcher: drifting, confirm: true });
    assert.equal(result.status, "applied", JSON.stringify(result));
    assert.equal(manifestCalls, 1);
    const written = JSON.parse(await readFile(path.join(root, "refresh-manifest.json"), "utf8"));
    assert.deepEqual(written.files, update.files);
    // Every path in the written manifest was staged and verified on disk.
    for (const name of Object.keys(written.files)) {
      assert.equal(await exists(root, ...name.split("/")), true, name);
    }
  });
});

test("a tampered state record cannot point a rollback outside root", async () => {
  await fixture(async (root) => {
    const outside = path.join(root, "..", `escape-${path.basename(root)}`);
    await mkdir(path.join(outside, "files"), { recursive: true });
    await writeFile(path.join(outside, "files", "README.md"), "attacker bytes");
    await mkdir(path.join(root, ".refresh"), { recursive: true });
    const tampered = [
      // backupPath escapes while still carrying the ".refresh/backups/" prefix.
      { runId: "run1", backupPath: `.refresh/backups/../../escape-${path.basename(root)}` },
      // runId itself carries the traversal, so a derived backupPath would escape too.
      { runId: `../../escape-${path.basename(root)}`, backupPath: `.refresh/backups/../../escape-${path.basename(root)}` },
      // A prefixed but unrelated backup directory is no longer accepted either.
      { runId: "run1", backupPath: ".refresh/backups/other-run" },
    ];
    try {
      for (const { runId, backupPath } of tampered) {
        await writeFile(path.join(root, ".refresh", "state.json"), JSON.stringify({
          schema: 1, runId, backupPath, manifestBackup: `${backupPath}/files/refresh-manifest.json`,
          entries: [{
            path: "README.md", action: "updated", backup: `${backupPath}/files/README.md`,
            previousHash: hash("attacker bytes"), newHash: hash("new"),
          }],
        }));
        const result = await rollback({ root });
        assert.equal(result.status, "state-invalid", backupPath);
        assert.deepEqual(result.restored, []);
        assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
      }
    } finally {
      await rm(outside, { recursive: true, force: true });
    }
  });
});

test("a partial rollback stays retryable until the skipped path is restored", async () => {
  await fixture(async (root) => {
    const published = makeManifest({ "README.md": hash("new") });
    const applied = await apply({
      root, confirm: true, fetcher: remote(published, { "README.md": "new" }),
    });
    assert.equal(applied.status, "applied");

    // A directory at the target path makes exactly one restore write fail, like a locked file.
    await rm(path.join(root, "README.md"));
    await mkdir(path.join(root, "README.md"));
    const partial = await rollback({ root });
    assert.equal(partial.status, "rolled-back-partial");
    assert.deepEqual(partial.restored, []);
    assert.deepEqual(partial.skipped.map((entry) => [entry.path, entry.reason]),
      [["README.md", "restore-failed"]]);
    assert.equal((await readState(root)).status, "rolled-back-partial");

    // Once the cause is resolved a second rollback retries instead of reporting completion.
    await rm(path.join(root, "README.md"), { recursive: true });
    const retry = await rollback({ root });
    assert.equal(retry.status, "rolled-back", JSON.stringify(retry));
    assert.deepEqual(retry.restored, ["README.md"]);
    assert.deepEqual(retry.skipped, []);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    assert.equal((await readState(root)).status, "rolled-back");
    assert.equal((await rollback({ root })).status, "already-rolled-back");
  });
});

test("an unreadable state record blocks rollback without changing files", async () => {
  await fixture(async (root) => {
    await mkdir(path.join(root, ".refresh"), { recursive: true });
    await writeFile(path.join(root, ".refresh", "state.json"), "{not json");
    const result = await rollback({ root });
    assert.equal(result.status, "state-invalid");
    assert.deepEqual(result.restored, []);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    await writeFile(path.join(root, ".refresh", "state.json"),
      JSON.stringify({ schema: 1, runId: "x", backupPath: "../escape", entries: [] }));
    assert.equal((await rollback({ root })).status, "state-invalid");
  });
});

test("agent-check caches a successful read-only result for 24 hours", async () => {
  await fixture(async (root) => {
    const counter = {};
    const fetcher = remote(update, updateContents, counter);
    const start = Date.parse("2026-09-20T10:00:00Z");
    const first = await agentCheck({ root, fetcher, now: () => start });
    assert.equal(first.status, "update-available");
    assert.equal(first.updateAvailable, true);
    assert.equal(first.cached, false);
    assert.match(first.advice, /apply --confirm/);
    const calls = counter.calls;

    const cached = await agentCheck({ root, fetcher, now: () => start + 60_000 });
    assert.equal(cached.cached, true);
    assert.equal(cached.status, "update-available");
    assert.equal(counter.calls, calls);

    const forced = await agentCheck({ root, fetcher, force: true, now: () => start + 60_000 });
    assert.equal(forced.cached, false);
    assert.ok(counter.calls > calls);

    const expired = await agentCheck({ root, fetcher, now: () => start + 25 * 60 * 60 * 1000 });
    assert.equal(expired.cached, false);

    // Read-only: no system file changed and no update was applied.
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
    assert.equal(await exists(root, "docs", "refresh.md"), false);
    assert.equal(await exists(root, ".refresh", "state.json"), false);
  });
});

test("agent-check fails open on a corrupt cache and when the network is down", async () => {
  await fixture(async (root) => {
    await mkdir(path.join(root, ".refresh"), { recursive: true });
    await writeFile(path.join(root, ".refresh", "agent-check.json"), "corrupt{");
    const recovered = await agentCheck({ root, fetcher: remote(update, updateContents) });
    assert.equal(recovered.status, "update-available");
    assert.equal(recovered.cached, false);

    const offline = await agentCheck({
      root, force: true, fetcher: async () => { throw new Error("network down"); },
    });
    assert.equal(offline.status, "offline");
    assert.equal(offline.updateAvailable, false);
    assert.equal(await readFile(path.join(root, "README.md"), "utf8"), "base");
  });
});

test("apply and rollback leave user data byte-identical", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "excel-ops-refresh-user-"));
  // Every protected user category from the manifest's PROTECTED_USER_PATHS list.
  const userFiles = {
    "templates/company-template.xlsx": Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0xff]),
    "templates/template-profile.json": Buffer.from('{"profile":"payroll"}', "utf8"),
    "recipes/project-recipe.json": Buffer.from('{"mapping":{"名称":"name"}}', "utf8"),
    "inputs/source-workbook.xlsx": Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x01]),
    "review/review-pack.json": Buffer.from('{"review":[1,2]}', "utf8"),
    "delivery/manifest.json": Buffer.from('{"delivered":true}', "utf8"),
    "delivery/2026-09/payroll.xlsx": Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x02]),
    "connectors/credentials.json": Buffer.from('{"token":"do-not-touch"}', "utf8"),
    "my-examples/business-rules.md": Buffer.from("local rule\r\n", "utf8"),
  };
  try {
    await writeFile(path.join(root, "README.md"), "base");
    await writeFile(path.join(root, "refresh-manifest.json"),
      JSON.stringify(makeManifest({ "README.md": hash("base") })));
    for (const [name, bytes] of Object.entries(userFiles)) {
      const filename = path.join(root, ...name.split("/"));
      await mkdir(path.dirname(filename), { recursive: true });
      await writeFile(filename, bytes);
    }
    const applied = await apply({ root, fetcher: remote(update, updateContents), confirm: true });
    assert.equal(applied.status, "applied");
    const undone = await rollback({ root });
    assert.equal(undone.status, "rolled-back");
    for (const [name, bytes] of Object.entries(userFiles)) {
      assert.deepEqual(await readFile(path.join(root, ...name.split("/"))), bytes, name);
    }
    // No user path, credential, or business value reaches the machine-readable output.
    const output = JSON.stringify(applied) + JSON.stringify(undone);
    for (const fragment of ["do-not-touch", "payroll", "credentials.json", "project-recipe.json",
      "company-template.xlsx", "local rule"]) {
      assert.equal(output.includes(fragment), false, fragment);
    }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("system paths and state entries stay platform independent", async () => {
  // A manifest may only use "/" separated relative paths on every platform.
  assert.equal(isSystemPath("docs\\refresh.md"), false);
  assert.equal(isSystemPath(path.join("docs", "refresh.md").split(path.sep).join("/")), true);
  assert.throws(() => validateManifest(makeManifest({ "docs\\refresh.md": hash("x") })));
  await fixture(async (root) => {
    const result = await apply({ root, fetcher: remote(update, updateContents), confirm: true });
    assert.equal(result.status, "applied");
    // Written through path.join, recorded with "/" regardless of the host separator.
    assert.equal(await readFile(path.join(root, "docs", "refresh.md"), "utf8"), "doc");
    const state = await readState(root);
    assert.ok(state.entries.every((entry) => !entry.path.includes("\\")));
    assert.ok(state.entries.every((entry) => !entry.path.includes(path.win32.sep)));
    assert.ok(state.entries.every((entry) => entry.backup === null || entry.backup.includes("/")));
    assert.ok(result.backupPath.startsWith(".refresh/backups/"));
    await rollback({ root });
  });
});

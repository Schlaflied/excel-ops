#!/usr/bin/env node
// Read-only system update check. Workbook processing remains in Python.
import { createHash, randomUUID } from "node:crypto";
import { lstat, mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const MANIFEST = "refresh-manifest.json";
const RAW = "https://raw.githubusercontent.com/Schlaflied/excel-ops/main/";
const REMOTE = RAW + MANIFEST;
const REPOSITORY = "Schlaflied/excel-ops";
// Local run data. Never a system path, so updates never rewrite their own backups.
const WORK_DIR = ".refresh";
const STATE = `${WORK_DIR}/state.json`;
const CACHE = `${WORK_DIR}/agent-check.json`;
const BACKUPS = `${WORK_DIR}/backups`;
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;
const CACHEABLE = new Set(["up-to-date", "update-available", "local-system-changed"]);
const HASH = /^[a-f0-9]{64}$/;
const TEXT_EXTENSIONS = new Set([".md", ".py", ".mjs", ".json", ".toml", ".yaml", ".yml"]);
const EXACT_SYSTEM = new Set([
  ".gitignore", "AGENTS.md", "LICENSE", "README.md", "README.zh-CN.md",
  "pyproject.toml", "refresh.mjs", "issue-refresh-mjs.md",
  "assets/logo.png",
]);
const SYSTEM_PREFIXES = [
  ".agents/skills/excel-agent/", "docs/", "examples/", "scripts/",
  "src/excel_ops/", "tests/",
];
export const PROTECTED_USER_PATHS = Object.freeze([
  "company templates and Template Profiles",
  "project Recipes, review decisions, and confirmed mappings",
  "input workbooks, review packs, Manifests, deliveries, and output directories",
  "cloud configuration, credentials, customer data, and employee data",
  "locally added examples and business rules",
]);

export function systemHash(name, bytes) {
  // Git checkouts may use CRLF on Windows and LF elsewhere.
  const isText = TEXT_EXTENSIONS.has(path.extname(name)) ||
    name === ".gitignore" || name === "LICENSE";
  const content = !isText ? bytes :
    Buffer.from(bytes.toString("utf8").replace(/\r\n/g, "\n"), "utf8");
  return createHash("sha256").update(content).digest("hex");
}

export function isSystemPath(name) {
  if (typeof name !== "string" || !name || name.includes("\\") ||
      name.startsWith("/") || /[\0:<>|?*]/.test(name)) return false;
  const parts = name.split("/");
  if (parts.some((part) => !part || part === "." || part === "..")) return false;
  return EXACT_SYSTEM.has(name) || SYSTEM_PREFIXES.some((prefix) => name.startsWith(prefix));
}

export function validateManifest(value) {
  if (!value || value.schema !== 1 || value.repository !== REPOSITORY ||
      typeof value.version !== "string" || !value.version ||
      !value.files || Array.isArray(value.files) || typeof value.files !== "object") {
    throw new Error("invalid-manifest");
  }
  for (const [name, hash] of Object.entries(value.files)) {
    if (!isSystemPath(name) || !HASH.test(hash)) throw new Error("invalid-manifest");
  }
  return value;
}

// Every filesystem write resolves through here, so a path must be relative and "/" separated.
// A ".refresh/" run path is checked as strictly as a system path: no separator, drive, or "."
// and ".." segment may survive, so no caller can escape root through path.join.
function at(root, name) {
  if (!isSystemPath(name) && !isWorkPath(name)) throw new Error("unsupported-path");
  return path.join(root, ...name.split("/"));
}

function isWorkPath(name) {
  if (typeof name !== "string" || !name.startsWith(`${WORK_DIR}/`) ||
      name.includes("\\") || /[\0:<>|?*]/.test(name)) return false;
  return !name.split("/").some((part) => !part || part === "." || part === "..");
}

// A missing path reports ENOENT on Windows but ENOTDIR on Unix when a parent is a file.
const absent = (error) => error.code === "ENOENT" || error.code === "ENOTDIR";

async function readIfPresent(filename) {
  try {
    const stat = await lstat(filename);
    if (!stat.isFile() || stat.isSymbolicLink()) throw new Error("unsupported-path");
    return await readFile(filename);
  } catch (error) {
    if (absent(error)) return null;
    throw error;
  }
}

async function writeFileAt(filename, bytes) {
  await mkdir(path.dirname(filename), { recursive: true });
  await writeFile(filename, bytes);
}

// A rejected restore destination: carries a detail string so the caller can record it in the
// same skip-entry shape a locked or unwritable path already uses.
class RestoreRejected extends Error {
  constructor(detail) {
    super("restore-rejected");
    this.detail = detail;
  }
}

// Rollback is the only place this tool writes over a path it did not itself just create, so it
// is the only place a planted symlink could redirect a write out of root. Both restore callers
// (entry files and the manifest) go through here; a second write path is exactly how the
// manifest missed this check the first time.
//
// Every component between root and the destination is inspected with lstat, never stat: stat
// follows the very link being looked for. A symlinked parent directory redirects a write just
// as effectively as a symlinked leaf, so both are rejected. Missing parents are created one
// component at a time rather than with a recursive mkdir, which would happily walk through an
// existing symlinked component.
//
// The write itself lands on a fresh temp file in the destination directory (flag "wx", so it
// can never open something already sitting at that name) and is then renamed into place.
// rename() replaces the directory entry and never writes through a link, so a symlink planted
// between the lstat check and the write receives nothing: it is replaced, not followed. That
// closes the TOCTOU gap a check-then-writeFile pair would leave open.
// The destination is passed already resolved, because the two callers resolve it differently:
// an entry path goes through at(), while the manifest is not itself a manifest-listed system
// path and so is joined directly. Containment in root is therefore re-checked here rather than
// assumed, so neither caller can hand this helper a path outside the sandboxed root.
async function restoreFileAt(root, filename, bytes) {
  const relative = path.relative(root, filename);
  if (!relative || path.isAbsolute(relative) ||
      relative.split(path.sep).some((part) => part === "..")) {
    throw new RestoreRejected("outside-root");
  }
  const parts = relative.split(path.sep);
  let walked = root;
  for (const part of parts.slice(0, -1)) {
    walked = path.join(walked, part);
    let stat = null;
    try {
      stat = await lstat(walked);
    } catch (error) {
      if (!absent(error)) throw new RestoreRejected(error.code || "unknown");
    }
    if (stat === null) {
      try {
        await mkdir(walked);
      } catch (error) {
        if (error.code !== "EEXIST") throw new RestoreRejected(error.code || "unknown");
      }
      continue;
    }
    if (stat.isSymbolicLink()) throw new RestoreRejected("symlink-parent-rejected");
    if (!stat.isDirectory()) throw new RestoreRejected("ENOTDIR");
  }

  try {
    const stat = await lstat(filename);
    if (stat.isSymbolicLink()) throw new RestoreRejected("symlink-destination-rejected");
    if (!stat.isFile()) throw new RestoreRejected("not-a-regular-file");
  } catch (error) {
    if (error instanceof RestoreRejected) throw error;
    if (!absent(error)) throw new RestoreRejected(error.code || "unknown");
  }

  const temporary = path.join(
    path.dirname(filename),
    `.${path.basename(filename)}.${randomUUID().slice(0, 8)}.tmp`,
  );
  try {
    await writeFile(temporary, bytes, { flag: "wx" });
    await rename(temporary, filename);
  } catch (error) {
    await rm(temporary, { force: true }).catch(() => {});
    throw new RestoreRejected(error.code || "unknown");
  }
}

async function localChanges(root, manifest) {
  const changes = [];
  for (const [name, expected] of Object.entries(manifest.files)) {
    const filename = path.join(root, ...name.split("/"));
    let actual;
    try {
      const stat = await lstat(filename);
      if (!stat.isFile() || stat.isSymbolicLink()) {
        changes.push(name);
        continue;
      }
      actual = systemHash(name, await readFile(filename));
    } catch (error) {
      if (!absent(error)) throw error;
    }
    if (actual !== expected) changes.push(name);
  }
  return changes.sort();
}

async function remoteManifest(fetcher) {
  const response = await fetcher(REMOTE, {
    headers: { Accept: "application/json" },
    signal: AbortSignal.timeout(4000),
    cache: "no-store",
  });
  if (!response.ok) throw new Error("remote-unavailable");
  let payload;
  try {
    payload = await response.json();
  } catch {
    throw new Error("invalid-manifest");
  }
  return validateManifest(payload);
}

async function remoteFile(fetcher, name) {
  const response = await fetcher(RAW + name.split("/").map(encodeURIComponent).join("/"), {
    signal: AbortSignal.timeout(8000),
    cache: "no-store",
  });
  if (!response.ok) throw new Error("download-failed");
  return Buffer.from(await response.arrayBuffer());
}

function manifestBytes(manifest) {
  const files = {};
  for (const name of Object.keys(manifest.files).sort()) files[name] = manifest.files[name];
  return Buffer.from(JSON.stringify({ ...manifest, files }, null, 2) + "\n", "utf8");
}

// onRemote receives the single validated remote manifest, so apply can reuse the exact object
// this inspect diffed instead of fetching a second, possibly different one.
export async function inspect({ mode = "check", root = ROOT, fetcher = fetch, onRemote } = {}) {
  if (mode !== "check" && mode !== "preview") throw new Error("unsupported-command");
  const result = {
    schema: 1,
    mode,
    status: "offline",
    localVersion: null,
    remoteVersion: null,
    systemFiles: { added: [], updated: [], removed: [], localModified: [], conflicts: [] },
    protectedUserPaths: PROTECTED_USER_PATHS,
  };
  let local;
  try {
    local = validateManifest(JSON.parse(await readFile(path.join(root, MANIFEST), "utf8")));
    result.localVersion = local.version;
    result.systemFiles.localModified = await localChanges(root, local);
  } catch {
    result.status = "local-manifest-invalid";
    return result;
  }

  let remote;
  try {
    remote = await remoteManifest(fetcher);
  } catch (error) {
    result.status = error.message === "invalid-manifest" ? "remote-manifest-invalid" : "offline";
    return result;
  }
  if (onRemote) onRemote(remote);
  result.remoteVersion = remote.version;
  const oldFiles = local.files;
  const newFiles = remote.files;
  result.systemFiles.added = Object.keys(newFiles).filter((name) => !(name in oldFiles)).sort();
  result.systemFiles.updated = Object.keys(newFiles)
    .filter((name) => name in oldFiles && newFiles[name] !== oldFiles[name]).sort();
  result.systemFiles.removed = Object.keys(oldFiles).filter((name) => !(name in newFiles)).sort();
  const incoming = new Set([...result.systemFiles.added, ...result.systemFiles.updated, ...result.systemFiles.removed]);
  result.systemFiles.conflicts = result.systemFiles.localModified.filter((name) => incoming.has(name));
  for (const name of result.systemFiles.added) {
    try {
      await lstat(path.join(root, ...name.split("/")));
      result.systemFiles.conflicts.push(name);
    } catch (error) {
      if (!absent(error)) throw error;
    }
  }
  result.systemFiles.conflicts.sort();
  result.status = result.systemFiles.localModified.length ? "local-system-changed" :
    incoming.size ? "update-available" : "up-to-date";
  return result;
}

async function doctor(root, manifest, written, removed) {
  // Minimal post-write sanity check: no truncated writes, no leftovers, manifest still valid.
  const failures = [];
  for (const [name, expected] of Object.entries(written)) {
    const bytes = await readIfPresent(at(root, name)).catch(() => null);
    if (bytes === null) failures.push({ path: name, reason: "missing-after-write" });
    else if (systemHash(name, bytes) !== expected) failures.push({ path: name, reason: "hash-mismatch" });
  }
  for (const name of removed) {
    if (await readIfPresent(at(root, name)).catch(() => null) !== null) {
      failures.push({ path: name, reason: "still-present-after-removal" });
    }
  }
  let localModified = [];
  try {
    const local = validateManifest(JSON.parse(await readFile(path.join(root, MANIFEST), "utf8")));
    if (local.version !== manifest.version) failures.push({ path: MANIFEST, reason: "version-mismatch" });
    localModified = await localChanges(root, local);
  } catch {
    failures.push({ path: MANIFEST, reason: "invalid-after-write" });
  }
  return { ok: failures.length === 0, failures, localModified };
}

export async function apply({ root = ROOT, fetcher = fetch, confirm = false, now = Date.now } = {}) {
  const result = {
    schema: 1,
    mode: "apply",
    status: "confirmation-required",
    message: "Updates are applied only with the explicit flag: node refresh.mjs apply --confirm. " +
      "Without it nothing is downloaded, written, backed up, or removed.",
    localVersion: null,
    remoteVersion: null,
    runId: null,
    backupPath: null,
    applied: { added: [], updated: [], removed: [] },
    conflicts: [],
    diagnostics: [],
    rollbackAvailable: false,
    protectedUserPaths: PROTECTED_USER_PATHS,
  };
  if (!confirm) return result;

  // One fetch, one validated manifest: the object that decided the conflict and staging lists
  // is the same one that supplies the hashes and is written at the end.
  let remote = null;
  const preview = await inspect({ mode: "preview", root, fetcher, onRemote: (value) => { remote = value; } });
  result.localVersion = preview.localVersion;
  result.remoteVersion = preview.remoteVersion;
  if (["offline", "local-manifest-invalid", "remote-manifest-invalid"].includes(preview.status)) {
    result.status = preview.status;
    result.message = "No update was applied and no file changed.";
    return result;
  }
  const { added, updated, removed, conflicts } = preview.systemFiles;
  if (conflicts.length) {
    result.status = "conflict-blocked";
    result.conflicts = conflicts;
    result.message = "Local system files differ from the recorded manifest at paths this update " +
      "would change. Resolve each listed path yourself; apply never picks a side.";
    return result;
  }
  if (!added.length && !updated.length && !removed.length) {
    result.status = "up-to-date";
    result.message = "No system file changed.";
    return result;
  }

  if (!remote) {
    // Defensive: a non-failing preview always yields a manifest, so this stays inside the
    // documented JSON contract instead of escaping as a rejection to stderr.
    result.status = "remote-manifest-invalid";
    result.message = "No update was applied and no file changed.";
    return result;
  }

  // Stage and verify every incoming byte before touching the working tree.
  const staged = new Map();
  for (const name of [...added, ...updated]) {
    let bytes;
    try {
      bytes = await remoteFile(fetcher, name);
    } catch {
      result.status = "download-failed";
      result.diagnostics.push({ path: name, reason: "download-failed" });
      result.message = "Download failed before any file changed.";
      return result;
    }
    if (systemHash(name, bytes) !== remote.files[name]) {
      result.status = "download-verification-failed";
      result.diagnostics.push({ path: name, reason: "hash-mismatch" });
      result.message = "A downloaded system file did not match the published hash. Nothing was written.";
      return result;
    }
    staged.set(name, bytes);
  }

  const runId = `${new Date(now()).toISOString().replace(/[:.]/g, "-")}-${randomUUID().slice(0, 8)}`;
  const backupPath = `${BACKUPS}/${runId}`;
  result.runId = runId;
  result.backupPath = backupPath;
  const entries = [];
  for (const name of [...updated, ...removed]) {
    const bytes = await readIfPresent(at(root, name));
    const entry = {
      path: name,
      action: removed.includes(name) ? "removed" : "updated",
      backup: null,
      previousHash: null,
      newHash: staged.has(name) ? remote.files[name] : null,
    };
    if (bytes !== null) {
      entry.backup = `${backupPath}/files/${name}`;
      entry.previousHash = systemHash(name, bytes);
      await writeFileAt(at(root, entry.backup), bytes);
    }
    entries.push(entry);
  }
  for (const name of added) {
    entries.push({
      path: name, action: "added", backup: null, previousHash: null, newHash: remote.files[name],
    });
  }
  const manifestBackup = `${backupPath}/files/${MANIFEST}`;
  await writeFileAt(at(root, manifestBackup), await readFile(path.join(root, MANIFEST)));

  // The state record exists before the first write, so an interrupted apply is still rollback-able.
  const state = {
    schema: 1,
    runId,
    startedAt: new Date(now()).toISOString(),
    status: "in-progress",
    fromVersion: preview.localVersion,
    toVersion: remote.version,
    backupPath,
    manifestBackup,
    entries,
    diagnostics: [],
  };
  const stateFile = path.join(root, ...STATE.split("/"));
  await writeFileAt(stateFile, Buffer.from(JSON.stringify(state, null, 2) + "\n", "utf8"));

  const written = {};
  try {
    for (const [name, bytes] of staged) {
      await writeFileAt(at(root, name), bytes);
      written[name] = remote.files[name];
    }
    for (const name of removed) await rm(at(root, name), { force: true });
    await writeFile(path.join(root, MANIFEST), manifestBytes(remote));
  } catch (error) {
    state.status = "failed";
    state.diagnostics = [{ reason: "write-failed", detail: error.code || "unknown" }];
    await writeFileAt(stateFile, Buffer.from(JSON.stringify(state, null, 2) + "\n", "utf8"));
    result.status = "apply-failed";
    result.diagnostics = state.diagnostics;
    result.rollbackAvailable = true;
    result.message = "Writing failed. Run node refresh.mjs rollback to restore the backed-up state.";
    return result;
  }

  const health = await doctor(root, remote, written, removed);
  result.applied = { added, updated, removed };
  result.integrity = { ok: health.ok, failures: health.failures, localModified: health.localModified };
  state.status = health.ok ? "applied" : "failed";
  state.diagnostics = health.failures;
  state.finishedAt = new Date(now()).toISOString();
  await writeFileAt(stateFile, Buffer.from(JSON.stringify(state, null, 2) + "\n", "utf8"));
  if (!health.ok) {
    result.status = "apply-failed";
    result.diagnostics = health.failures;
    result.rollbackAvailable = true;
    result.message = "The post-update check failed. Diagnostics are kept in " + STATE +
      "; run node refresh.mjs rollback to restore the previous system files.";
    return result;
  }
  result.status = "applied";
  result.rollbackAvailable = true;
  result.message = "System files were updated. Backups are in " + backupPath + ".";
  return result;
}

// A run id is the only free text in the state record, so it may not carry a path segment:
// the backup directory is then derived from it and must match exactly.
const RUN_ID = /^[A-Za-z0-9-]+$/;

function validateState(value) {
  if (!value || value.schema !== 1 || typeof value.runId !== "string" || !value.runId ||
      !RUN_ID.test(value.runId) || value.backupPath !== `${BACKUPS}/${value.runId}` ||
      !Array.isArray(value.entries)) {
    throw new Error("state-invalid");
  }
  const optionalHash = (hash) => hash === null || (typeof hash === "string" && HASH.test(hash));
  for (const entry of value.entries) {
    // A backup path is derived, never trusted: it must be this run's copy of that exact file.
    if (!entry || !isSystemPath(entry.path) ||
        !["added", "updated", "removed"].includes(entry.action) ||
        !optionalHash(entry.previousHash) || !optionalHash(entry.newHash) ||
        (entry.backup !== null && entry.backup !== `${value.backupPath}/files/${entry.path}`)) {
      throw new Error("state-invalid");
    }
  }
  return value;
}

export async function rollback({ root = ROOT, now = Date.now } = {}) {
  const result = {
    schema: 1,
    mode: "rollback",
    status: "no-apply-recorded",
    runId: null,
    restored: [],
    removed: [],
    skipped: [],
    message: "No recorded refresh.mjs update was found, so nothing was changed.",
    protectedUserPaths: PROTECTED_USER_PATHS,
  };
  const stateFile = path.join(root, ...STATE.split("/"));
  let state;
  try {
    state = validateState(JSON.parse(await readFile(stateFile, "utf8")));
  } catch (error) {
    if (error.code !== "ENOENT") {
      result.status = "state-invalid";
      result.message = "The update state record is unreadable. No file was changed; restore manually.";
    }
    return result;
  }
  result.runId = state.runId;
  if (state.status === "rolled-back") {
    result.status = "already-rolled-back";
    result.message = "The most recent update was already rolled back. No file was changed.";
    return result;
  }

  for (const entry of state.entries) {
    const filename = at(root, entry.path);
    const current = await readIfPresent(filename).catch(() => null);
    const currentHash = current === null ? null : systemHash(entry.path, current);
    // Anything changed after the apply stays as it is. A file still holding its pre-apply
    // bytes (an interrupted apply) is recognized too, so a failed run stays rollback-able.
    const known = currentHash === null || currentHash === entry.newHash ||
      currentHash === entry.previousHash;
    if (!known) {
      result.skipped.push({ path: entry.path, reason: "changed-after-apply" });
      continue;
    }
    if (entry.action === "added") {
      if (current !== null) {
        try {
          await rm(filename, { force: true });
        } catch (error) {
          // A locked or busy path stays recorded as retryable instead of aborting the run.
          result.skipped.push({ path: entry.path, reason: "restore-failed", detail: error.code || "unknown" });
          continue;
        }
        result.removed.push(entry.path);
      } else {
        result.skipped.push({ path: entry.path, reason: "already-absent" });
      }
      continue;
    }
    if (entry.backup === null) {
      result.skipped.push({ path: entry.path, reason: "no-backup-recorded" });
      continue;
    }
    if (entry.action === "removed" && current !== null) {
      result.skipped.push({ path: entry.path, reason: "recreated-after-apply" });
      continue;
    }
    const bytes = await readIfPresent(at(root, entry.backup));
    if (bytes === null) {
      result.skipped.push({ path: entry.path, reason: "backup-missing" });
      continue;
    }
    try {
      await restoreFileAt(root, filename, bytes);
    } catch (error) {
      result.skipped.push({
        path: entry.path, reason: "restore-failed", detail: error.detail || error.code || "unknown",
      });
      continue;
    }
    result.restored.push(entry.path);
  }

  let manifestRestored = false;
  if (state.manifestBackup === `${state.backupPath}/files/${MANIFEST}`) {
    const bytes = await readIfPresent(at(root, state.manifestBackup));
    if (bytes !== null) {
      // The manifest is restored through the same helper and the same failure path as every
      // other entry: a locked, unwritable, or symlinked manifest is recorded as retryable
      // instead of writing through the link or aborting the whole rollback.
      try {
        await restoreFileAt(root, path.join(root, MANIFEST), bytes);
        manifestRestored = true;
      } catch (error) {
        result.skipped.push({
          path: MANIFEST, reason: "restore-failed", detail: error.detail || error.code || "unknown",
        });
      }
    }
  }
  result.restored.sort();
  result.removed.sort();

  let integrity = { ok: false, reason: "local-manifest-invalid", localModified: [] };
  try {
    const local = validateManifest(JSON.parse(await readFile(path.join(root, MANIFEST), "utf8")));
    const localModified = await localChanges(root, local);
    integrity = { ok: localModified.length === 0, reason: localModified.length ? "local-system-changed" : "ok", localModified };
    result.localVersion = local.version;
  } catch { /* reported through integrity */ }
  result.integrity = integrity;
  result.manifestRestored = manifestRestored;
  result.status = result.skipped.length ? "rolled-back-partial" : integrity.ok ? "rolled-back" : "rolled-back-unverified";
  result.message = result.skipped.length
    ? "The recorded update was reverted except for paths changed after the apply; see skipped."
    : integrity.ok
      ? "The most recent recorded update was reverted."
      : "Files were restored but the post-rollback integrity check did not pass; " +
        "rerun node refresh.mjs rollback to retry and re-verify.";
  // Only a fully verified rollback closes the state record. A partial or unverified rollback
  // stays open so it can be rerun: rerunning is a no-op for entries already restored (their
  // current bytes match the recorded hashes), so only skipped paths and the check are retried.
  state.status = result.status;
  state.rolledBackAt = new Date(now()).toISOString();
  await writeFileAt(stateFile, Buffer.from(JSON.stringify(state, null, 2) + "\n", "utf8"));
  return result;
}

export async function agentCheck({ root = ROOT, fetcher = fetch, force = false, now = Date.now } = {}) {
  const cacheFile = path.join(root, ...CACHE.split("/"));
  const shape = (status, extra) => ({
    schema: 1,
    mode: "agent-check",
    status,
    updateAvailable: status === "update-available",
    cached: false,
    checkedAt: null,
    localVersion: null,
    remoteVersion: null,
    advice: status === "update-available"
      ? "A newer tested system snapshot exists. Tell the user and let them run node refresh.mjs preview, " +
        "then node refresh.mjs apply --confirm. Never apply an update on your own."
      : "Continue the spreadsheet workflow.",
    ...extra,
  });

  if (!force) {
    try {
      const cache = JSON.parse(await readFile(cacheFile, "utf8"));
      if (cache && cache.schema === 1 && typeof cache.checkedAt === "number" &&
          CACHEABLE.has(cache.status) && now() - cache.checkedAt < CACHE_TTL_MS &&
          now() - cache.checkedAt >= 0) {
        return shape(cache.status, {
          cached: true,
          checkedAt: cache.checkedAt,
          localVersion: cache.localVersion ?? null,
          remoteVersion: cache.remoteVersion ?? null,
        });
      }
    } catch { /* a missing, stale, or corrupted cache falls through to a fresh read-only check */ }
  }

  let check;
  try {
    check = await inspect({ mode: "check", root, fetcher });
  } catch {
    // Fail open: a broken check never blocks spreadsheet work.
    return shape("offline", { advice: "Continue the spreadsheet workflow." });
  }
  const checkedAt = now();
  if (CACHEABLE.has(check.status)) {
    try {
      await writeFileAt(cacheFile, Buffer.from(JSON.stringify({
        schema: 1, status: check.status, checkedAt,
        localVersion: check.localVersion, remoteVersion: check.remoteVersion,
      }, null, 2) + "\n", "utf8"));
    } catch { /* an unwritable cache only costs a repeated check */ }
  }
  return shape(check.status, {
    checkedAt,
    localVersion: check.localVersion,
    remoteVersion: check.remoteVersion,
  });
}

const FAILED = new Set([
  "conflict-blocked", "download-failed", "download-verification-failed", "apply-failed",
  "state-invalid", "rolled-back-unverified", "local-manifest-invalid", "remote-manifest-invalid",
]);

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const [command, ...flags] = process.argv.slice(2);
  const usage = "Supported commands: node refresh.mjs check | preview | apply --confirm | " +
    "rollback | agent-check [--force]\n";
  const runner = {
    check: () => flags.length === 0 && inspect({ mode: "check" }),
    preview: () => flags.length === 0 && inspect({ mode: "preview" }),
    apply: () => flags.every((flag) => flag === "--confirm") && flags.length <= 1 &&
      apply({ confirm: flags.includes("--confirm") }),
    rollback: () => flags.length === 0 && rollback(),
    "agent-check": () => flags.every((flag) => flag === "--force") && flags.length <= 1 &&
      agentCheck({ force: flags.includes("--force") }),
  }[command];
  const started = runner && runner();
  if (!started) {
    process.stderr.write(usage);
    process.exitCode = 2;
  } else {
    started.then(
      (result) => {
        process.stdout.write(JSON.stringify(result) + "\n");
        if (result.status === "confirmation-required") process.exitCode = 2;
        else if (FAILED.has(result.status) && result.mode !== "agent-check") process.exitCode = 1;
      },
      () => {
        process.stderr.write(`Refresh ${command} failed.\n`);
        process.exitCode = 1;
      },
    );
  }
}

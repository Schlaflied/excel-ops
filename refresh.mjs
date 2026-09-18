#!/usr/bin/env node
// Read-only system update check. Workbook processing remains in Python.
import { createHash } from "node:crypto";
import { lstat, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const MANIFEST = "refresh-manifest.json";
const REMOTE = "https://raw.githubusercontent.com/Schlaflied/excel-ops/main/refresh-manifest.json";
const REPOSITORY = "Schlaflied/excel-ops";
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
      if (error.code !== "ENOENT") throw error;
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

export async function inspect({ mode = "check", root = ROOT, fetcher = fetch } = {}) {
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
      if (error.code !== "ENOENT") throw error;
    }
  }
  result.systemFiles.conflicts.sort();
  result.status = result.systemFiles.localModified.length ? "local-system-changed" :
    incoming.size ? "update-available" : "up-to-date";
  return result;
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const command = process.argv[2];
  if (!["check", "preview"].includes(command) || process.argv.length !== 3) {
    process.stderr.write("Supported commands: node refresh.mjs check | preview\n");
    process.exitCode = 2;
  } else {
    inspect({ mode: command }).then(
      (result) => process.stdout.write(JSON.stringify(result) + "\n"),
      () => {
        process.stderr.write("Refresh check failed without changing files.\n");
        process.exitCode = 1;
      },
    );
  }
}

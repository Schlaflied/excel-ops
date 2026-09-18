#!/usr/bin/env node
// Maintainer-only command: refresh hashes for the explicit system-file list.
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { isSystemPath, systemHash, validateManifest } from "../refresh.mjs";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const filename = path.join(root, "refresh-manifest.json");
const manifest = validateManifest(JSON.parse(await readFile(filename, "utf8")));
const files = {};
for (const name of Object.keys(manifest.files).sort()) {
  if (!isSystemPath(name)) throw new Error("invalid system path");
  files[name] = systemHash(name, await readFile(path.join(root, ...name.split("/"))));
}
const content = JSON.stringify({ ...manifest, files }, null, 2) + "\n";
if (process.argv[2] === "--check") {
  if (content !== await readFile(filename, "utf8")) {
    process.stderr.write("refresh-manifest.json is stale; regenerate it.\n");
    process.exitCode = 1;
  }
} else if (process.argv.length === 2) {
  await writeFile(filename, content);
} else {
  process.stderr.write("Usage: node scripts/update-refresh-manifest.mjs [--check]\n");
  process.exitCode = 2;
}

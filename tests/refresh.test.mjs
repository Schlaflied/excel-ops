import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";
import { inspect, isSystemPath, systemHash, validateManifest } from "../refresh.mjs";

const hash = (value) => createHash("sha256").update(value).digest("hex");
const response = (manifest) => async () => ({ ok: true, json: async () => manifest });
const makeManifest = (files) => ({
  schema: 1, repository: "Schlaflied/excel-ops", version: "0.3.0", files,
});

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
    const remote = makeManifest({ "docs/refresh.md": hash("new") });
    const result = await inspect({ root, fetcher: response(remote) });
    assert.deepEqual(result.systemFiles.removed, ["README.md"]);
    assert.deepEqual(result.systemFiles.added, ["docs/refresh.md"]);
    assert.deepEqual(result.systemFiles.conflicts, ["docs/refresh.md"]);
  });
});

"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const { parseOptions } = require("../extractEmails");

const ROOT = path.join(__dirname, "..");

test("parseOptions defaults and validation", () => {
  const o = parseOptions([]);
  assert.equal(o.country, "PS");
  assert.equal(o.concurrency, 5);
  assert.equal(o.includeAttachments, true);
  assert.equal(path.basename(o.output), "extracted_emails.xlsx");

  const custom = parseOptions(["--no-attachments", "--country", "jo", "--concurrency", "2", "--query", "after:2026/01/01"]);
  assert.equal(custom.includeAttachments, false);
  assert.equal(custom.country, "JO");
  assert.equal(custom.query, "after:2026/01/01");

  assert.throws(() => parseOptions(["--country", "XX"]), /country/);
  assert.throws(() => parseOptions(["--concurrency", "0"]), /concurrency/);
  assert.throws(() => parseOptions(["--unknown"]));
});

test("a failed start leaves earlier results untouched", (t) => {
  // Run a copy of the tool in an empty folder, so no real credentials are found.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "extractor-cli-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  fs.copyFileSync(path.join(ROOT, "extractEmails.js"), path.join(dir, "extractEmails.js"));
  fs.cpSync(path.join(ROOT, "lib"), path.join(dir, "lib"), { recursive: true });
  fs.symlinkSync(path.join(ROOT, "node_modules"), path.join(dir, "node_modules"), "junction");
  const previous = path.join(dir, "extracted_emails.xlsx");
  fs.writeFileSync(previous, "earlier results");

  const run = spawnSync(process.execPath, [path.join(dir, "extractEmails.js")], { cwd: dir, encoding: "utf8" });

  assert.equal(run.status, 1);
  assert.match(run.stderr, /credentials\.json was not found/);
  assert.equal(fs.readFileSync(previous, "utf8"), "earlier results");
});

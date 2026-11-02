"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const os = require("os");
const path = require("path");
const ExcelJS = require("exceljs");
const { extractContacts } = require("../lib/run");
const { writeContactsWorkbook, partialPathFor } = require("../lib/output");
const { rawEmail, makeXlsx, httpError, fakeGmail } = require("./helpers");

async function sampleMailbox() {
  return {
    m1: { raw: rawEmail({ headers: { From: "Ali <ali@example.org>", To: "me@example.org", Date: "Mon, 05 Jan 2026 10:00:00 +0000" }, text: "Call 0599123456" }) },
    m2: { raw: rawEmail({ headers: { From: "ALI@example.org", Date: "Sun, 01 Mar 2026 10:00:00 +0000" }, text: "again" }) },
    m3: {
      raw: rawEmail({ headers: { From: "me@example.org", To: "Sara <sara@example.org>", Cc: "omar@example.org" }, text: "hello" }),
      labelIds: ["SENT", "IMPORTANT"],
    },
    m4: { raw: rawEmail({ headers: { From: "me@example.org", To: "draft@example.org" } }), labelIds: ["DRAFT"] },
    m5: httpError(404),
    m6: {
      raw: rawEmail({
        headers: { From: "huda@example.org" },
        attachments: [{ filename: "list.xlsx", content: await makeXlsx([["Nour", "nour@example.org", "0599222333"]]) }],
      }),
    },
  };
}

test("extractContacts builds one de-duplicated list and keeps going past failures", async () => {
  const messages = await sampleMailbox();
  const progress = [];
  const result = await extractContacts({
    gmail: fakeGmail({ messages }),
    ids: Object.keys(messages),
    ownAddresses: ["me@example.org"],
    onProgress: (p) => progress.push(p.processed),
  });

  assert.equal(result.stopped, false);
  assert.equal(result.processed, 6);
  assert.equal(result.skipped, 1);
  assert.deepEqual(result.failed.map((f) => f.id), ["m5"]);
  assert.deepEqual(progress, [1, 2, 3, 4, 5, 6]);
  assert.deepEqual(
    result.store.toRows().map((r) => [r.email, r.name, r.phone, r.source, r.messages, r.last_seen]),
    [
      ["ali@example.org", "Ali", "+970 599 123 456", "sender", 2, "2026-03-01"],
      ["huda@example.org", "", "", "sender", 1, ""],
      ["nour@example.org", "", "+970 599 222 333", "attachment", 1, ""],
      ["omar@example.org", "", "", "recipient", 1, ""],
      ["sara@example.org", "Sara", "", "recipient", 1, ""],
    ]
  );
});

test("extractContacts stops early when asked", async () => {
  const messages = await sampleMailbox();
  let stop = false;
  const result = await extractContacts({
    gmail: fakeGmail({ messages }),
    ids: Object.keys(messages),
    concurrency: 1,
    shouldStop: () => stop,
    onProgress: ({ processed }) => { if (processed === 2) stop = true; },
  });
  assert.equal(result.stopped, true);
  assert.equal(result.processed, 2);
});

test("writeContactsWorkbook saves a readable file and leaves no temporary file", async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "extractor-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const file = path.join(dir, "out.xlsx");
  const rows = [{ name: "=HYPERLINK(\"x\")", email: "a@example.org", phone: "", source: "sender", messages: 1, last_seen: "" }];
  await writeContactsWorkbook(rows, file);

  assert.deepEqual(fs.readdirSync(dir), ["out.xlsx"]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  const sheet = workbook.getWorksheet("Contacts");
  assert.deepEqual(sheet.getRow(1).values.slice(1), ["name", "email", "phone", "source", "messages", "last_seen"]);
  // Text that looks like a formula is stored as plain text.
  assert.equal(sheet.getCell("A2").value, "=HYPERLINK(\"x\")");
});

test("partialPathFor names the partial-results file", () => {
  assert.equal(partialPathFor("/x/extracted_emails.xlsx"), "/x/extracted_emails.partial.xlsx");
  assert.equal(partialPathFor("/x/results"), "/x/results.partial.xlsx");
});

"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { attachmentLines, attachmentType, MAX_ATTACHMENT_BYTES } = require("../lib/attachments");
const { makePdf, makeDocx, makeXlsx } = require("./helpers");

test("reads PDF attachments from memory", async () => {
  const lines = await attachmentLines({ filename: "list.pdf", content: makePdf("Ali ali@example.org 0599123456") });
  assert.ok(lines.some((l) => l.includes("ali@example.org") && l.includes("0599123456")));
});

test("reads Word attachments from memory", async () => {
  const lines = await attachmentLines({
    filename: "list.docx",
    content: makeDocx(["Participants", "Huda huda@example.org 0599222333"]),
  });
  assert.ok(lines.includes("Huda huda@example.org 0599222333"));
});

test("reads Excel attachments one row per line, including number cells", async () => {
  const content = await makeXlsx([
    ["Name", "Email", "Phone"],
    ["Omar", "omar@example.org", 599444555],
  ]);
  const lines = await attachmentLines({ filename: "list.xlsx", content });
  assert.deepEqual(lines, ["Name | Email | Phone", "Omar | omar@example.org | 599444555"]);
});

test("recognises attachments by content type when the name has no extension", () => {
  assert.equal(attachmentType({ filename: "scan", contentType: "application/pdf" }), "pdf");
  assert.equal(attachmentType({ filename: "photo.jpg", contentType: "image/jpeg" }), null);
});

test("skips unsupported and oversized attachments", async () => {
  assert.equal(await attachmentLines({ filename: "photo.jpg", content: Buffer.from("x") }), null);
  const huge = { filename: "big.pdf", content: { length: MAX_ATTACHMENT_BYTES + 1 } };
  assert.equal(await attachmentLines(huge), null);
});

test("throws on a damaged file so the caller can report it", async () => {
  await assert.rejects(attachmentLines({ filename: "bad.pdf", content: Buffer.from("not a pdf") }));
});

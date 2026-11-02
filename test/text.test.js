"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { extractEmails, extractPhoneNumbers, stripQuotedText, findContactsInLines } = require("../lib/text");

test("extractEmails keeps + addresses whole, drops trailing full stops, lower-cases", () => {
  assert.deepEqual(extractEmails("Write to John.Smith+unicef@Gmail.com. Or info@example.org."), [
    "john.smith+unicef@gmail.com",
    "info@example.org",
  ]);
});

test("extractEmails ignores image IDs and repeated addresses", () => {
  assert.deepEqual(extractEmails("cid:image001.png@01DA1234.5678ABCD"), []);
  assert.deepEqual(extractEmails("a@example.org, A@EXAMPLE.ORG"), ["a@example.org"]);
});

test("extractEmails and extractPhoneNumbers accept non-text input without crashing", () => {
  assert.deepEqual(extractEmails(undefined), []);
  assert.deepEqual(extractPhoneNumbers(["a@b.com"], "PS"), []);
  assert.deepEqual(extractPhoneNumbers(undefined, "PS"), []);
});

test("extractPhoneNumbers ignores dates, times and other numbers", () => {
  const text = "Meeting on 2024-03-15 at 10 30, invoice 4521, 250 participants. Call +970 59 123 4567";
  assert.deepEqual(
    extractPhoneNumbers(text, "PS").map((p) => p.e164),
    ["+970591234567"]
  );
});

test("extractPhoneNumbers reads local numbers and Arabic-Indic digits", () => {
  const arabicDigits = "جوال: " + "٠٥٩٩١٢٣٤٥٦";
  assert.deepEqual(extractPhoneNumbers(arabicDigits, "PS").map((p) => p.e164), ["+970599123456"]);
  const found = extractPhoneNumbers("Mobile 0599123456 or 059-912-3456", "PS");
  assert.deepEqual(found, [{ e164: "+970599123456", display: "+970 599 123 456" }]);
});

test("stripQuotedText removes an English quoted reply, even when the header wraps", () => {
  const text = "Thanks\r\nCall 0599123456\r\n\r\nOn Mon, Nov 2, 2026 at 10:00 AM Ahmed <\r\na@x.org> wrote:\r\n> old 0599000000";
  assert.equal(stripQuotedText(text).trim(), "Thanks\nCall 0599123456");
});

test("stripQuotedText removes an Arabic quoted reply", () => {
  const rlm = String.fromCharCode(0x200f);
  const text = `شكرا\nجوال 0599123456\n\n${rlm}في الأحد، 1 نوفمبر 2026 في 10:00 ص، كتب أحمد <a@x.org>:\n> قديم 0599000000`;
  assert.equal(stripQuotedText(text).trim(), "شكرا\nجوال 0599123456");
});

test("stripQuotedText removes Outlook headers, forwards and > lines", () => {
  assert.equal(stripQuotedText("Hi\n\nFrom: Sara <s@x.org>\nSent: Monday\nold").trim(), "Hi");
  assert.equal(stripQuotedText("FYI\n---------- Forwarded message ---------\nFrom: x").trim(), "FYI");
  assert.equal(stripQuotedText("new\n> old\n>> older").trim(), "new");
});

test("stripQuotedText keeps ordinary text that merely contains the word كتب", () => {
  const text = "قائمة الكتب:\nكتاب 1";
  assert.equal(stripQuotedText(text), text);
});

test("findContactsInLines pairs a phone with the only email on its line", () => {
  const result = findContactsInLines(
    ["Sara | sara@example.org | 0599123456", "a@example.org, b@example.org 0599111111", "no contacts here"],
    "PS"
  );
  assert.deepEqual(
    result.map((r) => [r.email, r.phones.map((p) => p.e164)]),
    [
      ["sara@example.org", ["+970599123456"]],
      ["a@example.org", []],
      ["b@example.org", []],
    ]
  );
});

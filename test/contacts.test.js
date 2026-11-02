"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const { simpleParser } = require("mailparser");
const { addressesOf, collectFromMessage, ContactStore } = require("../lib/contacts");
const { rawEmail, makeXlsx } = require("./helpers");

const parse = (spec) => simpleParser(rawEmail(spec));
const RECEIVED = { isSent: false, country: "PS", includeAttachments: true };
const SENT = { isSent: true, country: "PS", includeAttachments: true };
const summary = (contacts) => contacts.map((c) => [c.email, c.source, (c.phones || []).map((p) => p.e164)]);

test("addressesOf handles missing fields, lists and groups", async () => {
  assert.deepEqual(addressesOf(undefined), []);
  const parsed = await parse({ headers: { From: "x@example.org", To: "Team: a@example.org, B <b@example.org>;" } });
  assert.deepEqual(addressesOf(parsed.to).map((a) => a.address), ["a@example.org", "b@example.org"]);
});

test("received email: sender with phone numbers from their own text only", async () => {
  const parsed = await parse({
    headers: { From: "Ali <ali@example.org>", To: "me@example.org" },
    text: "Call me on 0599123456\n\nOn Mon, Nov 2, 2026, Sara <sara@example.org> wrote:\n> my number is 0599000000",
  });
  const { contacts } = await collectFromMessage(parsed, RECEIVED);
  assert.deepEqual(summary(contacts), [["ali@example.org", "sender", ["+970599123456"]]]);
  assert.equal(contacts[0].name, "Ali");
});

test("received email with an empty body still records the sender", async () => {
  const parsed = await parse({ headers: { From: "Ali <ali@example.org>" }, text: "" });
  const { contacts } = await collectFromMessage(parsed, RECEIVED);
  assert.deepEqual(summary(contacts), [["ali@example.org", "sender", []]]);
});

test("sent email: To, Cc and Bcc recipients, without your own phone number", async () => {
  const parsed = await parse({
    headers: { From: "me@example.org", Cc: "Omar <omar@example.org>", Bcc: "hidden@example.org" },
    text: "Regards, me. Mobile 0599777888",
  });
  const { contacts } = await collectFromMessage(parsed, SENT);
  assert.deepEqual(summary(contacts), [
    ["omar@example.org", "recipient", []],
    ["hidden@example.org", "recipient", []],
  ]);
});

test("attachments are read in memory, whatever their file name says", async (t) => {
  const writeFileSync = t.mock.method(fs, "writeFileSync");
  const createWriteStream = t.mock.method(fs, "createWriteStream");
  const parsed = await parse({
    headers: { From: "Ali <ali@example.org>" },
    text: "See attached",
    attachments: [
      {
        filename: "../../list.xlsx",
        contentType: "application/octet-stream",
        content: await makeXlsx([["Huda", "huda@example.org", "0599222333"]]),
      },
      { filename: "photo.jpg", contentType: "image/jpeg", content: Buffer.from("jpeg") },
    ],
  });
  const { contacts, warnings } = await collectFromMessage(parsed, RECEIVED);
  assert.deepEqual(summary(contacts), [
    ["ali@example.org", "sender", []],
    ["huda@example.org", "attachment", ["+970599222333"]],
  ]);
  assert.deepEqual(warnings, []);
  assert.equal(writeFileSync.mock.callCount(), 0);
  assert.equal(createWriteStream.mock.callCount(), 0);
});

test("a damaged attachment gives a warning and the rest of the email is kept", async () => {
  const parsed = await parse({
    headers: { From: "ali@example.org" },
    attachments: [{ filename: "broken.pdf", contentType: "application/pdf", content: Buffer.from("not a pdf") }],
  });
  const { contacts, warnings } = await collectFromMessage(parsed, RECEIVED);
  assert.deepEqual(summary(contacts), [["ali@example.org", "sender", []]]);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /broken\.pdf/);
});

test("attachments are skipped when turned off", async () => {
  const parsed = await parse({
    headers: { From: "ali@example.org" },
    attachments: [{ filename: "list.xlsx", content: await makeXlsx([["huda@example.org"]]) }],
  });
  const { contacts } = await collectFromMessage(parsed, { ...RECEIVED, includeAttachments: false });
  assert.deepEqual(summary(contacts), [["ali@example.org", "sender", []]]);
});

test("ContactStore merges duplicates and ignores your own addresses", () => {
  const store = new ContactStore({ ignore: ["Me@Example.org"] });
  const phone = { e164: "+970599123456", display: "+970 599 123 456" };
  store.addMessage(
    [
      { email: "Ali@Example.org", name: "", phones: [], source: "sender" },
      { email: "ali@example.org", name: "Ali Hassan", phones: [phone], source: "attachment" },
      { email: "me@example.org", name: "Me", phones: [], source: "recipient" },
    ],
    new Date("2026-01-05")
  );
  store.addMessage([{ email: "ali@example.org", name: "Ali H.", phones: [phone], source: "sender" }], new Date("2026-03-01"));
  store.addMessage([{ email: "sara@example.org", name: '"sara@example.org"', phones: [], source: "recipient" }], null);

  assert.deepEqual(store.toRows(), [
    {
      name: "Ali Hassan",
      email: "ali@example.org",
      phone: "+970 599 123 456",
      source: "attachment, sender",
      messages: 2,
      last_seen: "2026-03-01",
    },
    { name: "", email: "sara@example.org", phone: "", source: "recipient", messages: 1, last_seen: "" },
  ]);
});

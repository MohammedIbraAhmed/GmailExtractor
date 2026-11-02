"use strict";

const { normalizeEmail, extractEmails, extractPhoneNumbers, stripQuotedText, findContactsInLines } = require("./text");
const { attachmentLines } = require("./attachments");

/** Flattens a mailparser address field (object, array, or groups) into [{ address, name }]. */
function addressesOf(field) {
  if (!field) return [];
  const objects = Array.isArray(field) ? field : [field];
  const out = [];
  const visit = (entry) => {
    if (!entry) return;
    if (Array.isArray(entry.group)) entry.group.forEach(visit);
    else if (entry.address) out.push({ address: entry.address, name: entry.name || "" });
  };
  for (const obj of objects) (obj && obj.value ? obj.value : []).forEach(visit);
  return out;
}

/**
 * Collects the contacts in one parsed email.
 * - Sent emails: everyone in To, Cc and Bcc ("recipient"). Phone numbers in
 *   the body are your own, so they are not given to recipients.
 * - Received emails: the sender ("sender"), with phone numbers found in the
 *   text they wrote (quoted replies and forwards are ignored).
 * - Attachments (PDF, Word, Excel): every email address ("attachment"), with
 *   phone numbers that sit on the same line or row.
 * Returns { contacts: [{ email, name, phones, source }], warnings: [string] }.
 */
async function collectFromMessage(parsed, { isSent, country, includeAttachments }) {
  const contacts = [];
  const warnings = [];

  if (isSent) {
    for (const field of ["to", "cc", "bcc"]) {
      for (const a of addressesOf(parsed[field])) {
        contacts.push({ email: a.address, name: a.name, phones: [], source: "recipient" });
      }
    }
  } else {
    const senders = addressesOf(parsed.from);
    const phones = senders.length ? extractPhoneNumbers(stripQuotedText(parsed.text || ""), country) : [];
    senders.forEach((a, i) => {
      contacts.push({ email: a.address, name: a.name, phones: i === 0 ? phones : [], source: "sender" });
    });
  }

  if (includeAttachments) {
    for (const attachment of parsed.attachments || []) {
      let lines;
      try {
        lines = await attachmentLines(attachment);
      } catch (err) {
        warnings.push(`could not read attachment "${attachment.filename || "(no name)"}": ${err.message}`);
        continue;
      }
      if (!lines) continue;
      for (const found of findContactsInLines(lines, country)) {
        contacts.push({ email: found.email, name: "", phones: found.phones, source: "attachment" });
      }
    }
  }

  // Header addresses can be malformed; keep only ones that look like real addresses.
  const valid = contacts.filter((c) => extractEmails(c.email).length === 1);
  return { contacts: valid, warnings };
}

/** De-duplicated contact list, keyed by lower-cased email address. */
class ContactStore {
  constructor({ ignore = [] } = {}) {
    this.ignore = new Set(ignore.map(normalizeEmail));
    this.contacts = new Map();
  }

  get size() {
    return this.contacts.size;
  }

  /** Adds the contacts found in one email. Each contact's email count goes up by at most one. */
  addMessage(contacts, date) {
    const seenInThisMessage = new Set();
    for (const c of contacts) {
      const key = normalizeEmail(c.email);
      if (!key || this.ignore.has(key)) continue;
      let entry = this.contacts.get(key);
      if (!entry) {
        entry = { email: key, name: "", phones: new Map(), sources: new Set(), messages: 0, lastSeen: null };
        this.contacts.set(key, entry);
      }
      const name = cleanName(c.name, key);
      if (!entry.name && name) entry.name = name;
      for (const p of c.phones || []) if (!entry.phones.has(p.e164)) entry.phones.set(p.e164, p.display);
      if (c.source) entry.sources.add(c.source);
      if (!seenInThisMessage.has(key)) {
        seenInThisMessage.add(key);
        entry.messages += 1;
        if (date instanceof Date && !isNaN(date) && (!entry.lastSeen || date > entry.lastSeen)) {
          entry.lastSeen = date;
        }
      }
    }
  }

  /** Rows for the spreadsheet, most frequent contacts first. */
  toRows() {
    return [...this.contacts.values()]
      .sort((a, b) => b.messages - a.messages || a.email.localeCompare(b.email))
      .map((e) => ({
        name: e.name,
        email: e.email,
        phone: [...e.phones.values()].join("; "),
        source: [...e.sources].sort().join(", "),
        messages: e.messages,
        last_seen: e.lastSeen ? e.lastSeen.toISOString().slice(0, 10) : "",
      }));
  }
}

function cleanName(name, email) {
  const n = String(name || "").replace(/\s+/g, " ").replace(/^["']+|["']+$/g, "").trim();
  return n && normalizeEmail(n) !== email ? n : "";
}

module.exports = { addressesOf, collectFromMessage, ContactStore };

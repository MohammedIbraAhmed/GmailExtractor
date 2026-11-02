"use strict";

const { findPhoneNumbersInText } = require("libphonenumber-js/max");

// Local part allows the characters real addresses use (including "+").
// The domain must end in a letters-only top-level domain, so a full stop
// at the end of a sentence is not swallowed and IDs like
// "image001.png@01DA1234.5678ABCD" are not mistaken for addresses.
const EMAIL_RE =
  /[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,24}(?![A-Za-z0-9-])/g;

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

/** Returns the unique email addresses found in `text`, lower-cased. */
function extractEmails(text) {
  if (typeof text !== "string" || text === "") return [];
  return [...new Set((text.match(EMAIL_RE) || []).map(normalizeEmail))];
}

/**
 * Returns the unique, valid phone numbers found in `text`.
 * Numbers written without a country code are read as belonging to
 * `defaultCountry` (a two-letter code such as "PS").
 * Each result is { e164: "+970599123456", display: "+970 599 123 456" }.
 */
function extractPhoneNumbers(text, defaultCountry) {
  if (typeof text !== "string" || text === "") return [];
  const found = new Map();
  for (const match of findPhoneNumbersInText(text, defaultCountry)) {
    const number = match.number;
    if (!found.has(number.number)) {
      found.set(number.number, { e164: number.number, display: number.formatInternational() });
    }
  }
  return [...found.values()];
}

// Lines that mark the start of quoted or forwarded text in a reply.
const QUOTE_MARKERS = [
  /^On (?:[^\n]{0,300}\n)?[^\n]{0,300}?wrote:[ \t]*\r?$/m, // Gmail / Apple Mail (can wrap onto 2 lines)
  /^\u200f?في [^\n]{0,300}?كتب[^\n]{0,300}?:[ \t\u200e\u200f\u202a-\u202c]*\r?$/m, // Gmail in Arabic
  /^-{2,}\s*(?:Original Message|Forwarded message)\s*-{2,}/im,
  /^-{2,}\s*(?:الرسالة الأصلية|رسالة معاد توجيهها)\s*-{2,}/m,
  /^_{10,}[ \t]*$/m, // Outlook separator line
  /^\*?From:\*?[ \t].*\r?\n\*?(?:Sent|Date):\*?[ \t]/m, // Outlook header block
  /^\*?من:\*?[ \t].*\r?\n\*?(?:تاريخ الإرسال|التاريخ|أرسلت):\*?[ \t]/m,
];

/**
 * Removes quoted replies and forwarded text, so that phone numbers found
 * in an email belong to the person who wrote it rather than to someone
 * further down the thread.
 */
function stripQuotedText(text) {
  if (typeof text !== "string") return "";
  let cut = text.length;
  for (const marker of QUOTE_MARKERS) {
    const m = marker.exec(text);
    if (m && m.index < cut) cut = m.index;
  }
  return text
    .slice(0, cut)
    .split(/\r?\n/)
    .filter((line) => !/^\s*>/.test(line))
    .join("\n");
}

/**
 * For attachments: pairs phone numbers with an email address only when
 * they sit on the same line (or spreadsheet row) as exactly one address.
 * Returns [{ email, phones }].
 */
function findContactsInLines(lines, defaultCountry) {
  const results = [];
  for (const line of lines) {
    const emails = extractEmails(line);
    if (emails.length === 0) continue;
    const phones = emails.length === 1 ? extractPhoneNumbers(line, defaultCountry) : [];
    for (const email of emails) results.push({ email, phones });
  }
  return results;
}

module.exports = {
  normalizeEmail,
  extractEmails,
  extractPhoneNumbers,
  stripQuotedText,
  findContactsInLines,
};

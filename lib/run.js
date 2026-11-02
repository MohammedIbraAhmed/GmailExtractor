"use strict";

const { simpleParser } = require("mailparser");
const { fetchMessage, runPool } = require("./gmail");
const { collectFromMessage, ContactStore } = require("./contacts");

// Drafts were never sent and chats are not emails.
const SKIPPED_LABELS = ["DRAFT", "CHAT"];

const PARSER_OPTIONS = {
  skipTextToHtml: true, // only plain text is needed
  skipTextLinks: true,
  skipImageLinks: true,
};

/**
 * Reads every email in `ids` and builds the contact list.
 * Problems with single emails are recorded and skipped; they never stop the run.
 * Returns { store, processed, skipped, failed: [{ id, error }], warnings: [string], stopped }.
 */
async function extractContacts({
  gmail,
  ids,
  ownAddresses = [],
  country = "PS",
  includeAttachments = true,
  concurrency = 5,
  shouldStop = () => false,
  onProgress = () => {},
}) {
  const store = new ContactStore({ ignore: ownAddresses });
  const failed = [];
  const warnings = [];
  let processed = 0;
  let skipped = 0;

  await runPool(
    ids,
    concurrency,
    async (id) => {
      try {
        const message = await fetchMessage(gmail, id);
        if (!message || message.labelIds.some((l) => SKIPPED_LABELS.includes(l))) {
          skipped++;
          return;
        }
        const parsed = await simpleParser(message.raw, PARSER_OPTIONS);
        const { contacts, warnings: messageWarnings } = await collectFromMessage(parsed, {
          isSent: message.labelIds.includes("SENT"),
          country,
          includeAttachments,
        });
        store.addMessage(contacts, parsed.date);
        for (const w of messageWarnings) warnings.push(`email ${id}: ${w}`);
      } catch (err) {
        failed.push({ id, error: err.message || String(err) });
      } finally {
        processed++;
        onProgress({ processed, total: ids.length, contacts: store.size, failed: failed.length, store });
      }
    },
    shouldStop
  );

  return { store, processed, skipped, failed, warnings, stopped: processed < ids.length };
}

module.exports = { extractContacts };

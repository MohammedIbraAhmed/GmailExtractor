#!/usr/bin/env node
"use strict";

// Builds a contact list (names, email addresses, phone numbers) from a Gmail
// mailbox and its PDF, Word and Excel attachments. See README.md.

const fs = require("fs");
const path = require("path");
const { parseArgs } = require("util");
const { isSupportedCountry } = require("libphonenumber-js/max");
const { UserFacingError, connect, listMessageIds } = require("./lib/gmail");
const { extractContacts } = require("./lib/run");
const { writeContactsWorkbook, partialPathFor } = require("./lib/output");

const CREDENTIALS_PATH = path.join(__dirname, "credentials.json");
const TOKEN_PATH = path.join(__dirname, "token.json");
const DEFAULT_OUTPUT = path.join(__dirname, "extracted_emails.xlsx");
const PROGRESS_EVERY = 100; // print progress every N emails
const CHECKPOINT_EVERY = 500; // save partial results every N emails

const HELP = `Usage: node extractEmails.js [options]

Options:
  --query "<search>"   Only read emails matching a Gmail search,
                       e.g. --query "after:2026/01/01" or --query "label:partners"
  --output <file>      Where to save the results (default: extracted_emails.xlsx)
  --no-attachments     Don't read PDF, Word and Excel attachments
  --country <code>     Country for phone numbers written without a country code
                       (two letters, default: PS)
  --concurrency <n>    How many emails to download at the same time (1-20, default: 5)
  --help               Show this help
`;

function parseOptions(argv) {
  const { values } = parseArgs({
    args: argv,
    options: {
      query: { type: "string" },
      output: { type: "string" },
      "no-attachments": { type: "boolean", default: false },
      country: { type: "string", default: "PS" },
      concurrency: { type: "string", default: "5" },
      help: { type: "boolean", default: false },
    },
  });
  const country = values.country.toUpperCase();
  if (!isSupportedCountry(country)) {
    throw new UserFacingError(`Unknown country code "${values.country}". Use a two-letter code such as PS or JO.`);
  }
  const concurrency = Number(values.concurrency);
  if (!Number.isInteger(concurrency) || concurrency < 1 || concurrency > 20) {
    throw new UserFacingError("--concurrency must be a whole number from 1 to 20.");
  }
  return {
    help: values.help,
    query: values.query,
    output: values.output ? path.resolve(values.output) : DEFAULT_OUTPUT,
    includeAttachments: !values["no-attachments"],
    country,
    concurrency,
  };
}

async function main() {
  const options = parseOptions(process.argv.slice(2));
  if (options.help) {
    console.log(HELP);
    return;
  }
  const partialPath = partialPathFor(options.output);

  const { gmail, ownAddresses } = await connect({ credentialsPath: CREDENTIALS_PATH, tokenPath: TOKEN_PATH });
  console.log(`Signed in as ${ownAddresses[0]}.`);

  console.log(options.query ? `Finding emails matching: ${options.query}` : "Finding all emails (Spam and Trash are skipped)...");
  const ids = await listMessageIds(gmail, options.query, {
    onPage: (count) => console.log(`  ${count} found so far...`),
  });
  if (ids.length === 0) {
    console.log("No emails matched. Nothing was saved.");
    return;
  }

  let stopRequested = false;
  process.on("SIGINT", () => {
    if (stopRequested) {
      // Remove any half-written file before quitting.
      for (const file of [options.output, partialPath]) fs.rmSync(`${file}.tmp`, { force: true });
      console.log("\nStopped without saving.");
      process.exit(130);
    }
    stopRequested = true;
    console.log("\nStopping after the emails in progress, then saving. Press Ctrl+C again to quit without saving.");
  });

  console.log(`Reading ${ids.length} emails. Press Ctrl+C to stop early and keep what has been collected.`);
  let checkpoint = Promise.resolve();
  let checkpointWarned = false;
  const result = await extractContacts({
    gmail,
    ids,
    ownAddresses,
    country: options.country,
    includeAttachments: options.includeAttachments,
    concurrency: options.concurrency,
    shouldStop: () => stopRequested,
    onProgress: ({ processed, total, contacts, failed, store }) => {
      if (processed % PROGRESS_EVERY === 0 || processed === total) {
        console.log(`  ${processed}/${total} emails read, ${contacts} contacts, ${failed} failed`);
      }
      // Save partial results now and then, so a crash or power cut loses little.
      if (processed % CHECKPOINT_EVERY === 0 && processed < total) {
        const rows = store.toRows();
        checkpoint = checkpoint
          .then(() => writeContactsWorkbook(rows, partialPath))
          .catch((err) => {
            if (!checkpointWarned) console.warn(`  Could not save partial results: ${err.message}`);
            checkpointWarned = true;
          });
      }
    },
  });
  await checkpoint;

  const rows = result.store.toRows();
  if (result.stopped) {
    await writeContactsWorkbook(rows, partialPath);
    console.log(`\nStopped early after ${result.processed} of ${ids.length} emails.`);
    console.log(`Saved ${rows.length} contacts to ${partialPath}`);
    console.log(`(${options.output} was not changed.)`);
    process.exitCode = 130;
  } else {
    await writeContactsWorkbook(rows, options.output);
    fs.rmSync(partialPath, { force: true });
    console.log(`\nDone. Saved ${rows.length} contacts to ${options.output}`);
  }

  if (result.skipped) console.log(`Drafts and chats skipped: ${result.skipped}`);
  if (result.warnings.length) {
    console.log(`Attachments that could not be read: ${result.warnings.length}`);
    for (const w of result.warnings.slice(0, 5)) console.log(`  - ${w}`);
  }
  if (result.failed.length) {
    console.log(`Emails that could not be read (left out): ${result.failed.length}`);
    for (const f of result.failed.slice(0, 5)) console.log(`  - email ${f.id}: ${f.error}`);
  }
}

if (require.main === module) {
  main().catch((err) => {
    // Print only the message: full error objects can include request details.
    console.error(`Error: ${(err && err.message) || err}`);
    process.exitCode = 1;
  });
}

module.exports = { parseOptions };

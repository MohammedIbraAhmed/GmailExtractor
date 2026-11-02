"use strict";

const fs = require("fs");
const readline = require("readline");
const { gmail: createGmailClient, auth } = require("@googleapis/gmail");

const SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];

class UserFacingError extends Error {}

function readJson(filePath, description) {
  let text;
  try {
    text = fs.readFileSync(filePath, "utf8");
  } catch (err) {
    if (err.code === "ENOENT") return null;
    throw new UserFacingError(`Could not read ${description} (${filePath}): ${err.message}`);
  }
  try {
    return JSON.parse(text);
  } catch {
    throw new UserFacingError(`${description} (${filePath}) is damaged (not valid JSON). Delete it or download it again.`);
  }
}

/** Makes the token readable only by the current user: it gives access to the whole mailbox. */
function restrictToOwner(filePath) {
  try {
    fs.chmodSync(filePath, 0o600);
  } catch {
    // Not supported on every system (e.g. Windows).
  }
}

function saveToken(tokenPath, token) {
  fs.writeFileSync(tokenPath, JSON.stringify(token), { mode: 0o600 });
  restrictToOwner(tokenPath);
}

function ask(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

/** Accepts either the bare code or the whole address the browser was sent to. */
function codeFromInput(input) {
  try {
    const url = new URL(input);
    return url.searchParams.get("code") || "";
  } catch {
    return input;
  }
}

async function requestNewToken(client, tokenPath) {
  const url = client.generateAuthUrl({ access_type: "offline", prompt: "consent", scope: SCOPES });
  console.log("\nTo let this tool read your Gmail (read-only), open this link and sign in:\n");
  console.log(url);
  console.log(
    "\nAfter you approve, your browser goes to a page that may not load." +
      "\nCopy the full address from the browser's address bar and paste it here."
  );
  const code = codeFromInput(await ask("\nPaste the address (or just the code): "));
  if (!code) throw new UserFacingError("No authorization code was entered.");
  const { tokens } = await client.getToken(code);
  client.setCredentials(tokens);
  saveToken(tokenPath, tokens);
  console.log(`Sign-in saved to ${tokenPath}`);
}

function isInvalidGrant(err) {
  const data = err && err.response && err.response.data;
  return (data && data.error === "invalid_grant") || /invalid_grant/.test(String(err && err.message));
}

/**
 * Signs in and returns { gmail, ownAddresses }.
 * If the saved sign-in has expired or been revoked, asks to sign in again.
 */
async function connect({ credentialsPath, tokenPath }) {
  const credentials = readJson(credentialsPath, "credentials.json");
  if (!credentials) {
    throw new UserFacingError(
      `credentials.json was not found at ${credentialsPath}.\n` +
        "Download an OAuth client (Desktop app) from Google Cloud Console and save it there. See README.md."
    );
  }
  const keys = credentials.installed || credentials.web;
  if (!keys || !keys.client_id || !keys.client_secret) {
    throw new UserFacingError("credentials.json does not contain an OAuth client ID and secret.");
  }
  if (!Array.isArray(keys.redirect_uris) || keys.redirect_uris.length === 0) {
    throw new UserFacingError("Redirect URIs are not defined in credentials.json.");
  }

  const client = new auth.OAuth2(keys.client_id, keys.client_secret, keys.redirect_uris[0]);
  // Google sends new access tokens as old ones expire; keep the saved copy current.
  client.on("tokens", (tokens) => {
    const saved = readJson(tokenPath, "token.json") || {};
    saveToken(tokenPath, { ...saved, ...tokens });
  });

  // Retries are handled by withRetry() below, in one place.
  const gmail = createGmailClient({ version: "v1", auth: client, retry: false });

  const savedToken = readJson(tokenPath, "token.json");
  if (savedToken) {
    restrictToOwner(tokenPath); // older versions saved it readable by everyone
    client.setCredentials(savedToken);
  } else {
    await requestNewToken(client, tokenPath);
  }

  let profile;
  try {
    profile = await withRetry(() => gmail.users.getProfile({ userId: "me" }));
  } catch (err) {
    if (!savedToken || !isInvalidGrant(err)) throw err;
    console.log("Your saved Gmail sign-in has expired or was revoked. Please sign in again.");
    fs.rmSync(tokenPath, { force: true });
    await requestNewToken(client, tokenPath);
    profile = await withRetry(() => gmail.users.getProfile({ userId: "me" }));
  }

  const ownAddresses = [profile.data.emailAddress];
  try {
    const sendAs = await withRetry(() => gmail.users.settings.sendAs.list({ userId: "me" }));
    for (const s of sendAs.data.sendAs || []) ownAddresses.push(s.sendAsEmail);
  } catch {
    // Aliases are optional; the main address is enough.
  }
  return { gmail, ownAddresses: ownAddresses.filter(Boolean) };
}

const RATE_LIMIT_REASONS = new Set(["rateLimitExceeded", "userRateLimitExceeded"]);

/** True for errors worth retrying: rate limits, Google server errors, network failures. */
function isRetryable(err) {
  if (!err) return false;
  const status = Number((err.response && err.response.status) || err.status) || 0;
  if (status) {
    if (status === 408 || status === 429 || status >= 500) return true;
    if (status === 403) {
      const errors = (err.response && err.response.data && err.response.data.error && err.response.data.error.errors) || [];
      return errors.some((e) => RATE_LIMIT_REASONS.has(e.reason));
    }
    return false;
  }
  // No HTTP status: the request never got an answer (timeout, connection reset, DNS...).
  return Boolean(err.config) || /^(E[A-Z_]+|UND_ERR_\w+)$/.test(String(err.code || (err.cause && err.cause.code) || ""));
}

const sleepFor = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Runs `fn`, retrying temporary failures with growing waits
 * (about 1, 2, 4, 8, 16, 32 seconds). Permanent errors are thrown at once.
 */
async function withRetry(fn, { retries = 6, baseDelayMs = 1000, sleep = sleepFor } = {}) {
  for (let attempt = 0; ; attempt++) {
    try {
      return await fn();
    } catch (err) {
      if (attempt >= retries || !isRetryable(err)) throw err;
      const delay = Math.min(baseDelayMs * 2 ** attempt, 60000);
      await sleep(delay + Math.random() * baseDelayMs * 0.5);
    }
  }
}

/** Returns the IDs of all emails matching `query`, reading every page of results. */
async function listMessageIds(gmail, query, { onPage } = {}) {
  const ids = [];
  let pageToken;
  do {
    const res = await withRetry(() =>
      gmail.users.messages.list({ userId: "me", q: query || undefined, maxResults: 500, pageToken })
    );
    for (const m of res.data.messages || []) ids.push(m.id);
    pageToken = res.data.nextPageToken || undefined;
    if (onPage) onPage(ids.length);
  } while (pageToken);
  return ids;
}

/** Downloads one email in its original form. Returns { raw: Buffer, labelIds: [] } or null. */
async function fetchMessage(gmail, id) {
  const res = await withRetry(() => gmail.users.messages.get({ userId: "me", id, format: "raw" }));
  if (!res.data.raw) return null;
  return { raw: Buffer.from(res.data.raw, "base64url"), labelIds: res.data.labelIds || [] };
}

/**
 * Runs `worker` over `items`, at most `limit` at a time.
 * Stops starting new items once `shouldStop()` returns true.
 */
async function runPool(items, limit, worker, shouldStop = () => false) {
  let next = 0;
  const lane = async () => {
    while (next < items.length && !shouldStop()) {
      const index = next++;
      await worker(items[index], index);
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, lane));
}

module.exports = {
  UserFacingError,
  connect,
  codeFromInput,
  isRetryable,
  withRetry,
  listMessageIds,
  fetchMessage,
  runPool,
};

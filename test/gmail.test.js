"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { isRetryable, withRetry, listMessageIds, runPool, codeFromInput } = require("../lib/gmail");
const { httpError, fakeGmail } = require("./helpers");

test("isRetryable: rate limits, server errors and network failures only", () => {
  assert.equal(isRetryable(httpError(429)), true);
  assert.equal(isRetryable(httpError(503)), true);
  assert.equal(isRetryable(httpError(403, "userRateLimitExceeded")), true);
  assert.equal(isRetryable(httpError(403, "insufficientPermissions")), false);
  assert.equal(isRetryable(httpError(404)), false);
  assert.equal(isRetryable(httpError(400)), false);
  assert.equal(isRetryable(Object.assign(new Error("reset"), { code: "ECONNRESET" })), true);
  assert.equal(isRetryable(new TypeError("bug")), false);
});

test("withRetry waits longer each time and then succeeds", async () => {
  const delays = [];
  let calls = 0;
  const result = await withRetry(
    async () => {
      calls++;
      if (calls < 3) throw httpError(429);
      return "ok";
    },
    { baseDelayMs: 100, sleep: async (ms) => delays.push(ms) }
  );
  assert.equal(result, "ok");
  assert.equal(calls, 3);
  assert.ok(delays[0] >= 100 && delays[0] < 150);
  assert.ok(delays[1] >= 200 && delays[1] < 250);
});

test("withRetry does not retry permanent errors, and gives up eventually", async () => {
  let calls = 0;
  await assert.rejects(withRetry(async () => { calls++; throw httpError(404); }, { sleep: async () => {} }));
  assert.equal(calls, 1);

  calls = 0;
  await assert.rejects(withRetry(async () => { calls++; throw httpError(500); }, { retries: 3, sleep: async () => {} }));
  assert.equal(calls, 4);
});

test("listMessageIds reads every page", async () => {
  const gmail = fakeGmail({ pages: [["a", "b"], ["c"], ["d", "e"]] });
  const pageCounts = [];
  const ids = await listMessageIds(gmail, "", { onPage: (n) => pageCounts.push(n) });
  assert.deepEqual(ids, ["a", "b", "c", "d", "e"]);
  assert.deepEqual(pageCounts, [2, 3, 5]);
});

test("runPool never runs more than the limit at once, and can be stopped", async () => {
  let running = 0;
  let peak = 0;
  const done = [];
  await runPool([1, 2, 3, 4, 5, 6, 7], 3, async (n) => {
    running++;
    peak = Math.max(peak, running);
    await new Promise((r) => setTimeout(r, 5));
    running--;
    done.push(n);
  });
  assert.equal(peak, 3);
  assert.equal(done.length, 7);

  const started = [];
  await runPool([1, 2, 3, 4], 1, async (n) => started.push(n), () => started.length >= 2);
  assert.deepEqual(started, [1, 2]);
});

test("codeFromInput accepts the pasted address or the bare code", () => {
  assert.equal(codeFromInput("http://localhost/?code=4/abc&scope=x"), "4/abc");
  assert.equal(codeFromInput("4/abc"), "4/abc");
});

# Code Cleaning Report - GmailExtractor

**Date:** 2026-04-06
**File analyzed:** `extractEmails.js` (551 lines), `package.json`

---

## Summary

| Category | Issues Found | Severity |
|---|---|---|
| Dead Code | ~260 lines of commented-out code | Medium |
| Unused Dependencies | 3 packages | Low |
| Outdated Dependencies | 4 packages | Low |
| Security Vulnerabilities | 3 issues | **Critical** |
| Missing Error Handling | 5 locations | High |
| Code Style Inconsistencies | 6 issues | Low |
| Functions Needing Refactoring | 2 functions | Medium |
| Hardcoded Values | 6 values | Medium |
| Missing Infrastructure | 5 items | Medium |

---

## 1. Dead Code (Lines 291-551)

**Severity: Medium**

~260 lines of commented-out code starting at line 291. This is an earlier version of the same application that only extracted email addresses (without phone numbers) and used a `Set` for deduplication instead of an array.

**Recommendation:** Remove entirely. Use git history to retrieve old versions if needed.

```
// Line 291: "code below is working fine and it exctract email only"
// Lines 293-551: Full duplicate of every function in the active codebase
```

**Functions duplicated:**
- `authorize()` (lines 316-335)
- `getNewToken()` (lines 337-359)
- `listLabels()` (lines 361-374)
- `listMessagesFromLabel()` (lines 376-401)
- `getMessage()` (lines 403-458)
- `saveAttachment()` (lines 460-464)
- `parseDOCX()` (lines 476-484)
- `parsePDF()` (lines 486-494)
- `parseXLSX()` (lines 496-516)
- `extractEmailsFromText()` (lines 518-521)
- `parseAttachments()` (lines 523-539)
- `storeEmailsToFile()` (lines 541-547)

---

## 2. Unused Dependencies

**Severity: Low**

| Package | Version | Issue |
|---|---|---|
| `fs` | 0.0.1-security | Built-in Node.js module; should not be in `package.json` |
| `imap` | 0.8.17 | Never imported or used anywhere in the code |
| `node-fetch` | 3.3.2 | Never imported or used anywhere in the code |

**Recommendation:** Remove all three from `package.json` and run `npm install`.

---

## 3. Outdated Dependencies

**Severity: Low**

| Package | Current | Status |
|---|---|---|
| `googleapis` | 140.0.0 | Multiple major versions behind |
| `mailparser` | 3.7.1 | Minor versions behind |
| `mammoth` | 1.7.2 | Minor versions behind |
| `pdf-parse` | 1.1.1 | **Major version behind** (latest is 2.x) |

**Recommendation:** Run `npm outdated` to check exact latest versions, then update incrementally with testing.

---

## 4. Security Vulnerabilities

**Severity: Critical**

### 4.1 Path Traversal Vulnerability (Line 178)

```javascript
const filePath = path.join(__dirname, "attachments", attachment.filename);
fs.writeFileSync(filePath, attachment.content);
```

`attachment.filename` is used directly without sanitization. A malicious email attachment with a filename like `../../credentials.json` could write files outside the `attachments/` directory.

**Fix:** Sanitize the filename by stripping directory components:
```javascript
const safeName = path.basename(attachment.filename);
const filePath = path.join(__dirname, "attachments", safeName);
```

### 4.2 OAuth Token Logged to Console (Lines 51, 63)

```javascript
console.log("Authorize this app by visiting this url:", authUrl);  // Line 51
console.log("Token stored to", TOKEN_PATH);                        // Line 63
```

While the auth URL contains sensitive OAuth parameters, and the token path is logged. The `authUrl` on line 51 exposes the client ID and redirect URI.

**Fix:** Avoid logging sensitive URLs in full. Use a proper logger with configurable levels.

### 4.3 No Attachment Size Validation (Line 179)

No check on `attachment.content` size before writing to disk. A malicious email with an extremely large attachment could fill up the disk.

**Fix:** Add a size check before `fs.writeFileSync()`:
```javascript
const MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024; // 10MB
if (attachment.content.length > MAX_ATTACHMENT_SIZE) {
  console.warn(`Skipping oversized attachment: ${attachment.filename}`);
  return null;
}
```

---

## 5. Missing Error Handling

**Severity: High**

| Location | Issue | Risk |
|---|---|---|
| Line 131 | `parsed.from.value[0]` — no null check on `parsed.from` | Crash if email has no sender |
| Line 135 | `parsed.to.value` — no null check on `parsed.to` | Crash if email has no recipients |
| Line 144 | `extractPhoneNumbers(parsed.text)` — no check if `parsed.text` exists | Passes `undefined` to regex `.match()` |
| Line 152 | `saveAttachment()` called without try-catch | Crash if `attachments/` directory doesn't exist |
| Line 179 | `fs.writeFileSync()` with no error handling | Crash on disk full or permission errors |

**Example fix for lines 130-141:**
```javascript
let emails = [];
if (emailType === "from" && parsed.from?.value?.[0]) {
  const from = parsed.from.value[0];
  emails.push({ name: from.name, email: from.address });
} else if (emailType === "to" && parsed.to?.value) {
  const to = parsed.to.value.map((t) => ({
    name: t.name,
    email: t.address,
  }));
  emails = emails.concat(to);
}
```

---

## 6. Code Style Inconsistencies

**Severity: Low**

### 6.1 Mixed Async Patterns
- **Callback style:** `fs.readFile()` (line 20), `gmail.users.labels.list()` (line 72), `gmail.users.messages.list()` (line 87)
- **Async/await style:** `getMessage()` (line 112), `parseDOCX()` (line 183), `parsePDF()` (line 193)

**Recommendation:** Migrate all callback-based code to async/await for consistency.

### 6.2 Inconsistent File System Operations
- `fs.writeFile()` (async, line 61) vs `fs.writeFileSync()` (sync, line 179)
- `fs.unlink()` (async, line 272) used alongside sync operations

### 6.3 Excessive Console Logging
20+ `console.log`/`console.error` statements throughout the code. Production code should use a configurable logger (e.g., `winston` or `pino`).

**Lines with console output:** 21, 29-31, 51, 59, 62-63, 73, 76, 97, 100, 122, 133, 140, 156, 163, 166, 168, 188, 198, 220, 274, 276

### 6.4 Typo in Comment
Line 291: `"exctract"` should be `"extract"`

### 6.5 Variable Shadowing
Line 135: `parsed.to.value.map((to) => ...)` — parameter `to` shadows the outer variable `to`.

### 6.6 Bug on Line 140
```javascript
console.log("To:", to.join(", "));
```
`to` is an array of objects `{name, email}`, and `.join(", ")` will produce `[object Object], [object Object]`. This is a bug, not just a style issue.

---

## 7. Functions Needing Refactoring

**Severity: Medium**

### 7.1 `getMessage()` (Lines 112-175, 64 lines)

This function handles too many responsibilities:
1. API call with retry logic
2. Email parsing (from/to)
3. Phone number extraction
4. Attachment processing (save, parse, extract, store)
5. Error handling and retries

**Recommendation:** Split into:
- `fetchRawMessage(auth, messageId)` — API call with retries
- `processEmailContent(parsed, emailType)` — extract emails and phone numbers
- `processAttachments(parsed)` — handle attachment pipeline

### 7.2 `parseXLSX()` (Lines 203-223)

Three levels of nested `forEach` loops make this hard to follow.

**Recommendation:** Use `flatMap` for cleaner code:
```javascript
async function parseXLSX(filePath) {
  const workbook = xlsx.readFile(filePath);
  return workbook.SheetNames.flatMap((sheetName) => {
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    return data.flatMap((row) =>
      row.filter((cell) => typeof cell === "string").flatMap(extractEmailsFromText)
    );
  });
}
```

---

## 8. Hardcoded Values

**Severity: Medium**

| Line | Value | Recommendation |
|---|---|---|
| 10 | `TOKEN_PATH = "token.json"` | Move to environment variable or config file |
| 12 | `EMAIL_START_DATE = "2000/01/01"` | Accept as CLI argument or env variable |
| 20 | `"credentials.json"` | Move to constant/config |
| 178 | `"attachments"` directory | Move to constant |
| 267 | `"extracted_emails.xlsx"` | Accept as CLI argument |
| 3 | `package.json` `"main": "index.js"` | Should be `"extractEmails.js"` |

---

## 9. Missing Infrastructure

**Severity: Medium**

| Item | Status | Impact |
|---|---|---|
| Tests | None (`"test": "echo \"Error: no test specified\" && exit 1"`) | No confidence in changes |
| README | Missing | No setup/usage documentation |
| Linting | No ESLint/Prettier config | Inconsistent code style |
| CI/CD | No pipeline configured | No automated checks |
| `.env.example` | Missing | Users don't know what environment variables are needed |

---

## 10. Additional Observations

### 10.1 `package.json` Entry Point Mismatch
`"main": "index.js"` but the actual entry point is `extractEmails.js`.

### 10.2 Phone Number Regex Too Permissive (Line 232)
```javascript
/(\+?\d{1,4}[\s-]?)?(\(?\d{1,4}\)?[\s-]?)?\d{1,4}[\s-]?\d{1,4}[\s-]?\d{1,9}/gi
```
This regex is overly broad and will match arbitrary number sequences, dates, and IDs. Consider using a library like `libphonenumber-js` for reliable phone number detection.

### 10.3 No Pagination in Message Listing (Line 87-109)
`gmail.users.messages.list()` returns paginated results but the code only processes the first page. Large mailboxes will have incomplete results.

### 10.4 `process.on("exit")` Cleanup Issue (Lines 286-289)
The `exit` event handler calls `cleanUpAttachments()` which uses async `fs.unlink()`. Async operations in the `exit` handler will not complete because Node.js does not wait for them. Should use `fs.unlinkSync()` instead, or move cleanup to a `beforeExit` event.

### 10.5 No Email Deduplication
The active code pushes all emails to `emailData[]` without checking for duplicates. The old commented-out code used a `Set` for deduplication, but that was lost during the rewrite.

---

## Priority Action Items

1. **[Critical]** Fix path traversal vulnerability in `saveAttachment()` (line 178)
2. **[Critical]** Add null checks for `parsed.from`, `parsed.to`, `parsed.text` (lines 131-144)
3. **[Critical]** Fix async cleanup in `process.on("exit")` (line 286) — operations won't complete
4. **[High]** Remove 260 lines of dead code (lines 291-551)
5. **[High]** Remove unused dependencies (`fs`, `imap`, `node-fetch`)
6. **[High]** Fix `to.join()` bug on line 140
7. **[Medium]** Add pagination for message listing
8. **[Medium]** Add email deduplication
9. **[Medium]** Fix `package.json` main entry point
10. **[Low]** Migrate callbacks to async/await
11. **[Low]** Replace console.log with a proper logger
12. **[Low]** Add README, tests, and linting config

# Gmail contact extractor

Builds one Excel contact list from a Gmail mailbox:

- **People who emailed you**: the sender's name, address, and any phone
  numbers in the text they wrote. Quoted replies and forwarded text are
  ignored, so numbers belong to the right person.
- **People you emailed**: everyone in To, Cc and Bcc of your sent mail.
- **PDF, Word (.docx) and Excel (.xlsx) attachments**: every email address,
  plus phone numbers on the same line or spreadsheet row.

Each address appears once. Your own addresses, drafts, chats, Spam and Trash
are left out. Access to Gmail is read-only.

## Setup

1. Install [Node.js](https://nodejs.org/) 22 or newer.
2. In this folder, run:

   ```bash
   npm install
   ```

3. In [Google Cloud Console](https://console.cloud.google.com/):
   enable the **Gmail API**, create an **OAuth client ID** of type
   **Desktop app**, download it, and save it in this folder as
   `credentials.json`.

## Running

```bash
npm start
```

The first time, the tool prints a link. Open it, sign in, and approve
read-only access. Your browser then goes to a page that may not load: copy
the full address from the address bar and paste it into the terminal. The
sign-in is saved in `token.json` for next time.

Press **Ctrl+C** once to stop early. Emails already read are saved to
`extracted_emails.partial.xlsx`, and the previous `extracted_emails.xlsx`
is left unchanged.

### Options

Pass options after `--` when using `npm start`, for example
`npm start -- --query "after:2026/01/01"`.

| Option | What it does |
| --- | --- |
| `--query "<search>"` | Only read emails matching a Gmail search, e.g. `"after:2026/01/01"`, `"label:partners"`, `"from:example.org"` |
| `--output <file>` | Where to save results (default `extracted_emails.xlsx` in this folder) |
| `--no-attachments` | Skip PDF, Word and Excel attachments (faster) |
| `--country <code>` | Country for phone numbers written without a country code (default `PS`) |
| `--concurrency <n>` | Emails downloaded at the same time, 1–20 (default 5) |

## Output columns

| Column | Meaning |
| --- | --- |
| `name` | Display name from the email header (blank for attachment-only contacts) |
| `email` | Email address, lower-case |
| `phone` | Valid phone numbers found, in international format, separated by `;` |
| `source` | Where the contact was found: `sender`, `recipient`, `attachment` |
| `messages` | How many emails the contact appeared in |
| `last_seen` | Date of the most recent of those emails |

## Personal data

The results file and `token.json` are sensitive:

- `extracted_emails*.xlsx` holds names, email addresses and phone numbers.
  Store it somewhere access-controlled and delete it when it is no longer
  needed. All `.xlsx` files are listed in `.gitignore` so results are not
  committed by accident.
- `token.json` gives read access to the whole mailbox. To revoke it, delete
  the file and remove the app at
  <https://myaccount.google.com/permissions>.

Attachments are read in memory and never saved to disk.

## Tests

```bash
npm test
```

const fs = require("fs");
const path = require("path");
const readline = require("readline");
const { google } = require("googleapis");
const { simpleParser } = require("mailparser");
const mammoth = require("mammoth");
const pdf = require("pdf-parse");
const xlsx = require("xlsx");

const TOKEN_PATH = "token.json";
const MAX_LISTENERS = 20; // Set a higher limit for event listeners
const EMAIL_START_DATE = "2000/01/01"; // Set this to the earliest possible date to fetch all emails
const MAX_RETRIES = 5; // Number of retries for network errors
const RETRY_DELAY = 2000; // Delay between retries in milliseconds

let emailData = []; // To store email data
let attachmentPaths = []; // To store paths of attachments to delete

// Load client secrets from a local file.
fs.readFile("credentials.json", (err, content) => {
  if (err) return console.log("Error loading client secret file:", err);
  authorize(JSON.parse(content), listLabels);
});

function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } =
    credentials.installed || credentials.web;
  if (!redirect_uris || redirect_uris.length === 0) {
    return console.error(
      "Redirect URIs are not defined in the credentials.json"
    );
  }
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );

  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/gmail.readonly"],
  });
  console.log("Authorize this app by visiting this url:", authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question("Enter the code from that page here: ", (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error("Error retrieving access token", err);
      oAuth2Client.setCredentials(token);
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log("Token stored to", TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

function listLabels(auth) {
  const gmail = google.gmail({ version: "v1", auth });
  gmail.users.labels.list({ userId: "me" }, (err, res) => {
    if (err) return console.log("The API returned an error: " + err);
    const labels = res.data.labels;
    if (!labels || labels.length === 0) {
      console.log("No labels found.");
      return;
    }
    labels.forEach((label) => {
      listMessagesFromLabel(auth, label.id);
    });
  });
}

function listMessagesFromLabel(auth, labelId) {
  const gmail = google.gmail({ version: "v1", auth });
  gmail.users.messages.list(
    {
      userId: "me",
      labelIds: [labelId],
      q: `after:${EMAIL_START_DATE.replace(/\//g, "-")}`, // Fetch emails since the start date
    },
    (err, res) => {
      if (err) return console.log("The API returned an error: " + err);
      const messages = res.data.messages;
      if (!messages || messages.length === 0) {
        console.log(`No messages found in label ${labelId}.`);
        return;
      }
      console.log(`Messages in label ${labelId}:`);
      messages.forEach((message) => {
        if (labelId === "SENT") {
          getMessage(auth, message.id, "to");
        } else {
          getMessage(auth, message.id, "from");
        }
      });
    }
  );
}

async function getMessage(auth, messageId, emailType, retries = MAX_RETRIES) {
  const gmail = google.gmail({ version: "v1", auth });
  try {
    const res = await gmail.users.messages.get({
      userId: "me",
      id: messageId,
      format: "raw", // Ensure the raw format is requested
    });
    const msg = res.data;
    if (!msg.raw) {
      console.log("No raw data found for message:", messageId);
      return;
    }

    try {
      const parsed = await simpleParser(Buffer.from(msg.raw, "base64"));

      let emails = [];
      if (emailType === "from") {
        const from = parsed.from.value[0];
        emails.push({ name: from.name, email: from.address });
        console.log("From:", from);
      } else if (emailType === "to") {
        const to = parsed.to.value.map((to) => ({
          name: to.name,
          email: to.address,
        }));
        emails = emails.concat(to);
        console.log("To:", to.join(", "));
      }

      // Extract phone numbers
      const phoneNumbers = extractPhoneNumbers(parsed.text);

      // Store the emails in a data array
      storeEmailsToData(emails, phoneNumbers);

      if (parsed.attachments && parsed.attachments.length > 0) {
        for (const attachment of parsed.attachments) {
          if (attachment.filename) {
            const filePath = saveAttachment(attachment);
            const emailsInAttachment = await parseAttachments(filePath);
            const phoneNumbersInAttachment =
              extractPhoneNumbers(emailsInAttachment);
            console.log("Emails in attachment:", emailsInAttachment.join(", "));
            storeEmailsToData(emailsInAttachment, phoneNumbersInAttachment);
            attachmentPaths.push(filePath); // Collect the file path for later deletion
          }
        }
      }
    } catch (parseError) {
      console.error("Error parsing email:", parseError);
    }
  } catch (err) {
    console.error("The API returned an error:", err);
    if (retries > 0) {
      console.log(`Retrying... (${MAX_RETRIES - retries + 1})`);
      setTimeout(
        () => getMessage(auth, messageId, emailType, retries - 1),
        RETRY_DELAY
      ); // Retry after delay
    }
  }
}

function saveAttachment(attachment) {
  const filePath = path.join(__dirname, "attachments", attachment.filename);
  fs.writeFileSync(filePath, attachment.content);
  return filePath;
}

async function parseDOCX(filePath) {
  try {
    const result = await mammoth.extractRawText({ path: filePath });
    return extractEmailsFromText(result.value);
  } catch (err) {
    console.error("Error parsing DOCX:", err);
    return [];
  }
}

async function parsePDF(filePath) {
  try {
    const data = await pdf(fs.readFileSync(filePath));
    return extractEmailsFromText(data.text);
  } catch (err) {
    console.error("Error parsing PDF:", err);
    return [];
  }
}

async function parseXLSX(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    let emails = [];
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
      sheetData.forEach((row) => {
        row.forEach((cell) => {
          if (typeof cell === "string") {
            emails = emails.concat(extractEmailsFromText(cell));
          }
        });
      });
    });
    return emails;
  } catch (err) {
    console.error("Error parsing XLSX:", err);
    return [];
  }
}

function extractEmailsFromText(text) {
  const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi;
  return text.match(emailRegex) || [];
}

function extractPhoneNumbers(text) {
  const phoneRegex =
    /(\+?\d{1,4}[\s-]?)?(\(?\d{1,4}\)?[\s-]?)?\d{1,4}[\s-]?\d{1,4}[\s-]?\d{1,9}/gi;
  return text.match(phoneRegex) || [];
}

async function parseAttachments(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  let emails = [];
  switch (extension) {
    case ".docx":
      emails = await parseDOCX(filePath);
      break;
    case ".pdf":
      emails = await parsePDF(filePath);
      break;
    case ".xlsx":
      emails = await parseXLSX(filePath);
      break;
  }
  return emails;
}

function storeEmailsToData(emails, phoneNumbers) {
  emails.forEach((email) => {
    emailData.push({
      name: email.name || "",
      email: email.email,
      phone: phoneNumbers.join(", ") || "",
    });
  });
}

function storeEmailsToExcel() {
  const worksheet = xlsx.utils.json_to_sheet(emailData);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Emails");
  xlsx.writeFile(workbook, "extracted_emails.xlsx");
}

function cleanUpAttachments() {
  attachmentPaths.forEach((filePath) => {
    fs.unlink(filePath, (err) => {
      if (err) {
        console.error(`Error deleting attachment ${filePath}:`, err);
      } else {
        console.log(`Attachment ${filePath} deleted.`);
      }
    });
  });
}

// Increase the max number of listeners to avoid the warning
require("events").EventEmitter.defaultMaxListeners = MAX_LISTENERS;

// Call these functions at the end of processing
process.on("exit", () => {
  storeEmailsToExcel();
  cleanUpAttachments();
});

// code below is working fine and it exctract email only

// const fs = require("fs");
// const path = require("path");
// const readline = require("readline");
// const { google } = require("googleapis");
// const { simpleParser } = require("mailparser");
// const mammoth = require("mammoth");
// const pdf = require("pdf-parse");
// const xlsx = require("xlsx");

// const TOKEN_PATH = "token.json";
// const MAX_LISTENERS = 20; // Set a higher limit for event listeners
// const EMAIL_START_DATE = "2000/01/01"; // Set this to the earliest possible date to fetch all emails
// const MAX_RETRIES = 5; // Number of retries for network errors
// const RETRY_DELAY = 2000; // Delay between retries in milliseconds

// let emailSet = new Set(); // To store and filter duplicate emails

// // Load client secrets from a local file.
// fs.readFile("credentials.json", (err, content) => {
//   if (err) return console.log("Error loading client secret file:", err);
//   authorize(JSON.parse(content), listLabels);
// });

// function authorize(credentials, callback) {
//   const { client_secret, client_id, redirect_uris } =
//     credentials.installed || credentials.web;
//   if (!redirect_uris || redirect_uris.length === 0) {
//     return console.error(
//       "Redirect URIs are not defined in the credentials.json"
//     );
//   }
//   const oAuth2Client = new google.auth.OAuth2(
//     client_id,
//     client_secret,
//     redirect_uris[0]
//   );

//   fs.readFile(TOKEN_PATH, (err, token) => {
//     if (err) return getNewToken(oAuth2Client, callback);
//     oAuth2Client.setCredentials(JSON.parse(token));
//     callback(oAuth2Client);
//   });
// }

// function getNewToken(oAuth2Client, callback) {
//   const authUrl = oAuth2Client.generateAuthUrl({
//     access_type: "offline",
//     scope: ["https://www.googleapis.com/auth/gmail.readonly"],
//   });
//   console.log("Authorize this app by visiting this url:", authUrl);
//   const rl = readline.createInterface({
//     input: process.stdin,
//     output: process.stdout,
//   });
//   rl.question("Enter the code from that page here: ", (code) => {
//     rl.close();
//     oAuth2Client.getToken(code, (err, token) => {
//       if (err) return console.error("Error retrieving access token", err);
//       oAuth2Client.setCredentials(token);
//       fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
//         if (err) return console.error(err);
//         console.log("Token stored to", TOKEN_PATH);
//       });
//       callback(oAuth2Client);
//     });
//   });
// }

// function listLabels(auth) {
//   const gmail = google.gmail({ version: "v1", auth });
//   gmail.users.labels.list({ userId: "me" }, (err, res) => {
//     if (err) return console.log("The API returned an error: " + err);
//     const labels = res.data.labels;
//     if (!labels || labels.length === 0) {
//       console.log("No labels found.");
//       return;
//     }
//     labels.forEach((label) => {
//       listMessagesFromLabel(auth, label.id);
//     });
//   });
// }

// function listMessagesFromLabel(auth, labelId) {
//   const gmail = google.gmail({ version: "v1", auth });
//   gmail.users.messages.list(
//     {
//       userId: "me",
//       labelIds: [labelId],
//       q: `after:${EMAIL_START_DATE.replace(/\//g, "-")}`, // Fetch emails since the start date
//     },
//     (err, res) => {
//       if (err) return console.log("The API returned an error: " + err);
//       const messages = res.data.messages;
//       if (!messages || messages.length === 0) {
//         console.log(`No messages found in label ${labelId}.`);
//         return;
//       }
//       console.log(`Messages in label ${labelId}:`);
//       messages.forEach((message) => {
//         if (labelId === "SENT") {
//           getMessage(auth, message.id, "to");
//         } else {
//           getMessage(auth, message.id, "from");
//         }
//       });
//     }
//   );
// }

// async function getMessage(auth, messageId, emailType, retries = MAX_RETRIES) {
//   const gmail = google.gmail({ version: "v1", auth });
//   try {
//     const res = await gmail.users.messages.get({
//       userId: "me",
//       id: messageId,
//       format: "raw", // Ensure the raw format is requested
//     });
//     const msg = res.data;
//     if (!msg.raw) {
//       console.log("No raw data found for message:", messageId);
//       return;
//     }

//     try {
//       const parsed = await simpleParser(Buffer.from(msg.raw, "base64"));

//       let emails = [];
//       if (emailType === "from") {
//         const from = parsed.from.value[0].address;
//         emails.push(from);
//         console.log("From:", from);
//       } else if (emailType === "to") {
//         const to = parsed.to.value.map((to) => to.address);
//         emails = emails.concat(to);
//         console.log("To:", to.join(", "));
//       }

//       // Store the emails in a file
//       storeEmailsToFile(emails);

//       if (parsed.attachments && parsed.attachments.length > 0) {
//         for (const attachment of parsed.attachments) {
//           if (attachment.filename) {
//             const filePath = saveAttachment(attachment);
//             const emailsInAttachment = await parseAttachments(filePath);
//             console.log("Emails in attachment:", emailsInAttachment.join(", "));
//             storeEmailsToFile(emailsInAttachment);
//             cleanUpAttachment(filePath); // Clean up the attachment after processing
//           }
//         }
//       }
//     } catch (parseError) {
//       console.error("Error parsing email:", parseError);
//     }
//   } catch (err) {
//     console.error("The API returned an error:", err);
//     if (retries > 0) {
//       console.log(`Retrying... (${MAX_RETRIES - retries + 1})`);
//       setTimeout(
//         () => getMessage(auth, messageId, emailType, retries - 1),
//         RETRY_DELAY
//       ); // Retry after delay
//     }
//   }
// }

// function saveAttachment(attachment) {
//   const filePath = path.join(__dirname, "attachments", attachment.filename);
//   fs.writeFileSync(filePath, attachment.content);
//   return filePath;
// }

// function cleanUpAttachment(filePath) {
//   fs.unlink(filePath, (err) => {
//     if (err) {
//       console.error(`Error deleting attachment ${filePath}:`, err);
//     } else {
//       console.log(`Attachment ${filePath} deleted.`);
//     }
//   });
// }

// async function parseDOCX(filePath) {
//   try {
//     const result = await mammoth.extractRawText({ path: filePath });
//     return extractEmailsFromText(result.value);
//   } catch (err) {
//     console.error("Error parsing DOCX:", err);
//     return [];
//   }
// }

// async function parsePDF(filePath) {
//   try {
//     const data = await pdf(fs.readFileSync(filePath));
//     return extractEmailsFromText(data.text);
//   } catch (err) {
//     console.error("Error parsing PDF:", err);
//     return [];
//   }
// }

// async function parseXLSX(filePath) {
//   try {
//     const workbook = xlsx.readFile(filePath);
//     let emails = [];
//     workbook.SheetNames.forEach((sheetName) => {
//       const worksheet = workbook.Sheets[sheetName];
//       const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
//       sheetData.forEach((row) => {
//         row.forEach((cell) => {
//           if (typeof cell === "string") {
//             emails = emails.concat(extractEmailsFromText(cell));
//           }
//         });
//       });
//     });
//     return emails;
//   } catch (err) {
//     console.error("Error parsing XLSX:", err);
//     return [];
//   }
// }

// function extractEmailsFromText(text) {
//   const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi;
//   return text.match(emailRegex) || [];
// }

// async function parseAttachments(filePath) {
//   const extension = path.extname(filePath).toLowerCase();
//   let emails = [];
//   switch (extension) {
//     case ".docx":
//       emails = await parseDOCX(filePath);
//       break;
//     case ".pdf":
//       emails = await parsePDF(filePath);
//       break;
//     case ".xlsx":
//       emails = await parseXLSX(filePath);
//       break;
//   }
//   cleanUpAttachment(filePath); // Clean up the attachment after parsing
//   return emails;
// }

// function storeEmailsToFile(emails) {
//   emails.forEach((email) => {
//     emailSet.add(email);
//   });
//   const filePath = path.join(__dirname, "extracted_emails.txt");
//   fs.writeFileSync(filePath, Array.from(emailSet).join("\n"), "utf8");
// }

// // Increase the max number of listeners to avoid the warning
// require("events").EventEmitter.defaultMaxListeners = MAX_LISTENERS;

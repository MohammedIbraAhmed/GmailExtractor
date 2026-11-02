"use strict";

// Builds test emails, attachments and a fake Gmail client in memory.

const zlib = require("zlib");
const ExcelJS = require("exceljs");

/** A one-page PDF containing `text` (ASCII only). */
function makePdf(text) {
  const stream = `BT /F1 12 Tf 72 720 Td (${text}) Tj ET`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let out = "%PDF-1.4\n";
  const offsets = objects.map((obj, i) => {
    const offset = out.length;
    out += `${i + 1} 0 obj\n${obj}\nendobj\n`;
    return offset;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  out += offsets.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("");
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}

/** A minimal uncompressed zip file. `files` maps names to strings. */
function makeZip(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBuf = Buffer.from(name);
    const data = Buffer.from(content);
    const crc = zlib.crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, data);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + data.length;
  }
  const centralDir = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(centralDir.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, centralDir, end]);
}

/** A minimal Word document with one paragraph per entry in `paragraphs`. */
function makeDocx(paragraphs) {
  const xmlEscape = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const body = paragraphs.map((p) => `<w:p><w:r><w:t>${xmlEscape(p)}</w:t></w:r></w:p>`).join("");
  return makeZip({
    "[Content_Types].xml":
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    "_rels/.rels":
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/document.xml":
      '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:body>${body}</w:body></w:document>`,
  });
}

/** An Excel file with one sheet; `rows` is an array of arrays. */
async function makeXlsx(rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet("Sheet1").addRows(rows);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

/** A raw email (RFC 822). `headers` is an object; attachments are { filename, contentType, content }. */
function rawEmail({ headers = {}, text = "", attachments = [] }) {
  const head = Object.entries({ "MIME-Version": "1.0", Subject: "test", ...headers })
    .map(([k, v]) => `${k}: ${v}`)
    .join("\r\n");
  if (attachments.length === 0) {
    return Buffer.from(`${head}\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n${text}\r\n`);
  }
  const boundary = "TESTBOUNDARY";
  const parts = [`--${boundary}\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n${text}\r\n`];
  for (const a of attachments) {
    const encoded = a.content.toString("base64").replace(/.{76}/g, "$&\r\n");
    parts.push(
      `--${boundary}\r\nContent-Type: ${a.contentType || "application/octet-stream"}\r\n` +
        (a.filename ? `Content-Disposition: attachment; filename="${a.filename}"\r\n` : "Content-Disposition: attachment\r\n") +
        `Content-Transfer-Encoding: base64\r\n\r\n${encoded}\r\n`
    );
  }
  return Buffer.from(
    `${head}\r\nContent-Type: multipart/mixed; boundary="${boundary}"\r\n\r\n${parts.join("")}--${boundary}--\r\n`
  );
}

function httpError(status, reason) {
  const data = { error: { code: status, errors: reason ? [{ reason }] : [] } };
  return Object.assign(new Error(`HTTP ${status}`), { status, response: { status, data } });
}

/**
 * A fake Gmail client. `pages` is an array of arrays of message IDs;
 * `messages` maps an ID to { raw: Buffer, labelIds } or to an Error to throw.
 */
function fakeGmail({ pages = [], messages = {} }) {
  return {
    users: {
      messages: {
        list: async ({ pageToken }) => {
          const i = pageToken ? Number(pageToken) : 0;
          return {
            data: {
              messages: (pages[i] || []).map((id) => ({ id })),
              nextPageToken: i + 1 < pages.length ? String(i + 1) : undefined,
            },
          };
        },
        get: async ({ id }) => {
          const m = messages[id];
          if (m instanceof Error) throw m;
          return { data: { raw: m.raw.toString("base64url"), labelIds: m.labelIds || ["INBOX"] } };
        },
      },
    },
  };
}

module.exports = { makePdf, makeDocx, makeXlsx, rawEmail, httpError, fakeGmail };

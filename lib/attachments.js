"use strict";

// Reads text out of PDF, Word (.docx) and Excel (.xlsx) attachments.
// Everything happens in memory: attachments are never written to disk,
// so a sender cannot use a crafted file name to write files on this
// computer, and no copies of attachments are left behind.

const path = require("path");
const mammoth = require("mammoth");
const ExcelJS = require("exceljs");
const { PDFParse } = require("pdf-parse");

// Gmail does not accept attachments above 25 MB; anything bigger is skipped.
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;

const TYPES_BY_EXTENSION = { ".pdf": "pdf", ".docx": "docx", ".xlsx": "xlsx" };
const TYPES_BY_CONTENT_TYPE = {
  "application/pdf": "pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
};

/** Returns "pdf", "docx", "xlsx", or null if the attachment is not supported. */
function attachmentType(attachment) {
  const ext = path.extname(String(attachment.filename || "")).toLowerCase();
  const contentType = String(attachment.contentType || "").toLowerCase();
  return TYPES_BY_EXTENSION[ext] || TYPES_BY_CONTENT_TYPE[contentType] || null;
}

async function pdfLines(buffer) {
  // Copy the bytes: the PDF reader may take ownership of the array it is given.
  const parser = new PDFParse({ data: new Uint8Array(buffer) });
  try {
    const result = await parser.getText();
    return result.text.split(/\r?\n/);
  } finally {
    await parser.destroy();
  }
}

async function docxLines(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  return result.value.split(/\r?\n/);
}

async function xlsxLines(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const lines = [];
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      const cells = [];
      row.eachCell({ includeEmpty: false }, (cell) => {
        const text = cell.text;
        if (text) cells.push(text);
        const link = typeof cell.hyperlink === "string" ? cell.hyperlink : "";
        if (link.toLowerCase().startsWith("mailto:")) cells.push(link.slice(7));
      });
      // One line per row keeps an email next to the phone number in the same row.
      if (cells.length) lines.push(cells.join(" | "));
    });
  });
  return lines;
}

const READERS = { pdf: pdfLines, docx: docxLines, xlsx: xlsxLines };

/**
 * Returns the attachment's text as an array of lines, or null when the
 * attachment type is not supported or it is too large.
 * Throws if a supported file cannot be read (e.g. it is damaged).
 */
async function attachmentLines(attachment) {
  const type = attachmentType(attachment);
  if (!type || !attachment.content) return null;
  if (attachment.content.length > MAX_ATTACHMENT_BYTES) return null;
  return READERS[type](attachment.content);
}

module.exports = { attachmentType, attachmentLines, MAX_ATTACHMENT_BYTES };

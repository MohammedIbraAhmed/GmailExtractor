"use strict";

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const COLUMNS = [
  { header: "name", key: "name", width: 30 },
  { header: "email", key: "email", width: 36 },
  { header: "phone", key: "phone", width: 28 },
  { header: "source", key: "source", width: 28 },
  { header: "messages", key: "messages", width: 10 },
  { header: "last_seen", key: "last_seen", width: 12 },
];

/** "results.xlsx" -> "results.partial.xlsx" */
function partialPathFor(filePath) {
  const ext = path.extname(filePath);
  return filePath.slice(0, filePath.length - ext.length) + ".partial" + (ext || ".xlsx");
}

/**
 * Writes the contact rows to an Excel file. The file is written under a
 * temporary name first and then renamed, so an existing file is never left
 * half-written. Throws an error naming the temporary file if the rename fails
 * (for example because the file is open in Excel).
 */
async function writeContactsWorkbook(rows, filePath) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Contacts", { views: [{ state: "frozen", ySplit: 1 }] });
  sheet.columns = COLUMNS;
  sheet.getRow(1).font = { bold: true };
  sheet.autoFilter = { from: "A1", to: `${String.fromCharCode(64 + COLUMNS.length)}1` };
  sheet.addRows(rows);

  const tmpPath = `${filePath}.tmp`;
  await workbook.xlsx.writeFile(tmpPath);
  try {
    fs.chmodSync(tmpPath, 0o600); // the file holds personal data
  } catch {
    // Not supported on every system (e.g. Windows).
  }
  try {
    fs.renameSync(tmpPath, filePath);
  } catch (err) {
    throw new Error(
      `Could not replace ${filePath} (${err.code || err.message}). ` +
        `If it is open in Excel, close it. The results are saved in ${tmpPath}.`
    );
  }
}

module.exports = { writeContactsWorkbook, partialPathFor };

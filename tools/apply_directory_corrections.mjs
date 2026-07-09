import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = "C:/Users/so.a/Telephone_Directory";
const inputPath = `${root}/uploads/WR_DB_Ready_Final (1).xlsx`;
const outputDir = `${root}/outputs/directory_correction`;
const correctionsPath = `${outputDir}/corrections.json`;
const outputPath = `${outputDir}/WR_DB_Ready_Final_corrected_from_Word.xlsx`;

function colName(index) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function writeSheet(sheet, rows) {
  if (!rows.length) return;
  const used = sheet.getUsedRange();
  if (used) {
    used.clear({ applyTo: "contents" });
  }
  const lastCol = colName(rows[0].length - 1);
  const range = sheet.getRange(`A1:${lastCol}${rows.length}`);
  range.values = rows;
}

function styleAudit(sheet, rowCount, colCount) {
  const lastCol = colName(colCount - 1);
  sheet.getRange(`A1:${lastCol}1`).format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:${lastCol}${rowCount}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#D9E2F3",
  };
  sheet.getRange(`A1:${lastCol}${rowCount}`).format.wrapText = true;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  sheet.getRange("A:J").format.autofitColumns();
}

await fs.mkdir(outputDir, { recursive: true });
const corrections = JSON.parse(await fs.readFile(correctionsPath, "utf8"));

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

writeSheet(
  workbook.worksheets.getItem("employees"),
  [corrections.employeesHeader, ...corrections.employees],
);
writeSheet(
  workbook.worksheets.getItem("email_addresses"),
  [corrections.emailHeader, ...corrections.emails],
);
writeSheet(
  workbook.worksheets.getItem("phone_numbers"),
  [corrections.phoneHeader, ...corrections.phones],
);

const auditSheet = workbook.worksheets.getOrAdd("correction_audit");
writeSheet(auditSheet, corrections.audit);
styleAudit(auditSheet, corrections.audit.length, corrections.audit[0].length);

const summarySheet = workbook.worksheets.getOrAdd("correction_summary");
const summaryRows = [
  ["Metric", "Value"],
  ["Word records extracted", corrections.summary.wordRecords],
  ["Employee rows corrected", corrections.summary.employees],
  ["Email rows rebuilt", corrections.summary.emails],
  ["Phone rows rebuilt", corrections.summary.phones],
  ["Unique name matches", corrections.summary.matchStatusCounts["unique-name"] ?? 0],
  ["Name occurrence matches", corrections.summary.matchStatusCounts["name-occurrence"] ?? 0],
  ["Designation occurrence matches", corrections.summary.matchStatusCounts["designation-occurrence"] ?? 0],
  ["Partial occurrence matches", corrections.summary.matchStatusCounts["partial-occurrence"] ?? 0],
  ["Name and designation matches", corrections.summary.matchStatusCounts["name-and-designation"] ?? 0],
];
writeSheet(summarySheet, summaryRows);
styleAudit(summarySheet, summaryRows.length, summaryRows[0].length);

const employeesCheck = await workbook.inspect({
  kind: "region",
  sheetId: "employees",
  range: "A1:E12",
  maxChars: 2500,
});
console.log(employeesCheck.ndjson);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errorScan.ndjson);

const previewRanges = {
  employees: "A1:E40",
  phone_numbers: "A1:D45",
  email_addresses: "A1:C45",
  correction_summary: "A1:B10",
  correction_audit: "A1:J45",
};

for (const [sheetName, range] of Object.entries(previewRanges)) {
  const preview = await workbook.render({
    sheetName,
    range,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, `${sheetName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);

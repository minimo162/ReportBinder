import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function getArg(name, fallback = null) {
  const index = process.argv.indexOf(name);
  return index >= 0 && index + 1 < process.argv.length ? process.argv[index + 1] : fallback;
}

const outputDir = path.resolve(getArg("--output", "."));
const workbookCount = Number(getArg("--count", "3"));
const sheetCount = Number(getArg("--sheets", "8"));
const rowCount = Number(getArg("--rows", "45"));
const updateOnly = process.argv.includes("--update-only");
await fs.mkdir(outputDir, { recursive: true });

function styleSheet(sheet, workbookIndex, sheetIndex) {
  sheet.showGridLines = false;
  sheet.getRange("A1:J2").merge();
  sheet.getRange("A1:J2").values = [[`Scale benchmark workbook ${workbookIndex.toString().padStart(2, "0")} / Department ${sheetIndex.toString().padStart(2, "0")}`]];
  sheet.getRange("A1:J2").format = { fill: "#263A5A", font: { bold: true, color: "#FFFFFF", size: 15 }, rowHeight: 28 };
  sheet.getRange("A4:J4").values = [["ID", "Department", "Owner", "Budget", "Actual", "Variance", "Variance %", "Priority", "Status", "Notes"]];
  sheet.getRange("A4:J4").format = { fill: "#DCE6F1", font: { bold: true, color: "#1F2937" }, borders: { preset: "outside", style: "medium", color: "#708090" } };
  const values = Array.from({ length: rowCount }, (_, row) => [
    `${workbookIndex}-${sheetIndex}-${row + 1}`,
    `Department ${(row % 12) + 1}`,
    `Owner ${(row % 15) + 1}`,
    600000 + workbookIndex * 25000 + sheetIndex * 8000 + row * 1500,
    580000 + workbookIndex * 27000 + sheetIndex * 9000 + row * 1675,
    null,
    null,
    ["High", "Medium", "Low"][row % 3],
    null,
    row % 7 === 0 ? "Printed approval and supporting evidence required." : "",
  ]);
  sheet.getRange(`A5:J${rowCount + 4}`).values = values;
  sheet.getRange("F5").formulas = [["=D5-E5"]];
  sheet.getRange(`F5:F${rowCount + 4}`).fillDown();
  sheet.getRange("G5").formulas = [["=F5/D5"]];
  sheet.getRange(`G5:G${rowCount + 4}`).fillDown();
  sheet.getRange("I5").formulas = [["=IF(F5>=0,\"On plan\",\"Review\")"]];
  sheet.getRange(`I5:I${rowCount + 4}`).fillDown();
  sheet.getRange(`A5:J${rowCount + 4}`).format.borders = { preset: "inside", style: "thin", color: "#D4DCE6" };
  sheet.getRange(`D5:F${rowCount + 4}`).format.numberFormat = "#,##0";
  sheet.getRange(`G5:G${rowCount + 4}`).format.numberFormat = "0.0%";
  sheet.getRange("A:A").format.columnWidth = 12;
  sheet.getRange("B:B").format.columnWidth = 18;
  sheet.getRange("C:C").format.columnWidth = 16;
  sheet.getRange("D:G").format.columnWidth = 14;
  sheet.getRange("H:I").format.columnWidth = 12;
  sheet.getRange("J:J").format.columnWidth = 38;
  sheet.getRange(`J5:J${rowCount + 4}`).format.wrapText = true;
  sheet.freezePanes.freezeRows(4);
  sheet.freezePanes.freezeColumns(1);
}

async function verifyAndExport(workbook, workbookPath, previewPath = null) {
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "formula error scan",
  });
  if (errors.ndjson.includes('"matchCount":') && !errors.ndjson.includes('"matchCount":0')) throw new Error(errors.ndjson);
  if (previewPath) {
    const preview = await workbook.render({ sheetName: "Dept01", range: `A1:J${Math.min(rowCount + 4, 24)}`, scale: 1.2, format: "png" });
    await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(workbookPath);
}

if (updateOnly) {
  const workbookPath = path.join(outputDir, "large-excel-01.xlsx");
  const input = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  workbook.worksheets.getItem("Dept04").getRange("E20").values = [[777777]];
  await verifyAndExport(workbook, workbookPath, path.join(outputDir, "large-excel-01-v2-preview.png"));
  console.log(JSON.stringify({ updated: workbookPath, changedCell: "Dept04!E20" }));
} else {
  for (let workbookIndex = 1; workbookIndex <= workbookCount; workbookIndex += 1) {
    const workbook = Workbook.create();
    for (let sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex += 1) {
      const sheet = workbook.worksheets.add(`Dept${sheetIndex.toString().padStart(2, "0")}`);
      styleSheet(sheet, workbookIndex, sheetIndex);
    }
    const workbookPath = path.join(outputDir, `large-excel-${workbookIndex.toString().padStart(2, "0")}.xlsx`);
    const previewPath = workbookIndex === 1 ? path.join(outputDir, "large-excel-01-v1-preview.png") : null;
    await verifyAndExport(workbook, workbookPath, previewPath);
  }
  console.log(JSON.stringify({ outputDir, workbookCount, sheetCount, rowCount }));
}

// artifact-tool may set a non-zero process exit code after spilling a large inspect
// payload to an .inspect.ndjson file. Reaching this line means verification and all
// requested exports completed successfully.
process.exitCode = 0;

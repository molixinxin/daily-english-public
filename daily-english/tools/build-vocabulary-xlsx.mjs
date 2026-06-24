import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const root = path.resolve(__dirname, "..");
const dataPath = path.join(root, "data", "vocabulary.json");
const outputPath = path.join(root, "data", "advanced_english_vocabulary.xlsx");
const previewPath = path.join(root, "data", "advanced_english_vocabulary_preview.png");

const words = JSON.parse(await fs.readFile(dataPath, "utf8"));

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Vocabulary");
sheet.showGridLines = false;

sheet.getRange("A1:E1").values = [["日期", "单词", "音标", "解释", "例句"]];
sheet.getRange("A1:E1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
};

const rows = words.map((item) => [
  item.date,
  item.word,
  item.phonetic,
  item.definition,
  item.example,
]);

if (rows.length > 0) {
  const range = sheet.getRangeByIndexes(1, 0, rows.length, 5);
  range.values = rows;
  range.format = {
    borders: { preset: "inside", style: "thin", color: "#E5E7EB" },
    wrapText: true,
  };
}

sheet.getRange("A:A").format.columnWidth = 13;
sheet.getRange("B:B").format.columnWidth = 18;
sheet.getRange("C:C").format.columnWidth = 18;
sheet.getRange("D:D").format.columnWidth = 42;
sheet.getRange("E:E").format.columnWidth = 58;
sheet.getRange("A:E").format.font = { name: "Aptos", size: 11 };
sheet.getRange("A1:E1").format.font = { name: "Aptos", size: 11, bold: true, color: "#FFFFFF" };
sheet.freezePanes.freezeRows(1);

const tableRange = `A1:E${rows.length + 1}`;
const table = sheet.tables.add(tableRange, true, "VocabularyTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;

const inspect = await workbook.inspect({
  kind: "table",
  range: tableRange,
  tableMaxRows: Math.min(rows.length + 1, 15),
  tableMaxCols: 5,
  maxChars: 3000,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Vocabulary",
  range: `A1:E${Math.min(rows.length + 1, 13)}`,
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(`Saved ${outputPath}`);

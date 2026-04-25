import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const outputDir = path.join(__dirname, "outputs", "fao-study-tracker");
const workbook = Workbook.create();

const trackerSheet = workbook.worksheets.add("Daily Tracker");
const summarySheet = workbook.worksheets.add("Weekly Summary");
const phasesSheet = workbook.worksheets.add("Study Phases");
const reviewSheet = workbook.worksheets.add("Weekly Review");
const syllabusSheet = workbook.worksheets.add("Topic Checklist");
const testsSheet = workbook.worksheets.add("Mock Test Log");

trackerSheet.showGridLines = false;
summarySheet.showGridLines = false;
phasesSheet.showGridLines = false;
reviewSheet.showGridLines = false;
syllabusSheet.showGridLines = false;
testsSheet.showGridLines = false;

const tick = "\u2714";
const cross = "\u2718";
const tasks = [
  "Gym",
  "Maths/Stats (2 hrs)",
  "Commerce",
  "Reasoning (20 Q)",
  "Polity/Eco",
  "Current Affairs",
  "Revision",
  "Coaching",
];

const trackerHeaders = [
  "Day",
  "Date",
  ...tasks,
  "Daily Score",
  "Status",
  "Motivation",
];

trackerSheet.getRange("A1:M1").merge();
trackerSheet.getRange("A1").values = [["FAO Study Tracker"]];
trackerSheet.getRange("A2:M2").merge();
trackerSheet.getRange("A2").values = [[
  `Day-wise checklist with ${tick}/${cross} tracking. Each ${tick} is worth 10 points, for a total of 80 per day.`,
]];
trackerSheet.getRange("A4:M4").values = [trackerHeaders];

const trackerFormulaRows = [];
for (let row = 5; row <= 11; row += 1) {
  const previousDateRow = row - 1;
  trackerFormulaRows.push([
    null,
    row === 5 ? "=TODAY()-6" : `=B${previousDateRow}+1`,
    ...Array(tasks.length).fill(null),
    `=COUNTIF(C${row}:J${row},"${tick}")*10`,
    `=IF(K${row}>=70,"Topper Mode",IF(K${row}>=50,"Good","Improve"))`,
    `=IF(COUNTIF(C${row}:J${row},"${tick}")=8,"Excellent work - all tasks completed!","Keep going - every task counts.")`,
  ]);
}

trackerSheet.getRange("A5:A11").values = Array.from({ length: 7 }, (_, index) => [index + 1]);
trackerSheet.getRange("A5:M11").formulas = trackerFormulaRows;
trackerSheet.getRange("C5:J11").values = Array.from({ length: 7 }, () => Array(tasks.length).fill(cross));
trackerSheet.getRange("B5:B11").setNumberFormat("dd-mmm-yyyy");

trackerSheet.getRange("A1:M1").format = {
  fill: "#1F6B52",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
trackerSheet.getRange("A2:M2").format = {
  fill: "#D8EFE5",
  font: { color: "#1F2A21", italic: true },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
trackerSheet.getRange("A4:M4").format = {
  fill: "#E8DCC6",
  font: { bold: true, color: "#1F2A21" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
trackerSheet.getRange("A5:M11").format = {
  fill: "#FFFCF6",
  verticalAlignment: "center",
  wrapText: true,
};
trackerSheet.getRange("C5:J11").dataValidation = {
  rule: {
    type: "list",
    values: [tick, cross],
  },
};
trackerSheet.getRange("C5:J11").conditionalFormats.addCustom(`=C5="${tick}"`, {
  fill: "#D7F0E6",
  font: { color: "#176147", bold: true },
});
trackerSheet.getRange("C5:J11").conditionalFormats.addCustom(`=C5="${cross}"`, {
  fill: "#F8DDD7",
  font: { color: "#9C4334", bold: true },
});
trackerSheet.getRange("K5:K11").conditionalFormats.addDataBar({
  color: "#1F6B52",
  gradient: true,
});
trackerSheet.getRange("A1:M11").format.borders = {
  top: { style: "thin", color: "#D8CDBB" },
  bottom: { style: "thin", color: "#D8CDBB" },
  left: { style: "thin", color: "#D8CDBB" },
  right: { style: "thin", color: "#D8CDBB" },
};
trackerSheet.getRange("A:M").format.columnWidthPx = 108;
trackerSheet.getRange("A:A").format.columnWidthPx = 70;
trackerSheet.getRange("B:B").format.columnWidthPx = 110;
trackerSheet.getRange("K:K").format.columnWidthPx = 95;
trackerSheet.getRange("L:L").format.columnWidthPx = 120;
trackerSheet.getRange("M:M").format.columnWidthPx = 240;
trackerSheet.getRange("A1").format.rowHeightPx = 34;
trackerSheet.getRange("A2").format.rowHeightPx = 36;
trackerSheet.freezePanes.freezeRows(4);
trackerSheet.freezePanes.freezeColumns(2);

summarySheet.getRange("A1:F1").merge();
summarySheet.getRange("A1").values = [["Weekly Summary"]];
summarySheet.getRange("A2:F2").merge();
summarySheet.getRange("A2").values = [[
  "Scores and completion counts update automatically from the Daily Tracker sheet.",
]];
summarySheet.getRange("A4:F4").values = [[
  "Day",
  "Date",
  "Score / 80",
  "Status",
  "Completed Tasks",
  "Week Average",
]];

const summaryFormulas = [];
for (let row = 5; row <= 11; row += 1) {
  summaryFormulas.push([
    `='Daily Tracker'!A${row}`,
    `='Daily Tracker'!B${row}`,
    `='Daily Tracker'!K${row}`,
    `='Daily Tracker'!L${row}`,
    `=COUNTIF('Daily Tracker'!C${row}:J${row},"${tick}")`,
    `=AVERAGE('Daily Tracker'!K$5:K$11)`,
  ]);
}
summarySheet.getRange("A5:F11").formulas = summaryFormulas;
summarySheet.getRange("B5:B11").setNumberFormat("dd-mmm-yyyy");
summarySheet.getRange("A1:F1").format = {
  fill: "#1F2A21",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
summarySheet.getRange("A2:F2").format = {
  fill: "#F6E6BE",
  font: { color: "#1F2A21", italic: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
summarySheet.getRange("A4:F4").format = {
  fill: "#E8DCC6",
  font: { bold: true, color: "#1F2A21" },
  horizontalAlignment: "center",
};
summarySheet.getRange("A5:F11").format = {
  fill: "#FFFCF6",
  verticalAlignment: "center",
};
summarySheet.getRange("C5:C11").conditionalFormats.addDataBar({
  color: "#1F6B52",
  gradient: true,
});
summarySheet.getRange("D5:D11").conditionalFormats.addCustom('=D5="Topper Mode"', {
  fill: "#D7F0E6",
  font: { color: "#176147", bold: true },
});
summarySheet.getRange("D5:D11").conditionalFormats.addCustom('=D5="Good"', {
  fill: "#F6E6BE",
  font: { color: "#8A5B0F", bold: true },
});
summarySheet.getRange("D5:D11").conditionalFormats.addCustom('=D5="Improve"', {
  fill: "#F8DDD7",
  font: { color: "#9C4334", bold: true },
});
summarySheet.getRange("A:F").format.columnWidthPx = 120;
summarySheet.getRange("F:F").format.columnWidthPx = 100;
summarySheet.freezePanes.freezeRows(4);

const chart = summarySheet.charts.add("line", summarySheet.getRange("A4:C11"));
chart.title = "Weekly Score Trend";
chart.hasLegend = false;
chart.xAxis = { axisType: "textAxis" };
chart.yAxis = { numberFormatCode: "0" };
chart.setPosition("H4", "N20");

phasesSheet.getRange("A1:F1").merge();
phasesSheet.getRange("A1").values = [["FAO Study Phases"]];
phasesSheet.getRange("A2:F2").merge();
phasesSheet.getRange("A2").values = [[
  "Set your study start date once, then track your current phase automatically.",
]];
phasesSheet.getRange("A4:B6").values = [
  ["Study Start Date", new Date()],
  ["Current Study Day", null],
  ["Current Phase", null],
];
phasesSheet.getRange("B5").formulas = [["=TODAY()-B4+1"]];
phasesSheet.getRange("B6").formulas = [[
  '=IF(B5<=20,"Phase 1 (Foundation + Catch-up)",IF(B5<=50,"Phase 2 (Full Coverage + Practice)",IF(B5<=70,"Phase 3 (Mock Tests + Revision)","Post Plan Review")))',
]];
phasesSheet.getRange("A8:D8").values = [[
  "Phase",
  "Day Range",
  "Focus",
  "Notes",
]];
phasesSheet.getRange("A9:D11").values = [
  ["Phase 1", "Day 1-20", "Foundation + Catch-up", "Build basics and clear pending backlog."],
  ["Phase 2", "Day 21-50", "Full Coverage + Practice", "Cover the full syllabus and practice daily."],
  ["Phase 3", "Day 51-70", "Mock Tests + Revision", "Take mocks, revise weak areas, and improve speed."],
];
phasesSheet.getRange("A1:F1").format = {
  fill: "#1F6B52",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
phasesSheet.getRange("A2:F2").format = {
  fill: "#D8EFE5",
  font: { color: "#1F2A21", italic: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
phasesSheet.getRange("A4:A6").format = {
  fill: "#E8DCC6",
  font: { bold: true, color: "#1F2A21" },
};
phasesSheet.getRange("B4:B6").format = {
  fill: "#FFFCF6",
  font: { color: "#1F2A21", bold: true },
};
phasesSheet.getRange("B4").setNumberFormat("dd-mmm-yyyy");
phasesSheet.getRange("A8:D8").format = {
  fill: "#1F2A21",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
phasesSheet.getRange("A9:D11").format = {
  fill: "#FFFCF6",
  wrapText: true,
  verticalAlignment: "center",
};
phasesSheet.getRange("A9:A11").conditionalFormats.addCustom('=A9="Phase 1"', {
  fill: "#D7F0E6",
  font: { color: "#176147", bold: true },
});
phasesSheet.getRange("A9:A11").conditionalFormats.addCustom('=A9="Phase 2"', {
  fill: "#F6E6BE",
  font: { color: "#8A5B0F", bold: true },
});
phasesSheet.getRange("A9:A11").conditionalFormats.addCustom('=A9="Phase 3"', {
  fill: "#F8DDD7",
  font: { color: "#9C4334", bold: true },
});
phasesSheet.getRange("A:D").format.columnWidthPx = 180;
phasesSheet.getRange("A:A").format.columnWidthPx = 100;
phasesSheet.getRange("B:B").format.columnWidthPx = 140;
phasesSheet.getRange("C:C").format.columnWidthPx = 220;
phasesSheet.getRange("D:D").format.columnWidthPx = 260;
phasesSheet.getRange("A1:F11").format.borders = {
  top: { style: "thin", color: "#D8CDBB" },
  bottom: { style: "thin", color: "#D8CDBB" },
  left: { style: "thin", color: "#D8CDBB" },
  right: { style: "thin", color: "#D8CDBB" },
};
phasesSheet.freezePanes.freezeRows(8);

reviewSheet.getRange("A1:F1").merge();
reviewSheet.getRange("A1").values = [["Weekly Performance Review"]];
reviewSheet.getRange("A2:F2").merge();
reviewSheet.getRange("A2").values = [[
  "Use this sheet to review each week, note weak areas, and plan fixes.",
]];
reviewSheet.getRange("A4:F4").values = [[
  "Week",
  "Total Score",
  "Avg Score",
  "Mock Score",
  "Weak Area",
  "Fix Plan",
]];
reviewSheet.getRange("A5:F14").values = Array.from({ length: 10 }, (_, index) => [
  index + 1,
  null,
  null,
  null,
  "",
  "",
]);

reviewSheet.getRange("A1:F1").format = {
  fill: "#1F2A21",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
reviewSheet.getRange("A2:F2").format = {
  fill: "#F6E6BE",
  font: { color: "#1F2A21", italic: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
reviewSheet.getRange("A4:F4").format = {
  fill: "#E8DCC6",
  font: { bold: true, color: "#1F2A21" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
reviewSheet.getRange("A5:F14").format = {
  fill: "#FFFCF6",
  verticalAlignment: "center",
  wrapText: true,
};
reviewSheet.getRange("A5:A14").format = {
  fill: "#D8EFE5",
  font: { bold: true, color: "#176147" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
reviewSheet.getRange("B5:D14").conditionalFormats.addDataBar({
  color: "#1F6B52",
  gradient: true,
});
reviewSheet.getRange("A1:F14").format.borders = {
  top: { style: "thin", color: "#D8CDBB" },
  bottom: { style: "thin", color: "#D8CDBB" },
  left: { style: "thin", color: "#D8CDBB" },
  right: { style: "thin", color: "#D8CDBB" },
};
reviewSheet.getRange("A:F").format.columnWidthPx = 140;
reviewSheet.getRange("A:A").format.columnWidthPx = 70;
reviewSheet.getRange("E:E").format.columnWidthPx = 180;
reviewSheet.getRange("F:F").format.columnWidthPx = 260;
reviewSheet.freezePanes.freezeRows(4);

syllabusSheet.getRange("A1:C1").merge();
syllabusSheet.getRange("A1").values = [["FAO Topic Checklist"]];
syllabusSheet.getRange("A2:C2").merge();
syllabusSheet.getRange("A2").values = [[
  `Mark ${tick} when a topic is covered and ${cross} while pending.`,
]];
syllabusSheet.getRange("A4:C4").values = [["Subject", "Topic", "Status"]];

const syllabusRows = [
  ["Commercial Maths", "Percentage", cross],
  ["Commercial Maths", "Ratio", cross],
  ["Commercial Maths", "Mean", cross],
  ["Commercial Maths", "Median", cross],
  ["Commercial Maths", "SD", cross],
  ["Reasoning", "Series", cross],
  ["Reasoning", "Coding", cross],
  ["Reasoning", "Direction", cross],
  ["Reasoning", "Blood relation", cross],
  ["Commerce", "Business Organisation", cross],
  ["Commerce", "Banking", cross],
  ["Commerce", "Auditing", cross],
  ["Polity", "CAG", cross],
  ["Polity", "Finance Commission", cross],
  ["Polity", "Budget", cross],
  ["Economics", "GDP", cross],
  ["Economics", "Inflation", cross],
  ["Economics", "Monetary Policy", cross],
];
syllabusSheet.getRange(`A5:C${4 + syllabusRows.length}`).values = syllabusRows;

syllabusSheet.getRange("E4:F4").values = [["Subject", "Completed Topics"]];
syllabusSheet.getRange("E5:F9").values = [
  ["Commercial Maths", null],
  ["Reasoning", null],
  ["Commerce", null],
  ["Polity", null],
  ["Economics", null],
];
syllabusSheet.getRange("F5:F9").formulas = [
  [`=COUNTIFS($A$5:$A$22,E5,$C$5:$C$22,"${tick}")`],
  [`=COUNTIFS($A$5:$A$22,E6,$C$5:$C$22,"${tick}")`],
  [`=COUNTIFS($A$5:$A$22,E7,$C$5:$C$22,"${tick}")`],
  [`=COUNTIFS($A$5:$A$22,E8,$C$5:$C$22,"${tick}")`],
  [`=COUNTIFS($A$5:$A$22,E9,$C$5:$C$22,"${tick}")`],
];

syllabusSheet.getRange("A1:C1").format = {
  fill: "#1F6B52",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
syllabusSheet.getRange("A2:C2").format = {
  fill: "#D8EFE5",
  font: { color: "#1F2A21", italic: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
syllabusSheet.getRange("A4:C4").format = {
  fill: "#E8DCC6",
  font: { bold: true, color: "#1F2A21" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
syllabusSheet.getRange(`A5:C${4 + syllabusRows.length}`).format = {
  fill: "#FFFCF6",
  verticalAlignment: "center",
};
syllabusSheet.getRange("C5:C22").dataValidation = {
  rule: {
    type: "list",
    values: [tick, cross],
  },
};
syllabusSheet.getRange("C5:C22").conditionalFormats.addCustom(`=C5="${tick}"`, {
  fill: "#D7F0E6",
  font: { color: "#176147", bold: true },
});
syllabusSheet.getRange("C5:C22").conditionalFormats.addCustom(`=C5="${cross}"`, {
  fill: "#F8DDD7",
  font: { color: "#9C4334", bold: true },
});
syllabusSheet.getRange("A5:A22").conditionalFormats.addCustom('=A5="Commercial Maths"', {
  fill: "#F6E6BE",
  font: { color: "#8A5B0F", bold: true },
});
syllabusSheet.getRange("A5:A22").conditionalFormats.addCustom('=A5="Reasoning"', {
  fill: "#D7F0E6",
  font: { color: "#176147", bold: true },
});
syllabusSheet.getRange("A5:A22").conditionalFormats.addCustom('=A5="Commerce"', {
  fill: "#E6E0F8",
  font: { color: "#5B4694", bold: true },
});
syllabusSheet.getRange("A5:A22").conditionalFormats.addCustom('=A5="Polity"', {
  fill: "#F8DDD7",
  font: { color: "#9C4334", bold: true },
});
syllabusSheet.getRange("A5:A22").conditionalFormats.addCustom('=A5="Economics"', {
  fill: "#DCEAF8",
  font: { color: "#285F8F", bold: true },
});
syllabusSheet.getRange("E4:F4").format = {
  fill: "#1F2A21",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
syllabusSheet.getRange("E5:F9").format = {
  fill: "#FFFCF6",
  verticalAlignment: "center",
};
syllabusSheet.getRange("F5:F9").conditionalFormats.addDataBar({
  color: "#1F6B52",
  gradient: true,
});
syllabusSheet.getRange("A:F").format.columnWidthPx = 150;
syllabusSheet.getRange("A:A").format.columnWidthPx = 150;
syllabusSheet.getRange("B:B").format.columnWidthPx = 190;
syllabusSheet.getRange("C:C").format.columnWidthPx = 80;
syllabusSheet.getRange("E:E").format.columnWidthPx = 150;
syllabusSheet.getRange("F:F").format.columnWidthPx = 120;
syllabusSheet.getRange("A1:F22").format.borders = {
  top: { style: "thin", color: "#D8CDBB" },
  bottom: { style: "thin", color: "#D8CDBB" },
  left: { style: "thin", color: "#D8CDBB" },
  right: { style: "thin", color: "#D8CDBB" },
};
syllabusSheet.freezePanes.freezeRows(4);

testsSheet.getRange("A1:E1").merge();
testsSheet.getRange("A1").values = [["Mock Test Log"]];
testsSheet.getRange("A2:E2").merge();
testsSheet.getRange("A2").values = [[
  "Track each mock test with score, accuracy, weak area, and improvement notes.",
]];
testsSheet.getRange("A4:E4").values = [[
  "Test No",
  "Score",
  "Accuracy",
  "Weak Area",
  "Improvement",
]];
testsSheet.getRange("A5:E7").values = [
  [1, null, null, "", ""],
  [2, null, null, "", ""],
  [3, null, null, "", ""],
];

testsSheet.getRange("A1:E1").format = {
  fill: "#1F2A21",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
testsSheet.getRange("A2:E2").format = {
  fill: "#F6E6BE",
  font: { color: "#1F2A21", italic: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
testsSheet.getRange("A4:E4").format = {
  fill: "#E8DCC6",
  font: { bold: true, color: "#1F2A21" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
testsSheet.getRange("A5:E7").format = {
  fill: "#FFFCF6",
  verticalAlignment: "center",
  wrapText: true,
};
testsSheet.getRange("A5:A7").format = {
  fill: "#D8EFE5",
  font: { bold: true, color: "#176147" },
  horizontalAlignment: "center",
};
testsSheet.getRange("B5:C7").conditionalFormats.addDataBar({
  color: "#1F6B52",
  gradient: true,
});
testsSheet.getRange("C5:C7").format.numberFormat = "0.00%";
testsSheet.getRange("A:E").format.columnWidthPx = 150;
testsSheet.getRange("A:A").format.columnWidthPx = 80;
testsSheet.getRange("D:D").format.columnWidthPx = 180;
testsSheet.getRange("E:E").format.columnWidthPx = 240;
testsSheet.getRange("A1:E7").format.borders = {
  top: { style: "thin", color: "#D8CDBB" },
  bottom: { style: "thin", color: "#D8CDBB" },
  left: { style: "thin", color: "#D8CDBB" },
  right: { style: "thin", color: "#D8CDBB" },
};
testsSheet.freezePanes.freezeRows(4);

await fs.mkdir(outputDir, { recursive: true });

const trackerPreview = await workbook.render({
  sheetName: "Daily Tracker",
  range: "A1:M11",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "daily-tracker-preview.png"),
  new Uint8Array(await trackerPreview.arrayBuffer()),
);

const summaryPreview = await workbook.render({
  sheetName: "Weekly Summary",
  range: "A1:N20",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "weekly-summary-preview.png"),
  new Uint8Array(await summaryPreview.arrayBuffer()),
);

const phasesPreview = await workbook.render({
  sheetName: "Study Phases",
  range: "A1:F12",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "study-phases-preview.png"),
  new Uint8Array(await phasesPreview.arrayBuffer()),
);

const reviewPreview = await workbook.render({
  sheetName: "Weekly Review",
  range: "A1:F14",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "weekly-review-preview.png"),
  new Uint8Array(await reviewPreview.arrayBuffer()),
);

const syllabusPreview = await workbook.render({
  sheetName: "Topic Checklist",
  range: "A1:F22",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "topic-checklist-preview.png"),
  new Uint8Array(await syllabusPreview.arrayBuffer()),
);

const testsPreview = await workbook.render({
  sheetName: "Mock Test Log",
  range: "A1:E7",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "mock-test-log-preview.png"),
  new Uint8Array(await testsPreview.arrayBuffer()),
);

const inspection = await workbook.inspect({
  kind: "table",
  range: "Daily Tracker!A4:M11",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 13,
});
console.log(inspection.ndjson);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errorScan.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
const outputPath = path.join(outputDir, "fao-study-tracker.xlsx");
await output.save(outputPath);

console.log(`Saved workbook to ${outputPath}`);

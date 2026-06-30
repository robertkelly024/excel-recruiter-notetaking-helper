import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/recruiter_notetaking_helper";
const outputPath = `${outputDir}/Recruiter_Notetaking_Helper_v2.xlsx`;

const COLORS = {
  navy: "#243746",
  blue: "#2F5D7C",
  teal: "#0F766E",
  paleTeal: "#E7F3F1",
  paleBlue: "#EAF2F8",
  grayText: "#475569",
  lightGray: "#F3F6F8",
  line: "#CBD5E1",
  white: "#FFFFFF",
  warning: "#FFF7ED",
};

const scoreValues = ["1", "2", "3", "4", "5"];
const stageValues = [
  "Recruiter screen",
  "Hiring manager screen",
  "Panel",
  "Final",
  "Offer discussion",
  "Other",
];
const nextStepValues = [
  "Advance to hiring manager",
  "Schedule panel",
  "Follow up",
  "Hold / revisit",
  "Not moving forward",
  "Needs more info",
];

const employees = [
  ["100245", "MM-004812", "Avery Martinez", "Senior Financial Analyst", "2021-08-16", "Austin, TX", "Finance Analyst III", "Professional", "Jordan Lee", "Corporate Finance", "Planning & Analysis", "Finance"],
  ["100628", "MM-006104", "Priya Shah", "Product Operations Manager", "2020-03-02", "New York, NY", "Product Ops Manager II", "Manager", "Morgan Chen", "Product", "Operations", "Technology"],
  ["101117", "MM-007733", "Marcus Johnson", "Lead Recruiter", "2019-11-18", "Remote - US", "Talent Acquisition Lead", "Professional", "Casey Nguyen", "People", "Talent Acquisition", "HR"],
  ["101442", "MM-008512", "Elena Rossi", "Data Engineering Manager", "2018-06-25", "Chicago, IL", "Data Engineering Manager", "Manager", "Sam Patel", "Data", "Engineering", "Technology"],
  ["101879", "MM-009406", "Noah Kim", "Customer Success Specialist", "2022-01-10", "Seattle, WA", "Customer Success Specialist II", "Professional", "Riley Brooks", "Customer", "Customer Success", "Commercial"],
  ["102014", "MM-010102", "Danielle Green", "HR Business Partner", "2023-04-03", "Atlanta, GA", "HRBP II", "Professional", "Taylor Smith", "People", "Business Partners", "HR"],
];

const candidateRows = [
  [
    "NOTE-2026-001", "2026-06-02 10:12", "Recruiter Demo", "100245", "Matched on employee_id",
    "100245", "MM-004812", "Avery Martinez", "Senior Financial Analyst", "2021-08-16", "Austin, TX", "Finance Analyst III", "Professional", "Jordan Lee", "Corporate Finance", "Planning & Analysis", "Finance",
    "FP&A Manager", "Recruiter screen", "Recruiter Demo", "2026-06-02", "Morgan Chen", "Advance to hiring manager",
    4, "Strong analytical examples and clear ownership of planning cycles.",
    5, "Concise, structured responses.",
    4, "Clear interest in broader finance leadership.",
    4, "Relevant systems and forecast experience.",
    3, "Open to hybrid schedule, some timing constraints.",
    3, "Needs clarity on band before next step.",
    2, "Low overall risk; confirm comp expectations.",
    4,
    "Good internal mobility candidate for FP&A manager screen. Move to hiring-manager conversation.",
  ],
  [
    "NOTE-2026-002", "2026-06-03 14:40", "Recruiter Demo", "MM-008512", "Matched on mm_id",
    "101442", "MM-008512", "Elena Rossi", "Data Engineering Manager", "2018-06-25", "Chicago, IL", "Data Engineering Manager", "Manager", "Sam Patel", "Data", "Engineering", "Technology",
    "Director, Data Platform", "Recruiter screen", "Recruiter Demo", "2026-06-03", "Sam Patel", "Follow up",
    5, "Deep match to platform ownership and stakeholder management.",
    4, "Direct and clear; could provide more executive-level framing.",
    4, "Motivated by broader scope.",
    5, "Very strong experience match.",
    4, "Location and timing look workable.",
    3, "Likely needs level/comp calibration.",
    3, "Primary risk is scope alignment.",
    4.4,
    "High-potential internal candidate. Follow up on level expectations and executive stakeholder examples.",
  ],
];

const workbook = Workbook.create();
const intake = workbook.worksheets.getOrAdd("Intake", { renameFirstIfOnlyNewSpreadsheet: true });
const candidates = workbook.worksheets.add("Candidates");
const employeeSheet = workbook.worksheets.add("hd_employees");
const settings = workbook.worksheets.add("Settings");
const instructions = workbook.worksheets.add("Instructions");

function range(sheet, address) {
  return sheet.getRange(address);
}

function value(sheet, address, content) {
  range(sheet, address).values = [[content]];
}

function merge(sheet, address) {
  range(sheet, address).merge();
}

function style(sheet, address, fmt) {
  range(sheet, address).format = fmt;
}

function setWidths(sheet, widths) {
  widths.forEach(([address, px]) => {
    range(sheet, address).format.columnWidthPx = px;
  });
}

function setHeights(sheet, heights) {
  heights.forEach(([address, px]) => {
    range(sheet, address).format.rowHeightPx = px;
  });
}

function sectionTitle(sheet, address, title) {
  merge(sheet, address);
  const cell = address.split(":")[0];
  value(sheet, cell, title);
  style(sheet, address, {
    fill: COLORS.navy,
    font: { color: COLORS.white, bold: true, size: 13 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  });
}

function label(sheet, address, text) {
  value(sheet, address, text);
  style(sheet, address, {
    fill: COLORS.lightGray,
    font: { color: COLORS.grayText, bold: true, size: 10 },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: COLORS.line },
  });
}

function inputBlock(sheet, address, fill = COLORS.white) {
  style(sheet, address, {
    fill,
    font: { color: "#111827", size: 11 },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: COLORS.line },
  });
}

function outputBlock(sheet, address) {
  style(sheet, address, {
    fill: COLORS.paleBlue,
    font: { color: "#111827", size: 11 },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: COLORS.line },
  });
}

// Intake layout
intake.showGridLines = false;
setWidths(intake, [
  ["A:A", 24], ["B:B", 142], ["C:C", 128], ["D:D", 128], ["E:E", 126], ["F:F", 126],
  ["G:G", 126], ["H:H", 126], ["I:I", 126], ["J:J", 126], ["K:K", 126], ["L:L", 24],
]);
setHeights(intake, [
  ["1:1", 18], ["2:2", 40], ["3:3", 24], ["4:4", 16], ["5:5", 28], ["6:6", 12], ["7:7", 38],
  ["8:8", 12], ["9:17", 34], ["18:18", 14], ["19:19", 28], ["20:20", 12], ["21:24", 36],
  ["25:25", 14], ["26:26", 28], ["27:34", 44], ["35:35", 14], ["36:36", 28], ["37:43", 30],
  ["44:44", 16],
]);

merge(intake, "B2:K2");
value(intake, "B2", "Internal Candidate Interview Notes");
style(intake, "B2:K2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 20 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
});
merge(intake, "B3:K3");
value(intake, "B3", "Lookup an employee, capture screening notes, and submit the record to the Candidates table.");
style(intake, "B3:K3", {
  fill: COLORS.paleTeal,
  font: { color: COLORS.grayText, size: 11 },
  verticalAlignment: "center",
});

sectionTitle(intake, "B5:K5", "Candidate Details");
label(intake, "B7", "Lookup value");
merge(intake, "C7:F7");
inputBlock(intake, "C7:F7", COLORS.warning);
label(intake, "G7", "Lookup status");
merge(intake, "H7:K7");
outputBlock(intake, "H7:K7");

const detailRows = [
  ["B9", "Full name", "C9:D9", "E9", "Employee ID", "F9:G9", "H9", "MM ID", "I9:K9"],
  ["B11", "Business title", "C11:D11", "E11", "Recent hire date", "F11:G11", "H11", "Location", "I11:K11"],
  ["B13", "Job profile", "C13:D13", "E13", "Management level", "F13:G13", "H13", "Manager full name", "I13:K13"],
  ["B15", "Business group", "C15:D15", "E15", "Sub business group", "F15:G15", "H15", "Division", "I15:K15"],
];
for (const row of detailRows) {
  for (let i = 0; i < row.length; i += 3) {
    label(intake, row[i], row[i + 1]);
    merge(intake, row[i + 2]);
    outputBlock(intake, row[i + 2]);
  }
}

sectionTitle(intake, "B19:K19", "Screening Notes");
label(intake, "B21", "Open role / req");
merge(intake, "C21:D21");
inputBlock(intake, "C21:D21");
label(intake, "E21", "Stage");
merge(intake, "F21:G21");
inputBlock(intake, "F21:G21");
label(intake, "H21", "Recruiter");
merge(intake, "I21:K21");
inputBlock(intake, "I21:K21");

label(intake, "B23", "Screen date");
merge(intake, "C23:D23");
inputBlock(intake, "C23:D23");
label(intake, "E23", "Hiring manager");
merge(intake, "F23:G23");
inputBlock(intake, "F23:G23");
label(intake, "H23", "Next step");
merge(intake, "I23:K23");
inputBlock(intake, "I23:K23");

range(intake, "F21:G21").dataValidation = { allowBlank: false, list: { inCellDropDown: true, source: stageValues } };
range(intake, "I23:K23").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: nextStepValues } };

sectionTitle(intake, "B26:K26", "Recruiter Rubric");
range(intake, "B27:K27").values = [["Category", "Score", "Evidence / Notes", null, null, null, null, null, null, null]];
merge(intake, "D27:K27");
style(intake, "B27:K27", {
  fill: COLORS.blue,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: COLORS.blue },
});

const rubric = [
  "Role fit",
  "Communication",
  "Motivation",
  "Experience match",
  "Logistics fit",
  "Compensation fit",
  "Risk level",
];
rubric.forEach((name, idx) => {
  const row = 28 + idx;
  merge(intake, `B${row}:C${row}`);
  value(intake, `B${row}`, name);
  style(intake, `B${row}:C${row}`, {
    fill: COLORS.lightGray,
    font: { color: COLORS.grayText, bold: true, size: 10 },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: COLORS.line },
  });
  inputBlock(intake, `D${row}`);
  range(intake, `D${row}`).dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: scoreValues } };
  merge(intake, `E${row}:K${row}`);
  inputBlock(intake, `E${row}:K${row}`);
});

sectionTitle(intake, "B36:K36", "Summary Notes");
merge(intake, "B37:K43");
inputBlock(intake, "B37:K43");

// Candidates table
candidates.showGridLines = false;
const candidateHeaders = [
  "submission_id", "submitted_at", "submitted_by", "lookup_key", "lookup_status",
  "employee_id", "mm_id", "full_name", "business_title", "recent_hire_date", "location", "job_profile",
  "management_level", "manager_full_name", "business_group", "sub_business_group", "division",
  "open_role_req", "stage", "recruiter", "screen_date", "hiring_manager", "next_step",
  "role_fit_score", "role_fit_notes", "communication_score", "communication_notes", "motivation_score", "motivation_notes",
  "experience_match_score", "experience_match_notes", "logistics_fit_score", "logistics_fit_notes",
  "compensation_fit_score", "compensation_fit_notes", "risk_level_score", "risk_level_notes", "average_score", "summary_notes",
];
range(candidates, "A1:AM1").values = [candidateHeaders];
range(candidates, `A2:AM${candidateRows.length + 1}`).values = candidateRows;
const candidateTable = candidates.tables.add(`A1:AM${candidateRows.length + 1}`, true);
candidateTable.name = "tblCandidateNotes";
style(candidates, "A1:AM1", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
});
style(candidates, `A2:AM${candidateRows.length + 1}`, {
  fill: COLORS.white,
  font: { color: "#111827", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
setWidths(candidates, [
  ["A:E", 126], ["F:H", 128], ["I:Q", 146], ["R:W", 132], ["X:AL", 132], ["AM:AM", 260],
]);
candidates.freezePanes.freezeRows(1);

// hd_employees table
employeeSheet.showGridLines = false;
const employeeHeaders = [
  "employee_id", "mm_id", "full_name", "business_title", "recent_hire_date", "location", "job_profile",
  "management_level", "manager_full_name", "business_group", "sub_business_group", "division",
];
range(employeeSheet, "A1:L1").values = [employeeHeaders];
range(employeeSheet, `A2:L${employees.length + 1}`).values = employees;
const employeeTable = employeeSheet.tables.add(`A1:L${employees.length + 1}`, true);
employeeTable.name = "tblHdEmployees";
style(employeeSheet, "A1:L1", {
  fill: COLORS.teal,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
});
style(employeeSheet, `A2:L${employees.length + 1}`, {
  fill: COLORS.white,
  font: { color: "#111827", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
setWidths(employeeSheet, [["A:B", 120], ["C:C", 170], ["D:D", 210], ["E:E", 130], ["F:H", 150], ["I:I", 170], ["J:L", 170]]);
employeeSheet.freezePanes.freezeRows(1);

// Settings sheet
settings.showGridLines = false;
setWidths(settings, [["A:A", 28], ["B:B", 210], ["C:C", 300], ["D:D", 28], ["E:E", 220], ["F:F", 300]]);
merge(settings, "B2:F2");
value(settings, "B2", "Settings");
style(settings, "B2:F2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 18 },
  verticalAlignment: "center",
});
value(settings, "B4", "Score");
value(settings, "C4", "Meaning");
range(settings, "B5:C9").values = [
  ["1", "Lowest score / limited evidence"],
  ["2", "Below target"],
  ["3", "Meets target"],
  ["4", "Strong"],
  ["5", "Highest score / standout evidence"],
];
value(settings, "E4", "Stage values");
range(settings, `E5:E${stageValues.length + 4}`).values = stageValues.map((item) => [item]);
value(settings, "F4", "Next step values");
range(settings, `F5:F${nextStepValues.length + 4}`).values = nextStepValues.map((item) => [item]);
style(settings, "B4:C4", { fill: COLORS.blue, font: { color: COLORS.white, bold: true }, horizontalAlignment: "center" });
style(settings, "E4:F4", { fill: COLORS.blue, font: { color: COLORS.white, bold: true }, horizontalAlignment: "center" });
style(settings, "B5:C9", { borders: { preset: "outside", style: "thin", color: COLORS.line }, wrapText: true });
style(settings, `E5:F${Math.max(stageValues.length, nextStepValues.length) + 4}`, { borders: { preset: "outside", style: "thin", color: COLORS.line }, wrapText: true });

// Instructions sheet
instructions.showGridLines = false;
setWidths(instructions, [["A:A", 28], ["B:B", 220], ["C:C", 760], ["D:D", 28]]);
setHeights(instructions, [["1:1", 18], ["2:2", 42], ["3:3", 16], ["4:12", 46]]);
merge(instructions, "B2:C2");
value(instructions, "B2", "Recruiter Notetaking Helper - Setup Notes");
style(instructions, "B2:C2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 18 },
  verticalAlignment: "center",
});
const instructionRows = [
  ["1. Refresh roster", "A weekly Python job should replace the rows in the `tblHdEmployees` table on `hd_employees`. Keep the sheet name, table name, and column headers unchanged."],
  ["2. Lookup candidates", "On `Intake`, enter a full name, employee_id, or mm_id in the lookup box. Run `LookupEmployee.ts` from Excel Online Automate to populate Candidate Details."],
  ["3. Capture notes", "Complete the screening metadata, rubric scores, evidence notes, and summary notes. Scores are 1 through 5, where 1 is the lowest score and 5 is the highest."],
  ["4. Submit notes", "Run `SubmitCandidate.ts` to append the completed screen to `tblCandidateNotes` on the `Candidates` sheet. The script validates required fields before saving."],
  ["5. Reset form", "Run `ResetIntakeForm.ts` to clear editable intake fields without changing labels, formatting, dropdowns, or the underlying data tables."],
  ["6. Excel Online", "Add Automate buttons if desired after uploading the workbook to SharePoint or OneDrive. The intake sheet itself intentionally avoids setup text so it stays clean for recruiters."],
];
range(instructions, `B4:C${instructionRows.length + 3}`).values = instructionRows;
style(instructions, `B4:B${instructionRows.length + 3}`, {
  fill: COLORS.lightGray,
  font: { color: COLORS.grayText, bold: true },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
style(instructions, `C4:C${instructionRows.length + 3}`, {
  fill: COLORS.white,
  font: { color: "#111827" },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});

// General date formatting.
range(intake, "C23:D23").format.numberFormat = "yyyy-mm-dd";
range(candidates, "J:J").format.numberFormat = "yyyy-mm-dd";
range(candidates, "U:U").format.numberFormat = "yyyy-mm-dd";
range(employeeSheet, "E:E").format.numberFormat = "yyyy-mm-dd";

// Export workbook and visual previews.
await fs.mkdir(outputDir, { recursive: true });
for (const [sheetName, previewRange] of [
  ["Intake", "A1:L44"],
  ["Candidates", "A1:AM8"],
  ["hd_employees", "A1:L10"],
  ["Settings", "A1:F13"],
  ["Instructions", "A1:D11"],
]) {
  const blob = await workbook.render({ sheetName, range: previewRange, scale: 2 });
  await fs.writeFile(`${outputDir}/${sheetName}_v2.png`, Buffer.from(await blob.arrayBuffer()));
}
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);

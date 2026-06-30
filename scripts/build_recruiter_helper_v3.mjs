import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/recruiter_notetaking_helper";
const outputPath = `${outputDir}/Recruiter_Notetaking_Helper_v3.xlsx`;
const previewDir = "/tmp/recruiter_notetaking_helper_v3_previews";
const execFileAsync = promisify(execFile);

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

const nextStepValues = [
  "Advance to hiring manager",
  "Schedule panel",
  "Follow up",
  "Hold / revisit",
  "Not moving forward",
  "Needs more info",
];
const mobilityInterestValues = [
  "Exploratory",
  "Open to the right role",
  "Actively interested",
  "Not ready yet",
  "Unsure",
];
const yesNoUnknownValues = ["Yes", "No", "Unknown"];
const candidateAppliedValues = ["Yes", "No", "Unknown"];
const candidateStageValues = [
  "Not applied",
  "Applied",
  "Recruiter screen",
  "Hiring manager review",
  "Interview scheduled",
  "Offer",
  "Not moving forward",
  "Withdrawn",
];

const inputCatalogRows = [
  ["skill", "Financial modeling", "financial model | finance modeling | modeling", 3, "2026-06-16", "active"],
  ["skill", "Forecasting", "forecast | forecast planning | business forecasting", 4, "2026-06-16", "active"],
  ["skill", "Executive reporting", "exec reporting | leadership reporting | executive dashboards", 2, "2026-06-16", "active"],
  ["skill", "Stakeholder management", "stakeholder engagement | stakeholder communication | partner management", 2, "2026-06-16", "active"],
  ["skill", "People leadership", "people management | team leadership | coaching", 3, "2026-06-16", "active"],
  ["skill", "Data architecture", "data modeling | data platform architecture | architecture", 2, "2026-06-03", "active"],
  ["skill", "Executive communication", "exec communication | senior leader communication | leadership communication", 1, "2026-06-03", "active"],
  ["skill", "Cross-functional influence", "cross functional influence | influence without authority | cross-org influence", 1, "2026-06-16", "active"],
  ["role", "FP&A Manager", "fpa manager | financial planning manager | finance manager", 2, "2026-06-16", "active"],
  ["role", "Director, Data Platform", "data platform director | director data platform | data director", 1, "2026-06-03", "active"],
  ["role", "Senior HR Business Partner", "sr hrbp | senior hrbp | people business partner", 1, "2026-06-03", "active"],
  ["role", "Planning Lead", "planning leader | planning manager | business planning lead", 2, "2026-06-16", "active"],
  ["level", "Professional", "individual contributor | ic | senior professional", 2, "2026-06-16", "active"],
  ["level", "Manager", "people manager | management | mgr", 3, "2026-06-16", "active"],
  ["level", "Director", "dir | senior director | director level", 1, "2026-06-03", "active"],
  ["level", "Executive", "vp | vice president | executive level", 1, "2026-06-03", "active"],
  ["function", "Finance", "corporate finance | fpa | fp&a | planning and analysis", 4, "2026-06-16", "active"],
  ["function", "Data", "data engineering | analytics | data platform", 2, "2026-06-03", "active"],
  ["function", "People", "hr | human resources | talent", 2, "2026-06-03", "active"],
  ["function", "Product", "product management | product operations | product ops", 1, "2026-06-03", "active"],
];

const employees = [
  ["100245", "MM-004812", "2021-08-16", "2024-02-01", "Finance Analyst III", "Professional", "Senior Financial Analyst", "Avery Martinez", "Austin, TX", "900101", "Jordan Lee", "Corporate Finance", "Planning & Analysis", "Finance", "JP-FA3", 2026, 6],
  ["100628", "MM-006104", "2020-03-02", "2023-09-18", "Product Ops Manager II", "Manager", "Product Operations Manager", "Priya Shah", "New York, NY", "900202", "Morgan Chen", "Product", "Operations", "Technology", "JP-PO2", 2026, 6],
  ["101117", "MM-007733", "2019-11-18", "2022-07-11", "Talent Acquisition Lead", "Professional", "Lead Recruiter", "Marcus Johnson", "Remote - US", "900303", "Casey Nguyen", "People", "Talent Acquisition", "HR", "JP-TAL", 2026, 6],
  ["101442", "MM-008512", "2018-06-25", "2025-01-06", "Data Engineering Manager", "Manager", "Data Engineering Manager", "Elena Rossi", "Chicago, IL", "900404", "Sam Patel", "Data", "Engineering", "Technology", "JP-DEM", 2026, 6],
  ["101879", "MM-009406", "2022-01-10", "2022-01-10", "Customer Success Specialist II", "Professional", "Customer Success Specialist", "Noah Kim", "Seattle, WA", "900505", "Riley Brooks", "Customer", "Customer Success", "Commercial", "JP-CSS2", 2026, 6],
  ["102014", "MM-010102", "2023-04-03", "2024-11-04", "HRBP II", "Professional", "HR Business Partner", "Danielle Green", "Atlanta, GA", "900606", "Taylor Smith", "People", "Business Partners", "HR", "JP-HRBP2", 2026, 6],
];

const requisitions = [
  [
    "Employee", "Open", "REQ-2026-0142", "FP&A Manager", "Add to Staff", "Manager", 1,
    "Corporate Finance", "Planning & Analysis", "Finance", "Finance Leadership", "Recruiter Demo",
    "Austin, TX", "Morgan Chen",
    "Finance Manager",
  ],
  [
    "Employee", "Open", "REQ-2026-0198", "Director, Data Platform", "Replacement", "Director", 1,
    "Data", "Engineering", "Technology", "Data Platform Organization", "Recruiter Demo",
    "Chicago, IL", "Sam Patel",
    "Data Platform Director",
  ],
  [
    "Employee", "On Hold", "REQ-2026-0221", "Senior HR Business Partner", "Add to Staff", "Professional", 2,
    "People", "Business Partners", "HR", "People Partner Team", "Recruiter Demo",
    "Atlanta, GA", "Taylor Smith",
    "HRBP III",
  ],
];

const candidateRows = [
  [
    "NOTE-2026-001", "2026-06-02 10:12 UTC", "Recruiter Demo", "100245", "Matched on employee_id",
    "100245", "MM-004812", "2021-08-16", "2024-02-01", "Finance Analyst III", "Professional", "Senior Financial Analyst", "Avery Martinez", "Austin, TX", "900101", "Jordan Lee", "Corporate Finance", "Planning & Analysis", "Finance", "JP-FA3", 2026, 6,
    "REQ-2026-0142", "Matched on requisition_id: FP&A Manager",
    "Employee", "Open", "REQ-2026-0142", "FP&A Manager", "Add to Staff", "Manager", 1,
    "Corporate Finance", "Planning & Analysis", "Finance", "Finance Leadership", "Recruiter Demo",
    "Austin, TX", "Morgan Chen", "Finance Manager", "Yes", "Recruiter screen",
    "Recruiter Demo", "2026-06-02", "Advance to hiring manager",
    "FP&A Manager | Finance Manager | Planning Lead",
    "Manager",
    "Finance",
    "Actively interested",
    "Yes",
    "forecasting | executive reporting | stakeholder management | financial modeling | people leadership | strategic planning cadence | finance transformation | enterprise leadership",
    "Yes",
    "Yes",
    "Strong internal mobility candidate. Confirm readiness for people leadership and align with current manager before next step.",
  ],
  [
    "NOTE-2026-002", "2026-06-03 14:40 UTC", "Recruiter Demo", "MM-008512", "Matched on mm_id",
    "101442", "MM-008512", "2018-06-25", "2025-01-06", "Data Engineering Manager", "Manager", "Data Engineering Manager", "Elena Rossi", "Chicago, IL", "900404", "Sam Patel", "Data", "Engineering", "Technology", "JP-DEM", 2026, 6,
    "Director, Data Platform", "Matched on job_posting_title: Director, Data Platform",
    "Employee", "Open", "REQ-2026-0198", "Director, Data Platform", "Replacement", "Director", 1,
    "Data", "Engineering", "Technology", "Data Platform Organization", "Recruiter Demo",
    "Chicago, IL", "Sam Patel", "Data Platform Director", "Unknown", "Hiring manager review",
    "Recruiter Demo", "2026-06-03", "Follow up",
    "Director, Data Platform | Senior Engineering Leadership",
    "Director | Executive",
    "Data | Product",
    "Open to the right role",
    "Unknown",
    "platform strategy | engineering leadership | data architecture | executive communication | enterprise operating model | chief data officer scope | product strategy",
    "Unknown",
    "Yes",
    "High-potential conversation. Follow up on level expectations, timing, and examples of cross-org influence.",
  ],
  [
    "NOTE-2026-003", "2026-06-16 15:25 UTC", "Recruiter Demo", "100245", "Matched on employee_id",
    "100245", "MM-004812", "2021-08-16", "2024-02-01", "Finance Analyst III", "Professional", "Senior Financial Analyst", "Avery Martinez", "Austin, TX", "900101", "Jordan Lee", "Corporate Finance", "Planning & Analysis", "Finance", "JP-FA3", 2026, 6,
    "REQ-2026-0142", "Matched on requisition_id: FP&A Manager",
    "Employee", "Open", "REQ-2026-0142", "FP&A Manager", "Add to Staff", "Manager", 1,
    "Corporate Finance", "Planning & Analysis", "Finance", "Finance Leadership", "Recruiter Demo",
    "Austin, TX", "Morgan Chen", "Finance Manager", "Yes", "Interview scheduled",
    "Recruiter Demo", "2026-06-16", "Schedule panel",
    "FP&A Manager | Planning Lead",
    "Manager | Director",
    "Finance",
    "Actively interested",
    "Yes",
    "forecasting | executive reporting | stakeholder management | financial modeling | coaching | people leadership | cross-functional influence | finance transformation | enterprise leadership",
    "Yes",
    "Yes",
    "Follow-up conversation confirmed continued interest. Avery shared stronger examples of coaching peers and leading the quarterly planning cadence; schedule a panel focused on leadership readiness.",
  ],
];

const workbook = Workbook.create();
const intake = workbook.worksheets.getOrAdd("Intake", { renameFirstIfOnlyNewSpreadsheet: true });
const pipeline = workbook.worksheets.add("Pipeline");
const candidates = workbook.worksheets.add("Candidates");
const employeeSheet = workbook.worksheets.add("hd_employees");
const requisitionSheet = workbook.worksheets.add("hd_requisitions");
const inputCatalog = workbook.worksheets.add("input_catalog");
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

function columnName(index) {
  let name = "";
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function unique(values) {
  return Array.from(new Set(values));
}

function cloneUnlockedXf(xf) {
  const prefixMatch = xf.match(/^<(\w+:)?xf\b/);
  const prefix = prefixMatch?.[1] ?? "";
  let next = xf.replace(new RegExp(`<${prefix}protection\\b[^>]*\\/>`, "g"), "");
  next = /applyProtection=/.test(next)
    ? next.replace(/applyProtection="[^"]*"/, 'applyProtection="1"')
    : next.replace(new RegExp(`^<${prefix}xf\\b`), `<${prefix}xf applyProtection="1"`);
  if (next.endsWith("/>")) {
    return next.replace(/\s*\/>$/, `><${prefix}protection locked="0"/></${prefix}xf>`);
  }
  return next.replace(new RegExp(`</${prefix}xf>$`), `<${prefix}protection locked="0"/></${prefix}xf>`);
}

function styleIdForCell(sheetXml, cellRef) {
  const match = sheetXml.match(new RegExp(`<(?:\\w+:)?c\\b[^>]*\\br="${cellRef}"[^>]*>`));
  if (!match) return 0;
  const styleMatch = match[0].match(/\bs="(\d+)"/);
  return styleMatch ? Number(styleMatch[1]) : 0;
}

function setStyleIdForCell(sheetXml, cellRef, styleId) {
  return sheetXml.replace(new RegExp(`(<(?:\\w+:)?c\\b[^>]*\\br="${cellRef}"[^>]*)(>)`), (match, start, end) => {
    const next = /\bs="\d+"/.test(start)
      ? start.replace(/\bs="\d+"/, `s="${styleId}"`)
      : `${start} s="${styleId}"`;
    return `${next}${end}`;
  });
}

async function protectIntakeSheet(xlsxPath) {
  const tempDir = `${previewDir}/xlsx_protection`;
  await fs.rm(tempDir, { recursive: true, force: true });
  await fs.mkdir(tempDir, { recursive: true });
  await execFileAsync("/usr/bin/unzip", ["-q", xlsxPath, "-d", tempDir]);

  const stylesPath = path.join(tempDir, "xl/styles.xml");
  const sheetPath = path.join(tempDir, "xl/worksheets/sheet1.xml");
  let stylesXml = await fs.readFile(stylesPath, "utf8");
  let sheetXml = await fs.readFile(sheetPath, "utf8");

  const unlockedCells = ["C3", "F3", "I3", "C7", "C21", "F27", "I27", "C30", "E31", "C32", "E33", "C34", "E35", "C36", "E37", "C40", "F40", "I40", "K40", "B43"];
  const styleIds = unique(unlockedCells.map((cellRef) => styleIdForCell(sheetXml, cellRef)));
  const cellXfsMatch = stylesXml.match(/<(\w+:)?cellXfs count="(\d+)">([\s\S]*?)<\/(\w+:)?cellXfs>/);
  if (!cellXfsMatch) throw new Error("Unable to locate workbook cellXfs for sheet protection patch.");

  const prefix = cellXfsMatch[1] ?? "";
  const xfs = cellXfsMatch[3].match(
    new RegExp(
      `<${prefix}xf\\b(?=[^>]*\\/\\s*>)[^>]*>|<${prefix}xf\\b(?![^>]*\\/\\s*>)[^>]*>[\\s\\S]*?<\\/${prefix}xf>`,
      "g",
    ),
  ) ?? [];
  const styleMap = new Map();
  for (const styleId of styleIds) {
    const sourceXf = xfs[styleId] ?? xfs[0];
    styleMap.set(styleId, xfs.length);
    xfs.push(cloneUnlockedXf(sourceXf));
  }

  const replacementCellXfs = `<${prefix}cellXfs count="${xfs.length}">${xfs.join("")}</${prefix}cellXfs>`;
  stylesXml = stylesXml.replace(new RegExp(`<${prefix}cellXfs count="\\d+">[\\s\\S]*?<\\/${prefix}cellXfs>`), replacementCellXfs);

  for (const cellRef of unlockedCells) {
    const originalStyleId = styleIdForCell(sheetXml, cellRef);
    sheetXml = setStyleIdForCell(sheetXml, cellRef, styleMap.get(originalStyleId));
  }

  for (const rowNumber of [31, 33, 35, 37]) {
    sheetXml = sheetXml.replace(
      new RegExp(`(<(?:\\w+:)?row\\b[^>]*\\br="${rowNumber}"[^>]*)(>)`),
      (match, start, end) => {
        const next = /\bhidden="[^"]*"/.test(start)
          ? start.replace(/\bhidden="[^"]*"/, 'hidden="1"')
          : `${start} hidden="1"`;
        return `${next}${end}`;
      },
    );
  }

  const sheetPrefix = sheetXml.match(/<(\w+:)?worksheet\b/)?.[1] ?? "";
  const hiddenHelperColumn = `<${sheetPrefix}col min="13" max="13" width="2" hidden="1" customWidth="1"/>`;
  if (!new RegExp(`<${sheetPrefix}col\\b[^>]*\\bmin="13"[^>]*>`).test(sheetXml)) {
    sheetXml = sheetXml.replace(
      new RegExp(`</${sheetPrefix}cols>`),
      `${hiddenHelperColumn}</${sheetPrefix}cols>`,
    );
  }
  const protectionXml = `<${sheetPrefix}sheetProtection sheet="1" objects="1" scenarios="1"/>`;
  sheetXml = new RegExp(`<${sheetPrefix}sheetProtection\\b[^>]*\\/>`).test(sheetXml)
    ? sheetXml.replace(new RegExp(`<${sheetPrefix}sheetProtection\\b[^>]*\\/>`), protectionXml)
    : sheetXml.replace(new RegExp(`</${sheetPrefix}sheetData>`), `</${sheetPrefix}sheetData>${protectionXml}`);

  await fs.writeFile(stylesPath, stylesXml);
  await fs.writeFile(sheetPath, sheetXml);
  await execFileAsync("/usr/bin/zip", ["-qr", xlsxPath, "."], { cwd: tempDir });
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

function resolutionRow(sheet, rowNumber, labelText) {
  merge(sheet, `B${rowNumber}:D${rowNumber}`);
  value(sheet, `B${rowNumber}`, labelText);
  style(sheet, `B${rowNumber}:D${rowNumber}`, {
    fill: COLORS.lightGray,
    font: { color: COLORS.grayText, bold: true, size: 10 },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: COLORS.line },
  });
  merge(sheet, `E${rowNumber}:K${rowNumber}`);
  outputBlock(sheet, `E${rowNumber}:K${rowNumber}`);
  range(sheet, `E${rowNumber}`).dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: "=$M$2:$M$5" } };
}

// Intake layout
intake.showGridLines = false;
setWidths(intake, [
  ["A:A", 24], ["B:B", 142], ["C:C", 128], ["D:D", 128], ["E:E", 126], ["F:F", 126],
  ["G:G", 126], ["H:H", 126], ["I:I", 126], ["J:J", 126], ["K:K", 126], ["L:L", 24],
]);
setHeights(intake, [
  ["1:1", 18], ["2:2", 40], ["3:3", 36], ["4:4", 28], ["5:5", 28], ["6:6", 12], ["7:7", 38],
  ["8:8", 12], ["9:17", 34], ["18:18", 14], ["19:19", 28], ["20:20", 12], ["21:25", 36],
  ["26:26", 10], ["27:27", 36], ["28:28", 28], ["29:29", 12], ["30:30", 40],
  ["31:31", 38], ["32:32", 40], ["33:33", 38], ["34:34", 40], ["35:35", 38],
  ["36:36", 40], ["37:37", 38], ["38:39", 12], ["40:40", 40], ["41:41", 16],
  ["42:42", 28], ["43:48", 32], ["49:49", 16], ["50:50", 28], ["51:51", 34],
  ["52:52", 10], ["53:53", 30],
  ["54:54", 36], ["55:55", 52], ["56:56", 8],
  ["57:57", 36], ["58:58", 52], ["59:59", 8],
  ["60:60", 36], ["61:61", 52], ["62:62", 8],
  ["63:63", 36], ["64:64", 52], ["65:65", 8],
  ["66:66", 36], ["67:67", 52], ["68:68", 16],
]);

merge(intake, "B2:K2");
value(intake, "B2", "Internal Candidate Interview Notes");
style(intake, "B2:K2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 20 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
});
label(intake, "B3", "Screener");
merge(intake, "C3:D3");
inputBlock(intake, "C3:D3");
label(intake, "E3", "Screen date");
merge(intake, "F3:G3");
inputBlock(intake, "F3:G3");
label(intake, "H3", "Next step");
merge(intake, "I3:K3");
inputBlock(intake, "I3:K3");
range(intake, "I3:K3").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: nextStepValues } };
range(intake, "F3").dataValidation = {
  allowBlank: true,
  rule: {
    type: "date",
    operator: "between",
    formula1: "=DATE(2000,1,1)",
    formula2: "=DATE(2100,12,31)",
  },
  errorAlert: {
    style: "stop",
    title: "Invalid screen date",
    message: "Enter a valid date between 2000 and 2100.",
  },
};
range(intake, "C3").formulas = [["=Settings!$C$16"]];

merge(intake, "B4:K4");
value(intake, "B4", "Ready to submit when required candidate, requisition, and screener fields are complete.");
style(intake, "B4:K4", {
  fill: COLORS.paleTeal,
  font: { color: COLORS.grayText, italic: true, size: 10 },
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
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
  ["B11", "Business title", "C11:D11", "E11", "Date of last mobility event", "F11:G11", "H11", "Location", "I11:K11"],
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

sectionTitle(intake, "B19:K19", "Requisition Details");
label(intake, "B21", "Req lookup");
merge(intake, "C21:F21");
inputBlock(intake, "C21:F21", COLORS.warning);
label(intake, "G21", "Req status");
merge(intake, "H21:K21");
outputBlock(intake, "H21:K21");

label(intake, "B23", "Req ID");
merge(intake, "C23:D23");
outputBlock(intake, "C23:D23");
label(intake, "E23", "Job title");
merge(intake, "F23:G23");
outputBlock(intake, "F23:G23");
label(intake, "H23", "Job profile");
merge(intake, "I23:K23");
outputBlock(intake, "I23:K23");

label(intake, "B25", "Recruiter");
merge(intake, "C25:D25");
outputBlock(intake, "C25:D25");
label(intake, "E25", "Location");
merge(intake, "F25:G25");
outputBlock(intake, "F25:G25");
label(intake, "H25", "Hiring manager");
merge(intake, "I25:K25");
outputBlock(intake, "I25:K25");

label(intake, "B27", "Job level");
merge(intake, "C27:D27");
outputBlock(intake, "C27:D27");
label(intake, "E27", "Has candidate applied?");
merge(intake, "F27:G27");
inputBlock(intake, "F27:G27");
label(intake, "H27", "Candidate stage");
merge(intake, "I27:K27");
inputBlock(intake, "I27:K27");
range(intake, "F27:G27").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: `=Settings!$E$5:$E$${candidateAppliedValues.length + 4}` } };
range(intake, "I27:K27").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: `=Settings!$F$5:$F$${candidateStageValues.length + 4}` } };

sectionTitle(intake, "B28:K28", "Exploratory Career Conversation");
label(intake, "B30", "Role input");
merge(intake, "C30:E30");
inputBlock(intake, "C30:E30", COLORS.warning);
label(intake, "F30", "Desired roles");
merge(intake, "G30:K30");
outputBlock(intake, "G30:K30");

resolutionRow(intake, 31, "Similar roles found. Select intended role from dropdown:");

label(intake, "B32", "Level input");
merge(intake, "C32:E32");
inputBlock(intake, "C32:E32", COLORS.warning);
label(intake, "F32", "Desired levels");
merge(intake, "G32:K32");
outputBlock(intake, "G32:K32");

resolutionRow(intake, 33, "Similar desired levels found. Select intended level from dropdown:");

label(intake, "B34", "Function input");
merge(intake, "C34:E34");
inputBlock(intake, "C34:E34", COLORS.warning);
label(intake, "F34", "Desired functions");
merge(intake, "G34:K34");
outputBlock(intake, "G34:K34");

resolutionRow(intake, 35, "Similar desired functions found. Select intended function from dropdown:");

label(intake, "B36", "Skill input");
merge(intake, "C36:E36");
inputBlock(intake, "C36:E36", COLORS.warning);
label(intake, "F36", "Skills");
merge(intake, "G36:K36");
outputBlock(intake, "G36:K36");

resolutionRow(intake, 37, "Similar skills found. Select intended skill from dropdown:");

label(intake, "B40", "Mobility interest");
merge(intake, "C40:D40");
inputBlock(intake, "C40:D40");
label(intake, "E40", "Willing to relocate?");
merge(intake, "F40:G40");
inputBlock(intake, "F40:G40");
label(intake, "H40", "Career profile complete?");
inputBlock(intake, "I40");
label(intake, "J40", "Development goals?");
inputBlock(intake, "K40");

range(intake, "C40:D40").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: mobilityInterestValues } };
range(intake, "F40:G40").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: yesNoUnknownValues } };
range(intake, "I40").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: yesNoUnknownValues } };
range(intake, "K40").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: yesNoUnknownValues } };

sectionTitle(intake, "B42:K42", "Recruiter Synthesis / Follow-Up Notes");
merge(intake, "B43:K48");
inputBlock(intake, "B43:K48");

sectionTitle(intake, "B50:K50", "Candidate Conversation Timeline");
merge(intake, "B51:K51");
value(intake, "B51", "Run Lookup Employee to display recent conversations. Use Open Candidate Notes to view this candidate in tblCandidateNotes.");
style(intake, "B51:K51", {
  fill: COLORS.paleTeal,
  font: { color: COLORS.grayText, italic: true, size: 10 },
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});

const timelineHeaders = [
  ["B53:C53", "Screen date"],
  ["D53:F53", "Current BG / Title"],
  ["G53:H53", "Role / requisition"],
  ["I53:K53", "Next step"],
];
for (const [address, text] of timelineHeaders) {
  merge(intake, address);
  value(intake, address.split(":")[0], text);
  style(intake, address, {
    fill: COLORS.blue,
    font: { color: COLORS.white, bold: true, size: 10 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.line },
  });
}

for (let index = 0; index < 5; index += 1) {
  const summaryRow = 54 + index * 3;
  const notesRow = summaryRow + 1;
  const separatorRow = summaryRow + 2;
  const fill = index % 2 === 0 ? COLORS.white : COLORS.lightGray;
  for (const address of [
    `B${summaryRow}:C${summaryRow}`,
    `D${summaryRow}:F${summaryRow}`,
    `G${summaryRow}:H${summaryRow}`,
    `I${summaryRow}:K${summaryRow}`,
  ]) {
    merge(intake, address);
    style(intake, address, {
      fill,
      font: { color: "#111827", size: 10 },
      verticalAlignment: "top",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: COLORS.line },
    });
  }
  merge(intake, `B${notesRow}:K${notesRow}`);
  style(intake, `B${notesRow}:K${notesRow}`, {
    fill,
    font: { color: COLORS.grayText, italic: true, size: 10 },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.line },
  });
  if (index < 4) {
    merge(intake, `B${separatorRow}:K${separatorRow}`);
    style(intake, `B${separatorRow}:K${separatorRow}`, {
      fill: COLORS.line,
      borders: { preset: "none" },
    });
  }
}

// Pipeline layout
pipeline.showGridLines = false;
setWidths(pipeline, [
  ["A:A", 24], ["B:B", 142], ["C:C", 128], ["D:D", 128], ["E:E", 126], ["F:F", 126],
  ["G:G", 126], ["H:H", 126], ["I:I", 126], ["J:J", 126], ["K:K", 126], ["L:L", 24],
  ["M:M", 2],
]);
setHeights(pipeline, [
  ["1:1", 18], ["2:2", 40], ["3:3", 36], ["4:4", 16], ["5:5", 28], ["6:6", 38],
  ["7:7", 12], ["8:10", 34], ["11:11", 16], ["12:12", 28], ["13:13", 12],
  ["14:14", 38], ["15:15", 12], ["16:16", 34], ["17:17", 16], ["18:18", 28],
  ["19:19", 34], ["20:20", 12], ["21:21", 30], ["22:40", 42], ["41:41", 16],
]);

merge(pipeline, "B2:K2");
value(pipeline, "B2", "Internal Candidate Pipeline");
style(pipeline, "B2:K2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 20 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
});
label(pipeline, "B3", "Screener");
merge(pipeline, "C3:D3");
inputBlock(pipeline, "C3:D3");
range(pipeline, "C3").formulas = [["=Settings!$C$16"]];
label(pipeline, "E3", "Pipeline status");
merge(pipeline, "F3:K3");
outputBlock(pipeline, "F3:K3");
value(pipeline, "F3", "Select a requisition, then add or review internal candidates.");

sectionTitle(pipeline, "B5:K5", "Requisition");
label(pipeline, "B6", "Req lookup");
merge(pipeline, "C6:F6");
inputBlock(pipeline, "C6:F6", COLORS.warning);
label(pipeline, "G6", "Req status");
merge(pipeline, "H6:K6");
outputBlock(pipeline, "H6:K6");

label(pipeline, "B8", "Req ID");
merge(pipeline, "C8:D8");
outputBlock(pipeline, "C8:D8");
label(pipeline, "E8", "Job title");
merge(pipeline, "F8:G8");
outputBlock(pipeline, "F8:G8");
label(pipeline, "H8", "Job profile");
merge(pipeline, "I8:K8");
outputBlock(pipeline, "I8:K8");

label(pipeline, "B10", "Recruiter");
merge(pipeline, "C10:D10");
outputBlock(pipeline, "C10:D10");
label(pipeline, "E10", "Location");
merge(pipeline, "F10:G10");
outputBlock(pipeline, "F10:G10");
label(pipeline, "H10", "Hiring manager");
merge(pipeline, "I10:K10");
outputBlock(pipeline, "I10:K10");

sectionTitle(pipeline, "B12:K12", "Quick Add - Internal Candidate");
label(pipeline, "B14", "Candidate lookup");
merge(pipeline, "C14:F14");
inputBlock(pipeline, "C14:F14", COLORS.warning);
label(pipeline, "G14", "Lookup status");
merge(pipeline, "H14:K14");
outputBlock(pipeline, "H14:K14");

const pipelineCandidateRows = [
  ["B16", "Employee ID", "C16:D16", "E16", "Full name", "F16:G16", "H16", "Business title", "I16:K16"],
];
for (const row of pipelineCandidateRows) {
  for (let i = 0; i < row.length; i += 3) {
    label(pipeline, row[i], row[i + 1]);
    merge(pipeline, row[i + 2]);
    outputBlock(pipeline, row[i + 2]);
  }
}

sectionTitle(pipeline, "B18:K18", "Requisition Internal Pipeline");
merge(pipeline, "B19:K19");
value(pipeline, "B19", "Run Lookup Requisition to populate this list from tblCandidateNotes.");
style(pipeline, "B19:K19", {
  fill: COLORS.paleTeal,
  font: { color: COLORS.grayText, italic: true, size: 10 },
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
const pipelineHeaders = [
  ["B21:C21", "Employee ID"],
  ["D21:E21", "Full name"],
  ["F21:G21", "Current title"],
  ["H21", "Screened?"],
  ["I21:K21", "Recruiter synthesis notes"],
];
for (const [address, text] of pipelineHeaders) {
  if (address.includes(":")) merge(pipeline, address);
  value(pipeline, address.split(":")[0], text);
  style(pipeline, address, {
    fill: COLORS.blue,
    font: { color: COLORS.white, bold: true, size: 10 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.line },
  });
}
for (let row = 22; row <= 40; row += 1) {
  const fill = row % 2 === 0 ? COLORS.white : COLORS.lightGray;
  for (const address of [`B${row}:C${row}`, `D${row}:E${row}`, `F${row}:G${row}`, `H${row}`, `I${row}:K${row}`]) {
    if (address.includes(":") && !address.match(/([A-Z]+)\d+:\1\d+/)) merge(pipeline, address);
    style(pipeline, address, {
      fill,
      font: { color: "#111827", size: 10 },
      verticalAlignment: "top",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: COLORS.line },
    });
  }
}
pipeline.freezePanes.freezeRows(21);

// Candidates table
candidates.showGridLines = false;
const candidateHeaders = [
  "submission_id", "added_datetime", "submitted_by", "candidate_lookup_key", "candidate_lookup_status",
  "employee_id", "mm_id", "recent_hire_dt", "date_of_last_mobility_event", "job_profile_nm", "management_level_desc", "business_title_txt",
  "preferred_full_nm", "location_group_desc", "manager_employee_id", "manager_preferred_full_nm",
  "business_group_nm", "sub_business_unit_nm", "division_nm", "job_profile_id", "year_nr", "month_nr",
  "requisition_lookup_key", "requisition_lookup_status", "position_worker_type", "job_requisition_status",
  "requisition_id", "job_posting_title", "add_to_staff_or_replacement", "requisition_job_level",
  "number_of_openings_available", "requisition_bg", "requisition_sbu", "requisition_div",
  "supervisory_organization", "primary_recruiter", "requisition_location", "hiring_managers", "requisition_job_profile",
  "has_candidate_applied", "candidate_stage",
  "screener", "screen_date", "next_step",
  "desired_roles", "desired_level", "desired_function", "mobility_interest", "willing_to_relocate", "skills",
  "career_profile_complete", "development_goals_documented", "recruiter_synthesis_notes",
];
const candidateEndCol = columnName(candidateHeaders.length - 1);
range(candidates, `A1:${candidateEndCol}1`).values = [candidateHeaders];
range(candidates, "B:B").format.numberFormat = "@";
range(candidates, `A2:${candidateEndCol}${candidateRows.length + 1}`).values = candidateRows;
const candidateTable = candidates.tables.add(`A1:${candidateEndCol}${candidateRows.length + 1}`, true);
candidateTable.name = "tblCandidateNotes";
style(candidates, `A1:${candidateEndCol}1`, {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
});
style(candidates, `A2:${candidateEndCol}${candidateRows.length + 1}`, {
  fill: COLORS.white,
  font: { color: "#111827", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
setWidths(candidates, [
  ["A:E", 136], ["F:I", 128], ["J:V", 146], ["W:AM", 142], ["AN:AZ", 132], ["BA:BA", 300],
]);
candidates.freezePanes.freezeRows(1);

// hd_employees table
employeeSheet.showGridLines = false;
const employeeHeaders = [
  "employee_id", "mm_id", "recent_hire_dt", "date_of_last_mobility_event", "job_profile_nm", "management_level_desc", "business_title_txt",
  "preferred_full_nm", "location_group_desc", "manager_employee_id", "manager_preferred_full_nm",
  "business_group_nm", "sub_business_unit_nm", "division_nm", "job_profile_id", "year_nr", "month_nr",
];
range(employeeSheet, "A1:Q1").values = [employeeHeaders];
range(employeeSheet, `A2:Q${employees.length + 1}`).values = employees;
const employeeTable = employeeSheet.tables.add(`A1:Q${employees.length + 1}`, true);
employeeTable.name = "tblHdEmployees";
style(employeeSheet, "A1:Q1", {
  fill: COLORS.teal,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
});
style(employeeSheet, `A2:Q${employees.length + 1}`, {
  fill: COLORS.white,
  font: { color: "#111827", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
setWidths(employeeSheet, [["A:B", 120], ["C:C", 170], ["D:D", 210], ["E:E", 130], ["F:H", 150], ["I:I", 170], ["J:M", 170]]);
employeeSheet.freezePanes.freezeRows(1);

// hd_requisitions table
requisitionSheet.showGridLines = false;
const requisitionHeaders = [
  "Position_Worker_Type", "Job_Requisition_Status", "Requisition_ID", "Job_Posting_Title",
  "Add_to_Staff_or_Replacement", "Job_Level", "Number_of_Openings_Available", "BG", "SBU", "DIV",
  "Supervisory_Organization", "Primary_Recruiter", "Location", "Hiring_Managers", "Job_Profile",
];
const requisitionEndCol = columnName(requisitionHeaders.length - 1);
range(requisitionSheet, `A1:${requisitionEndCol}1`).values = [requisitionHeaders];
range(requisitionSheet, `A2:${requisitionEndCol}${requisitions.length + 1}`).values = requisitions;
const requisitionTable = requisitionSheet.tables.add(`A1:${requisitionEndCol}${requisitions.length + 1}`, true);
requisitionTable.name = "tblHdRequisitions";
style(requisitionSheet, `A1:${requisitionEndCol}1`, {
  fill: COLORS.teal,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
});
style(requisitionSheet, `A2:${requisitionEndCol}${requisitions.length + 1}`, {
  fill: COLORS.white,
  font: { color: "#111827", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
setWidths(requisitionSheet, [
  ["A:C", 130], ["D:D", 220], ["E:G", 150], ["H:J", 120], ["K:K", 210], ["L:N", 170], ["O:O", 180],
]);
requisitionSheet.freezePanes.freezeRows(1);

// input_catalog table
inputCatalog.showGridLines = false;
const inputCatalogHeaders = ["type", "canonical_skill", "aliases", "usage_count", "last_used", "status"];
range(inputCatalog, "A1:F1").values = [inputCatalogHeaders];
range(inputCatalog, `A2:F${inputCatalogRows.length + 1}`).values = inputCatalogRows;
const inputCatalogTable = inputCatalog.tables.add(`A1:F${inputCatalogRows.length + 1}`, true);
inputCatalogTable.name = "tblInputCatalog";
style(inputCatalog, "A1:F1", {
  fill: COLORS.teal,
  font: { color: COLORS.white, bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
});
style(inputCatalog, `A2:F${inputCatalogRows.length + 1}`, {
  fill: COLORS.white,
  font: { color: "#111827", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});
setWidths(inputCatalog, [["A:A", 110], ["B:B", 220], ["C:C", 430], ["D:D", 120], ["E:E", 120], ["F:F", 120]]);
range(inputCatalog, "D:D").format.numberFormat = "0";
range(inputCatalog, "E:E").format.numberFormat = "yyyy-mm-dd";
inputCatalog.freezePanes.freezeRows(1);

// Settings sheet
settings.showGridLines = false;
setWidths(settings, [["A:A", 28], ["B:B", 230], ["C:C", 260], ["D:D", 220], ["E:E", 220], ["F:F", 260], ["G:G", 28]]);
merge(settings, "B2:G2");
value(settings, "B2", "Settings");
style(settings, "B2:G2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 18 },
  verticalAlignment: "center",
});
value(settings, "B4", "Next step values");
range(settings, `B5:B${nextStepValues.length + 4}`).values = nextStepValues.map((item) => [item]);
value(settings, "C4", "Mobility interest");
range(settings, `C5:C${mobilityInterestValues.length + 4}`).values = mobilityInterestValues.map((item) => [item]);
value(settings, "D4", "Yes / No values");
range(settings, `D5:D${yesNoUnknownValues.length + 4}`).values = yesNoUnknownValues.map((item) => [item]);
value(settings, "E4", "Candidate applied");
range(settings, `E5:E${candidateAppliedValues.length + 4}`).values = candidateAppliedValues.map((item) => [item]);
value(settings, "F4", "Candidate stage");
range(settings, `F5:F${candidateStageValues.length + 4}`).values = candidateStageValues.map((item) => [item]);
style(settings, "B4:F4", { fill: COLORS.blue, font: { color: COLORS.white, bold: true }, horizontalAlignment: "center" });
style(settings, `B5:F${Math.max(nextStepValues.length, mobilityInterestValues.length, yesNoUnknownValues.length, candidateAppliedValues.length, candidateStageValues.length) + 4}`, {
  borders: { preset: "outside", style: "thin", color: COLORS.line },
  wrapText: true,
});
value(settings, "B16", "Default screener");
value(settings, "C16", "Recruiter Demo");
style(settings, "B16", {
  fill: COLORS.blue,
  font: { color: COLORS.white, bold: true },
  verticalAlignment: "center",
});
style(settings, "C16", {
  fill: COLORS.warning,
  font: { color: "#111827" },
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: COLORS.line },
});

// Instructions sheet
instructions.showGridLines = false;
setWidths(instructions, [["A:A", 28], ["B:B", 220], ["C:C", 760], ["D:D", 28]]);
setHeights(instructions, [["1:1", 18], ["2:2", 42], ["3:3", 16], ["4:17", 66]]);
merge(instructions, "B2:C2");
value(instructions, "B2", "Recruiter Notetaking Helper - Setup Notes");
style(instructions, "B2:C2", {
  fill: COLORS.navy,
  font: { color: COLORS.white, bold: true, size: 18 },
  verticalAlignment: "center",
});
const instructionRows = [
  ["1. Refresh data", "Weekly Python jobs should replace rows in `tblHdEmployees` on `hd_employees` and `tblHdRequisitions` on `hd_requisitions`. Keep sheet names, table names, and column headers unchanged. `tblHdEmployees` uses the HR source field names, including `preferred_full_nm`, `business_title_txt`, `recent_hire_dt`, `date_of_last_mobility_event`, manager fields, `job_profile_id`, `year_nr`, and `month_nr`."],
  ["2. Lookup candidate", "On `Intake`, enter a full name, employee_id, or mm_id in the Candidate Details lookup box. Run `LookupEmployee.ts` to populate employee information and the five most recent conversation records."],
  ["3. Lookup requisition", "Enter a Requisition_ID or Job_Posting_Title in the Requisition Details lookup box. Run `LookupRequisition.ts` to populate requisition information, including job level. Use the applied/stage dropdowns to capture candidate-requisition status."],
  ["4. Add/delete desired roles", "Enter one role and run `AddDesiredRole.ts`. Fuzzy matches reveal the Similar row; choose an option and run the add script again. Run `DeleteRole.ts` to remove an exact normalized value, or show a not-found message when absent."],
  ["5. Add/delete desired levels/functions", "Enter one level or function and run `AddDesiredLevel.ts` or `AddDesiredFunction.ts`. Run `DeleteLevel.ts` or `DeleteFunction.ts` to remove exact normalized values, or show a not-found message when absent."],
  ["6. Add/delete skills", "Enter one skill and run `AddSkill.ts`. Fuzzy matches reveal the Similar row; choose an option and run the add script again. Run `DeleteSkill.ts` to remove an exact normalized value, or show a not-found message when absent."],
  ["7. Input catalog", "`input_catalog` / `tblInputCatalog` stores canonical values, aliases, usage counts, last-used dates, and status. The `type` column separates skill, role, level, and function records."],
  ["8. Capture notes", "Complete next step, mobility interest, relocation willingness, career profile status, development goals, and recruiter synthesis notes. Employee lookup restores desired roles, desired levels, desired functions, and skills from the latest conversation."],
  ["9. Date and screener", "Screen date sits at the top of the intake form, is blank by default, and supports Excel's date picker. If it remains blank, `SubmitCandidate.ts` saves today's date. Set the reusable screener default in `Settings!C16`."],
  ["10. Submit notes", "Run `SubmitCandidate.ts` to append a new conversation row to `tblCandidateNotes`. Missing required fields appear in the Intake status strip before the script stops. Each submission carries the referenced employee and requisition source fields, receives an added_datetime audit timestamp, and the timeline sorts by screen_date."],
  ["11. Pipeline", "On `Pipeline`, run `LookupRequisition.ts` to populate requisition details and internal candidates from `tblCandidateNotes`. Run `LookupEmployee_Pipeline.ts`, then `SubmitCandidate_Pipeline.ts`, to add preliminary candidate rows with all referenced employee and requisition source fields but without screen notes. Missing required fields appear in Pipeline status before the script stops."],
  ["12. Open candidate notes", "Run `OpenCandidateNotes.ts` after employee lookup to activate `Candidates` and pre-filter `tblCandidateNotes` for the current candidate by employee_id, mm_id, or full name."],
  ["13. Reset form", "Run `ResetIntakeForm.ts` to clear editable intake fields and timeline results without changing labels, formatting, dropdowns, or the underlying data tables."],
  ["14. Excel Online", "Add Automate buttons if desired after uploading the workbook to SharePoint or OneDrive. Candidate detail outputs are locked; use the lookup input and scripts to populate them."],
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
range(intake, "F3").format.numberFormat = "yyyy-mm-dd";
range(intake, "F11").format.numberFormat = "yyyy-mm-dd";
range(intake, "M11").format.numberFormat = "yyyy-mm-dd";
range(pipeline, "M8").format.numberFormat = "yyyy-mm-dd";
range(pipeline, "M20").format.numberFormat = "yyyy-mm-dd";
range(candidates, "B:B").format.numberFormat = "@";
range(candidates, "H:H").format.numberFormat = "yyyy-mm-dd";
range(candidates, "I:I").format.numberFormat = "yyyy-mm-dd";
range(candidates, "AQ:AQ").format.numberFormat = "yyyy-mm-dd";
range(employeeSheet, "C:C").format.numberFormat = "yyyy-mm-dd";
range(employeeSheet, "D:D").format.numberFormat = "yyyy-mm-dd";

// Export workbook and visual previews.
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });
for (const [sheetName, previewRange] of [
  ["Intake", "A1:L70"],
  ["Pipeline", "A1:L45"],
  ["Candidates", "A1:BA8"],
  ["hd_employees", "A1:Q10"],
  ["hd_requisitions", "A1:O8"],
  ["input_catalog", "A1:F24"],
  ["Settings", "A1:G21"],
  ["Instructions", "A1:D16"],
]) {
  const blob = await workbook.render({ sheetName, range: previewRange, scale: 2 });
  await fs.writeFile(`${previewDir}/${sheetName}_v3.png`, Buffer.from(await blob.arrayBuffer()));
}
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
await protectIntakeSheet(path.resolve(outputPath));
console.log(outputPath);

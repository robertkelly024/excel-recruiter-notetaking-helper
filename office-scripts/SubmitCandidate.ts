function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const candidates = workbook.getWorksheet("Candidates");
  const table = candidates.getTable("tblCandidateNotes");
  const protection = intake.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
  setIntakeStatus(intake, "Checking required fields...", "info");

  const text = (address: string): string => {
    const value = intake.getRange(address).getValue();
    return value === null || value === undefined ? "" : String(value).trim();
  };

  const required: [string, string][] = [
    ["Candidate lookup", "C7"],
    ["Full name", "C9"],
    ["Employee ID or MM ID", "F9"],
    ["Requisition lookup", "C21"],
    ["Requisition ID or job title", "C23"],
    ["Screener", "C3"],
  ];

  const missing = required
    .filter(([label, address]) => {
      if (label === "Employee ID or MM ID") {
        return !text("F9") && !text("I9");
      }
      if (label === "Requisition ID or job title") {
        return !text("C23") && !text("F23");
      }
      return !text(address);
    })
    .map(([label]) => label);

  if (missing.length > 0) {
    const message = `Missing required field${missing.length === 1 ? "" : "s"}: ${missing.join(", ")}.`;
    setIntakeStatus(intake, message, "error");
    throw new Error(message);
  }

  const existingRows = table.getRangeBetweenHeaderAndTotal().getRowCount();
  const year = new Date().getFullYear();
  const submissionId = `NOTE-${year}-${String(existingRows + 1).padStart(3, "0")}`;
  const addedDateTime = `${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`;
  const submittedBy = text("C3") || "Screener";
  let screenDate = "";
  try {
    screenDate = dateAsIso(intake.getRange("F3"));
  } catch (error) {
    const message = "Screen date must be blank or a valid date.";
    setIntakeStatus(intake, message, "error");
    throw new Error(message);
  }

  const row: string[] = [
    submissionId,
    addedDateTime,
    submittedBy,
    text("C7"),
    text("H7"),
    text("F9"),
    text("I9"),
    text("C9"),
    text("C11"),
    text("M7"),
    text("F11"),
    text("I11"),
    text("C13"),
    text("F13"),
    text("I13"),
    text("C15"),
    text("F15"),
    text("I15"),
    text("C21"),
    text("H21"),
    text("C23"),
    text("F23"),
    text("C25"),
    text("F25"),
    text("I25"),
    text("I23"),
    text("C27"),
    text("F27"),
    text("I27"),
    text("C3"),
    screenDate,
    text("I3"),
    text("G30"),
    text("G32"),
    text("G34"),
    text("C40"),
    text("F40"),
    text("G36"),
    text("I40"),
    text("K40"),
    text("B43"),
  ];

  table.addRow(-1, row);
  clearIntake(intake);
  setIntakeStatus(intake, `Submitted ${submissionId} at ${formatConversationDate(addedDateTime)}.`, "success");
  intake.getRange("H7").setValue(`Submitted ${submissionId} at ${formatConversationDate(addedDateTime)}.`);
  intake.activate();
  } finally {
    if (wasProtected) protection.protect();
  }
}

function clearIntake(sheet: ExcelScript.Worksheet) {
  [
    "C7", "H7",
    "C9", "F9", "I9",
    "C11", "F11", "I11", "M7",
    "C13", "F13", "I13",
    "C15", "F15", "I15",
    "C21", "H21",
    "C23", "F23", "I23",
    "C25", "F25", "I25",
    "C27", "F27", "I27",
    "C3", "F3", "I3",
    "C30", "G30",
    "C32", "G32",
    "C34", "G34",
    "C36", "G36",
    "E31", "E33", "E35", "E37", "M2", "M3", "M4", "M5", "M6",
    "C40", "F40", "I40", "K40",
    "B43",
    "B51",
    "B54", "D54", "G54", "I54", "B55",
    "B57", "D57", "G57", "I57", "B58",
    "B60", "D60", "G60", "I60", "B61",
    "B63", "D63", "G63", "I63", "B64",
    "B66", "D66", "G66", "I66", "B67",
  ].forEach((address) => sheet.getRange(address).setValue(""));

  sheet.getRange("C3").setFormula("=Settings!$C$16");
  setIntakeStatus(sheet, "Ready to submit when required candidate, requisition, and screener fields are complete.", "info");
  sheet.getRange("B31").setValue("Similar roles found. Select intended role from dropdown:");
  sheet.getRange("B33").setValue("Similar desired levels found. Select intended level from dropdown:");
  sheet.getRange("B35").setValue("Similar desired functions found. Select intended function from dropdown:");
  sheet.getRange("B37").setValue("Similar skills found. Select intended skill from dropdown:");
  ["E31", "E33", "E35", "E37"].forEach((address) => {
    sheet.getRange(address).getFormat().getFill().setColor("#EAF2F8");
  });
  ["31:31", "33:33", "35:35", "37:37"].forEach((address) => {
    sheet.getRange(address).setRowHidden(true);
  });
  sheet.getRange("B51").setValue("Run Lookup Employee to display recent conversations. Use Open Candidate Notes to view this candidate in tblCandidateNotes.");
}

function setIntakeStatus(sheet: ExcelScript.Worksheet, message: string, tone: string) {
  const statusRange = sheet.getRange("B4:K4");
  const statusCell = sheet.getRange("B4");
  statusCell.setValue(message);
  statusRange.getFormat().setWrapText(true);
  statusRange.getFormat().getFill().setColor(tone === "error" ? "#FDECEC" : "#E7F3F1");
  statusRange.getFormat().getFont().setColor(tone === "error" ? "#7F1D1D" : "#475569");
  statusRange.getFormat().getFont().setItalic(tone !== "error");
}

function dateAsIso(dateRange: ExcelScript.Range): string {
  const raw = dateRange.getCell(0, 0).getValue();
  if (typeof raw === "number" && Number.isFinite(raw)) {
    const milliseconds = Math.round(raw * 86400000);
    return new Date(Date.UTC(1899, 11, 30) + milliseconds).toISOString().slice(0, 10);
  }

  const value = raw === null || raw === undefined ? "" : String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;

  const parsed = value ? new Date(value) : new Date();
  if (Number.isNaN(parsed.getTime())) {
    throw new Error("Screen date must be a valid date.");
  }
  return parsed.toISOString().slice(0, 10);
}

function formatConversationDate(value: string): string {
  return value
    .replace("T", " ")
    .replace(/:\d{2}(?:\.\d+)?Z?$/, "")
    .trim();
}

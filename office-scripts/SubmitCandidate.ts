function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const candidates = workbook.getWorksheet("Candidates");
  const table = candidates.getTable("tblCandidateNotes");

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
    ["Stage", "C32"],
    ["Screener", "F32"],
    ["Desired role(s)", "C39"],
    ["Mobility interest", "C41"],
    ["Recruiter synthesis / follow-up notes", "B46"],
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
    throw new Error(`Complete required fields before submitting: ${missing.join(", ")}.`);
  }

  const existingRows = table.getRangeBetweenHeaderAndTotal().getRowCount();
  const year = new Date().getFullYear();
  const submissionId = `NOTE-${year}-${String(existingRows + 1).padStart(3, "0")}`;
  const conversationDateTime = `${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`;
  const submittedBy = text("F32") || "Screener";
  const screenDate = dateAsIso(intake.getRange("I32"));

  const row: string[] = [
    submissionId,
    conversationDateTime,
    submittedBy,
    text("C7"),
    text("H7"),
    text("F9"),
    text("I9"),
    text("C9"),
    text("C11"),
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
    text("C32"),
    text("F32"),
    screenDate,
    text("C34"),
    text("C39"),
    text("H39"),
    text("C41"),
    text("F43"),
    text("F41"),
    text("I41"),
    text("B46"),
  ];

  table.addRow(-1, row);
  clearIntake(intake);
  intake.getRange("H7").setValue(`Submitted ${submissionId} at ${formatConversationDate(conversationDateTime)}.`);
  intake.activate();
}

function clearIntake(sheet: ExcelScript.Worksheet) {
  [
    "C7", "H7",
    "C9", "F9", "I9",
    "C11", "F11", "I11",
    "C13", "F13", "I13",
    "C15", "F15", "I15",
    "C21", "H21",
    "C23", "F23", "I23",
    "C25", "F25", "I25",
    "C27",
    "C32", "F32", "I32",
    "C34",
    "C39", "H39",
    "C41", "F41", "I41",
    "C43", "F43",
    "B46",
    "B54",
    "B57", "D57", "F57", "I57", "B58",
    "B59", "D59", "F59", "I59", "B60",
    "B61", "D61", "F61", "I61", "B62",
    "B63", "D63", "F63", "I63", "B64",
    "B65", "D65", "F65", "I65", "B66",
  ].forEach((address) => sheet.getRange(address).setValue(""));

  sheet.getRange("F32").setFormula("=Settings!$C$16");
  sheet.getRange("I32").setFormula("=TODAY()");
  sheet.getRange("B54").setValue("Run Lookup Employee to display the five most recent conversations for this candidate.");
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

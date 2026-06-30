function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const protection = intake.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
  [
    "C7", "H7",
    "C9", "F9", "I9",
    "C11", "F11", "I11", "M7", "M8", "M9", "M10", "M11",
    "C13", "F13", "I13",
    "C15", "F15", "I15",
    "C21", "H21",
    "M21", "M22", "M23", "M24", "M25", "M26", "M27", "M28",
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
  ].forEach((address) => intake.getRange(address).setValue(""));

  intake.getRange("C3").setFormula("=Settings!$C$16");
  setIntakeStatus(intake, "Ready to submit when required candidate, requisition, and screener fields are complete.", "info");
  intake.getRange("B31").setValue("Similar roles found. Select intended role from dropdown:");
  intake.getRange("B33").setValue("Similar desired levels found. Select intended level from dropdown:");
  intake.getRange("B35").setValue("Similar desired functions found. Select intended function from dropdown:");
  intake.getRange("B37").setValue("Similar skills found. Select intended skill from dropdown:");
  ["E31", "E33", "E35", "E37"].forEach((address) => {
    intake.getRange(address).getFormat().getFill().setColor("#EAF2F8");
  });
  ["31:31", "33:33", "35:35", "37:37"].forEach((address) => {
    intake.getRange(address).setRowHidden(true);
  });
  intake.getRange("B51").setValue("Run Lookup Employee to display recent conversations. Use Open Candidate Notes to view this candidate in tblCandidateNotes.");
  intake.activate();
  } finally {
    if (wasProtected) protection.protect();
  }
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

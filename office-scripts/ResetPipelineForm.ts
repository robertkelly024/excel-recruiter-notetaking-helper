function main(workbook: ExcelScript.Workbook) {
  const pipeline = workbook.getWorksheet("Pipeline");
  const protection = pipeline.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
    [
      "C6", "H6",
      "C8", "F8", "I8",
      "C10", "F10", "I10",
      "M3", "M4", "M5",
      "C14", "H14",
      "C16", "F16", "I16",
      "M7", "M8", "M9", "M10", "M11", "M12", "M13", "M14", "M15", "M16", "M17", "M18", "M19", "M20",
      "M21", "M22", "M23", "M24", "M25", "M26", "M27", "M28",
    ].forEach((address) => pipeline.getRange(address).setValue(""));

    pipeline.getRange("C3").setFormula("=Settings!$C$16");
    clearPipelineRows(pipeline);
    pipeline.getRange("B19").setValue("Run Lookup Requisition to populate this list from tblCandidateNotes.");
    setPipelineStatus(pipeline, "Pipeline form reset. Run Lookup Requisition to begin.", "info");
    pipeline.activate();
  } finally {
    if (wasProtected) protection.protect();
  }
}

function clearPipelineRows(sheet: ExcelScript.Worksheet) {
  for (let row = 22; row <= 40; row += 1) {
    ["B", "D", "F", "H", "I"].forEach((column) => {
      sheet.getRange(`${column}${row}`).setValue("");
    });
  }
}

function setPipelineStatus(sheet: ExcelScript.Worksheet, message: string, tone: string) {
  const statusRange = sheet.getRange("F3:K3");
  const statusCell = sheet.getRange("F3");
  statusCell.setValue(message);
  statusRange.getFormat().setWrapText(true);
  statusRange.getFormat().getFill().setColor(tone === "error" ? "#FDECEC" : "#E7F3F1");
  statusRange.getFormat().getFont().setColor(tone === "error" ? "#7F1D1D" : "#475569");
  statusRange.getFormat().getFont().setItalic(tone !== "error");
}

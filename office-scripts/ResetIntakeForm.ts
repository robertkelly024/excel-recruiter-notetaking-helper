function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");

  [
    "C7:F7", "H7:K7",
    "C9:D9", "F9:G9", "I9:K9",
    "C11:D11", "F11:G11", "I11:K11",
    "C13:D13", "F13:G13", "I13:K13",
    "C15:D15", "F15:G15", "I15:K15",
    "C21:D21", "F21:G21", "I21:K21",
    "C23:D23", "F23:G23", "I23:K23",
    "D28:D34", "E28:K34",
    "B37:K43",
  ].forEach((address) => intake.getRange(address).clear(ExcelScript.ClearApplyTo.contents));

  intake.activate();
}

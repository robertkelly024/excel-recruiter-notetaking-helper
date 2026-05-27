function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  [
    "C6:D6", "F6:G6", "I6",
    "C8:D8", "F8:G8", "I8",
    "C10:D10", "F10:G10", "I10",
    "C12:D12", "F12:G12", "I12",
    "C18:C24", "D18:D24",
    "F18:I18", "F20:I20", "F22:I22",
  ].forEach((address) => intake.getRange(address).clear(ExcelScript.ClearApplyTo.contents));
  intake.activate();
}

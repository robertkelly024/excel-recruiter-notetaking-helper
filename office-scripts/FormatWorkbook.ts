function main(workbook: ExcelScript.Workbook) {
  const candidates = workbook.getWorksheet("Candidates");
  const intake = workbook.getWorksheet("Intake");
  const dashboard = workbook.getWorksheet("Dashboard");
  const settings = workbook.getWorksheet("Settings");
  const instructions = workbook.getWorksheet("Instructions");

  [dashboard, intake, candidates, settings, instructions].forEach((sheet) => {
    sheet.setShowGridlines(false);
  });

  const candidateTable = candidates.getTable("Candidates");
  candidateTable.setPredefinedTableStyle("TableStyleMedium4");
  candidateTable.setShowFilterButton(true);

  candidates.getFreezePanes().freezeRows(4);
  intake.getFreezePanes().freezeRows(5);
  dashboard.getFreezePanes().freezeRows(4);

  setWidths(candidates, {
    A: 120, B: 170, C: 145, D: 155, E: 120, F: 145, G: 145, H: 150, I: 130, J: 120, K: 170,
    L: 105, M: 105, N: 105, O: 105, P: 125, Q: 135, R: 105, S: 120, T: 150, U: 260, V: 260, W: 160, X: 165, Y: 170,
  });
  setWidths(intake, { B: 170, C: 120, D: 300, E: 120, F: 170, G: 170, H: 170, I: 170 });
  setWidths(dashboard, { B: 155, C: 125, E: 150, F: 110, H: 150, I: 130, J: 150, K: 150 });

  candidates.getRange("H:H").setNumberFormat("yyyy-mm-dd");
  candidates.getRange("X:X").setNumberFormat("yyyy-mm-dd hh:mm");
  candidates.getRange("L:R").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  candidates.getRange("A4:Y4").getFormat().getFont().setBold(true);

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  intake.activate();
  intake.getRange("C6").select();
}

function setWidths(sheet: ExcelScript.Worksheet, widths: Record<string, number>) {
  Object.keys(widths).forEach((column) => {
    sheet.getRange(`${column}:${column}`).getFormat().setColumnWidth(widths[column]);
  });
}

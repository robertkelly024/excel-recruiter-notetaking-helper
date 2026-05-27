function main(workbook: ExcelScript.Workbook) {
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  const dashboard = workbook.getWorksheet("Dashboard");
  dashboard.activate();
  dashboard.getRange("B2").select();
}

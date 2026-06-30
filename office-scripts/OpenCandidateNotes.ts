function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const candidates = workbook.getWorksheet("Candidates");
  const table = candidates.getTable("tblCandidateNotes");

  const employeeId = clean(intake.getRange("F9").getValue());
  const mmId = clean(intake.getRange("I9").getValue());
  const fullName = clean(intake.getRange("C9").getValue());

  if (!employeeId && !mmId && !fullName) {
    throw new Error("Lookup an employee before opening candidate notes.");
  }

  table.getAutoFilter().clearCriteria();

  if (employeeId) {
    table.getColumnByName("employee_id").getFilter().applyValuesFilter([employeeId]);
  } else if (mmId) {
    table.getColumnByName("mm_id").getFilter().applyValuesFilter([mmId]);
  } else {
    table.getColumnByName("full_name").getFilter().applyValuesFilter([fullName]);
  }

  candidates.activate();
  table.getRange().getCell(0, 0).select();
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const employees = workbook.getWorksheet("hd_employees");
  const table = employees.getTable("tblHdEmployees");

  const lookupValue = normalize(intake.getRange("C7").getValue());
  clearCandidateDetails(intake);

  if (!lookupValue) {
    intake.getRange("H7").setValue("Enter a full name, employee_id, or mm_id before running lookup.");
    return;
  }

  const headers = table.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const col = (name: string): number => {
    const index = headers.indexOf(name);
    if (index < 0) throw new Error(`Missing required column in tblHdEmployees: ${name}`);
    return index;
  };

  const indexes = {
    employeeId: col("employee_id"),
    mmId: col("mm_id"),
    fullName: col("full_name"),
    businessTitle: col("business_title"),
    recentHireDate: col("recent_hire_date"),
    location: col("location"),
    jobProfile: col("job_profile"),
    managementLevel: col("management_level"),
    managerFullName: col("manager_full_name"),
    businessGroup: col("business_group"),
    subBusinessGroup: col("sub_business_group"),
    division: col("division"),
  };

  const exactIdMatches = rows.filter((row) =>
    normalize(row[indexes.employeeId]) === lookupValue || normalize(row[indexes.mmId]) === lookupValue
  );
  const exactNameMatches = rows.filter((row) => normalize(row[indexes.fullName]) === lookupValue);
  const partialNameMatches = rows.filter((row) => normalize(row[indexes.fullName]).includes(lookupValue));

  const matches = exactIdMatches.length > 0
    ? exactIdMatches
    : exactNameMatches.length > 0
      ? exactNameMatches
      : partialNameMatches;

  if (matches.length === 0) {
    intake.getRange("H7").setValue("No employee found. Check the spelling, employee_id, or mm_id.");
    return;
  }

  if (matches.length > 1) {
    const preview = matches
      .slice(0, 5)
      .map((row) => `${asText(row[indexes.fullName])} (${asText(row[indexes.employeeId])})`)
      .join("; ");
    intake.getRange("H7").setValue(`Multiple matches found: ${preview}. Search by employee_id or mm_id.`);
    return;
  }

  const row = matches[0];
  intake.getRange("C9").setValue(asText(row[indexes.fullName]));
  intake.getRange("F9").setValue(asText(row[indexes.employeeId]));
  intake.getRange("I9").setValue(asText(row[indexes.mmId]));
  intake.getRange("C11").setValue(asText(row[indexes.businessTitle]));
  intake.getRange("F11").setValue(asText(row[indexes.recentHireDate]));
  intake.getRange("I11").setValue(asText(row[indexes.location]));
  intake.getRange("C13").setValue(asText(row[indexes.jobProfile]));
  intake.getRange("F13").setValue(asText(row[indexes.managementLevel]));
  intake.getRange("I13").setValue(asText(row[indexes.managerFullName]));
  intake.getRange("C15").setValue(asText(row[indexes.businessGroup]));
  intake.getRange("F15").setValue(asText(row[indexes.subBusinessGroup]));
  intake.getRange("I15").setValue(asText(row[indexes.division]));

  const matchType = exactIdMatches.length > 0
    ? "employee_id/mm_id"
    : exactNameMatches.length > 0
      ? "full name"
      : "partial full name";
  intake.getRange("H7").setValue(`Matched on ${matchType}: ${asText(row[indexes.fullName])}`);
  intake.activate();
}

function normalize(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ").toLowerCase();
}

function asText(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function clearCandidateDetails(sheet: ExcelScript.Worksheet) {
  [
    "H7:K7",
    "C9:D9", "F9:G9", "I9:K9",
    "C11:D11", "F11:G11", "I11:K11",
    "C13:D13", "F13:G13", "I13:K13",
    "C15:D15", "F15:G15", "I15:K15",
  ].forEach((address) => sheet.getRange(address).clear(ExcelScript.ClearApplyTo.contents));
}

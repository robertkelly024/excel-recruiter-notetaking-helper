function main(workbook: ExcelScript.Workbook) {
  const pipeline = workbook.getWorksheet("Pipeline");
  const employees = workbook.getWorksheet("hd_employees");
  const table = employees.getTable("tblHdEmployees");

  const lookupValue = normalize(pipeline.getRange("C14").getValue());
  clearCandidateDetails(pipeline);

  if (!lookupValue) {
    pipeline.getRange("H14").setValue("Enter a full name, employee_id, or mm_id before running lookup.");
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
    dateOfLastMobilityEvent: col("date_of_last_mobility_event"),
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
    pipeline.getRange("H14").setValue("No employee found. Check the spelling, employee_id, or mm_id.");
    return;
  }

  if (matches.length > 1) {
    const preview = matches
      .slice(0, 5)
      .map((row) => `${asText(row[indexes.fullName])} (${asText(row[indexes.employeeId])})`)
      .join("; ");
    pipeline.getRange("H14").setValue(`Multiple matches found: ${preview}. Search by employee_id or mm_id.`);
    return;
  }

  const row = matches[0];
  pipeline.getRange("C16").setValue(asText(row[indexes.employeeId]));
  pipeline.getRange("F16").setValue(asText(row[indexes.fullName]));
  pipeline.getRange("I16").setValue(asText(row[indexes.businessTitle]));
  pipeline.getRange("M7").setValue(asText(row[indexes.recentHireDate]));
  pipeline.getRange("M8").setValue(asText(row[indexes.mmId]));
  pipeline.getRange("M9").setValue(asText(row[indexes.dateOfLastMobilityEvent]));
  pipeline.getRange("M10").setValue(asText(row[indexes.location]));
  pipeline.getRange("M11").setValue(asText(row[indexes.jobProfile]));
  pipeline.getRange("M12").setValue(asText(row[indexes.managementLevel]));
  pipeline.getRange("M13").setValue(asText(row[indexes.managerFullName]));
  pipeline.getRange("M14").setValue(asText(row[indexes.businessGroup]));
  pipeline.getRange("M15").setValue(asText(row[indexes.subBusinessGroup]));
  pipeline.getRange("M16").setValue(asText(row[indexes.division]));

  const matchType = exactIdMatches.length > 0
    ? "employee_id/mm_id"
    : exactNameMatches.length > 0
      ? "full name"
      : "partial full name";
  pipeline.getRange("H14").setValue(`Matched on ${matchType}: ${asText(row[indexes.fullName])}`);
  pipeline.activate();
}

function clearCandidateDetails(sheet: ExcelScript.Worksheet) {
  [
    "H14",
    "C16", "F16", "I16",
    "M7", "M8", "M9", "M10", "M11", "M12", "M13", "M14", "M15", "M16",
  ].forEach((address) => sheet.getRange(address).setValue(""));
}

function normalize(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ").toLowerCase();
}

function asText(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

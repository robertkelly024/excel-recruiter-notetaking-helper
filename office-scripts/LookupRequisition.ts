function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const requisitions = workbook.getWorksheet("hd_requisitions");
  const table = requisitions.getTable("tblHdRequisitions");

  const lookupValue = normalize(intake.getRange("C21").getValue());
  clearRequisitionDetails(intake);

  if (!lookupValue) {
    intake.getRange("H21").setValue("Enter a Requisition_ID or Job_Posting_Title before running lookup.");
    return;
  }

  const headers = table.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const col = (name: string): number => {
    const index = headers.indexOf(name);
    if (index < 0) throw new Error(`Missing required column in tblHdRequisitions: ${name}`);
    return index;
  };

  const indexes = {
    requisitionId: col("Requisition_ID"),
    jobPostingTitle: col("Job_Posting_Title"),
    primaryRecruiter: col("Primary_Recruiter"),
    location: col("Location"),
    hiringManagers: col("Hiring_Managers"),
    jobDescription: col("job_description"),
    jobProfile: col("Job_Profile"),
  };

  const exactIdMatches = rows.filter((row) => normalize(row[indexes.requisitionId]) === lookupValue);
  const exactTitleMatches = rows.filter((row) => normalize(row[indexes.jobPostingTitle]) === lookupValue);
  const partialTitleMatches = rows.filter((row) => normalize(row[indexes.jobPostingTitle]).includes(lookupValue));

  const matches = exactIdMatches.length > 0
    ? exactIdMatches
    : exactTitleMatches.length > 0
      ? exactTitleMatches
      : partialTitleMatches;

  if (matches.length === 0) {
    intake.getRange("H21").setValue("No requisition found. Check the Requisition_ID or Job_Posting_Title.");
    return;
  }

  if (matches.length > 1) {
    const preview = matches
      .slice(0, 5)
      .map((row) => `${asText(row[indexes.jobPostingTitle])} (${asText(row[indexes.requisitionId])})`)
      .join("; ");
    intake.getRange("H21").setValue(`Multiple matches found: ${preview}. Search by Requisition_ID.`);
    return;
  }

  const row = matches[0];
  intake.getRange("C23").setValue(asText(row[indexes.requisitionId]));
  intake.getRange("F23").setValue(asText(row[indexes.jobPostingTitle]));
  intake.getRange("C25").setValue(asText(row[indexes.primaryRecruiter]));
  intake.getRange("F25").setValue(asText(row[indexes.location]));
  intake.getRange("I25").setValue(asText(row[indexes.hiringManagers]));
  intake.getRange("I23").setValue(asText(row[indexes.jobProfile]));
  intake.getRange("C27").setValue(asText(row[indexes.jobDescription]));

  const matchType = exactIdMatches.length > 0
    ? "requisition_id"
    : exactTitleMatches.length > 0
      ? "job_posting_title"
      : "partial job_posting_title";
  intake.getRange("H21").setValue(`Matched on ${matchType}: ${asText(row[indexes.jobPostingTitle])}`);
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

function clearRequisitionDetails(sheet: ExcelScript.Worksheet) {
  [
    "H21:K21",
    "C23:D23", "F23:G23", "I23:J23",
    "C25:D25", "F25:G25", "I25:K25",
    "C27:K27",
  ].forEach((address) => sheet.getRange(address).clear(ExcelScript.ClearApplyTo.contents));
}

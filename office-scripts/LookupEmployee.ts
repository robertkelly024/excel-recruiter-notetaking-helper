function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const employees = workbook.getWorksheet("hd_employees");
  const candidates = workbook.getWorksheet("Candidates");
  const table = employees.getTable("tblHdEmployees");
  const notesTable = candidates.getTable("tblCandidateNotes");

  const lookupValue = normalize(intake.getRange("C7").getValue());
  clearCandidateDetails(intake);
  clearTimeline(intake);

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
  populateTimeline(
    intake,
    notesTable,
    asText(row[indexes.employeeId]),
    asText(row[indexes.mmId]),
    asText(row[indexes.businessGroup]),
    asText(row[indexes.businessTitle])
  );
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
    "H7",
    "C9", "F9", "I9",
    "C11", "F11", "I11",
    "C13", "F13", "I13",
    "C15", "F15", "I15",
  ].forEach((address) => sheet.getRange(address).setValue(""));
}

function clearTimeline(sheet: ExcelScript.Worksheet) {
  sheet.getRange("G39").setValue("");
  sheet.getRange("C45").setValue("");
  sheet.getRange("C47").setValue("");
  sheet.getRange("C49").setValue("");
  sheet.getRange("B60").setValue("Run Lookup Employee to display the five most recent conversations for this candidate.");
  for (let index = 0; index < 5; index += 1) {
    const summaryRow = 63 + index * 2;
    sheet.getRange(`B${summaryRow}`).setValue("");
    sheet.getRange(`D${summaryRow}`).setValue("");
    sheet.getRange(`G${summaryRow}`).setValue("");
    sheet.getRange(`I${summaryRow}`).setValue("");
    sheet.getRange(`B${summaryRow + 1}`).setValue("");
  }
}

function populateTimeline(
  sheet: ExcelScript.Worksheet,
  table: ExcelScript.Table,
  employeeId: string,
  mmId: string,
  businessGroup: string,
  businessTitle: string
) {
  const headers = table.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const col = (name: string): number => {
    const index = headers.indexOf(name);
    if (index < 0) throw new Error(`Missing required column in tblCandidateNotes: ${name}`);
    return index;
  };

  const indexes = {
    conversationDateTime: col("conversation_datetime"),
    employeeId: col("employee_id"),
    mmId: col("mm_id"),
    requisitionId: col("requisition_id"),
    jobTitle: col("job_posting_title"),
    stage: col("stage"),
    nextStep: col("next_step"),
    desiredRoles: col("desired_roles"),
    currentSkills: col("current_skills"),
    developingSkills: col("developing_skills"),
    aspirationalSkills: col("aspirational_skills"),
    notes: col("recruiter_synthesis_notes"),
  };

  const normalizedEmployeeId = normalize(employeeId);
  const normalizedMmId = normalize(mmId);
  const matches = rows
    .map((row, sourceIndex) => ({ row, sourceIndex }))
    .filter(({ row }) =>
      (normalizedEmployeeId && normalize(row[indexes.employeeId]) === normalizedEmployeeId) ||
      (normalizedMmId && normalize(row[indexes.mmId]) === normalizedMmId)
    )
    .sort((left, right) => {
      const rightTime = conversationTimestamp(right.row[indexes.conversationDateTime]);
      const leftTime = conversationTimestamp(left.row[indexes.conversationDateTime]);
      if (rightTime !== null && leftTime !== null) return rightTime - leftTime;
      return right.sourceIndex - left.sourceIndex;
    });

  if (matches.length === 0) {
    sheet.getRange("B60").setValue("No prior conversations found. The next submitted conversation will appear here.");
    return;
  }

  const visible = matches.slice(0, 5);
  const latest = matches[0].row;
  sheet.getRange("G39").setValue(asText(latest[indexes.desiredRoles]));
  sheet.getRange("C45").setValue(asText(latest[indexes.currentSkills]));
  sheet.getRange("C47").setValue(asText(latest[indexes.developingSkills]));
  sheet.getRange("C49").setValue(asText(latest[indexes.aspirationalSkills]));

  sheet.getRange("B60").setValue(
    `Showing ${visible.length} of ${matches.length} conversation${matches.length === 1 ? "" : "s"}, most recent first.`
  );

  visible.forEach(({ row }, index) => {
    const summaryRow = 63 + index * 2;
    const title = asText(row[indexes.jobTitle]);
    const requisitionId = asText(row[indexes.requisitionId]);
    const role = title && requisitionId ? `${title} (${requisitionId})` : title || requisitionId;
    const currentBgAndTitle = [businessGroup, businessTitle]
      .filter((item) => item.length > 0)
      .join(" | ");
    const stage = asText(row[indexes.stage]);
    const nextStep = asText(row[indexes.nextStep]);
    const stageAndNext = [stage, nextStep ? `Next: ${nextStep}` : ""]
      .filter((item) => item.length > 0)
      .join(" | ");

    sheet.getRange(`B${summaryRow}`).setValue(formatConversationDate(row[indexes.conversationDateTime]));
    sheet.getRange(`D${summaryRow}`).setValue(currentBgAndTitle);
    sheet.getRange(`G${summaryRow}`).setValue(role);
    sheet.getRange(`I${summaryRow}`).setValue(stageAndNext);
    sheet.getRange(`B${summaryRow + 1}`).setValue(asText(row[indexes.notes]));
  });
}

function conversationTimestamp(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Date.UTC(1899, 11, 30) + Math.round(value * 86400000);
  }
  const parsed = Date.parse(asText(value));
  return Number.isNaN(parsed) ? null : parsed;
}

function formatConversationDate(value: unknown): string {
  const timestamp = conversationTimestamp(value);
  if (timestamp !== null) {
    return new Date(timestamp).toISOString().slice(0, 10);
  }
  return asText(value).slice(0, 10);
}

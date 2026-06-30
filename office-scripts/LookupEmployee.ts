function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const employees = workbook.getWorksheet("hd_employees");
  const candidates = workbook.getWorksheet("Candidates");
  const table = employees.getTable("tblHdEmployees");
  const notesTable = candidates.getTable("tblCandidateNotes");
  const protection = intake.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
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
  intake.getRange("F11").setValue(asText(row[indexes.dateOfLastMobilityEvent]));
  intake.getRange("M7").setValue(asText(row[indexes.recentHireDate]));
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
  } finally {
    if (wasProtected) protection.protect();
  }
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
    "C11", "F11", "I11", "M7",
    "C13", "F13", "I13",
    "C15", "F15", "I15",
  ].forEach((address) => sheet.getRange(address).setValue(""));
}

function clearTimeline(sheet: ExcelScript.Worksheet) {
  sheet.getRange("G30").setValue("");
  sheet.getRange("G34").setValue("");
  sheet.getRange("G32").setValue("");
  sheet.getRange("G36").setValue("");
  ["E31", "E33", "E35", "E37", "M2", "M3", "M4", "M5", "M6"].forEach((address) => {
    sheet.getRange(address).setValue("");
  });
  sheet.getRange("B31").setValue("Similar roles found. Select intended role from dropdown:");
  sheet.getRange("B33").setValue("Similar desired levels found. Select intended level from dropdown:");
  sheet.getRange("B35").setValue("Similar desired functions found. Select intended function from dropdown:");
  sheet.getRange("B37").setValue("Similar skills found. Select intended skill from dropdown:");
  ["E31", "E33", "E35", "E37"].forEach((address) => {
    sheet.getRange(address).getFormat().getFill().setColor("#EAF2F8");
  });
  ["31:31", "33:33", "35:35", "37:37"].forEach((address) => {
    sheet.getRange(address).setRowHidden(true);
  });
  sheet.getRange("B51").setValue("Run Lookup Employee to display recent conversations. Use Open Candidate Notes to view this candidate in tblCandidateNotes.");
  for (let index = 0; index < 5; index += 1) {
    const summaryRow = 54 + index * 3;
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
    addedDateTime: col("added_datetime"),
    screenDate: col("screen_date"),
    employeeId: col("employee_id"),
    mmId: col("mm_id"),
    requisitionId: col("requisition_id"),
    jobTitle: col("job_posting_title"),
    nextStep: col("next_step"),
    desiredRoles: col("desired_roles"),
    desiredLevel: col("desired_level"),
    desiredFunction: col("desired_function"),
    skills: col("skills"),
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
      const rightTime = conversationTimestamp(right.row[indexes.screenDate]) ?? conversationTimestamp(right.row[indexes.addedDateTime]);
      const leftTime = conversationTimestamp(left.row[indexes.screenDate]) ?? conversationTimestamp(left.row[indexes.addedDateTime]);
      if (rightTime !== null && leftTime !== null) return rightTime - leftTime;
      return right.sourceIndex - left.sourceIndex;
    });

  if (matches.length === 0) {
    sheet.getRange("B51").setValue("No prior conversations found. The next submitted conversation will appear here.");
    return;
  }

  const visible = matches.slice(0, 5);
  const latest = matches[0].row;
  sheet.getRange("G30").setValue(asText(latest[indexes.desiredRoles]));
  sheet.getRange("G32").setValue(asText(latest[indexes.desiredLevel]));
  sheet.getRange("G34").setValue(asText(latest[indexes.desiredFunction]));
  sheet.getRange("G36").setValue(asText(latest[indexes.skills]));

  sheet.getRange("B51").setValue(
    `Showing ${visible.length} of ${matches.length} conversation${matches.length === 1 ? "" : "s"}, most recent first. Use Open Candidate Notes to view the filtered table.`
  );

  visible.forEach(({ row }, index) => {
    const summaryRow = 54 + index * 3;
    const title = asText(row[indexes.jobTitle]);
    const requisitionId = asText(row[indexes.requisitionId]);
    const role = title && requisitionId ? `${title} (${requisitionId})` : title || requisitionId;
    const currentBgAndTitle = [businessGroup, businessTitle]
      .filter((item) => item.length > 0)
      .join(" | ");
    const nextStep = asText(row[indexes.nextStep]);

    sheet.getRange(`B${summaryRow}`).setValue(formatConversationDate(row[indexes.screenDate]));
    sheet.getRange(`D${summaryRow}`).setValue(currentBgAndTitle);
    sheet.getRange(`G${summaryRow}`).setValue(role);
    sheet.getRange(`I${summaryRow}`).setValue(nextStep);
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

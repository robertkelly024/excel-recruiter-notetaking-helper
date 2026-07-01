function main(workbook: ExcelScript.Workbook) {
  const activeSheet = workbook.getActiveWorksheet();
  const target = activeSheet.getName() === "Pipeline"
    ? workbook.getWorksheet("Pipeline")
    : workbook.getWorksheet("Intake");
  const config = requisitionConfig(target.getName());
  const requisitions = workbook.getWorksheet("hd_requisitions");
  const table = requisitions.getTable("tblHdRequisitions");
  const protection = target.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
    const lookupValue = normalize(target.getRange(config.lookupCell).getValue());
    clearRequisitionDetails(target, config);

    if (!lookupValue) {
      target.getRange(config.statusCell).setValue("Enter a Requisition_ID or Job_Posting_Title before running lookup.");
      if (target.getName() === "Pipeline") clearPipelineList(target, "Run Lookup Requisition to populate this list from tblCandidateNotes.");
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
      positionWorkerType: col("Position_Worker_Type"),
      jobRequisitionStatus: col("Job_Requisition_Status"),
      requisitionId: col("Requisition_ID"),
      jobPostingTitle: col("Job_Posting_Title"),
      addToStaffOrReplacement: col("Add_to_Staff_or_Replacement"),
      jobLevel: col("Job_Level"),
      numberOfOpeningsAvailable: col("Number_of_Openings_Available"),
      bg: col("BG"),
      sbu: col("SBU"),
      div: col("DIV"),
      supervisoryOrganization: col("Supervisory_Organization"),
      primaryRecruiter: col("Primary_Recruiter"),
      location: col("Location"),
      hiringManagers: col("Hiring_Managers"),
      jobProfile: col("Job_Profile"),
    };

    const exactIdMatches = rows.filter((row) => normalize(row[indexes.requisitionId]) === lookupValue);
    const exactTitleMatches = rows.filter((row) => normalize(row[indexes.jobPostingTitle]) === lookupValue);
    const partialTitleMatches = rows.filter((row) => flexibleTextMatch(row[indexes.jobPostingTitle], lookupValue));

    const matches = exactIdMatches.length > 0
      ? exactIdMatches
      : exactTitleMatches.length > 0
        ? exactTitleMatches
        : partialTitleMatches;

    if (matches.length === 0) {
      target.getRange(config.statusCell).setValue("No requisition found (0 matches). Check the Requisition_ID or Job_Posting_Title.");
      if (target.getName() === "Pipeline") clearPipelineList(target, "No requisition selected.");
      return;
    }

    if (matches.length > 1) {
      const previewCount = Math.min(matches.length, 5);
      const preview = matches
        .slice(0, previewCount)
        .map((row) => `${asText(row[indexes.jobPostingTitle])} (${asText(row[indexes.requisitionId])})`)
        .join("; ");
      target.getRange(config.statusCell).setValue(`Multiple matches found (${matches.length} total; showing first ${previewCount}): ${preview}. Search by Requisition_ID.`);
      if (target.getName() === "Pipeline") clearPipelineList(target, "Resolve the requisition match before viewing the pipeline.");
      return;
    }

    const row = matches[0];
    target.getRange(config.requisitionIdCell).setValue(asText(row[indexes.requisitionId]));
    target.getRange(config.jobTitleCell).setValue(asText(row[indexes.jobPostingTitle]));
    target.getRange(config.primaryRecruiterCell).setValue(asText(row[indexes.primaryRecruiter]));
    target.getRange(config.locationCell).setValue(asText(row[indexes.location]));
    target.getRange(config.hiringManagersCell).setValue(asText(row[indexes.hiringManagers]));
    target.getRange(config.jobProfileCell).setValue(asText(row[indexes.jobProfile]));
    target.getRange(config.jobLevelCell).setValue(asText(row[indexes.jobLevel]));
    target.getRange("M21").setValue(asText(row[indexes.positionWorkerType]));
    target.getRange("M22").setValue(asText(row[indexes.jobRequisitionStatus]));
    target.getRange("M23").setValue(asText(row[indexes.addToStaffOrReplacement]));
    target.getRange("M24").setValue(asText(row[indexes.numberOfOpeningsAvailable]));
    target.getRange("M25").setValue(asText(row[indexes.bg]));
    target.getRange("M26").setValue(asText(row[indexes.sbu]));
    target.getRange("M27").setValue(asText(row[indexes.div]));
    target.getRange("M28").setValue(asText(row[indexes.supervisoryOrganization]));

    const matchType = exactIdMatches.length > 0
      ? "requisition_id"
      : exactTitleMatches.length > 0
        ? "job_posting_title"
        : "flexible job_posting_title";
    target.getRange(config.statusCell).setValue(`Matched 1 result on ${matchType}: ${asText(row[indexes.jobPostingTitle])}`);

    if (target.getName() === "Pipeline") {
      populatePipelineList(workbook, target);
    }
    target.activate();
  } finally {
    if (wasProtected) protection.protect();
  }
}

type RequisitionConfig = {
  lookupCell: string;
  statusCell: string;
  requisitionIdCell: string;
  jobTitleCell: string;
  primaryRecruiterCell: string;
  locationCell: string;
  hiringManagersCell: string;
  jobProfileCell: string;
  jobLevelCell: string;
  candidateAppliedCell: string;
  candidateStageCell: string;
  clearRanges: string[];
};

type PipelineCandidate = {
  employeeId: string;
  fullName: string;
  businessTitle: string;
  screened: string;
  notes: string;
  latestTime: number;
};

function requisitionConfig(sheetName: string): RequisitionConfig {
  if (sheetName === "Pipeline") {
    return {
      lookupCell: "C6",
      statusCell: "H6",
      requisitionIdCell: "C8",
      jobTitleCell: "F8",
      primaryRecruiterCell: "C10",
      locationCell: "F10",
      hiringManagersCell: "I10",
      jobProfileCell: "I8",
      jobLevelCell: "M3",
      candidateAppliedCell: "M4",
      candidateStageCell: "M5",
      clearRanges: ["H6:K6", "C8:D8", "F8:G8", "I8:K8", "C10:D10", "F10:G10", "I10:K10", "M3", "M4", "M5", "M21:M28"],
    };
  }

  return {
    lookupCell: "C21",
    statusCell: "H21",
    requisitionIdCell: "C23",
    jobTitleCell: "F23",
    primaryRecruiterCell: "C25",
    locationCell: "F25",
    hiringManagersCell: "I25",
    jobProfileCell: "I23",
    jobLevelCell: "C27",
    candidateAppliedCell: "F27",
    candidateStageCell: "I27",
    clearRanges: ["H21:K21", "C23:D23", "F23:G23", "I23:K23", "C25:D25", "F25:G25", "I25:K25", "C27:D27", "F27:G27", "I27:K27", "M21:M28"],
  };
}

function populatePipelineList(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet) {
  const reqId = normalize(sheet.getRange("C8").getValue());
  const jobTitle = normalize(sheet.getRange("F8").getValue());
  clearPipelineRows(sheet);

  if (!reqId && !jobTitle) {
    sheet.getRange("B19").setValue("Run Lookup Requisition to populate this list from tblCandidateNotes.");
    return;
  }

  const table = workbook.getWorksheet("Candidates").getTable("tblCandidateNotes");
  const headers = table.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const col = (name: string): number => {
    const index = headers.indexOf(name);
    if (index < 0) throw new Error(`Missing required column in tblCandidateNotes: ${name}`);
    return index;
  };
  const indexes = {
    addedDateTime: col("added_datetime"),
    employeeId: col("employee_id"),
    mmId: col("mm_id"),
    fullName: col("preferred_full_nm"),
    businessTitle: col("business_title_txt"),
    requisitionId: col("requisition_id"),
    jobTitle: col("job_posting_title"),
    screenDate: col("screen_date"),
    notes: col("recruiter_synthesis_notes"),
  };

  const matches = rows
    .filter((row) =>
      (reqId && normalize(row[indexes.requisitionId]) === reqId) ||
      (!reqId && jobTitle && normalize(row[indexes.jobTitle]) === jobTitle)
    )
    .map((row, sourceIndex) => ({ row, sourceIndex }))
    .sort((left, right) => {
      const rightTime = conversationTimestamp(right.row[indexes.addedDateTime]);
      const leftTime = conversationTimestamp(left.row[indexes.addedDateTime]);
      if (rightTime !== leftTime) return rightTime - leftTime;
      return right.sourceIndex - left.sourceIndex;
    });

  const byCandidate: { [key: string]: PipelineCandidate } = {};
  matches.forEach(({ row }) => {
    const key = normalize(row[indexes.employeeId]) || normalize(row[indexes.mmId]) || normalize(row[indexes.fullName]);
    if (!key) return;
    const existing = byCandidate[key];
    const screened = asText(row[indexes.screenDate]) ? "Yes" : "No";
    const notes = asText(row[indexes.notes]);
    const latestTime = conversationTimestamp(row[indexes.addedDateTime]);
    if (!existing) {
      byCandidate[key] = {
        employeeId: asText(row[indexes.employeeId]),
        fullName: asText(row[indexes.fullName]),
        businessTitle: asText(row[indexes.businessTitle]),
        screened,
        notes,
        latestTime,
      };
      return;
    }
    if (screened === "Yes") existing.screened = "Yes";
    if (!existing.notes && notes) existing.notes = notes;
  });

  const candidates = Object.keys(byCandidate)
    .map((key) => byCandidate[key])
    .sort((left, right) => right.latestTime - left.latestTime);
  const visible = candidates.slice(0, 15);
  visible.forEach((candidate, index) => {
    const row = 22 + index;
    sheet.getRange(`B${row}`).setValue(candidate.employeeId);
    sheet.getRange(`D${row}`).setValue(candidate.fullName);
    sheet.getRange(`F${row}`).setValue(candidate.businessTitle);
    sheet.getRange(`H${row}`).setValue(candidate.screened);
    sheet.getRange(`I${row}`).setValue(candidate.notes);
  });

  const reqLabel = asText(sheet.getRange("C8").getValue()) || asText(sheet.getRange("F8").getValue());
  if (candidates.length === 0) {
    sheet.getRange("B19").setValue(`No internal candidates found for ${reqLabel}.`);
  } else {
    sheet.getRange("B19").setValue(`Showing ${visible.length} of ${candidates.length} internal candidate${candidates.length === 1 ? "" : "s"} for ${reqLabel}.`);
  }
}

function clearRequisitionDetails(sheet: ExcelScript.Worksheet, config: RequisitionConfig) {
  config.clearRanges.forEach((address) => sheet.getRange(address).clear(ExcelScript.ClearApplyTo.contents));
}

function clearPipelineList(sheet: ExcelScript.Worksheet, message: string) {
  clearPipelineRows(sheet);
  sheet.getRange("B19").setValue(message);
}

function clearPipelineRows(sheet: ExcelScript.Worksheet) {
  for (let row = 22; row <= 40; row += 1) {
    ["B", "D", "F", "H", "I"].forEach((column) => {
      sheet.getRange(`${column}${row}`).setValue("");
    });
  }
}

function normalize(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ").toLowerCase();
}

function flexibleTextMatch(value: unknown, query: string): boolean {
  const normalizedValue = normalize(value);
  const normalizedQuery = normalize(query);
  if (!normalizedQuery) return false;
  if (normalizedValue.includes(normalizedQuery)) return true;

  const queryTokens = normalizedQuery.split(" ").filter((token) => token.length > 0);
  const valueTokens = normalizedValue.split(" ").filter((token) => token.length > 0);
  if (queryTokens.length > 0 && orderedTokenMatch(queryTokens, valueTokens)) return true;

  const queryCompact = compact(normalizedQuery);
  const valueCompact = compact(normalizedValue);
  return queryCompact.length >= 3 && charactersInOrder(queryCompact, valueCompact);
}

function orderedTokenMatch(queryTokens: string[], valueTokens: string[]): boolean {
  let startIndex = 0;
  for (const queryToken of queryTokens) {
    let found = false;
    for (let index = startIndex; index < valueTokens.length; index += 1) {
      const valueToken = valueTokens[index];
      if (
        valueToken.startsWith(queryToken) ||
        valueToken.includes(queryToken) ||
        charactersInOrder(queryToken, valueToken)
      ) {
        found = true;
        startIndex = index + 1;
        break;
      }
    }
    if (!found) return false;
  }
  return true;
}

function compact(value: unknown): string {
  return normalize(value).replace(/[^a-z0-9]/g, "");
}

function charactersInOrder(needle: string, haystack: string): boolean {
  if (!needle) return false;
  let position = 0;
  for (let index = 0; index < haystack.length && position < needle.length; index += 1) {
    if (haystack[index] === needle[position]) position += 1;
  }
  return position === needle.length;
}

function asText(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function conversationTimestamp(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Date.UTC(1899, 11, 30) + Math.round(value * 86400000);
  }
  const parsed = Date.parse(asText(value));
  return Number.isNaN(parsed) ? 0 : parsed;
}

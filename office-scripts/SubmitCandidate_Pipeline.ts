function main(workbook: ExcelScript.Workbook) {
  const pipeline = workbook.getWorksheet("Pipeline");
  const table = workbook.getWorksheet("Candidates").getTable("tblCandidateNotes");
  setPipelineStatus(pipeline, "Checking required fields...", "info");

  const required: [string, string][] = [
    ["Requisition lookup", "C6"],
    ["Requisition ID or job title", "C8"],
    ["Candidate lookup", "C14"],
    ["Full name", "F16"],
    ["Employee ID or MM ID", "C16"],
    ["Screener", "C3"],
  ];

  const missing = required
    .filter(([label, address]) => {
      if (label === "Requisition ID or job title") {
        return !text(pipeline, "C8") && !text(pipeline, "F8");
      }
      if (label === "Employee ID or MM ID") {
        return !text(pipeline, "C16") && !text(pipeline, "M8");
      }
      return !text(pipeline, address);
    })
    .map(([label]) => label);

  if (missing.length > 0) {
    const message = `Missing required field${missing.length === 1 ? "" : "s"}: ${missing.join(", ")}.`;
    setPipelineStatus(pipeline, message, "error");
    throw new Error(message);
  }

  const existingRows = table.getRangeBetweenHeaderAndTotal().getRowCount();
  const year = new Date().getFullYear();
  const submissionId = `NOTE-${year}-${String(existingRows + 1).padStart(3, "0")}`;
  const addedDateTime = `${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`;
  const headers = table.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());

  const values: { [key: string]: string } = {
    submission_id: submissionId,
    added_datetime: addedDateTime,
    submitted_by: text(pipeline, "C3") || "Screener",
    candidate_lookup_key: text(pipeline, "C14"),
    candidate_lookup_status: text(pipeline, "H14"),
    employee_id: text(pipeline, "C16"),
    mm_id: text(pipeline, "M8"),
    full_name: text(pipeline, "F16"),
    business_title: text(pipeline, "I16"),
    recent_hire_date: text(pipeline, "M7"),
    date_of_last_mobility_event: text(pipeline, "M9"),
    location: text(pipeline, "M10"),
    job_profile: text(pipeline, "M11"),
    management_level: text(pipeline, "M12"),
    manager_full_name: text(pipeline, "M13"),
    business_group: text(pipeline, "M14"),
    sub_business_group: text(pipeline, "M15"),
    division: text(pipeline, "M16"),
    requisition_lookup_key: text(pipeline, "C6"),
    requisition_lookup_status: text(pipeline, "H6"),
    requisition_id: text(pipeline, "C8"),
    job_posting_title: text(pipeline, "F8"),
    primary_recruiter: text(pipeline, "C10"),
    requisition_location: text(pipeline, "F10"),
    hiring_managers: text(pipeline, "I10"),
    requisition_job_profile: text(pipeline, "I8"),
    requisition_job_level: text(pipeline, "M3"),
    has_candidate_applied: "",
    candidate_stage: "",
    screener: text(pipeline, "C3"),
    screen_date: "",
  };

  table.addRow(-1, headers.map((header) => values[header] ?? ""));
  clearCandidateEntry(pipeline);
  setPipelineStatus(pipeline, `Added ${submissionId}. Requisition remains selected.`, "success");
  populatePipelineList(workbook, pipeline);
  pipeline.activate();
}

function setPipelineStatus(sheet: ExcelScript.Worksheet, message: string, tone: string) {
  const statusRange = sheet.getRange("F3:K3");
  const statusCell = sheet.getRange("F3");
  statusCell.setValue(message);
  statusRange.getFormat().setWrapText(true);
  statusRange.getFormat().getFill().setColor(tone === "error" ? "#FDECEC" : "#E7F3F1");
  statusRange.getFormat().getFont().setColor(tone === "error" ? "#7F1D1D" : "#475569");
  statusRange.getFormat().getFont().setItalic(tone !== "error");
}

function clearCandidateEntry(sheet: ExcelScript.Worksheet) {
  [
    "C14", "H14",
    "C16", "F16", "I16",
    "M7", "M8", "M9", "M10", "M11", "M12", "M13", "M14", "M15", "M16",
  ].forEach((address) => sheet.getRange(address).setValue(""));
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
    fullName: col("full_name"),
    businessTitle: col("business_title"),
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

type PipelineCandidate = {
  employeeId: string;
  fullName: string;
  businessTitle: string;
  screened: string;
  notes: string;
  latestTime: number;
};

function clearPipelineRows(sheet: ExcelScript.Worksheet) {
  for (let row = 22; row <= 40; row += 1) {
    ["B", "D", "F", "H", "I"].forEach((column) => {
      sheet.getRange(`${column}${row}`).setValue("");
    });
  }
}

function text(sheet: ExcelScript.Worksheet, address: string): string {
  const value = sheet.getRange(address).getValue();
  return value === null || value === undefined ? "" : String(value).trim();
}

function normalize(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ").toLowerCase();
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

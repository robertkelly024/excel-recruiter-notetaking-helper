function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const candidates = workbook.getWorksheet("Candidates");
  const table = candidates.getTable("tblCandidateNotes");

  const text = (address: string): string => {
    const value = intake.getRange(address).getValue();
    return value === null || value === undefined ? "" : String(value).trim();
  };

  const score = (address: string): number | "" => {
    const raw = text(address);
    if (!raw) return "";
    const parsed = Number(raw);
    if (!Number.isInteger(parsed) || parsed < 1 || parsed > 5) {
      throw new Error(`Score in ${address} must be a whole number from 1 to 5.`);
    }
    return parsed;
  };

  const required: [string, string][] = [
    ["Candidate lookup", "C7"],
    ["Full name", "C9"],
    ["Employee ID or MM ID", "F9"],
    ["Open role / req", "C21"],
    ["Stage", "F21"],
    ["Recruiter", "I21"],
    ["Screen date", "C23"],
    ["Summary notes", "B37"],
  ];

  const missing = required
    .filter(([label, address]) => {
      if (label === "Employee ID or MM ID") {
        return !text("F9") && !text("I9");
      }
      return !text(address);
    })
    .map(([label]) => label);

  if (missing.length > 0) {
    throw new Error(`Complete required fields before submitting: ${missing.join(", ")}.`);
  }

  const roleFit = score("D28");
  const communication = score("D29");
  const motivation = score("D30");
  const experienceMatch = score("D31");
  const logisticsFit = score("D32");
  const compensationFit = score("D33");
  const riskLevel = score("D34");

  const averageInputs = [roleFit, communication, motivation, experienceMatch, logisticsFit, compensationFit]
    .filter((value): value is number => typeof value === "number");
  const averageScore = averageInputs.length > 0
    ? Math.round((averageInputs.reduce((sum, value) => sum + value, 0) / averageInputs.length) * 100) / 100
    : "";

  const existingRows = table.getRangeBetweenHeaderAndTotal().getRowCount();
  const year = new Date().getFullYear();
  const submissionId = `NOTE-${year}-${String(existingRows + 1).padStart(3, "0")}`;
  const submittedAt = new Date().toISOString().slice(0, 16).replace("T", " ");
  const submittedBy = text("I21") || "Recruiter";

  const row: (string | number)[] = [
    submissionId,
    submittedAt,
    submittedBy,
    text("C7"),
    text("H7"),
    text("F9"),
    text("I9"),
    text("C9"),
    text("C11"),
    text("F11"),
    text("I11"),
    text("C13"),
    text("F13"),
    text("I13"),
    text("C15"),
    text("F15"),
    text("I15"),
    text("C21"),
    text("F21"),
    text("I21"),
    text("C23"),
    text("F23"),
    text("I23"),
    roleFit,
    text("E28"),
    communication,
    text("E29"),
    motivation,
    text("E30"),
    experienceMatch,
    text("E31"),
    logisticsFit,
    text("E32"),
    compensationFit,
    text("E33"),
    riskLevel,
    text("E34"),
    averageScore,
    text("B37"),
  ];

  table.addRow(-1, row);
  clearIntake(intake);
  intake.getRange("H7").setValue(`Submitted ${submissionId} at ${submittedAt}.`);
  intake.activate();
}

function clearIntake(sheet: ExcelScript.Worksheet) {
  [
    "C7:F7", "H7:K7",
    "C9:D9", "F9:G9", "I9:K9",
    "C11:D11", "F11:G11", "I11:K11",
    "C13:D13", "F13:G13", "I13:K13",
    "C15:D15", "F15:G15", "I15:K15",
    "C21:D21", "F21:G21", "I21:K21",
    "C23:D23", "F23:G23", "I23:K23",
    "D28:D34", "E28:K34",
    "B37:K43",
  ].forEach((address) => sheet.getRange(address).clear(ExcelScript.ClearApplyTo.contents));
}

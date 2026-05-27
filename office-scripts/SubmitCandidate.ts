function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const candidates = workbook.getWorksheet("Candidates");
  const table = candidates.getTable("Candidates");

  const text = (address: string): string => {
    const value = intake.getRange(address).getValue();
    return value === null || value === undefined ? "" : String(value).trim();
  };

  const score = (address: string): number | "" => {
    const value = intake.getRange(address).getValue();
    if (value === null || value === undefined || value === "") return "";
    const parsed = Number(value);
    if (!Number.isInteger(parsed) || parsed < 1 || parsed > 5) {
      throw new Error(`Score in ${address} must be a whole number from 1 to 5.`);
    }
    return parsed;
  };

  const required: [string, string][] = [
    ["Candidate name", "C6"],
    ["Role", "F6"],
    ["Stage", "I6"],
    ["Interview date", "C10"],
    ["Interviewer", "F10"],
  ];
  const missing = required.filter(([_, address]) => !text(address)).map(([name]) => name);
  if (missing.length > 0) {
    throw new Error(`Complete required fields before submitting: ${missing.join(", ")}.`);
  }

  const roleFit = score("C18");
  const communication = score("C19");
  const motivation = score("C20");
  const experienceMatch = score("C21");
  const logisticsFit = score("C22");
  const compensationFit = score("C23");
  const riskLevel = score("C24");

  const scoringValues = [roleFit, communication, motivation, experienceMatch, logisticsFit, compensationFit]
    .filter((value): value is number => typeof value === "number");
  const averageScore = scoringValues.length > 0
    ? Math.round((scoringValues.reduce((sum, value) => sum + value, 0) / scoringValues.length) * 100) / 100
    : "";

  const existingRows = table.getRangeBetweenHeaderAndTotal().getRowCount();
  const year = new Date().getFullYear();
  const candidateId = `CAND-${year}-${String(existingRows + 1).padStart(3, "0")}`;
  const submittedBy = text("I12") || "Recruiter";
  const submittedAt = new Date().toISOString().slice(0, 16).replace("T", " ");

  const row: (string | number)[] = [
    candidateId,
    text("C6"),
    text("F6"),
    text("I6"),
    text("I8"),
    text("C8"),
    text("F8"),
    text("C10"),
    text("F10"),
    text("I10"),
    text("C12"),
    roleFit,
    communication,
    motivation,
    experienceMatch,
    logisticsFit,
    compensationFit,
    riskLevel,
    averageScore,
    text("F22"),
    text("F18"),
    text("F20"),
    text("F12"),
    submittedAt,
    submittedBy,
  ];

  table.addRow(-1, row);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  clearIntake(intake);
  candidates.activate();
}

function clearIntake(sheet: ExcelScript.Worksheet) {
  [
    "C6:D6", "F6:G6", "I6",
    "C8:D8", "F8:G8", "I8",
    "C10:D10", "F10:G10", "I10",
    "C12:D12", "F12:G12", "I12",
    "C18:C24", "D18:D24",
    "F18:I18", "F20:I20", "F22:I22",
  ].forEach((address) => sheet.getRange(address).clear(ExcelScript.ClearApplyTo.contents));
}

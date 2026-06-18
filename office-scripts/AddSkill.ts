function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const inputCell = intake.getRange("C43");
  const chainCell = intake.getRange("F43");

  const skill = clean(inputCell.getValue());
  if (!skill) {
    inputCell.setValue("");
    inputCell.select();
    return;
  }

  const existing = clean(chainCell.getValue());
  const skills = existing
    ? existing.split("|").map((item) => item.trim()).filter((item) => item.length > 0)
    : [];

  const alreadyExists = skills.some((item) => item.toLowerCase() === skill.toLowerCase());
  if (!alreadyExists) {
    skills.push(skill);
    chainCell.setValue(skills.join(" | "));
  }

  inputCell.setValue("");
  inputCell.select();
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ");
}

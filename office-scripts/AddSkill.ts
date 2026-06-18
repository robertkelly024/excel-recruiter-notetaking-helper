function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const inputCell = intake.getRange("C43");
  const categoryCell = intake.getRange("F43");
  const chainCells: { [key: string]: ExcelScript.Range } = {
    current: intake.getRange("C45"),
    developing: intake.getRange("C47"),
    aspirational: intake.getRange("C49"),
  };

  const skill = clean(inputCell.getValue());
  if (!skill) {
    inputCell.setValue("");
    inputCell.select();
    return;
  }

  const category = clean(categoryCell.getValue()).toLowerCase();
  const destination = chainCells[category];
  if (!destination) {
    throw new Error("Choose Current, Developing, or Aspirational before adding the skill.");
  }

  Object.keys(chainCells).forEach((key) => {
    const cell = chainCells[key];
    const remaining = parseChain(cell.getValue())
      .filter((item) => item.toLowerCase() !== skill.toLowerCase());
    cell.setValue(remaining.join(" | "));
  });

  const destinationSkills = parseChain(destination.getValue());
  destinationSkills.push(skill);
  destination.setValue(destinationSkills.join(" | "));
  inputCell.setValue("");
  inputCell.select();
}

function parseChain(value: unknown): string[] {
  const existing = clean(value);
  return existing
    ? existing.split("|").map((item) => item.trim()).filter((item) => item.length > 0)
    : [];
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ");
}

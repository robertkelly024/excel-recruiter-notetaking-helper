function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const inputCell = intake.getRange("C39");
  const chainCell = intake.getRange("G39");

  const role = clean(inputCell.getValue());
  if (!role) {
    inputCell.setValue("");
    inputCell.select();
    return;
  }

  const roles = parseChain(chainCell.getValue());
  const alreadyExists = roles.some((item) => item.toLowerCase() === role.toLowerCase());
  if (!alreadyExists) {
    roles.push(role);
    chainCell.setValue(roles.join(" | "));
  }

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

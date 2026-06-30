function main(workbook: ExcelScript.Workbook) {
  deleteChainValue(workbook, {
    itemLabel: "desired level",
    chainLabel: "Desired levels",
    inputCell: "C32",
    chainCell: "G32",
    messageCell: "E33",
    messageLabelCell: "B33",
    messageRow: "33:33",
  });
}

type DeleteChainConfig = {
  itemLabel: string;
  chainLabel: string;
  inputCell: string;
  chainCell: string;
  messageCell: string;
  messageLabelCell: string;
  messageRow: string;
};

function deleteChainValue(workbook: ExcelScript.Workbook, config: DeleteChainConfig) {
  const intake = workbook.getWorksheet("Intake");
  const protection = intake.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
    const inputCell = intake.getRange(config.inputCell);
    const chainCell = intake.getRange(config.chainCell);
    const typedValue = clean(inputCell.getValue());
    clearResolution(intake);

    if (!typedValue) {
      writeMessage(intake, config, `Enter a ${config.itemLabel} to delete.`);
      inputCell.select();
      return;
    }

    const typedKey = normalizeInput(typedValue);
    const values = parseChain(chainCell.getValue());
    const remaining: string[] = [];
    let removed = false;

    values.forEach((value) => {
      if (normalizeInput(value) === typedKey) {
        removed = true;
      } else {
        remaining.push(value);
      }
    });

    if (!removed) {
      writeMessage(intake, config, `"${typedValue}" does not exist in ${config.chainLabel}.`);
      inputCell.select();
      return;
    }

    chainCell.setValue(remaining.join(" | "));
    inputCell.setValue("");
    inputCell.select();
  } finally {
    if (wasProtected) protection.protect();
  }
}

function writeMessage(sheet: ExcelScript.Worksheet, config: DeleteChainConfig, message: string) {
  sheet.getRange(config.messageLabelCell).setValue(`${config.chainLabel} update:`);
  const messageCell = sheet.getRange(config.messageCell);
  messageCell.getDataValidation().clear();
  messageCell.setValue(message);
  messageCell.getFormat().getFill().setColor("#FFF7ED");
  sheet.getRange(config.messageRow).setRowHidden(false);
}

function clearResolution(sheet: ExcelScript.Worksheet) {
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
}

function parseChain(value: unknown): string[] {
  const existing = clean(value);
  return existing
    ? existing.split("|").map((item) => item.trim()).filter((item) => item.length > 0)
    : [];
}

function normalizeInput(value: unknown): string {
  return clean(value)
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9+#.\s-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ");
}

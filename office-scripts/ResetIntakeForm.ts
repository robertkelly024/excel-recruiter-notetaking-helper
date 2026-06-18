function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");

  [
    "C7", "H7",
    "C9", "F9", "I9",
    "C11", "F11", "I11",
    "C13", "F13", "I13",
    "C15", "F15", "I15",
    "C21", "H21",
    "C23", "F23", "I23",
    "C25", "F25", "I25",
    "C27",
    "C32", "F32", "I32",
    "C34",
    "C39", "G39",
    "C41", "F41", "I41",
    "C43", "F43",
    "C45", "C47", "C49",
    "B52",
    "B60",
    "B63", "D63", "G63", "I63", "B64",
    "B65", "D65", "G65", "I65", "B66",
    "B67", "D67", "G67", "I67", "B68",
    "B69", "D69", "G69", "I69", "B70",
    "B71", "D71", "G71", "I71", "B72",
  ].forEach((address) => intake.getRange(address).setValue(""));

  intake.getRange("F32").setFormula("=Settings!$C$16");
  intake.getRange("B60").setValue("Run Lookup Employee to display the five most recent conversations for this candidate.");
  intake.activate();
}

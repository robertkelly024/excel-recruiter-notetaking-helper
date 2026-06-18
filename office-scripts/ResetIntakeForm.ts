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
    "C39", "H39",
    "C41", "F41", "I41",
    "C43", "F43",
    "B46",
    "B54",
    "B57", "D57", "F57", "I57", "B58",
    "B59", "D59", "F59", "I59", "B60",
    "B61", "D61", "F61", "I61", "B62",
    "B63", "D63", "F63", "I63", "B64",
    "B65", "D65", "F65", "I65", "B66",
  ].forEach((address) => intake.getRange(address).setValue(""));

  intake.getRange("F32").setFormula("=Settings!$C$16");
  intake.getRange("I32").setFormula("=TODAY()");
  intake.getRange("B54").setValue("Run Lookup Employee to display the five most recent conversations for this candidate.");
  intake.activate();
}

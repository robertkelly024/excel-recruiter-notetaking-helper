function main(workbook: ExcelScript.Workbook) {
  const intake = workbook.getWorksheet("Intake");
  const descriptionRange = intake.getRange("C27:K27");
  const rowRange = intake.getRange("27:27");
  const format = rowRange.getFormat();

  const defaultHeight = 58;
  const maxHeight = 409;
  const currentHeight = format.getRowHeight();

  descriptionRange.getFormat().setWrapText(true);
  descriptionRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);

  if (currentHeight > defaultHeight + 6) {
    format.setRowHeight(defaultHeight);
    return;
  }

  const description = String(descriptionRange.getCell(0, 0).getValue() ?? "").trim();
  if (!description) {
    format.setRowHeight(defaultHeight);
    return;
  }

  const explicitLines = description.split(/\r\n|\r|\n/).length;
  const wrappedLines = Math.ceil(description.length / 135);
  const estimatedLines = Math.max(2, explicitLines, wrappedLines);
  const expandedHeight = Math.min(maxHeight, Math.max(defaultHeight, 22 + estimatedLines * 17));

  format.setRowHeight(expandedHeight);
}

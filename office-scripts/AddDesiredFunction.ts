function main(workbook: ExcelScript.Workbook) {
  addCatalogBackedInput(workbook, {
    itemType: "function",
    inputCell: "C34",
    chainCell: "G34",
    historyColumn: "desired_function",
    promptLabel: "Similar desired functions found. Select intended function from dropdown:",
    resolutionCell: "E35",
    resolutionLabelCell: "B35",
    resolutionRow: "35:35",
  });
}

type AddInputConfig = {
  itemType: string;
  inputCell: string;
  chainCell: string;
  historyColumn: string;
  promptLabel: string;
  resolutionCell: string;
  resolutionLabelCell: string;
  resolutionRow: string;
};

type InputCandidate = {
  canonicalValue: string;
  aliases: string[];
};

type InputMatch = {
  canonicalValue: string;
  score: number;
  reason: string;
};

function addCatalogBackedInput(workbook: ExcelScript.Workbook, config: AddInputConfig) {
  const intake = workbook.getWorksheet("Intake");
  const protection = intake.getProtection();
  const wasProtected = protection.getProtected();
  if (wasProtected) protection.unprotect();

  try {
    const inputCell = intake.getRange(config.inputCell);
    const chainCell = intake.getRange(config.chainCell);
    const typedValue = clean(inputCell.getValue());
    const normalizedValue = normalizeInput(typedValue);
    const selectedChoice = clean(intake.getRange(config.resolutionCell).getValue());
    const pendingInput = clean(intake.getRange("M6").getValue());

    if (selectedChoice && normalizedValue === pendingInput) {
      const chosenValue = valueFromChoice(selectedChoice);
      if (chosenValue) {
        appendValue(chainCell, chosenValue);
        upsertCatalogValue(getInputCatalogTable(workbook), config.itemType, chosenValue, typedValue);
        clearResolution(intake);
        inputCell.setValue("");
        inputCell.select();
        return;
      }
    }

    if (pendingInput && normalizedValue !== pendingInput) {
      clearResolution(intake);
    }

    if (!normalizedValue) {
      clearResolution(intake);
      inputCell.setValue("");
      inputCell.select();
      return;
    }

    const catalogTable = getInputCatalogTable(workbook);
    const candidateTable = workbook.getWorksheet("Candidates").getTable("tblCandidateNotes");
    const candidates = buildCandidates(catalogTable, candidateTable, chainCell, config.itemType, config.historyColumn);
    const matches = rankMatches(normalizedValue, candidates);
    const highConfidence = matches.find((match) => match.score >= 95);

    if (highConfidence) {
      appendValue(chainCell, highConfidence.canonicalValue);
      upsertCatalogValue(catalogTable, config.itemType, highConfidence.canonicalValue, typedValue);
      clearResolution(intake);
      inputCell.setValue("");
      inputCell.select();
      return;
    }

    const suggestions = matches
      .filter((match) => match.score >= 80)
      .slice(0, 3);

    if (suggestions.length > 0) {
      writeResolution(intake, suggestions, normalizedValue, config);
      intake.getRange(config.resolutionCell).select();
      return;
    }

    appendValue(chainCell, normalizedValue);
    upsertCatalogValue(catalogTable, config.itemType, normalizedValue, typedValue);
    clearResolution(intake);
    inputCell.setValue("");
    inputCell.select();
  } finally {
    if (wasProtected) protection.protect();
  }
}

function getInputCatalogTable(workbook: ExcelScript.Workbook): ExcelScript.Table {
  return workbook.getWorksheet("input_catalog").getTable("tblInputCatalog");
}

function buildCandidates(
  catalogTable: ExcelScript.Table,
  candidateTable: ExcelScript.Table,
  chainCell: ExcelScript.Range,
  itemType: string,
  historyColumn: string
): InputCandidate[] {
  const byKey: { [key: string]: InputCandidate } = {};
  const catalogHeaders = catalogTable.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const catalogRows = catalogTable.getRangeBetweenHeaderAndTotal().getValues();
  const catalogCol = (name: string): number => catalogHeaders.indexOf(name);
  const typeIndex = catalogCol("type");
  const canonicalIndex = catalogCol("canonical_skill");
  const aliasesIndex = catalogCol("aliases");
  const statusIndex = catalogCol("status");

  catalogRows.forEach((row) => {
    const type = normalizeInput(row[typeIndex]);
    const status = normalizeInput(row[statusIndex]);
    const canonicalValue = clean(row[canonicalIndex]);
    if (type !== itemType || !canonicalValue || status === "inactive") return;
    const key = normalizeInput(canonicalValue);
    byKey[key] = {
      canonicalValue,
      aliases: parseAliases(row[aliasesIndex]),
    };
  });

  parseChain(chainCell.getValue()).forEach((value) => {
    const key = normalizeInput(value);
    if (key && !byKey[key]) byKey[key] = { canonicalValue: value, aliases: [] };
  });

  const notesHeaders = candidateTable.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const historyIndex = notesHeaders.indexOf(historyColumn);
  if (historyIndex >= 0) {
    candidateTable.getRangeBetweenHeaderAndTotal().getValues().forEach((row) => {
      parseChain(row[historyIndex]).forEach((value) => {
        const key = normalizeInput(value);
        if (key && !byKey[key]) byKey[key] = { canonicalValue: value, aliases: [] };
      });
    });
  }

  return Object.keys(byKey).map((key) => byKey[key]);
}

function rankMatches(input: string, candidates: InputCandidate[]): InputMatch[] {
  const bestByValue: { [key: string]: InputMatch } = {};
  candidates.forEach((candidate) => {
    const match = scoreValue(input, candidate);
    if (match.score < 60) return;
    const key = normalizeInput(match.canonicalValue);
    if (!bestByValue[key] || match.score > bestByValue[key].score) {
      bestByValue[key] = match;
    }
  });
  return Object.keys(bestByValue)
    .map((key) => bestByValue[key])
    .sort((left, right) => right.score - left.score || left.canonicalValue.localeCompare(right.canonicalValue));
}

function scoreValue(input: string, candidate: InputCandidate): InputMatch {
  const canonical = normalizeInput(candidate.canonicalValue);
  if (input === canonical) {
    return { canonicalValue: candidate.canonicalValue, score: 100, reason: "100 - exact canonical match" };
  }

  const exactAlias = candidate.aliases.find((alias) => normalizeInput(alias) === input);
  if (exactAlias) {
    return { canonicalValue: candidate.canonicalValue, score: 95, reason: `95 - exact alias match: ${exactAlias}` };
  }

  let best: InputMatch = { canonicalValue: candidate.canonicalValue, score: 0, reason: "" };
  const terms = [candidate.canonicalValue].concat(candidate.aliases);
  terms.forEach((term) => {
    const normalizedTerm = normalizeInput(term);
    if (!normalizedTerm) return;

    const containsScore = containsSimilarity(input, normalizedTerm);
    if (containsScore > best.score) {
      best = { canonicalValue: candidate.canonicalValue, score: containsScore, reason: `${containsScore} - one term contains the other: ${term}` };
    }

    const overlapScore = tokenOverlapScore(input, normalizedTerm);
    if (overlapScore > best.score) {
      best = { canonicalValue: candidate.canonicalValue, score: overlapScore, reason: `${overlapScore} - strong token overlap: ${term}` };
    }

    const editScore = editDistanceScore(input, normalizedTerm);
    if (editScore > best.score) {
      best = { canonicalValue: candidate.canonicalValue, score: editScore, reason: `${editScore} - edit-distance similarity: ${term}` };
    }
  });
  return best;
}

function containsSimilarity(left: string, right: string): number {
  const shorter = left.length <= right.length ? left : right;
  const longer = left.length > right.length ? left : right;
  const meaningful = shorter.length >= 4 && tokens(shorter).length > 0;
  if (!meaningful || !longer.includes(shorter)) return 0;
  return tokens(shorter).length >= 2 ? 90 : 85;
}

function tokenOverlapScore(left: string, right: string): number {
  const leftTokens = tokens(left);
  const rightTokens = tokens(right);
  if (leftTokens.length === 0 || rightTokens.length === 0) return 0;
  const common = leftTokens.filter((token) => rightTokens.indexOf(token) >= 0).length;
  const smaller = Math.min(leftTokens.length, rightTokens.length);
  const union = uniqueStrings(leftTokens.concat(rightTokens)).length;
  const containment = common / smaller;
  const jaccard = common / union;
  const strength = Math.max(containment, jaccard);
  if (common < 1 || strength < 0.5) return 0;
  return Math.min(85, Math.max(70, Math.round(70 + strength * 15)));
}

function editDistanceScore(left: string, right: string): number {
  const maxLength = Math.max(left.length, right.length);
  if (maxLength < 5) return 0;
  const distance = levenshtein(left, right);
  const similarity = 1 - distance / maxLength;
  if (similarity < 0.72) return 0;
  return Math.min(75, Math.max(60, Math.round(60 + ((similarity - 0.72) / 0.28) * 15)));
}

function writeResolution(sheet: ExcelScript.Worksheet, suggestions: InputMatch[], normalizedValue: string, config: AddInputConfig) {
  clearResolution(sheet);
  const choices = suggestions.map((suggestion) => `Use: ${suggestion.canonicalValue}`);
  choices.push(`Add new: ${normalizedValue}`);
  while (choices.length < 4) choices.push("");

  sheet.getRange("M2:M5").setValues(choices.slice(0, 4).map((choice) => [choice]));
  sheet.getRange("M6").setValue(normalizedValue);
  sheet.getRange(config.resolutionLabelCell).setValue(config.promptLabel);
  sheet.getRange(config.resolutionRow).setRowHidden(false);
  const selected = sheet.getRange(config.resolutionCell);
  selected.setValue("");
  selected.getFormat().getFill().setColor("#FFF7ED");
  selected.getDataValidation().setRule({
    list: {
      inCellDropDown: true,
      source: "=$M$2:$M$5",
    },
  });
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

function valueFromChoice(choice: string): string {
  const usePrefix = "Use:";
  const addPrefix = "Add new:";
  if (choice.toLowerCase().startsWith(usePrefix.toLowerCase())) {
    return clean(choice.slice(usePrefix.length));
  }
  if (choice.toLowerCase().startsWith(addPrefix.toLowerCase())) {
    return normalizeInput(choice.slice(addPrefix.length));
  }
  return "";
}

function appendValue(chainCell: ExcelScript.Range, value: string) {
  const normalized = normalizeInput(value);
  const values = parseChain(chainCell.getValue());
  const alreadyExists = values.some((item) => normalizeInput(item) === normalized);
  if (!alreadyExists) {
    values.push(value);
    chainCell.setValue(values.join(" | "));
  }
}

function upsertCatalogValue(table: ExcelScript.Table, itemType: string, canonicalValue: string, typedValue: string) {
  const headers = table.getHeaderRowRange().getValues()[0].map((value) => String(value).trim());
  const rowsRange = table.getRangeBetweenHeaderAndTotal();
  const rows = rowsRange.getValues();
  const indexes = {
    type: headers.indexOf("type"),
    canonical: headers.indexOf("canonical_skill"),
    aliases: headers.indexOf("aliases"),
    usage: headers.indexOf("usage_count"),
    lastUsed: headers.indexOf("last_used"),
    status: headers.indexOf("status"),
  };
  const canonicalKey = normalizeInput(canonicalValue);
  const typedKey = normalizeInput(typedValue);

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    if (normalizeInput(row[indexes.type]) !== itemType) continue;
    if (normalizeInput(row[indexes.canonical]) !== canonicalKey) continue;

    const currentUsage = Number(row[indexes.usage]) || 0;
    rowsRange.getCell(rowIndex, indexes.usage).setValue(currentUsage + 1);
    rowsRange.getCell(rowIndex, indexes.lastUsed).setValue(todayIso());
    rowsRange.getCell(rowIndex, indexes.status).setValue("active");

    if (typedKey && typedKey !== canonicalKey) {
      const aliases = parseAliases(row[indexes.aliases]);
      const hasAlias = aliases.some((alias) => normalizeInput(alias) === typedKey);
      if (!hasAlias) {
        aliases.push(typedKey);
        rowsRange.getCell(rowIndex, indexes.aliases).setValue(aliases.join(" | "));
      }
    }
    return;
  }

  const aliases = typedKey && typedKey !== canonicalKey ? typedKey : "";
  table.addRow(-1, [itemType, canonicalValue, aliases, 1, todayIso(), "active"]);
}

function parseAliases(value: unknown): string[] {
  return parseChain(value);
}

function parseChain(value: unknown): string[] {
  const existing = clean(value);
  return existing
    ? existing.split("|").map((item) => item.trim()).filter((item) => item.length > 0)
    : [];
}

function tokens(value: string): string[] {
  const stopWords = ["and", "or", "the", "a", "an", "of", "to", "in", "for", "with"];
  return uniqueStrings(
    value
      .split(" ")
      .map((item) => item.trim())
      .filter((item) => item.length > 1 && stopWords.indexOf(item) < 0)
  );
}

function uniqueStrings(values: string[]): string[] {
  const seen: { [key: string]: boolean } = {};
  const result: string[] = [];
  values.forEach((value) => {
    if (!value || seen[value]) return;
    seen[value] = true;
    result.push(value);
  });
  return result;
}

function levenshtein(left: string, right: string): number {
  const previous: number[] = [];
  const current: number[] = [];
  for (let column = 0; column <= right.length; column += 1) previous[column] = column;
  for (let row = 1; row <= left.length; row += 1) {
    current[0] = row;
    for (let column = 1; column <= right.length; column += 1) {
      const substitution = previous[column - 1] + (left.charAt(row - 1) === right.charAt(column - 1) ? 0 : 1);
      current[column] = Math.min(previous[column] + 1, current[column - 1] + 1, substitution);
    }
    for (let column = 0; column <= right.length; column += 1) previous[column] = current[column];
  }
  return previous[right.length];
}

function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
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

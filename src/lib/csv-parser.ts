export interface ParsedCsv {
  headers: string[];
  rows: string[][];
  rawCount: number;
}

/**
 * Parse pasted CSV or TSV text into headers + rows.
 * Handles comma-separated, tab-separated, and pipe-separated formats.
 * Strips common bank statement junk lines (e.g. "Opening Balance", empty rows).
 */
export function parsePastedText(raw: string): ParsedCsv | null {
  const lines = raw
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  if (lines.length < 2) return null;

  // Detect delimiter: whichever of \t , | appears most in first line
  const firstLine = lines[0];
  const tabCount = (firstLine.match(/\t/g) || []).length;
  const commaCount = (firstLine.match(/,/g) || []).length;
  const pipeCount = (firstLine.match(/\|/g) || []).length;

  let delimiter = ",";
  if (tabCount >= commaCount && tabCount >= pipeCount) delimiter = "\t";
  else if (pipeCount > commaCount) delimiter = "|";

  const splitLine = (line: string): string[] =>
    line.split(delimiter).map((cell) => cell.replace(/^["']|["']$/g, "").trim());

  const allRows = lines.map(splitLine);

  // Find the header row: look for a row that contains common column keywords
  let headerIdx = 0;
  for (let i = 0; i < Math.min(allRows.length, 8); i++) {
    const joined = allRows[i].join(" ").toLowerCase();
    if (
      joined.includes("date") ||
      joined.includes("description") ||
      joined.includes("narration") ||
      joined.includes("amount") ||
      joined.includes("debit") ||
      joined.includes("credit")
    ) {
      headerIdx = i;
      break;
    }
  }

  const headers = allRows[headerIdx];
  const dataRows = allRows
    .slice(headerIdx + 1)
    .filter((row) => {
      if (row.every((cell) => cell === "")) return false;
      // Skip summary rows like "Opening Balance", "Closing Balance", "Total"
      const firstCell = row[0].toLowerCase();
      if (
        firstCell.includes("opening") ||
        firstCell.includes("closing") ||
        firstCell.includes("total") ||
        firstCell.includes("balance b/f") ||
        firstCell.includes("balance c/f")
      )
        return false;
      return true;
    });

  if (dataRows.length === 0) return null;

  // Normalize every row to exactly the same width as the header.
  // Rows that are shorter get padded; rows that are wider get truncated.
  const colCount = headers.length;
  const normalizedRows = dataRows.map((row) => {
    const r = [...row];
    while (r.length < colCount) r.push("");
    return r.slice(0, colCount);
  });

  return { headers, rows: normalizedRows, rawCount: normalizedRows.length };
}

/**
 * Write parsed CSV data into a new (or existing) Excel sheet and return the sheet name.
 */
export async function writeToExcelSheet(
  parsed: ParsedCsv,
  context: Excel.RequestContext,
  sheetName = "BSA Import"
): Promise<string> {
  // Remove existing import sheet if present
  try {
    const existing = context.workbook.worksheets.getItem(sheetName);
    existing.delete();
    await context.sync();
  } catch {
    // Sheet didn't exist
  }

  const sheet = context.workbook.worksheets.add(sheetName);
  sheet.activate();

  const colCount = parsed.headers.length;
  // Ensure every row (including header) is exactly colCount wide
  const allData = [parsed.headers, ...parsed.rows].map((row) => {
    const r = [...row];
    while (r.length < colCount) r.push("");
    return r.slice(0, colCount);
  });
  const range = sheet.getRange(`A1:${colLetter(colCount)}${allData.length}`);
  range.values = allData as Excel.RangeValueType[][];

  // Bold + fill the header row
  const headerRange = sheet.getRange(`A1:${colLetter(parsed.headers.length)}1`);
  headerRange.format.font.bold = true;
  headerRange.format.fill.color = "#dbeafe";

  // Auto-fit columns
  sheet.getRange("A:Z").format.autofitColumns();

  await context.sync();
  return sheetName;
}

function colLetter(n: number): string {
  // n is 1-based column count → last column letter
  let col = "";
  let num = n;
  while (num > 0) {
    const rem = (num - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

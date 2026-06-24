import { categorize, categorizeByName, buildSummary, type Transaction, type Summary } from "./categorizer";

declare const Office: typeof import("@microsoft/office-js");

export type ColumnMap = {
  date: number;
  description: number;
  // Single-column format (amount + optional type)
  amount: number | null;
  type: number | null;
  // Split-column format (separate debit / credit / balance)
  debit: number | null;
  credit: number | null;
  balance: number | null;
  // Optional pre-assigned category column
  category: number | null;
};

export function fmt(n: number, symbol = "₦"): string {
  const abs = Math.abs(n);
  return `${symbol}${abs.toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
}

export function fmtShort(n: number, symbol = "₦"): string {
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return `${symbol}${(abs / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `${symbol}${(abs / 1_000).toFixed(0)}K`;
  return `${symbol}${abs.toFixed(0)}`;
}

export async function detectColumns(sheet: Excel.Worksheet): Promise<ColumnMap | null> {
  const headerRange = sheet.getRange("A1:L1");
  headerRange.load("values");
  await (sheet.context as Excel.RequestContext).sync();

  const headers: string[] = (headerRange.values[0] as string[]).map((h) =>
    String(h || "").toLowerCase().trim()
  );

  const find = (...terms: string[]) =>
    headers.findIndex((h) => terms.some((t) => h === t || h.includes(t)));

  const date        = find("date", "tran date", "value date", "trans date");
  const description = find("description", "narration", "details", "particulars", "memo", "remark", "transaction");

  if (date === -1 || description === -1) return null;

  // Category column (optional — present in split-format exports from banks)
  const category = find("category", "cat");

  // ── Split-column format: separate Debit / Credit columns ──────────────────
  const debit  = find("debit");
  const credit = find("credit");

  if (debit !== -1 && credit !== -1) {
    const balance = find("balance", "running balance", "bal");
    return {
      date, description,
      amount: null, type: null,
      debit, credit,
      balance: balance === -1 ? null : balance,
      category: category === -1 ? null : category,
    };
  }

  // ── Single-column format: one Amount column ───────────────────────────────
  const amount = find("amount", "value");
  const type   = find("type", "dr/cr", "debit/credit", "transaction type", "cr/dr");

  if (amount === -1) return null;

  return {
    date, description,
    amount, type: type === -1 ? null : type,
    debit: null, credit: null, balance: null,
    category: category === -1 ? null : category,
  };
}

export async function readTransactions(
  sheet: Excel.Worksheet,
  columnMap: ColumnMap
): Promise<Transaction[]> {
  const usedRange = sheet.getUsedRange();
  usedRange.load("values,rowCount");
  await (sheet.context as Excel.RequestContext).sync();

  const rows = usedRange.values as (string | number | boolean)[][];
  const transactions: Transaction[] = [];

  const isSplitFormat = columnMap.debit !== null && columnMap.credit !== null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rawDate = row[columnMap.date];
    const rawDesc = row[columnMap.description];

    if (!rawDesc) continue;

    const date = rawDate ? String(rawDate) : "";
    const description = String(rawDesc || "");

    let amount = 0;
    let type: "credit" | "debit" = "debit";

    if (isSplitFormat) {
      // ── Split Debit / Credit format ────────────────────────────────────────
      const rawDebit  = columnMap.debit  !== null ? row[columnMap.debit]  : "";
      const rawCredit = columnMap.credit !== null ? row[columnMap.credit] : "";

      const debitVal  = rawDebit  !== "" && rawDebit  !== null ? Math.abs(Number(rawDebit))  : 0;
      const creditVal = rawCredit !== "" && rawCredit !== null ? Math.abs(Number(rawCredit)) : 0;

      if (isNaN(debitVal) && isNaN(creditVal)) continue;

      if (creditVal > 0) {
        amount = creditVal;
        type = "credit";
      } else if (debitVal > 0) {
        amount = debitVal;
        type = "debit";
      } else {
        continue; // both zero — skip (e.g. opening balance header rows)
      }
    } else {
      // ── Single Amount column format ────────────────────────────────────────
      const rawAmount = columnMap.amount !== null ? row[columnMap.amount] : null;
      if (rawAmount === "" || rawAmount === null || rawAmount === undefined) continue;

      amount = Math.abs(Number(rawAmount));
      if (isNaN(amount) || amount === 0) continue;

      if (columnMap.type !== null) {
        const rawType = String(row[columnMap.type] || "").toLowerCase();
        type = rawType.includes("cr") || rawType.includes("credit") ? "credit" : "debit";
      } else {
        type = Number(rawAmount) > 0 ? "credit" : "debit";
      }
    }

    // Use pre-assigned category from the sheet if present, else auto-detect
    const rawCat = columnMap.category !== null ? String(row[columnMap.category] ?? "").trim() : "";
    const category = (rawCat ? categorizeByName(rawCat) : null) ?? categorize(description, amount, type);

    transactions.push({ row: i + 1, date, description, amount, type, category });
  }

  return transactions;
}

/** Blend a hex color with white at `opacity` (0–1) to produce a light tint.
 *  Excel fill.color only accepts 6-digit RGB hex — no alpha channel. */
function tint(hex: string, opacity: number): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const blend = (c: number) => Math.round(c * opacity + 255 * (1 - opacity));
  return "#" + [r, g, b].map(blend).map((c) => c.toString(16).padStart(2, "0")).join("");
}

export async function highlightTransactions(
  sheet: Excel.Worksheet,
  transactions: Transaction[],
  columnMap: ColumnMap
): Promise<void> {
  const ctx = sheet.context as Excel.RequestContext;

  // Load used range so we know how many columns to color per row
  const usedRange = sheet.getUsedRange();
  usedRange.load("columnCount");
  await ctx.sync();

  const colCount = Math.max(usedRange.columnCount, 1);

  for (const tx of transactions) {
    // Highlight the entire row (0-indexed row = tx.row - 1)
    const rowRange = sheet.getRangeByIndexes(tx.row - 1, 0, 1, colCount);
    rowRange.format.fill.color = tint(tx.category.color, 0.25);
    rowRange.format.font.color = "#1e293b";
  }

  await ctx.sync();
}

export async function clearHighlights(sheet: Excel.Worksheet): Promise<void> {
  const ctx = sheet.context as Excel.RequestContext;
  const usedRange = sheet.getUsedRange();
  usedRange.load("rowCount,columnCount");
  await ctx.sync();
  usedRange.format.fill.clear();
  await ctx.sync();
}

export async function createSummarySheet(summary: Summary, context: Excel.RequestContext, currencySymbol = "₦"): Promise<void> {
  const sheetName = "BSA Summary";

  try {
    const existing = context.workbook.worksheets.getItem(sheetName);
    existing.delete();
    await context.sync();
  } catch {
    // Sheet didn't exist — that's fine
  }

  const summarySheet = context.workbook.worksheets.add(sheetName);
  summarySheet.activate();

  const now = new Date().toLocaleDateString("en-US");

  const data: (string | number)[][] = [
    ["Bank Statement Analyzer Pro — Summary Report"],
    [`Generated: ${now}`],
    [],
    ["OVERVIEW"],
    ["Total Income", summary.totalIncome],
    ["Total Expenses", summary.totalExpenses],
    ["Net Savings", summary.net],
    ["Savings Rate (%)", summary.savingsRate],
    ["Health Score (/100)", summary.healthScore],
    [],
    ["SPENDING BY CATEGORY"],
    [`Category`, `Amount (${currencySymbol})`, "Transactions", "% of Expenses"],
  ];

  const sortedCats = Object.entries(summary.byCategory).sort((a, b) => b[1].total - a[1].total);
  for (const [name, info] of sortedCats) {
    const pct = summary.totalExpenses > 0 ? Math.round((info.total / summary.totalExpenses) * 100) : 0;
    data.push([name, info.total, info.count, pct]);
  }

  if (summary.monthly.length > 1) {
    data.push([]);
    data.push(["MONTHLY BREAKDOWN"]);
    data.push(["Month", `Income (${currencySymbol})`, `Expenses (${currencySymbol})`, `Net (${currencySymbol})`]);
    for (const m of summary.monthly) {
      data.push([m.month, m.income, m.expenses, m.net]);
    }
  }

  if (summary.recurring.length > 0) {
    data.push([]);
    data.push(["RECURRING TRANSACTIONS"]);
    data.push(["Description", "Occurrences", `Avg Amount (${currencySymbol})`, `Total (${currencySymbol})`]);
    for (const r of summary.recurring) {
      data.push([r.description, r.count, Math.round(r.avgAmount), Math.round(r.totalAmount)]);
    }
  }

  data.push([]);
  data.push(["ALL TRANSACTIONS"]);
  data.push(["Date", "Description", `Amount (${currencySymbol})`, "Type", "Category"]);
  for (const tx of summary.transactions) {
    data.push([tx.date, tx.description, tx.amount, tx.type.toUpperCase(), tx.category.name]);
  }

  // Normalize every row to exactly 5 columns so the range dimensions match
  const normalizedData = data.map((row) => {
    const r = [...row] as (string | number)[];
    while (r.length < 5) r.push("");
    return r.slice(0, 5);
  });
  const range = summarySheet.getRange(`A1:E${normalizedData.length}`);
  range.values = normalizedData as Excel.RangeValueType[][];

  // Style header
  const titleCell = summarySheet.getRange("A1");
  titleCell.format.font.bold = true;
  titleCell.format.font.size = 14;
  titleCell.format.font.color = "#1e3a8a";

  const overviewHeader = summarySheet.getRange("A4");
  overviewHeader.format.font.bold = true;
  overviewHeader.format.fill.color = "#dbeafe";

  const catHeaderRow = 11;
  summarySheet.getRange(`A${catHeaderRow}`).format.font.bold = true;
  summarySheet.getRange(`A${catHeaderRow}`).format.fill.color = "#dbeafe";
  summarySheet.getRange(`A${catHeaderRow + 1}:D${catHeaderRow + 1}`).format.font.bold = true;
  summarySheet.getRange(`A${catHeaderRow + 1}:D${catHeaderRow + 1}`).format.fill.color = "#e2e8f0";

  summarySheet.getRange("A:E").format.autofitColumns();

  await context.sync();
}

export function exportToCsv(summary: Summary, notes: Record<number, string> = {}, flaggedRows: Set<number> = new Set()): void {
  const rows: string[][] = [
    ["Date", "Description", "Amount", "Type", "Category", "Flagged?", "Duplicate?", "Notes"],
  ];
  for (const tx of summary.transactions) {
    const note = notes[tx.row] ?? "";
    rows.push([
      tx.date,
      `"${tx.description.replace(/"/g, '""')}"`,
      String(tx.amount),
      tx.type.toUpperCase(),
      tx.category.name,
      flaggedRows.has(tx.row) ? "Flagged" : "",
      summary.duplicateRows.has(tx.row) ? "Possible Duplicate" : "",
      `"${note.replace(/"/g, '""')}"`,
    ]);
  }

  const csv = rows.map((r) => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `bank-statement-${new Date().toISOString().split("T")[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

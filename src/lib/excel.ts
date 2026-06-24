import { categorize, buildSummary, type Transaction, type Summary } from "./categorizer";

declare const Office: typeof import("@microsoft/office-js");

export type ColumnMap = {
  date: number;
  description: number;
  amount: number;
  type: number | null;
};

export function fmt(n: number): string {
  return new Intl.NumberFormat("en-NG", { style: "currency", currency: "NGN", maximumFractionDigits: 0 }).format(n);
}

export function fmtShort(n: number): string {
  if (n >= 1_000_000) return `₦${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000) return `₦${(n / 1_000).toFixed(0)}K`;
  return `₦${n.toFixed(0)}`;
}

export async function detectColumns(sheet: Excel.Worksheet): Promise<ColumnMap | null> {
  const headerRange = sheet.getRange("A1:J1");
  headerRange.load("values");
  await (sheet.context as Excel.RequestContext).sync();

  const headers: string[] = (headerRange.values[0] as string[]).map((h) =>
    String(h || "").toLowerCase().trim()
  );

  const find = (...terms: string[]) =>
    headers.findIndex((h) => terms.some((t) => h.includes(t)));

  const date = find("date", "tran date", "value date");
  const description = find("description", "narration", "details", "particulars", "memo", "remark");
  const amount = find("amount", "debit", "credit", "value");
  const type = find("type", "dr/cr", "debit/credit", "transaction type");

  if (date === -1 || description === -1 || amount === -1) return null;

  return { date, description, amount, type: type === -1 ? null : type };
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

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rawDate = row[columnMap.date];
    const rawDesc = row[columnMap.description];
    const rawAmount = row[columnMap.amount];

    if (!rawDesc || rawAmount === "" || rawAmount === null) continue;

    const amount = Math.abs(Number(rawAmount));
    if (isNaN(amount) || amount === 0) continue;

    const date = rawDate ? String(rawDate) : "";
    const description = String(rawDesc || "");

    let type: "credit" | "debit" = "debit";
    if (columnMap.type !== null) {
      const rawType = String(row[columnMap.type] || "").toLowerCase();
      if (rawType.includes("cr") || rawType.includes("credit")) type = "credit";
    } else {
      type = Number(rawAmount) > 0 ? "credit" : "debit";
    }

    const category = categorize(description, amount, type);
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

  for (const tx of transactions) {
    const colLetter = String.fromCharCode(65 + columnMap.description);
    const cell = sheet.getRange(`${colLetter}${tx.row}`);
    cell.format.fill.color = tint(tx.category.color, 0.25);
    cell.format.font.color = "#1e293b";
  }

  await ctx.sync();
}

export async function createSummarySheet(summary: Summary, context: Excel.RequestContext): Promise<void> {
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

  const now = new Date().toLocaleDateString("en-NG");

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
    ["Category", "Amount (₦)", "Transactions", "% of Expenses"],
  ];

  const sortedCats = Object.entries(summary.byCategory).sort((a, b) => b[1].total - a[1].total);
  for (const [name, info] of sortedCats) {
    const pct = summary.totalExpenses > 0 ? Math.round((info.total / summary.totalExpenses) * 100) : 0;
    data.push([name, info.total, info.count, pct]);
  }

  if (summary.monthly.length > 1) {
    data.push([]);
    data.push(["MONTHLY BREAKDOWN"]);
    data.push(["Month", "Income (₦)", "Expenses (₦)", "Net (₦)"]);
    for (const m of summary.monthly) {
      data.push([m.month, m.income, m.expenses, m.net]);
    }
  }

  if (summary.recurring.length > 0) {
    data.push([]);
    data.push(["RECURRING TRANSACTIONS"]);
    data.push(["Description", "Occurrences", "Avg Amount (₦)", "Total (₦)"]);
    for (const r of summary.recurring) {
      data.push([r.description, r.count, Math.round(r.avgAmount), Math.round(r.totalAmount)]);
    }
  }

  data.push([]);
  data.push(["ALL TRANSACTIONS"]);
  data.push(["Date", "Description", "Amount (₦)", "Type", "Category"]);
  for (const tx of summary.transactions) {
    data.push([tx.date, tx.description, tx.amount, tx.type.toUpperCase(), tx.category.name]);
  }

  const range = summarySheet.getRange(`A1:E${data.length}`);
  range.values = data as Excel.RangeValueType[][];

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

export function exportToCsv(summary: Summary): void {
  const rows: string[][] = [
    ["Date", "Description", "Amount", "Type", "Category", "Duplicate?"],
  ];
  for (const tx of summary.transactions) {
    rows.push([
      tx.date,
      `"${tx.description.replace(/"/g, '""')}"`,
      String(tx.amount),
      tx.type.toUpperCase(),
      tx.category.name,
      summary.duplicateRows.has(tx.row) ? "Possible Duplicate" : "",
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

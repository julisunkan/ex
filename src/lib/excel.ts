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
      // Heuristic: if amount column header is "credit" and positive, it's income
      type = Number(rawAmount) > 0 ? "credit" : "debit";
    }

    const category = categorize(description, amount, type);
    transactions.push({ row: i + 1, date, description, amount, type, category });
  }

  return transactions;
}

export async function highlightTransactions(
  sheet: Excel.Worksheet,
  transactions: Transaction[],
  columnMap: ColumnMap
): Promise<void> {
  const ctx = sheet.context as Excel.RequestContext;

  for (const tx of transactions) {
    const cellAddress = `${String.fromCharCode(65 + columnMap.description)}${tx.row}`;
    const cell = sheet.getRange(cellAddress);
    cell.format.fill.color = tx.category.color + "33"; // 20% opacity hex
    cell.format.font.color = "#1e293b";
  }

  await ctx.sync();
}

export async function createSummarySheet(summary: Summary, context: Excel.RequestContext): Promise<void> {
  const sheetName = "BSA Summary";

  // Remove existing summary sheet if present
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
    [],
    ["SPENDING BY CATEGORY"],
    ["Category", "Amount (₦)", "Transactions"],
  ];

  const sortedCats = Object.entries(summary.byCategory).sort((a, b) => b[1].total - a[1].total);
  for (const [name, info] of sortedCats) {
    data.push([name, info.total, info.count]);
  }

  data.push([]);
  data.push(["TRANSACTIONS"]);
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

  const catHeader = summarySheet.getRange("A10");
  catHeader.format.font.bold = true;
  catHeader.format.fill.color = "#dbeafe";

  summarySheet.getRange("A11:C11").format.font.bold = true;
  summarySheet.getRange("A11:C11").format.fill.color = "#e2e8f0";

  const txHeaderRow = 11 + sortedCats.length + 3;
  summarySheet.getRange(`A${txHeaderRow}:E${txHeaderRow}`).format.font.bold = true;
  summarySheet.getRange(`A${txHeaderRow}:E${txHeaderRow}`).format.fill.color = "#e2e8f0";

  // Auto-fit columns A through E
  summarySheet.getRange("A:E").format.autofitColumns();

  await context.sync();
}

export type TransactionType = "credit" | "debit";

export interface Category {
  name: string;
  color: string;
  className: string;
  keywords: string[];
  type: "income" | "expense" | "transfer" | "any";
}

export const CATEGORIES: Category[] = [
  {
    name: "Salary",
    color: "#22c55e",
    className: "cat-salary",
    keywords: ["salary", "payroll", "wage", "income", "payment from", "acme", "employer", "stipend", "commission"],
    type: "income",
  },
  {
    name: "Groceries",
    color: "#10b981",
    className: "cat-grocery",
    keywords: ["shoprite", "grocery", "supermarket", "market", "spar", "walmart", "whole foods", "sainsbury", "checkers", "pick n pay", "nakumatt", "carrefour"],
    type: "expense",
  },
  {
    name: "Food & Dining",
    color: "#f59e0b",
    className: "cat-food",
    keywords: ["kfc", "mcdonalds", "restaurant", "cafe", "pizza", "burger", "domino", "chicken republic", "cafeteria", "lunch", "dinner", "food", "eatery", "suya", "shawarma", "chinese", "sushi", "starbucks"],
    type: "expense",
  },
  {
    name: "Transport",
    color: "#8b5cf6",
    className: "cat-transport",
    keywords: ["uber", "bolt", "lyft", "taxi", "transport", "ride", "bus", "train", "metro", "trip", "indriver", "gokada", "okada"],
    type: "expense",
  },
  {
    name: "Fuel",
    color: "#ef4444",
    className: "cat-transport",
    keywords: ["petrol", "fuel", "gas station", "total", "oando", "shell", "mobil", "filling station", "nnpc"],
    type: "expense",
  },
  {
    name: "Utilities",
    color: "#f97316",
    className: "cat-utility",
    keywords: ["electricity", "water", "gas", "internet", "airtime", "mtn", "glo", "airtel", "dstv", "electric", "ekedc", "lawma", "utility", "bill", "ikedc", "phed", "starlink", "wifi", "broadband"],
    type: "expense",
  },
  {
    name: "Shopping",
    color: "#ec4899",
    className: "cat-shopping",
    keywords: ["amazon", "jumia", "konga", "shop", "purchase", "order", "buy", "aliexpress", "shein", "zara", "h&m", "stores", "mall"],
    type: "expense",
  },
  {
    name: "Entertainment",
    color: "#6366f1",
    className: "cat-entertain",
    keywords: ["netflix", "spotify", "apple", "disney", "youtube", "subscription", "stream", "prime", "showmax", "canopy", "cinema", "movies", "gaming", "playstation", "xbox"],
    type: "expense",
  },
  {
    name: "Healthcare",
    color: "#14b8a6",
    className: "cat-health",
    keywords: ["pharmacy", "hospital", "clinic", "doctor", "health", "medical", "drug", "pharma", "lab", "test", "scan", "dental", "optician", "wellness"],
    type: "expense",
  },
  {
    name: "Savings",
    color: "#0ea5e9",
    className: "cat-savings",
    keywords: ["savings", "save", "piggy", "cowrywise", "fixed deposit", "fd ", "treasury"],
    type: "expense",
  },
  {
    name: "Investments",
    color: "#84cc16",
    className: "cat-invest",
    keywords: ["invest", "stock", "fund", "mutual", "stanbic", "asset", "portfolio", "shares", "etf", "bond", "reit"],
    type: "expense",
  },
  {
    name: "Loans",
    color: "#9333ea",
    className: "cat-loan",
    keywords: ["loan", "repayment", "mortgage", "credit card", "debt", "emi", "instalm", "fairmoney", "carbon", "branch", "renmoney"],
    type: "expense",
  },
  {
    name: "Education",
    color: "#f43f5e",
    className: "cat-education",
    keywords: ["school", "tuition", "university", "college", "course", "training", "udemy", "coursera", "book", "exam", "study"],
    type: "expense",
  },
  {
    name: "Rent & Housing",
    color: "#a16207",
    className: "cat-rent",
    keywords: ["rent", "landlord", "housing", "estate", "property", "accommodation", "caution", "agent fee"],
    type: "expense",
  },
];

/**
 * Look up a Category by its exact name (case-insensitive).
 * Returns null if no match — caller should fall back to `categorize()`.
 */
export function categorizeByName(name: string): Category | null {
  if (!name) return null;
  const normalized = name.toLowerCase().trim();
  return CATEGORIES.find((c) => c.name.toLowerCase() === normalized) ?? null;
}

export function categorize(description: string, amount: number, type: TransactionType): Category {
  const desc = description.toLowerCase();

  for (const cat of CATEGORIES) {
    if (cat.type === "income" && type === "debit") continue;
    if (cat.type === "expense" && type === "credit") continue;
    if (cat.keywords.some((kw) => desc.includes(kw))) return cat;
  }

  return {
    name: "Other",
    color: "#64748b",
    className: "cat-other",
    keywords: [],
    type: "any",
  };
}

export interface Transaction {
  row: number;
  date: string;
  description: string;
  amount: number;
  type: TransactionType;
  category: Category;
}

export interface MonthlyPoint {
  month: string;
  income: number;
  expenses: number;
  net: number;
}

export interface RecurringGroup {
  description: string;
  category: Category;
  count: number;
  avgAmount: number;
  totalAmount: number;
  transactions: Transaction[];
}

export interface MomChange {
  category: string;
  color: string;
  className: string;
  current: number;
  previous: number;
  delta: number;
  pctChange: number;
  isNew: boolean;
  isGone: boolean;
}

export interface Summary {
  totalIncome: number;
  totalExpenses: number;
  net: number;
  savingsRate: number;
  byCategory: Record<string, { total: number; count: number; color: string; className: string }>;
  transactions: Transaction[];
  monthly: MonthlyPoint[];
  monthlyByCategory: Record<string, Record<string, number>>; // month -> category -> total
  momChanges: MomChange[];  // category-level changes between the two most recent months
  momMonths: [string, string] | null;  // [current month label, previous month label]
  recurring: RecurringGroup[];
  duplicateRows: Set<number>;
  topMerchants: { name: string; total: number; count: number; category: Category }[];
  healthScore: number;
  healthTips: string[];
}

function parseMonthKey(dateStr: string): string {
  if (!dateStr) return "Unknown";
  const s = dateStr.trim();

  // Try YYYY-MM-DD
  const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    const d = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, 1);
    return d.toLocaleString("en", { month: "short", year: "2-digit" });
  }

  // Try DD/MM/YYYY or MM/DD/YYYY
  const slashMatch = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (slashMatch) {
    const a = Number(slashMatch[1]);
    const b = Number(slashMatch[2]);
    const yr = Number(slashMatch[3]);
    const year = yr < 100 ? 2000 + yr : yr;
    // Heuristic: if first number > 12, it's DD/MM
    const month = a > 12 ? b : a <= 12 && b <= 12 ? b : a;
    const d = new Date(year, month - 1, 1);
    return d.toLocaleString("en", { month: "short", year: "2-digit" });
  }

  // Excel serial number (days since 1899-12-30)
  const serial = Number(s);
  if (!isNaN(serial) && serial > 40000 && serial < 60000) {
    const d = new Date((serial - 25569) * 86400000);
    return d.toLocaleString("en", { month: "short", year: "2-digit" });
  }

  return "Other";
}

function normalizeDesc(desc: string): string {
  return desc.toLowerCase()
    .replace(/\d+/g, "") // strip numbers
    .replace(/[^a-z\s]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .substring(0, 30);
}

function detectRecurring(transactions: Transaction[]): RecurringGroup[] {
  const groups: Record<string, Transaction[]> = {};

  for (const tx of transactions) {
    if (tx.type !== "debit") continue; // Only track expense recurring
    const key = normalizeDesc(tx.description);
    if (key.length < 4) continue;
    if (!groups[key]) groups[key] = [];
    groups[key].push(tx);
  }

  const result: RecurringGroup[] = [];
  for (const [, txs] of Object.entries(groups)) {
    if (txs.length < 2) continue;
    const totalAmount = txs.reduce((s, t) => s + t.amount, 0);
    const avgAmount = totalAmount / txs.length;
    result.push({
      description: txs[0].description,
      category: txs[0].category,
      count: txs.length,
      avgAmount,
      totalAmount,
      transactions: txs,
    });
  }

  return result.sort((a, b) => b.totalAmount - a.totalAmount).slice(0, 8);
}

function detectDuplicates(transactions: Transaction[]): Set<number> {
  const dupeRows = new Set<number>();
  for (let i = 0; i < transactions.length; i++) {
    for (let j = i + 1; j < transactions.length; j++) {
      const a = transactions[i];
      const b = transactions[j];
      if (
        Math.abs(a.amount - b.amount) < 0.01 &&
        a.type === b.type &&
        normalizeDesc(a.description) === normalizeDesc(b.description)
      ) {
        dupeRows.add(a.row);
        dupeRows.add(b.row);
      }
    }
  }
  return dupeRows;
}

function computeTopMerchants(transactions: Transaction[]): Summary["topMerchants"] {
  const map: Record<string, { total: number; count: number; category: Category }> = {};
  for (const tx of transactions) {
    if (tx.type !== "debit") continue;
    const key = tx.description.substring(0, 30).toUpperCase().trim();
    if (!map[key]) map[key] = { total: 0, count: 0, category: tx.category };
    map[key].total += tx.amount;
    map[key].count++;
  }
  return Object.entries(map)
    .map(([name, v]) => ({ name, ...v }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 6);
}

function computeHealthScore(totalIncome: number, totalExpenses: number, savingsRate: number, byCategory: Summary["byCategory"]): { score: number; tips: string[] } {
  let score = 50;
  const tips: string[] = [];

  // Savings rate: 0-30 points
  if (savingsRate >= 30) score += 30;
  else if (savingsRate >= 20) score += 22;
  else if (savingsRate >= 10) score += 12;
  else if (savingsRate >= 5) score += 5;
  else { score -= 5; tips.push("Boost savings — aim for at least 10% of income saved each month."); }

  // Expense ratio: 0-20 points
  const expRatio = totalIncome > 0 ? totalExpenses / totalIncome : 1;
  if (expRatio < 0.7) score += 20;
  else if (expRatio < 0.85) score += 10;
  else if (expRatio > 1) { score -= 10; tips.push("You're spending more than you earn — review essential vs non-essential costs."); }

  // Category diversity
  const catCount = Object.keys(byCategory).length;
  if (catCount >= 6) score += 5;

  // Specific tips
  const entertainTotal = byCategory["Entertainment"]?.total ?? 0;
  const diningTotal = byCategory["Food & Dining"]?.total ?? 0;
  const loanTotal = byCategory["Loans"]?.total ?? 0;
  if (totalExpenses > 0 && entertainTotal / totalExpenses > 0.15)
    tips.push("Entertainment is >15% of spending — consider trimming subscriptions.");
  if (totalExpenses > 0 && diningTotal / totalExpenses > 0.2)
    tips.push("Food & Dining is high — cooking at home more often can reduce this.");
  if (loanTotal > 0 && totalIncome > 0 && loanTotal / totalIncome > 0.3)
    tips.push("Loan repayments exceed 30% of income — focus on debt reduction.");
  if (savingsRate >= 20 && tips.length === 0)
    tips.push("Great job! Keep building your emergency fund (3–6 months of expenses).");
  if (byCategory["Investments"])
    tips.push("You're investing — stay consistent to build long-term wealth.");

  return { score: Math.max(0, Math.min(100, score)), tips: tips.slice(0, 3) };
}

function computeMomChanges(
  monthlyByCategory: Record<string, Record<string, number>>,
  monthOrder: string[],
  byCategory: Summary["byCategory"]
): { changes: MomChange[]; months: [string, string] | null } {
  if (monthOrder.length < 2) return { changes: [], months: null };

  const curLabel = monthOrder[monthOrder.length - 1];
  const prevLabel = monthOrder[monthOrder.length - 2];
  const curCats = monthlyByCategory[curLabel] ?? {};
  const prevCats = monthlyByCategory[prevLabel] ?? {};

  const allCats = new Set([...Object.keys(curCats), ...Object.keys(prevCats)]);
  const changes: MomChange[] = [];

  for (const cat of allCats) {
    if (cat === "Salary" || cat === "Other") continue; // skip income & uncategorized
    const cur = curCats[cat] ?? 0;
    const prev = prevCats[cat] ?? 0;
    const delta = cur - prev;
    const pctChange = prev > 0 ? Math.round((delta / prev) * 100) : 100;
    const info = byCategory[cat];
    changes.push({
      category: cat,
      color: info?.color ?? "#64748b",
      className: info?.className ?? "cat-other",
      current: cur,
      previous: prev,
      delta,
      pctChange,
      isNew: prev === 0 && cur > 0,
      isGone: cur === 0 && prev > 0,
    });
  }

  // Sort: biggest increases first, then decreases, then new/gone
  changes.sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));

  return { changes, months: [curLabel, prevLabel] };
}

export function buildSummary(transactions: Transaction[]): Summary {
  let totalIncome = 0;
  let totalExpenses = 0;
  const byCategory: Summary["byCategory"] = {};
  const monthlyMap: Record<string, MonthlyPoint> = {};
  const monthlyByCategory: Record<string, Record<string, number>> = {};
  const monthOrder: string[] = [];

  for (const tx of transactions) {
    if (tx.type === "credit") {
      totalIncome += tx.amount;
    } else {
      totalExpenses += tx.amount;
    }

    const key = tx.category.name;
    if (!byCategory[key]) {
      byCategory[key] = { total: 0, count: 0, color: tx.category.color, className: tx.category.className };
    }
    byCategory[key].total += tx.amount;
    byCategory[key].count++;

    const monthKey = parseMonthKey(tx.date);
    if (!monthlyMap[monthKey]) {
      monthlyMap[monthKey] = { month: monthKey, income: 0, expenses: 0, net: 0 };
      monthOrder.push(monthKey);
    }
    if (tx.type === "credit") monthlyMap[monthKey].income += tx.amount;
    else monthlyMap[monthKey].expenses += tx.amount;

    // Track per-category per-month (expenses only)
    if (tx.type === "debit") {
      if (!monthlyByCategory[monthKey]) monthlyByCategory[monthKey] = {};
      if (!monthlyByCategory[monthKey][key]) monthlyByCategory[monthKey][key] = 0;
      monthlyByCategory[monthKey][key] += tx.amount;
    }
  }

  for (const pt of Object.values(monthlyMap)) pt.net = pt.income - pt.expenses;

  const net = totalIncome - totalExpenses;
  const savingsRate = totalIncome > 0 ? Math.round((net / totalIncome) * 100) : 0;
  const monthly = Object.values(monthlyMap);
  const { changes: momChanges, months: momMonths } = computeMomChanges(monthlyByCategory, monthOrder, byCategory);
  const recurring = detectRecurring(transactions);
  const duplicateRows = detectDuplicates(transactions);
  const topMerchants = computeTopMerchants(transactions);
  const { score: healthScore, tips: healthTips } = computeHealthScore(totalIncome, totalExpenses, savingsRate, byCategory);

  return {
    totalIncome, totalExpenses, net, savingsRate, byCategory, transactions,
    monthly, monthlyByCategory, momChanges, momMonths,
    recurring, duplicateRows, topMerchants, healthScore, healthTips,
  };
}

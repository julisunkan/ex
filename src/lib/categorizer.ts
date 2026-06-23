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
    keywords: ["salary", "payroll", "wage", "income", "payment from", "acme", "employer"],
    type: "income",
  },
  {
    name: "Groceries",
    color: "#10b981",
    className: "cat-grocery",
    keywords: ["shoprite", "grocery", "supermarket", "market", "spar", "walmart", "whole foods", "sainsbury"],
    type: "expense",
  },
  {
    name: "Food & Dining",
    color: "#f59e0b",
    className: "cat-food",
    keywords: ["kfc", "mcdonalds", "restaurant", "cafe", "pizza", "burger", "domino", "chicken republic", "cafeteria", "lunch", "dinner", "food"],
    type: "expense",
  },
  {
    name: "Transport",
    color: "#8b5cf6",
    className: "cat-transport",
    keywords: ["uber", "bolt", "lyft", "taxi", "transport", "ride", "bus", "train", "metro", "trip"],
    type: "expense",
  },
  {
    name: "Fuel",
    color: "#ef4444",
    className: "cat-transport",
    keywords: ["petrol", "fuel", "gas station", "total", "oando", "shell", "mobil", "filling station"],
    type: "expense",
  },
  {
    name: "Utilities",
    color: "#f97316",
    className: "cat-utility",
    keywords: ["electricity", "water", "gas", "internet", "airtime", "mtn", "glo", "airtel", "dstv", "electric", "ekedc", "lawma", "utility", "bill"],
    type: "expense",
  },
  {
    name: "Shopping",
    color: "#ec4899",
    className: "cat-shopping",
    keywords: ["amazon", "jumia", "konga", "shop", "purchase", "order", "buy"],
    type: "expense",
  },
  {
    name: "Entertainment",
    color: "#6366f1",
    className: "cat-entertain",
    keywords: ["netflix", "spotify", "apple", "disney", "youtube", "subscription", "stream", "prime"],
    type: "expense",
  },
  {
    name: "Healthcare",
    color: "#14b8a6",
    className: "cat-health",
    keywords: ["pharmacy", "hospital", "clinic", "doctor", "health", "medical", "drug", "pharma"],
    type: "expense",
  },
  {
    name: "Savings",
    color: "#0ea5e9",
    className: "cat-savings",
    keywords: ["savings", "save", "piggy", "cowrywise"],
    type: "expense",
  },
  {
    name: "Investments",
    color: "#84cc16",
    className: "cat-invest",
    keywords: ["invest", "stock", "fund", "mutual", "stanbic", "asset", "portfolio"],
    type: "expense",
  },
  {
    name: "Loans",
    color: "#9333ea",
    className: "cat-loan",
    keywords: ["loan", "repayment", "mortgage", "credit card", "debt", "emi"],
    type: "expense",
  },
];

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

export interface Summary {
  totalIncome: number;
  totalExpenses: number;
  net: number;
  savingsRate: number;
  byCategory: Record<string, { total: number; count: number; color: string; className: string }>;
  transactions: Transaction[];
}

export function buildSummary(transactions: Transaction[]): Summary {
  let totalIncome = 0;
  let totalExpenses = 0;
  const byCategory: Summary["byCategory"] = {};

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
  }

  const net = totalIncome - totalExpenses;
  const savingsRate = totalIncome > 0 ? Math.round((net / totalIncome) * 100) : 0;

  return { totalIncome, totalExpenses, net, savingsRate, byCategory, transactions };
}

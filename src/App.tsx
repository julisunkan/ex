import { useState, useCallback, useRef, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import {
  detectColumns,
  readTransactions,
  highlightTransactions,
  clearHighlights,
  createSummarySheet,
  exportToCsv,
  type ColumnMap,
} from "./lib/excel";
import { useCurrency, CURRENCIES } from "./hooks/use-currency";
import { buildSummary, CATEGORIES, type Category, type Summary, type Transaction } from "./lib/categorizer";
import { parsePastedText, writeToExcelSheet } from "./lib/csv-parser";
import { exportToPdf, buildReportHtml } from "./lib/pdf";
import { getLicense, checkLicenseValid } from "./lib/payment";
import PaymentGate from "./components/PaymentGate";
import SubscriptionDashboard from "./components/SubscriptionDashboard";
import { useAppConfig } from "./context/AppConfigContext";
import iconLogo from "@assets/icons/icon-logo.png";
import iconAnalyze from "@assets/icons/icon-analyze.png";
import iconPaste from "@assets/icons/icon-paste.png";
import iconIncome from "@assets/icons/icon-income.png";
import iconExpenses from "@assets/icons/icon-expenses.png";
import iconSavings from "@assets/icons/icon-savings.png";
import iconRate from "@assets/icons/icon-rate.png";
import iconHighlight from "@assets/icons/icon-highlight.png";
import iconExport from "@assets/icons/icon-export.png";
import iconPro from "@assets/icons/icon-pro.png";

declare const Excel: typeof import("@microsoft/office-js").Excel;
declare const Office: typeof import("@microsoft/office-js");

type Step = "idle" | "paste" | "importing" | "loading" | "results" | "error" | "subscription";
type ResultTab = "overview" | "categories" | "transactions" | "budget";

const isOfficeAvailable = () =>
  typeof Office !== "undefined" && typeof Excel !== "undefined";

function runExcel<T>(fn: (context: Excel.RequestContext) => Promise<T>): Promise<T> {
  if (!isOfficeAvailable()) {
    return Promise.reject(new Error("Office.js not available — open this add-in inside Excel."));
  }
  return Excel.run(fn);
}

// ── KPI Card ─────────────────────────────────────────────────────────────────
function KpiCard({ label, value, color, icon }: { label: string; value: string; color: string; icon: string }) {
  return (
    <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
      <div className="flex items-center gap-1.5 mb-2">
        <img src={icon} alt={label} className="w-5 h-5 object-contain" />
        <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">{label}</p>
      </div>
      <p className={`text-xl font-extrabold tracking-tight ${color}`}>{value}</p>
    </div>
  );
}

// ── Action Button ─────────────────────────────────────────────────────────────
function ActionBtn({
  onClick, disabled, children, variant = "primary", size = "md"
}: {
  onClick: () => void; disabled?: boolean; children: React.ReactNode; variant?: "primary" | "secondary" | "ghost"; size?: "md" | "sm";
}) {
  const base = size === "sm"
    ? "flex items-center justify-center gap-1.5 text-xs font-semibold py-1.5 px-3 rounded-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed"
    : "flex-1 flex items-center justify-center gap-2 text-sm font-semibold py-2.5 rounded-xl transition-all disabled:opacity-50 disabled:cursor-not-allowed";
  const styles = variant === "primary"
    ? `${base} bg-primary text-primary-foreground hover:opacity-90 shadow-sm`
    : variant === "ghost"
    ? `${base} text-muted-foreground hover:text-foreground hover:bg-muted`
    : `${base} bg-secondary text-secondary-foreground hover:bg-secondary/70 border border-border`;
  return <button onClick={onClick} disabled={disabled} className={styles}>{children}</button>;
}

// ── Pro Lock Badge ─────────────────────────────────────────────────────────────
function ProBadge({ onClick }: { onClick: () => void }) {
  return (
    <button onClick={onClick} className="flex items-center gap-1 text-xs font-bold bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full border border-amber-300 hover:bg-amber-200 transition-colors shrink-0">
      🔒 PRO
    </button>
  );
}

// ── Health Score Ring ─────────────────────────────────────────────────────────
function HealthRing({ score }: { score: number }) {
  const color = score >= 75 ? "#22c55e" : score >= 50 ? "#f59e0b" : "#ef4444";
  const label = score >= 75 ? "Excellent" : score >= 60 ? "Good" : score >= 45 ? "Fair" : "Needs Work";
  const r = 28;
  const circ = 2 * Math.PI * r;
  const dash = (score / 100) * circ;
  return (
    <div className="flex items-center gap-3">
      <div className="relative w-16 h-16 shrink-0">
        <svg className="w-16 h-16 -rotate-90" viewBox="0 0 72 72">
          <circle cx="36" cy="36" r={r} fill="none" stroke="#e2e8f0" strokeWidth="6" />
          <circle cx="36" cy="36" r={r} fill="none" stroke={color} strokeWidth="6"
            strokeDasharray={`${dash} ${circ}`} strokeLinecap="round" />
        </svg>
        <div className="absolute inset-0 flex items-center justify-center">
          <span className="text-sm font-extrabold" style={{ color }}>{score}</span>
        </div>
      </div>
      <div>
        <p className="text-sm font-bold text-foreground">Financial Health</p>
        <p className="text-xs font-semibold" style={{ color }}>{label}</p>
      </div>
    </div>
  );
}

// ── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const config = useAppConfig();
  const appName = config.appearance.name;
  const appTagline = config.appearance.tagline;
  const { symbol, setSymbol, fmt, fmtShort } = useCurrency();
  const [showCurrencyPicker, setShowCurrencyPicker] = useState(false);

  const [step, setStep] = useState<Step>("idle");
  const [error, setError] = useState<string>("");
  const [summary, setSummary] = useState<Summary | null>(null);
  const [highlighting, setHighlighting] = useState(false);
  const [highlightDone, setHighlightDone] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [clearDone, setClearDone] = useState(false);
  const [reAnalyzing, setReAnalyzing] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [activeTab, setActiveTab] = useState<ResultTab>("overview");
  const [csvText, setCsvText] = useState("");
  const [csvError, setCsvError] = useState("");
  const [isPro, setIsPro] = useState(false);
  const [showPayment, setShowPayment] = useState(false);
  const [paymentMode, setPaymentMode] = useState<"pay" | "key">("pay");
  const [pendingAction, setPendingAction] = useState<"highlight" | "export" | "csv" | "pdf" | "budget" | "recurring" | null>(null);
  const [txSearch, setTxSearch] = useState("");
  const [txCategoryFilter, setTxCategoryFilter] = useState("All");
  const [txTypeFilter, setTxTypeFilter] = useState<"all" | "credit" | "debit">("all");
  const [budgets, setBudgets] = useState<Record<string, string>>({});
  const [showDuplicates, setShowDuplicates] = useState(false);
  const [showRecurring, setShowRecurring] = useState(false);
  const [showPdfDialog, setShowPdfDialog] = useState(false);
  const [pdfEmail, setPdfEmail] = useState("");
  const [pdfEmailSending, setPdfEmailSending] = useState(false);
  const [pdfEmailSent, setPdfEmailSent] = useState(false);
  const [pdfEmailError, setPdfEmailError] = useState("");
  const [actionError, setActionError] = useState("");
  const [categoryOverrides, setCategoryOverrides] = useState<Record<number, Category>>({});
  const [expandedTxRow, setExpandedTxRow] = useState<number | null>(null);
  const [txNotes, setTxNotes] = useState<Record<number, string>>({});
  const [flaggedRows, setFlaggedRows] = useState<Set<number>>(new Set());
  const [showFlagged, setShowFlagged] = useState(false);

  const lowestPlanPrice = config.plans.length > 0 ? Math.min(...config.plans.map((p) => p.price)) : 5;
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    const key = getLicense();
    if (key) checkLicenseValid(key).then((valid) => setIsPro(valid));
  }, []);

  const requirePro = useCallback((action: typeof pendingAction) => {
    if (isPro) return true;
    setPendingAction(action);
    setShowPayment(true);
    return false;
  }, [isPro]);

  const onPaymentUnlocked = useCallback(() => {
    setIsPro(true);
    setShowPayment(false);
    const action = pendingAction;
    setPendingAction(null);
    if (action === "highlight") doHighlight();
    if (action === "export") doExport();
    if (action === "csv") doExportCsv();
    if (action === "pdf") setShowPdfDialog(true);
    if (action === "budget") setActiveTab("budget");
    if (action === "recurring") setShowRecurring(true);
  }, [pendingAction]); // eslint-disable-line react-hooks/exhaustive-deps

  const analyzeSheet = useCallback(async () => {
    setStep("loading"); setError(""); setSummary(null); setExportDone(false);
    try {
      const txns: Transaction[] = await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await ctx.sync();
        const columnMap: ColumnMap | null = await detectColumns(sheet);
        if (!columnMap) throw new Error("Could not find required columns in row 1.\n\nSupported formats:\n• Date, Description, Amount\n• Date, Description, Debit, Credit (with optional Balance)\n\nMake sure your sheet has a header row with these column names.");
        return await readTransactions(sheet, columnMap);
      });
      if (txns.length === 0) throw new Error("No transactions found. Check that the sheet has data rows below the header.");
      setSummary(buildSummary(txns));
      setStep("results");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setStep("error");
    }
  }, []);

  const doReAnalyze = useCallback(async () => {
    setReAnalyzing(true);
    setActionError("");
    try {
      const txns: Transaction[] = await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await ctx.sync();
        const columnMap: ColumnMap | null = await detectColumns(sheet);
        if (!columnMap) throw new Error("Could not find required columns (Date, Description, Amount) in row 1.");
        return await readTransactions(sheet, columnMap);
      });
      if (txns.length === 0) throw new Error("No transactions found in the active sheet.");
      setSummary(buildSummary(txns));
      setExportDone(false);
      setHighlightDone(false);
      setClearDone(false);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally {
      setReAnalyzing(false);
    }
  }, []);

  const importAndAnalyzeCsv = useCallback(async () => {
    setCsvError("");
    const parsed = parsePastedText(csvText);
    if (!parsed) { setCsvError("Could not parse. Make sure there's a header row and at least one data row."); return; }
    setStep("importing");
    try {
      await runExcel(async (ctx) => { await writeToExcelSheet(parsed, ctx); });
      setStep("loading");
      const txns: Transaction[] = await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        await ctx.sync();
        const columnMap = await detectColumns(sheet);
        if (!columnMap) throw new Error("Columns could not be mapped after import.");
        return await readTransactions(sheet, columnMap);
      });
      if (txns.length === 0) throw new Error("No transactions were parsed from the pasted data.");
      setSummary(buildSummary(txns));
      setStep("results");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setStep("error");
    }
  }, [csvText]);

  const doHighlight = useCallback(async () => {
    if (!summary) return;
    setHighlighting(true);
    setHighlightDone(false);
    setActionError("");
    try {
      await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const columnMap = await detectColumns(sheet);
        if (columnMap) await highlightTransactions(sheet, summary.transactions, columnMap);
        else throw new Error("Could not detect columns on the active sheet.");
      });
      setHighlightDone(true);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally { setHighlighting(false); }
  }, [summary]);

  const doClearHighlights = useCallback(async () => {
    setClearing(true);
    setClearDone(false);
    setActionError("");
    try {
      await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        await clearHighlights(sheet);
      });
      setClearDone(true);
      setHighlightDone(false);
      setTimeout(() => setClearDone(false), 2000);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally {
      setClearing(false);
    }
  }, []);

  const doExport = useCallback(async () => {
    if (!summary) return;
    setExporting(true);
    setActionError("");
    try {
      await runExcel(async (ctx) => { await createSummarySheet(summary, ctx, symbol); });
      setExportDone(true);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally { setExporting(false); }
  }, [summary]);

  const doExportCsv = useCallback(() => {
    if (!summary) return;
    exportToCsv(summary, txNotes, flaggedRows);
  }, [summary, txNotes, flaggedRows]);

  const doExportPdf = useCallback(() => {
    if (!summary) return;
    exportToPdf(summary, appName, symbol, txNotes, flaggedRows);
  }, [summary, appName, symbol, txNotes, flaggedRows]);

  const doSendReportByEmail = useCallback(async (email: string) => {
    if (!summary) return;
    setPdfEmailSending(true);
    setPdfEmailError("");
    try {
      const html = buildReportHtml(summary, appName, symbol, txNotes, flaggedRows);
      const apiBase = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");
      const res = await fetch(`${apiBase}/api/send-report`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, html, appName }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || `Server error (${res.status})`);
      setPdfEmailSent(true);
    } catch (e: unknown) {
      setPdfEmailError(e instanceof Error ? e.message : String(e));
    } finally {
      setPdfEmailSending(false);
    }
  }, [summary, appName]);

  const handleHighlight = useCallback(() => { if (requirePro("highlight")) doHighlight(); }, [requirePro, doHighlight]);
  const handleExport = useCallback(() => { if (requirePro("export")) doExport(); }, [requirePro, doExport]);
  const handleExportCsv = useCallback(() => { if (requirePro("csv")) doExportCsv(); }, [requirePro, doExportCsv]);
  const handleExportPdf = useCallback(() => {
    if (requirePro("pdf")) {
      setPdfEmailSent(false); setPdfEmailError(""); setPdfEmail("");
      setShowPdfDialog(true);
    }
  }, [requirePro]);
  const handleBudgetTab = useCallback(() => { if (requirePro("budget")) setActiveTab("budget"); }, [requirePro]);
  const handleToggleRecurring = useCallback(() => { if (requirePro("recurring")) setShowRecurring((v) => !v); }, [requirePro]);

  const handleRecategorize = useCallback((row: number, newCat: Category) => {
    if (!summary) return;
    const newOverrides = { ...categoryOverrides, [row]: newCat };
    setCategoryOverrides(newOverrides);
    setExpandedTxRow(null);
    const updatedTxns = summary.transactions.map((tx) =>
      newOverrides[tx.row] ? { ...tx, category: newOverrides[tx.row] } : tx
    );
    setSummary(buildSummary(updatedTxns));
  }, [summary, categoryOverrides]);

  const reset = () => {
    setStep("idle"); setSummary(null); setError("");
    setExportDone(false); setActiveTab("overview");
    setCsvText(""); setCsvError("");
    setTxSearch(""); setTxCategoryFilter("All"); setTxTypeFilter("all");
    setBudgets({}); setShowDuplicates(false); setShowRecurring(false);
    setActionError(""); setShowPdfDialog(false); setHighlightDone(false);
    setClearing(false); setClearDone(false);
    setCategoryOverrides({}); setExpandedTxRow(null); setTxNotes({});
    setFlaggedRows(new Set()); setShowFlagged(false);
  };

  const openSubscription = () => setStep("subscription");

  const topCategories = summary
    ? Object.entries(summary.byCategory).sort((a, b) => b[1].total - a[1].total)
    : [];
  const maxCatTotal = topCategories.length > 0 ? topCategories[0][1].total : 1;

  const filteredTxns: Transaction[] = summary
    ? summary.transactions.filter((tx) => {
        const matchSearch = txSearch === "" || tx.description.toLowerCase().includes(txSearch.toLowerCase());
        const matchCat = txCategoryFilter === "All" || tx.category.name === txCategoryFilter;
        const matchType = txTypeFilter === "all" || tx.type === txTypeFilter;
        const matchDupe = !showDuplicates || summary.duplicateRows.has(tx.row);
        const matchFlag = !showFlagged || flaggedRows.has(tx.row);
        return matchSearch && matchCat && matchType && matchDupe && matchFlag;
      })
    : [];

  const categoryNames = summary ? ["All", ...Object.keys(summary.byCategory).sort()] : ["All"];

  return (
    <div className="flex flex-col h-screen bg-background text-foreground">
      {/* Payment overlay */}
      {showPayment && (
        <PaymentGate
          initialMode={paymentMode}
          onUnlocked={onPaymentUnlocked}
          onDismiss={() => { setShowPayment(false); setPendingAction(null); setPaymentMode("pay"); }}
        />
      )}

      {/* PDF / Email dialog */}
      {showPdfDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-[340px] overflow-hidden">
            {/* Header */}
            <div className="flex items-center justify-between px-5 py-4 border-b border-border">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-lg bg-primary/10 flex items-center justify-center">
                  <svg className="w-4 h-4 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                    <line x1="9" y1="15" x2="15" y2="15"/><line x1="9" y1="11" x2="15" y2="11"/>
                  </svg>
                </div>
                <p className="text-sm font-bold text-foreground">Export Report</p>
              </div>
              <button
                onClick={() => setShowPdfDialog(false)}
                className="text-muted-foreground hover:text-foreground transition-colors"
                data-testid="button-close-pdf-dialog"
              >
                <svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>

            <div className="p-5 space-y-3">
              {/* Open PDF button */}
              <button
                onClick={() => { doExportPdf(); setShowPdfDialog(false); }}
                className="w-full flex items-center gap-3.5 bg-primary text-white rounded-xl px-4 py-3.5 hover:bg-primary/90 transition-colors text-left group"
                data-testid="button-open-pdf"
              >
                <div className="w-9 h-9 rounded-lg bg-white/20 flex items-center justify-center shrink-0">
                  <svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <polyline points="6 9 12 15 18 9"/>
                  </svg>
                </div>
                <div>
                  <p className="text-sm font-bold">Open PDF</p>
                  <p className="text-xs opacity-80 mt-0.5">Opens print dialog to save as PDF</p>
                </div>
              </button>

              {/* Divider */}
              <div className="flex items-center gap-3">
                <div className="flex-1 h-px bg-border" />
                <span className="text-xs text-muted-foreground font-medium">or</span>
                <div className="flex-1 h-px bg-border" />
              </div>

              {/* Email section */}
              {pdfEmailSent ? (
                <div className="flex flex-col items-center gap-2 py-4 text-center">
                  <div className="w-11 h-11 rounded-full bg-green-100 flex items-center justify-center">
                    <svg className="w-6 h-6 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="20 6 9 17 4 12"/>
                    </svg>
                  </div>
                  <p className="text-sm font-bold text-green-700">Report sent!</p>
                  <p className="text-xs text-muted-foreground">Check <span className="font-semibold">{pdfEmail}</span> for your report</p>
                  <button
                    onClick={() => { setPdfEmailSent(false); setPdfEmail(""); }}
                    className="text-xs text-primary hover:underline mt-1"
                  >Send to another email</button>
                </div>
              ) : (
                <div className="space-y-2.5">
                  <p className="text-xs font-semibold text-foreground">Send report to email</p>
                  <input
                    type="email"
                    value={pdfEmail}
                    onChange={(e) => { setPdfEmail(e.target.value); setPdfEmailError(""); }}
                    placeholder="you@example.com"
                    className="w-full rounded-xl border border-border bg-background px-3.5 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                    onKeyDown={(e) => { if (e.key === "Enter" && pdfEmail.includes("@")) doSendReportByEmail(pdfEmail); }}
                    data-testid="input-pdf-email"
                    disabled={pdfEmailSending}
                  />
                  {pdfEmailError && (
                    <p className="text-xs text-destructive font-medium">{pdfEmailError}</p>
                  )}
                  <button
                    onClick={() => doSendReportByEmail(pdfEmail)}
                    disabled={pdfEmailSending || !pdfEmail.includes("@")}
                    className="w-full flex items-center justify-center gap-2 bg-accent text-white rounded-xl px-4 py-2.5 text-sm font-bold hover:bg-accent/90 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                    data-testid="button-send-report-email"
                  >
                    {pdfEmailSending ? (
                      <><span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />Sending…</>
                    ) : (
                      <><svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                        <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
                      </svg>Send Report</>
                    )}
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── Header ── */}
      <header className="flex items-center gap-3 px-4 py-3 bg-white border-b border-border shadow-sm shrink-0">
        <div className="flex items-center justify-center w-9 h-9 rounded-xl overflow-hidden shadow-sm">
          <img src={iconLogo} alt="App logo" className="w-9 h-9 object-cover" />
        </div>
        <div className="flex-1 min-w-0">
          <div className="font-bold text-[15px] leading-tight text-foreground truncate">{appName}</div>
          <div className="text-xs text-muted-foreground font-medium">Excel Add-in</div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {/* Currency picker */}
          <div className="relative">
            <button
              onClick={() => setShowCurrencyPicker((v) => !v)}
              data-testid="button-currency-picker"
              className="flex items-center gap-1 text-xs font-bold bg-muted text-muted-foreground hover:bg-muted/70 hover:text-foreground px-2 py-1 rounded-lg border border-border transition-colors"
              title="Change currency symbol"
            >
              <span className="font-mono">{symbol}</span>
              <svg className="w-3 h-3 opacity-60" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="6 9 12 15 18 9" />
              </svg>
            </button>
            {showCurrencyPicker && (
              <>
                <div className="fixed inset-0 z-40" onClick={() => setShowCurrencyPicker(false)} />
                <div className="absolute right-0 top-full mt-1.5 z-50 bg-white border border-border rounded-xl shadow-xl w-52 overflow-hidden">
                  <div className="px-3 py-2 border-b border-border">
                    <p className="text-[11px] font-bold text-muted-foreground uppercase tracking-wider">Currency Symbol</p>
                  </div>
                  <div className="max-h-64 overflow-y-auto py-1">
                    {CURRENCIES.map((c) => (
                      <button
                        key={c.symbol}
                        data-testid={`button-currency-${c.symbol}`}
                        onClick={() => { setSymbol(c.symbol); setShowCurrencyPicker(false); }}
                        className={`w-full flex items-center gap-3 px-3 py-2 text-left hover:bg-muted/50 transition-colors ${symbol === c.symbol ? "bg-primary/5" : ""}`}
                      >
                        <span className={`font-mono text-sm font-bold w-8 shrink-0 ${symbol === c.symbol ? "text-primary" : "text-foreground"}`}>{c.symbol}</span>
                        <span className="text-xs text-muted-foreground truncate">{c.label}</span>
                        {symbol === c.symbol && (
                          <svg className="w-3.5 h-3.5 text-primary shrink-0 ml-auto" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                            <polyline points="20 6 9 17 4 12" />
                          </svg>
                        )}
                      </button>
                    ))}
                  </div>
                </div>
              </>
            )}
          </div>

          {isPro && (
            <button
              onClick={openSubscription}
              className="flex items-center gap-1 text-xs font-bold bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full border border-amber-200 tracking-wide hover:bg-amber-200 transition-colors"
            >
              <img src={iconPro} alt="Pro" className="w-3.5 h-3.5 object-contain" /> PRO
            </button>
          )}
          {step === "results" && (
            <button
              onClick={doReAnalyze}
              disabled={reAnalyzing}
              title="Re-read the active sheet and refresh results"
              data-testid="button-reanalyze"
              className="flex items-center justify-center w-8 h-8 rounded-lg text-muted-foreground hover:text-foreground hover:bg-muted transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
            >
              <svg
                className={`w-4 h-4 ${reAnalyzing ? "animate-spin" : ""}`}
                viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"
              >
                <polyline points="23 4 23 10 17 10" />
                <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10" />
              </svg>
            </button>
          )}
          {(step === "results" || step === "error" || step === "paste" || step === "subscription") && (
            <button onClick={reset}
              className="flex items-center gap-1 text-sm font-semibold text-muted-foreground hover:text-foreground px-2.5 py-1.5 rounded-lg hover:bg-muted transition-colors">
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="15 18 9 12 15 6" />
              </svg>
              Back
            </button>
          )}
        </div>
      </header>

      {/* ── Body ── */}
      <div className="flex-1 overflow-y-auto">

        {/* ── IDLE ── */}
        {step === "idle" && (
          <div className="flex flex-col items-center justify-center min-h-full px-5 py-8 text-center gap-7">
            {/* Hero */}
            <div className="flex flex-col items-center gap-3">
              <div className="w-16 h-16 rounded-2xl overflow-hidden shadow-md">
                <img src={iconLogo} alt="Bank Statement Analyzer" className="w-16 h-16 object-cover" />
              </div>
              <div>
                <h1 className="text-xl font-extrabold text-foreground tracking-tight mb-1">{appName}</h1>
                <p className="text-sm text-muted-foreground max-w-[230px] mx-auto leading-relaxed">{appTagline}</p>
              </div>
            </div>

            {/* Action cards */}
            <div className="w-full max-w-[300px] space-y-3">
              <button onClick={analyzeSheet}
                className="w-full text-left bg-white border-2 border-border rounded-xl p-4 hover:border-primary/60 hover:shadow-md transition-all group">
                <div className="flex items-center gap-3.5">
                  <div className="w-10 h-10 rounded-xl bg-primary/10 flex items-center justify-center shrink-0 group-hover:bg-primary/20 transition-colors overflow-hidden">
                    <img src={iconAnalyze} alt="Analyze sheet" className="w-8 h-8 object-contain" />
                  </div>
                  <div>
                    <p className="text-sm font-bold text-foreground">Analyze Active Sheet</p>
                    <p className="text-xs text-muted-foreground mt-0.5">Read data already in Excel</p>
                  </div>
                  <svg className="w-4 h-4 text-muted-foreground ml-auto shrink-0 group-hover:text-primary transition-colors" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <polyline points="9 18 15 12 9 6" />
                  </svg>
                </div>
              </button>

              <button onClick={() => setStep("paste")}
                className="w-full text-left bg-white border-2 border-border rounded-xl p-4 hover:border-primary/60 hover:shadow-md transition-all group">
                <div className="flex items-center gap-3.5">
                  <div className="w-10 h-10 rounded-xl bg-accent/10 flex items-center justify-center shrink-0 group-hover:bg-accent/20 transition-colors overflow-hidden">
                    <img src={iconPaste} alt="Paste CSV" className="w-8 h-8 object-contain" />
                  </div>
                  <div>
                    <p className="text-sm font-bold text-foreground">Paste CSV / Text</p>
                    <p className="text-xs text-muted-foreground mt-0.5">Import from your bank portal</p>
                  </div>
                  <svg className="w-4 h-4 text-muted-foreground ml-auto shrink-0 group-hover:text-primary transition-colors" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <polyline points="9 18 15 12 9 6" />
                  </svg>
                </div>
              </button>
            </div>

            {/* Free vs Pro */}
            {!isPro && (
              <div className="w-full max-w-[300px] rounded-xl border-2 border-amber-200 overflow-hidden shadow-sm">
                <div className="bg-gradient-to-r from-amber-50 to-orange-50 px-4 py-2.5 flex items-center justify-between border-b border-amber-200">
                  <p className="text-sm font-bold text-amber-800">Free vs Pro</p>
                  <button onClick={() => { setPaymentMode("pay"); setShowPayment(true); }}
                    className="text-xs font-bold text-white bg-amber-500 hover:bg-amber-600 px-3 py-1 rounded-lg transition-colors shadow-sm">
                    From ${lowestPlanPrice} USDT
                  </button>
                </div>
                <div className="p-3 space-y-1.5 bg-white">
                  {[
                    { label: "Analyze transactions", free: true },
                    { label: "Auto-categorize (14 categories)", free: true },
                    { label: "Monthly trend chart", free: true },
                    { label: "Financial health score", free: true },
                    { label: "Search & filter transactions", free: true },
                    { label: "Top merchants breakdown", free: true },
                    { label: "Highlight cells by category", free: false },
                    { label: "Export Excel summary sheet", free: false },
                    { label: "Download CSV with categories", free: false },
                    { label: "Export color-coded PDF report", free: false },
                    { label: "Recurring subscriptions detector", free: false },
                    { label: "Duplicate transaction flags", free: false },
                    { label: "Category budget tracker", free: false },
                  ].map((f) => (
                    <div key={f.label} className="flex items-center gap-2.5">
                      <span className={`text-sm shrink-0 ${f.free ? "text-green-600" : "text-muted-foreground/40"}`}>
                        {f.free ? "✓" : "🔒"}
                      </span>
                      <span className={`text-xs ${f.free ? "text-foreground font-medium" : "text-muted-foreground/60"}`}>
                        {f.label}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {!isPro && (
              <button
                onClick={() => { setPaymentMode("key"); setShowPayment(true); }}
                className="text-xs text-muted-foreground hover:text-primary transition-colors underline underline-offset-2"
              >
                Already have a license key? Enter it here
              </button>
            )}
          </div>
        )}

        {/* ── PASTE CSV ── */}
        {step === "paste" && (
          <div className="flex flex-col h-full p-4 gap-4">
            <div>
              <h2 className="text-base font-bold text-foreground mb-1">Paste Bank Statement</h2>
              <p className="text-sm text-muted-foreground leading-relaxed">
                Copy your statement from your bank's portal or CSV export and paste it below.
              </p>
            </div>
            <textarea
              ref={textareaRef}
              value={csvText}
              onChange={(e) => { setCsvText(e.target.value); setCsvError(""); }}
              placeholder={"Format 1 — Debit/Credit columns (recommended):\nDate,Description,Category,Debit,Credit,Balance\n6/1/2026,Opening Balance,,,10000,10000\n6/2/2026,Payroll,Income,,4500,14500\n6/2/2026,Starbucks,Dining,6.75,,14493.25\n\nFormat 2 — Single Amount column:\nDate,Description,Amount,Type\n01/06/2026,SALARY JUNE,650000,CR\n02/06/2026,SHOPRITE,-45000,DR"}
              className="flex-1 w-full rounded-xl border-2 border-border bg-white p-3 text-sm font-mono resize-none focus:outline-none focus:border-primary placeholder:text-muted-foreground/40 transition-colors"
              spellCheck={false}
            />
            {csvError && (
              <div className="flex items-start gap-2 bg-destructive/10 border border-destructive/20 rounded-xl px-3 py-2.5">
                <svg className="w-4 h-4 text-destructive shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
                <p className="text-sm text-destructive font-medium">{csvError}</p>
              </div>
            )}
            {csvText.trim().length > 0 && (() => {
              const parsed = parsePastedText(csvText);
              return parsed ? (
                <p className="text-sm text-muted-foreground">
                  Detected <span className="font-bold text-foreground">{parsed.rawCount} rows</span> ·{" "}
                  <span className="font-bold text-foreground">{parsed.headers.length} columns</span>:{" "}
                  {parsed.headers.join(", ")}
                </p>
              ) : (
                <p className="text-sm text-amber-600 font-medium">⚠ Need at least a header row and one data row.</p>
              );
            })()}
            <div className="flex gap-3 shrink-0">
              <ActionBtn variant="secondary" onClick={() => { setCsvText(""); setCsvError(""); }}>
                Clear
              </ActionBtn>
              <ActionBtn onClick={importAndAnalyzeCsv} disabled={!csvText.trim()}>
                <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="16 16 12 12 8 16" /><line x1="12" y1="12" x2="12" y2="21" />
                  <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3" />
                </svg>
                Import &amp; Analyze
              </ActionBtn>
            </div>
          </div>
        )}

        {/* ── LOADING / IMPORTING ── */}
        {(step === "importing" || step === "loading") && (
          <div className="flex flex-col items-center justify-center h-full gap-4">
            <div className="relative w-14 h-14">
              <div className="w-14 h-14 border-4 border-primary/20 rounded-full" />
              <div className="absolute inset-0 w-14 h-14 border-4 border-primary border-t-transparent rounded-full animate-spin" />
            </div>
            <div className="text-center">
              <p className="text-base font-bold text-foreground">
                {step === "importing" ? "Writing to Excel…" : "Analyzing transactions…"}
              </p>
              <p className="text-sm text-muted-foreground mt-1">This will only take a moment</p>
            </div>
          </div>
        )}

        {/* ── ERROR ── */}
        {step === "error" && (
          <div className="flex flex-col items-center justify-center h-full px-5 py-8 text-center">
            <div className="w-14 h-14 rounded-2xl bg-destructive/10 flex items-center justify-center mb-4">
              <svg className="w-7 h-7 text-destructive" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
              </svg>
            </div>
            <h3 className="text-base font-bold text-destructive mb-2">Analysis Failed</h3>
            <p className="text-sm text-muted-foreground mb-5 whitespace-pre-wrap leading-relaxed max-w-xs">{error}</p>
            <button onClick={reset}
              className="text-sm font-bold bg-muted hover:bg-muted/70 text-foreground px-6 py-2.5 rounded-xl transition-colors">
              ← Try Again
            </button>
          </div>
        )}

        {/* ── RESULTS ── */}
        {step === "results" && summary && (
          <div className="flex flex-col h-full">

            {/* Action buttons row */}
            <div className="flex gap-2 px-4 pt-4 pb-2 shrink-0 flex-wrap">
              <ActionBtn variant="secondary" onClick={handleHighlight} disabled={highlighting || clearing} size="sm">
                {highlighting
                  ? <span className="w-3.5 h-3.5 border-2 border-current border-t-transparent rounded-full animate-spin" />
                  : highlightDone
                  ? <svg className="w-3.5 h-3.5 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
                  : <img src={iconHighlight} alt="Highlight" className="w-3.5 h-3.5 object-contain opacity-90" />
                }
                {highlightDone ? "Done!" : "Highlight"}
                {!isPro && <span className="text-[10px] opacity-60">🔒</span>}
              </ActionBtn>
              <ActionBtn variant="ghost" onClick={doClearHighlights} disabled={clearing || highlighting} size="sm">
                {clearing
                  ? <span className="w-3.5 h-3.5 border-2 border-current border-t-transparent rounded-full animate-spin" />
                  : clearDone
                  ? <svg className="w-3.5 h-3.5 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
                  : <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M20 5H9l-7 7 7 7h11a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2z"/><line x1="18" y1="9" x2="12" y2="15"/><line x1="12" y1="9" x2="18" y2="15"/></svg>
                }
                {clearDone ? "Cleared!" : "Clear"}
              </ActionBtn>
              <ActionBtn onClick={handleExport} disabled={exporting} size="sm">
                {exporting
                  ? <span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  : exportDone
                  ? <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
                  : <img src={iconExport} alt="Export" className="w-3.5 h-3.5 object-contain opacity-90" />
                }
                {exportDone ? "Exported!" : "Export"}
                {!isPro && <span className="text-[10px] opacity-60">🔒</span>}
              </ActionBtn>
              <ActionBtn variant="secondary" onClick={handleExportCsv} size="sm">
                <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="7 10 12 15 17 10" />
                  <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                CSV
                {!isPro && <span className="text-[10px] opacity-60">🔒</span>}
              </ActionBtn>
              <ActionBtn variant="secondary" onClick={handleExportPdf} size="sm">
                <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                  <polyline points="14 2 14 8 20 8"/>
                  <line x1="9" y1="15" x2="15" y2="15"/>
                  <line x1="9" y1="11" x2="15" y2="11"/>
                </svg>
                PDF
                {!isPro && <span className="text-[10px] opacity-60">🔒</span>}
              </ActionBtn>
            </div>

            {/* Action error banner */}
            {actionError && (
              <div className="mx-4 mb-1 flex items-start gap-2 bg-destructive/10 border border-destructive/20 rounded-xl px-3 py-2.5 shrink-0">
                <svg className="w-4 h-4 text-destructive shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <p className="text-xs text-destructive font-medium flex-1 leading-relaxed">{actionError}</p>
                <button onClick={() => setActionError("")} className="text-destructive/60 hover:text-destructive shrink-0 transition-colors">
                  <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                  </svg>
                </button>
              </div>
            )}

            {/* Upsell banner */}
            {!isPro && (
              <button onClick={() => setShowPayment(true)}
                className="mx-4 mb-2 flex items-center justify-between bg-gradient-to-r from-amber-50 to-orange-50 border-2 border-amber-200 rounded-xl px-4 py-2.5 hover:border-amber-400 transition-all group">
                <div className="text-left">
                  <p className="text-xs font-bold text-amber-800">Unlock Pro — from ${lowestPlanPrice} USDT</p>
                  <p className="text-[11px] text-amber-600 mt-0.5">Budgets · CSV · Recurring · Duplicates · Export</p>
                </div>
                <svg className="w-4 h-4 text-amber-500 shrink-0 group-hover:translate-x-0.5 transition-transform" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="9 18 15 12 9 6" />
                </svg>
              </button>
            )}

            {/* Tabs */}
            <div className="flex border-b border-border shrink-0 px-1">
              {(["overview", "categories", "transactions", "budget"] as const).map((tab) => {
                const isProTab = tab === "budget";
                return (
                  <button key={tab} onClick={() => { if (tab === "budget") handleBudgetTab(); else setActiveTab(tab); }}
                    className={`flex-1 py-2 text-xs font-semibold capitalize transition-all flex items-center justify-center gap-1 ${activeTab === tab ? "border-b-2 border-primary text-primary" : "text-muted-foreground hover:text-foreground"}`}>
                    {tab}
                    {isProTab && !isPro && <span className="text-[9px]">🔒</span>}
                  </button>
                );
              })}
            </div>

            {/* Tab content */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4">

              {/* ── OVERVIEW TAB ── */}
              {activeTab === "overview" && (
                <>
                  {/* KPI grid */}
                  <div className="grid grid-cols-2 gap-3">
                    <KpiCard label="Income" value={fmt(summary.totalIncome)} color="text-green-600" icon={iconIncome} />
                    <KpiCard label="Expenses" value={fmt(summary.totalExpenses)} color="text-red-500" icon={iconExpenses} />
                    <KpiCard label="Net Savings" value={fmt(summary.net)} color={summary.net >= 0 ? "text-blue-600" : "text-red-500"} icon={iconSavings} />
                    <KpiCard label="Savings Rate" value={`${summary.savingsRate}%`} color={summary.savingsRate >= 20 ? "text-green-600" : summary.savingsRate >= 10 ? "text-yellow-600" : "text-red-500"} icon={iconRate} />
                  </div>

                  {/* Health Score */}
                  <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
                    <HealthRing score={summary.healthScore} />
                    {summary.healthTips.length > 0 && (
                      <div className="mt-3 space-y-1.5">
                        {summary.healthTips.map((tip, i) => (
                          <div key={i} className="flex items-start gap-2">
                            <span className="text-xs shrink-0 mt-0.5">{i === 0 && summary.healthScore >= 60 ? "💡" : "⚠️"}</span>
                            <p className="text-xs text-muted-foreground leading-relaxed">{tip}</p>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Monthly Trend Chart */}
                  {summary.monthly.length > 1 && (
                    <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
                      <p className="text-xs font-bold text-muted-foreground uppercase tracking-widest mb-3">Monthly Trend</p>
                      <ResponsiveContainer width="100%" height={120}>
                        <BarChart data={summary.monthly} barGap={2} barCategoryGap="30%">
                          <XAxis dataKey="month" tick={{ fontSize: 10 }} tickLine={false} axisLine={false} />
                          <YAxis hide tickFormatter={(v) => fmtShort(v)} />
                          <Tooltip
                            formatter={(value: number, name: string) => [fmt(value), name === "income" ? "Income" : "Expenses"]}
                            contentStyle={{ fontSize: 11, borderRadius: 8, border: "1px solid #e2e8f0" }}
                          />
                          <Bar dataKey="income" fill="#22c55e" radius={[3, 3, 0, 0]} />
                          <Bar dataKey="expenses" fill="#ef4444" radius={[3, 3, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                      <div className="flex items-center justify-center gap-4 mt-2">
                        <div className="flex items-center gap-1.5"><div className="w-2.5 h-2.5 rounded-full bg-green-500" /><span className="text-[11px] text-muted-foreground font-medium">Income</span></div>
                        <div className="flex items-center gap-1.5"><div className="w-2.5 h-2.5 rounded-full bg-red-500" /><span className="text-[11px] text-muted-foreground font-medium">Expenses</span></div>
                      </div>
                    </div>
                  )}

                  {/* Month-over-Month Comparison */}
                  {summary.momMonths && summary.momChanges.length > 0 && (
                    <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
                      <div className="flex items-center justify-between mb-3">
                        <p className="text-xs font-bold text-muted-foreground uppercase tracking-widest">Month vs Month</p>
                        <div className="flex items-center gap-1.5">
                          <span className="text-[11px] font-semibold text-muted-foreground bg-muted px-2 py-0.5 rounded-full">{summary.momMonths[1]}</span>
                          <svg className="w-3 h-3 text-muted-foreground" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="9 18 15 12 9 6"/></svg>
                          <span className="text-[11px] font-bold text-primary bg-primary/10 px-2 py-0.5 rounded-full">{summary.momMonths[0]}</span>
                        </div>
                      </div>

                      {/* Total expenses delta */}
                      {(() => {
                        const prevTotal = summary.momChanges.reduce((s, c) => s + c.previous, 0);
                        const curTotal = summary.momChanges.reduce((s, c) => s + c.current, 0);
                        const totalDelta = curTotal - prevTotal;
                        const totalPct = prevTotal > 0 ? Math.round((totalDelta / prevTotal) * 100) : 0;
                        const isUp = totalDelta > 0;
                        return (
                          <div className={`flex items-center justify-between rounded-xl px-3 py-2.5 mb-3 ${isUp ? "bg-red-50 border border-red-100" : "bg-green-50 border border-green-100"}`}>
                            <span className="text-xs font-semibold text-foreground">Total Expenses</span>
                            <div className="flex items-center gap-1.5">
                              <span className={`text-xs font-extrabold ${isUp ? "text-red-600" : "text-green-600"}`}>{fmt(curTotal)}</span>
                              <span className={`flex items-center gap-0.5 text-[11px] font-bold px-1.5 py-0.5 rounded-full ${isUp ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"}`}>
                                {isUp ? "▲" : "▼"} {Math.abs(totalPct)}%
                              </span>
                            </div>
                          </div>
                        );
                      })()}

                      <div className="space-y-2.5">
                        {summary.momChanges.slice(0, 6).map((c) => {
                          const isUp = c.delta > 0;
                          const isNew = c.isNew;
                          const isGone = c.isGone;
                          return (
                            <div key={c.category} className="flex items-center gap-2">
                              <span className={`text-xs px-1.5 py-0.5 rounded-md shrink-0 ${c.className}`}>{c.category}</span>
                              <div className="flex-1 min-w-0">
                                {isNew ? (
                                  <div className="h-1.5 bg-muted rounded-full overflow-hidden">
                                    <div className="h-1.5 rounded-full bg-amber-400" style={{ width: "100%" }} />
                                  </div>
                                ) : isGone ? (
                                  <div className="h-1.5 bg-muted rounded-full" />
                                ) : (
                                  <div className="relative h-1.5 bg-muted rounded-full overflow-hidden">
                                    <div className="absolute inset-y-0 left-0 rounded-full" style={{ width: `${Math.min((Math.min(c.previous, c.current) / Math.max(c.previous, c.current)) * 100, 100)}%`, backgroundColor: c.color + "66" }} />
                                    <div className="absolute inset-y-0 rounded-full" style={{
                                      left: isUp ? `${Math.min((c.previous / Math.max(c.previous, c.current)) * 100, 100)}%` : `${Math.min((c.current / Math.max(c.previous, c.current)) * 100, 100)}%`,
                                      width: `${Math.abs(((c.delta) / Math.max(c.previous, c.current))) * 100}%`,
                                      backgroundColor: isUp ? "#ef4444" : "#22c55e",
                                    }} />
                                  </div>
                                )}
                              </div>
                              <div className="flex items-center gap-1 shrink-0">
                                {isNew ? (
                                  <span className="text-[10px] font-bold text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">NEW</span>
                                ) : isGone ? (
                                  <span className="text-[10px] font-bold text-muted-foreground bg-muted px-1.5 py-0.5 rounded">GONE</span>
                                ) : (
                                  <span className={`text-[11px] font-bold flex items-center gap-0.5 ${isUp ? "text-red-600" : "text-green-600"}`}>
                                    {isUp ? "▲" : "▼"}{Math.abs(c.pctChange)}%
                                  </span>
                                )}
                                <span className="text-[11px] font-semibold text-muted-foreground">{fmtShort(c.current)}</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      {summary.momChanges.length > 6 && (
                        <p className="text-[11px] text-muted-foreground text-center mt-3">
                          +{summary.momChanges.length - 6} more categories in the Categories tab
                        </p>
                      )}
                    </div>
                  )}

                  {/* Top Spending */}
                  <div>
                    <p className="text-xs font-bold text-muted-foreground uppercase tracking-widest mb-3">Top Spending</p>
                    {topCategories.slice(0, 4).map(([name, info]) => (
                      <div key={name} className="mb-3">
                        <div className="flex justify-between items-center mb-1.5">
                          <span className={`text-xs px-2 py-0.5 rounded-md ${info.className}`}>{name}</span>
                          <span className="text-sm font-bold">{fmt(info.total)}</span>
                        </div>
                        <div className="h-2 bg-muted rounded-full overflow-hidden">
                          <div className="h-2 rounded-full transition-all" style={{ width: `${(info.total / maxCatTotal) * 100}%`, backgroundColor: info.color }} />
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Duplicates alert */}
                  {summary.duplicateRows.size > 0 && (
                    <div className="flex items-center justify-between bg-amber-50 border border-amber-200 rounded-xl px-4 py-3">
                      <div className="flex items-center gap-2">
                        <span className="text-sm">⚠️</span>
                        <div>
                          <p className="text-xs font-bold text-amber-800">{summary.duplicateRows.size} possible duplicates</p>
                          <p className="text-[11px] text-amber-600">Check the Transactions tab</p>
                        </div>
                      </div>
                      {isPro
                        ? <ActionBtn size="sm" variant="secondary" onClick={() => { setActiveTab("transactions"); setShowDuplicates(true); }}>Review</ActionBtn>
                        : <ProBadge onClick={() => setShowPayment(true)} />
                      }
                    </div>
                  )}

                  <p className="text-xs text-muted-foreground font-medium text-right">
                    {summary.transactions.length} transactions analyzed
                  </p>
                </>
              )}

              {/* ── CATEGORIES TAB ── */}
              {activeTab === "categories" && (
                <div className="space-y-3">
                  {/* Top Merchants */}
                  <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
                    <p className="text-xs font-bold text-muted-foreground uppercase tracking-widest mb-3">Top Merchants</p>
                    <div className="space-y-2">
                      {summary.topMerchants.slice(0, 5).map((m, i) => (
                        <div key={i} className="flex items-center gap-2">
                          <span className="text-xs font-bold text-muted-foreground w-4 shrink-0">#{i + 1}</span>
                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-semibold truncate">{m.name}</p>
                            <span className={`text-[10px] px-1.5 py-0.5 rounded ${m.category.className}`}>{m.category.name}</span>
                          </div>
                          <div className="text-right shrink-0">
                            <p className="text-xs font-extrabold text-red-500">{fmtShort(m.total)}</p>
                            <p className="text-[10px] text-muted-foreground">{m.count}×</p>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Recurring Subscriptions */}
                  <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
                    <div className="flex items-center justify-between mb-3">
                      <p className="text-xs font-bold text-muted-foreground uppercase tracking-widest">Recurring / Subscriptions</p>
                      {!isPro
                        ? <ProBadge onClick={() => setShowPayment(true)} />
                        : <ActionBtn size="sm" variant="ghost" onClick={() => setShowRecurring((v) => !v)}>
                            {showRecurring ? "Hide" : "Show"}
                          </ActionBtn>
                      }
                    </div>
                    {isPro && showRecurring && summary.recurring.length > 0 && (
                      <div className="space-y-2">
                        {summary.recurring.map((r, i) => (
                          <div key={i} className="flex items-center gap-2 py-1.5 border-b border-border last:border-0">
                            <div className="flex-1 min-w-0">
                              <p className="text-xs font-semibold truncate">{r.description}</p>
                              <div className="flex items-center gap-1.5 mt-0.5">
                                <span className={`text-[10px] px-1.5 py-0.5 rounded ${r.category.className}`}>{r.category.name}</span>
                                <span className="text-[10px] text-muted-foreground">{r.count}× detected</span>
                              </div>
                            </div>
                            <div className="text-right shrink-0">
                              <p className="text-xs font-extrabold text-red-500">{fmtShort(r.totalAmount)}</p>
                              <p className="text-[10px] text-muted-foreground">~{fmtShort(r.avgAmount)}/time</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                    {isPro && showRecurring && summary.recurring.length === 0 && (
                      <p className="text-xs text-muted-foreground text-center py-2">No recurring patterns detected.</p>
                    )}
                    {isPro && !showRecurring && (
                      <p className="text-xs text-muted-foreground">
                        {summary.recurring.length} recurring pattern{summary.recurring.length !== 1 ? "s" : ""} found — click Show to review.
                      </p>
                    )}
                    {!isPro && (
                      <p className="text-xs text-muted-foreground">Upgrade to see subscriptions & recurring charges auto-detected.</p>
                    )}
                  </div>

                  {/* Category breakdown */}
                  {topCategories.map(([name, info], idx) => {
                    const medal = idx === 0 ? "🥇" : idx === 1 ? "🥈" : idx === 2 ? "🥉" : null;
                    return (
                    <div key={name} className={`bg-white border rounded-xl p-4 shadow-sm ${idx < 3 ? "border-amber-200" : "border-border"}`}>
                      <div className="flex items-center justify-between mb-2.5">
                        <div className="flex items-center gap-2 min-w-0">
                          {medal && <span className="text-base shrink-0 leading-none">{medal}</span>}
                          <span className={`text-xs px-2 py-0.5 rounded-md shrink-0 ${info.className}`}>{name}</span>
                          <span className="text-[11px] font-bold px-1.5 py-0.5 rounded-full bg-muted text-muted-foreground shrink-0">
                            {info.count} tx
                          </span>
                        </div>
                        <span className="text-base font-extrabold shrink-0">{fmt(info.total)}</span>
                      </div>
                      <div className="h-2 bg-muted rounded-full overflow-hidden mb-2">
                        <div className="h-2 rounded-full" style={{ width: `${(info.total / maxCatTotal) * 100}%`, backgroundColor: info.color }} />
                      </div>
                      <div className="flex items-center justify-between">
                        <p className="text-xs text-muted-foreground font-semibold">avg {fmt(Math.round(info.total / info.count))} / tx</p>
                        {summary.totalExpenses > 0 && (
                          <p className="text-xs text-muted-foreground font-semibold">{Math.round((info.total / summary.totalExpenses) * 100)}% of expenses</p>
                        )}
                      </div>
                    </div>
                  );})}
                </div>
              )}

              {/* ── TRANSACTIONS TAB ── */}
              {activeTab === "transactions" && (
                <div className="space-y-3">
                  {/* Search & Filters */}
                  <div className="space-y-2">
                    <input
                      type="text"
                      placeholder="Search transactions…"
                      value={txSearch}
                      onChange={(e) => setTxSearch(e.target.value)}
                      className="w-full rounded-xl border border-border bg-white px-3 py-2 text-sm focus:outline-none focus:border-primary transition-colors"
                    />
                    <div className="flex gap-2">
                      <select
                        value={txCategoryFilter}
                        onChange={(e) => setTxCategoryFilter(e.target.value)}
                        className="flex-1 rounded-xl border border-border bg-white px-3 py-1.5 text-xs focus:outline-none focus:border-primary transition-colors"
                      >
                        {categoryNames.map((c) => <option key={c}>{c}</option>)}
                      </select>
                      <select
                        value={txTypeFilter}
                        onChange={(e) => setTxTypeFilter(e.target.value as "all" | "credit" | "debit")}
                        className="flex-1 rounded-xl border border-border bg-white px-3 py-1.5 text-xs focus:outline-none focus:border-primary transition-colors"
                      >
                        <option value="all">All Types</option>
                        <option value="credit">Income only</option>
                        <option value="debit">Expenses only</option>
                      </select>
                    </div>
                    {/* Flagged filter */}
                    {flaggedRows.size > 0 && (
                      <div className="flex items-center justify-between bg-white border border-border rounded-xl px-3 py-2">
                        <span className="text-xs font-semibold text-foreground">🚩 Show flagged only ({flaggedRows.size})</span>
                        <button
                          data-testid="toggle-flagged-filter"
                          onClick={() => setShowFlagged((v) => !v)}
                          className={`relative w-8 h-4 rounded-full transition-colors ${showFlagged ? "bg-red-400" : "bg-muted-foreground/30"}`}
                        >
                          <div className={`absolute top-0.5 w-3 h-3 bg-white rounded-full shadow transition-transform ${showFlagged ? "translate-x-4" : "translate-x-0.5"}`} />
                        </button>
                      </div>
                    )}
                    {/* Duplicate filter */}
                    {summary.duplicateRows.size > 0 && (
                      <div className="flex items-center justify-between bg-white border border-border rounded-xl px-3 py-2">
                        <span className="text-xs font-semibold text-foreground">Show duplicates only ({summary.duplicateRows.size})</span>
                        {isPro ? (
                          <button
                            onClick={() => setShowDuplicates((v) => !v)}
                            className={`relative w-8 h-4 rounded-full transition-colors ${showDuplicates ? "bg-primary" : "bg-muted-foreground/30"}`}
                          >
                            <div className={`absolute top-0.5 w-3 h-3 bg-white rounded-full shadow transition-transform ${showDuplicates ? "translate-x-4" : "translate-x-0.5"}`} />
                          </button>
                        ) : (
                          <ProBadge onClick={() => setShowPayment(true)} />
                        )}
                      </div>
                    )}
                  </div>

                  {/* Transaction count */}
                  <p className="text-xs text-muted-foreground font-semibold">{filteredTxns.length} of {summary.transactions.length} transactions</p>

                  {/* Transaction list */}
                  {filteredTxns.length === 0 ? (
                    <div className="text-center py-8 text-muted-foreground">
                      <p className="text-2xl mb-2">🔍</p>
                      <p className="text-sm font-semibold">No transactions match</p>
                      <button onClick={() => { setTxSearch(""); setTxCategoryFilter("All"); setTxTypeFilter("all"); setShowDuplicates(false); }}
                        className="text-xs text-primary mt-2 hover:underline">Clear filters</button>
                    </div>
                  ) : (
                    filteredTxns.map((tx, i) => {
                      const isDupe = summary.duplicateRows.has(tx.row);
                      const isExpanded = expandedTxRow === tx.row;
                      const wasOverridden = !!categoryOverrides[tx.row];
                      const isFlagged = flaggedRows.has(tx.row);
                      const toggleFlag = () => setFlaggedRows((prev) => {
                        const next = new Set(prev);
                        if (next.has(tx.row)) next.delete(tx.row); else next.add(tx.row);
                        return next;
                      });
                      return (
                        <div key={i} className={`bg-white border rounded-xl shadow-sm overflow-hidden ${isFlagged ? "border-red-300 bg-red-50/30" : isDupe && isPro ? "border-amber-300 bg-amber-50/40" : isExpanded ? "border-primary" : "border-border"}`}>
                          {/* Main row */}
                          <div className="flex items-start gap-3 p-3">
                            <div className={`w-2 h-2 rounded-full shrink-0 mt-1.5 ${tx.type === "credit" ? "bg-green-500" : "bg-red-500"}`} />
                            <div className="flex-1 min-w-0">
                              <p className="text-sm font-semibold truncate">{tx.description}</p>
                              <div className="flex flex-wrap items-center gap-1.5 mt-1">
                                <span className="text-xs text-muted-foreground font-medium">{tx.date}</span>
                                <button
                                  data-testid={`btn-recategorize-${tx.row}`}
                                  onClick={() => setExpandedTxRow(isExpanded ? null : tx.row)}
                                  title="Tap to change category"
                                  className={`text-xs px-1.5 py-0.5 rounded-md transition-opacity hover:opacity-80 active:opacity-60 ${tx.category.className} ${wasOverridden ? "ring-1 ring-offset-1 ring-primary/40" : ""}`}
                                >
                                  {tx.category.name} {isExpanded ? "▲" : "▼"}
                                </button>
                                {isFlagged && (
                                  <span className="text-[10px] font-bold bg-red-100 text-red-600 px-1.5 py-0.5 rounded border border-red-200">🚩 Flagged</span>
                                )}
                                {isDupe && isPro && (
                                  <span className="text-[10px] font-bold bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded border border-amber-200">⚠ Possible Duplicate</span>
                                )}
                              </div>
                            </div>
                            <div className="flex flex-col items-end gap-1.5 shrink-0">
                              <span className={`text-sm font-extrabold ${tx.type === "credit" ? "text-green-600" : "text-red-500"}`}>
                                {tx.type === "credit" ? "+" : "−"}{fmt(tx.amount)}
                              </span>
                              <button
                                data-testid={`btn-flag-${tx.row}`}
                                onClick={toggleFlag}
                                title={isFlagged ? "Remove flag" : "Flag for review"}
                                className={`text-base leading-none transition-all hover:scale-110 active:scale-95 ${isFlagged ? "opacity-100" : "opacity-25 hover:opacity-60"}`}
                              >
                                🚩
                              </button>
                            </div>
                          </div>
                          {/* Category picker */}
                          {isExpanded && (
                            <div className="border-t border-border px-3 pb-3 pt-2 bg-muted/30">
                              <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-wider mb-2">Change category</p>
                              <div className="flex flex-wrap gap-1.5">
                                {CATEGORIES.map((cat) => (
                                  <button
                                    key={cat.name}
                                    data-testid={`btn-cat-${cat.name.replace(/\s+/g, "-").toLowerCase()}-${tx.row}`}
                                    onClick={() => handleRecategorize(tx.row, cat)}
                                    className={`text-xs px-2 py-1 rounded-lg font-medium transition-all hover:opacity-80 active:scale-95 ${cat.className} ${tx.category.name === cat.name ? "ring-2 ring-offset-1 ring-foreground/30" : ""}`}
                                  >
                                    {cat.name}
                                  </button>
                                ))}
                                <button
                                  data-testid={`btn-cat-other-${tx.row}`}
                                  onClick={() => handleRecategorize(tx.row, { name: "Other", color: "#64748b", className: "cat-other", keywords: [], type: "any" })}
                                  className={`text-xs px-2 py-1 rounded-lg font-medium transition-all hover:opacity-80 active:scale-95 cat-other ${tx.category.name === "Other" ? "ring-2 ring-offset-1 ring-foreground/30" : ""}`}
                                >
                                  Other
                                </button>
                              </div>
                              {/* Notes field */}
                              <div className="mt-3">
                                <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-wider mb-1.5">Note</p>
                                <input
                                  data-testid={`input-note-${tx.row}`}
                                  type="text"
                                  placeholder="Add a memo for this transaction…"
                                  maxLength={140}
                                  value={txNotes[tx.row] ?? ""}
                                  onChange={(e) => setTxNotes((prev) => ({ ...prev, [tx.row]: e.target.value }))}
                                  className="w-full rounded-lg border border-border bg-white px-2.5 py-1.5 text-xs focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                                />
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })
                  )}
                </div>
              )}

              {/* ── BUDGET TAB (PRO) ── */}
              {activeTab === "budget" && isPro && (
                <div className="space-y-3">
                  <div className="bg-primary/5 border border-primary/20 rounded-xl px-4 py-3">
                    <p className="text-xs font-bold text-primary mb-0.5">Category Budget Tracker</p>
                    <p className="text-xs text-muted-foreground">Set a monthly budget for each category and track your spending against it.</p>
                  </div>

                  {topCategories
                    .filter(([name]) => {
                      const cat = summary.byCategory[name];
                      return cat && name !== "Salary"; // Skip income categories
                    })
                    .map(([name, info]) => {
                      const budget = Number(budgets[name] || 0);
                      const spent = info.total;
                      const pct = budget > 0 ? Math.min((spent / budget) * 100, 100) : 0;
                      const over = budget > 0 && spent > budget;
                      const barColor = over ? "#ef4444" : pct > 80 ? "#f59e0b" : info.color;
                      return (
                        <div key={name} className="bg-white border border-border rounded-xl p-4 shadow-sm">
                          <div className="flex items-center justify-between mb-2">
                            <span className={`text-xs px-2 py-0.5 rounded-md ${info.className}`}>{name}</span>
                            <span className="text-xs font-bold text-foreground">{fmt(spent)}</span>
                          </div>
                          <div className="flex items-center gap-2 mb-2">
                            <span className="text-xs text-muted-foreground shrink-0">Budget {symbol}</span>
                            <input
                              type="number"
                              placeholder="0"
                              value={budgets[name] || ""}
                              onChange={(e) => setBudgets((prev) => ({ ...prev, [name]: e.target.value }))}
                              className="flex-1 rounded-lg border border-border bg-muted/30 px-2 py-1 text-xs focus:outline-none focus:border-primary transition-colors"
                            />
                          </div>
                          {budget > 0 && (
                            <>
                              <div className="h-2 bg-muted rounded-full overflow-hidden mb-1.5">
                                <div className="h-2 rounded-full transition-all" style={{ width: `${pct}%`, backgroundColor: barColor }} />
                              </div>
                              <div className="flex justify-between items-center">
                                <p className="text-[11px] font-semibold" style={{ color: barColor }}>
                                  {over ? `Over by ${fmt(spent - budget)}` : `${Math.round(pct)}% used`}
                                </p>
                                {!over && <p className="text-[11px] text-muted-foreground">{fmt(budget - spent)} left</p>}
                              </div>
                            </>
                          )}
                        </div>
                      );
                    })}

                  {/* Budget summary */}
                  {Object.values(budgets).some((v) => Number(v) > 0) && (
                    <div className="bg-white border border-border rounded-xl p-4 shadow-sm">
                      <p className="text-xs font-bold text-muted-foreground uppercase tracking-widest mb-3">Budget Summary</p>
                      {(() => {
                        const totalBudget = Object.entries(budgets).reduce((s, [, v]) => s + Number(v || 0), 0);
                        const totalSpent = Object.entries(budgets).reduce((s, [name]) => s + (summary.byCategory[name]?.total ?? 0), 0);
                        const remaining = totalBudget - totalSpent;
                        const pct = totalBudget > 0 ? Math.min((totalSpent / totalBudget) * 100, 100) : 0;
                        return (
                          <div className="space-y-2">
                            <div className="flex justify-between text-xs font-semibold">
                              <span>Total Budget</span><span>{fmt(totalBudget)}</span>
                            </div>
                            <div className="flex justify-between text-xs font-semibold">
                              <span>Total Spent</span><span className="text-red-500">{fmt(totalSpent)}</span>
                            </div>
                            <div className="h-2 bg-muted rounded-full overflow-hidden">
                              <div className="h-2 rounded-full transition-all" style={{ width: `${pct}%`, backgroundColor: remaining >= 0 ? "#3b82f6" : "#ef4444" }} />
                            </div>
                            <div className="flex justify-between text-xs font-bold">
                              <span>{remaining >= 0 ? "Remaining" : "Overspent"}</span>
                              <span className={remaining >= 0 ? "text-blue-600" : "text-red-500"}>{fmt(Math.abs(remaining))}</span>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  )}
                </div>
              )}

            </div>
          </div>
        )}

        {/* ── SUBSCRIPTION DASHBOARD ── */}
        {step === "subscription" && (
          <div className="flex flex-col h-full">
            <div className="px-4 pt-4 pb-2 border-b border-border shrink-0">
              <h2 className="text-base font-bold text-foreground">My Subscription</h2>
              <p className="text-xs text-muted-foreground mt-0.5">Manage your Pro plan and license</p>
            </div>
            <div className="flex-1 overflow-y-auto">
              <SubscriptionDashboard
                onStatusChange={(pro) => setIsPro(pro)}
                onUpgrade={() => { setPaymentMode("pay"); setShowPayment(true); }}
              />
            </div>
          </div>
        )}
      </div>

      {/* ── Footer ── */}
      <footer className="px-4 py-2.5 border-t border-border bg-white shrink-0">
        <div className="flex items-center justify-center gap-3 text-[11px] text-muted-foreground">
          <a href="/eula" className="hover:text-foreground hover:underline transition-colors">EULA</a>
          <span className="opacity-40">·</span>
          <a href="/privacy" className="hover:text-foreground hover:underline transition-colors">Privacy Policy</a>
          <span className="opacity-40">·</span>
          <a href="/support" className="hover:text-foreground hover:underline transition-colors">Support</a>
        </div>
        <p className="text-center text-[10px] text-muted-foreground/50 mt-1">{appName} · Excel Add-in</p>
      </footer>
    </div>
  );
}

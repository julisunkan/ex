import { useState, useCallback, useRef, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from "recharts";
import {
  detectColumns,
  readTransactions,
  highlightTransactions,
  createSummarySheet,
  exportToCsv,
  fmt,
  fmtShort,
  type ColumnMap,
} from "./lib/excel";
import { buildSummary, type Summary, type Transaction } from "./lib/categorizer";
import { parsePastedText, writeToExcelSheet } from "./lib/csv-parser";
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

  const [step, setStep] = useState<Step>("idle");
  const [error, setError] = useState<string>("");
  const [summary, setSummary] = useState<Summary | null>(null);
  const [highlighting, setHighlighting] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [activeTab, setActiveTab] = useState<ResultTab>("overview");
  const [csvText, setCsvText] = useState("");
  const [csvError, setCsvError] = useState("");
  const [isPro, setIsPro] = useState(false);
  const [showPayment, setShowPayment] = useState(false);
  const [paymentMode, setPaymentMode] = useState<"pay" | "key">("pay");
  const [pendingAction, setPendingAction] = useState<"highlight" | "export" | "csv" | "budget" | "recurring" | null>(null);
  const [txSearch, setTxSearch] = useState("");
  const [txCategoryFilter, setTxCategoryFilter] = useState("All");
  const [txTypeFilter, setTxTypeFilter] = useState<"all" | "credit" | "debit">("all");
  const [budgets, setBudgets] = useState<Record<string, string>>({});
  const [showDuplicates, setShowDuplicates] = useState(false);
  const [showRecurring, setShowRecurring] = useState(false);

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
        if (!columnMap) throw new Error("Could not find required columns (Date, Description, Amount) in row 1.\nMake sure the active sheet has column headers.");
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
    try {
      await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const columnMap = await detectColumns(sheet);
        if (columnMap) await highlightTransactions(sheet, summary.transactions, columnMap);
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally { setHighlighting(false); }
  }, [summary]);

  const doExport = useCallback(async () => {
    if (!summary) return;
    setExporting(true);
    try {
      await runExcel(async (ctx) => { await createSummarySheet(summary, ctx); });
      setExportDone(true);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally { setExporting(false); }
  }, [summary]);

  const doExportCsv = useCallback(() => {
    if (!summary) return;
    exportToCsv(summary);
  }, [summary]);

  const handleHighlight = useCallback(() => { if (requirePro("highlight")) doHighlight(); }, [requirePro, doHighlight]);
  const handleExport = useCallback(() => { if (requirePro("export")) doExport(); }, [requirePro, doExport]);
  const handleExportCsv = useCallback(() => { if (requirePro("csv")) doExportCsv(); }, [requirePro, doExportCsv]);
  const handleBudgetTab = useCallback(() => { if (requirePro("budget")) setActiveTab("budget"); }, [requirePro]);
  const handleToggleRecurring = useCallback(() => { if (requirePro("recurring")) setShowRecurring((v) => !v); }, [requirePro]);

  const reset = () => {
    setStep("idle"); setSummary(null); setError("");
    setExportDone(false); setActiveTab("overview");
    setCsvText(""); setCsvError("");
    setTxSearch(""); setTxCategoryFilter("All"); setTxTypeFilter("all");
    setBudgets({}); setShowDuplicates(false); setShowRecurring(false);
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
        return matchSearch && matchCat && matchType && matchDupe;
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
          {isPro && (
            <button
              onClick={openSubscription}
              className="flex items-center gap-1 text-xs font-bold bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full border border-amber-200 tracking-wide hover:bg-amber-200 transition-colors"
            >
              <img src={iconPro} alt="Pro" className="w-3.5 h-3.5 object-contain" /> PRO
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
              placeholder={"Date,Description,Amount,Type\n01/06/2026,SALARY JUNE,650000,CR\n02/06/2026,SHOPRITE,-45000,DR\n..."}
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
            <div className="flex gap-2 px-4 pt-4 pb-2 shrink-0">
              <ActionBtn variant="secondary" onClick={handleHighlight} disabled={highlighting} size="sm">
                {highlighting
                  ? <span className="w-3.5 h-3.5 border-2 border-current border-t-transparent rounded-full animate-spin" />
                  : <img src={iconHighlight} alt="Highlight" className="w-3.5 h-3.5 object-contain opacity-90" />
                }
                Highlight
                {!isPro && <span className="text-[10px] opacity-60">🔒</span>}
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
            </div>

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
                  {topCategories.map(([name, info]) => (
                    <div key={name} className="bg-white border border-border rounded-xl p-4 shadow-sm">
                      <div className="flex items-center justify-between mb-2.5">
                        <span className={`text-xs px-2 py-0.5 rounded-md ${info.className}`}>{name}</span>
                        <span className="text-base font-extrabold">{fmt(info.total)}</span>
                      </div>
                      <div className="h-2 bg-muted rounded-full overflow-hidden mb-2">
                        <div className="h-2 rounded-full" style={{ width: `${(info.total / maxCatTotal) * 100}%`, backgroundColor: info.color }} />
                      </div>
                      <div className="flex items-center justify-between">
                        <p className="text-xs text-muted-foreground font-semibold">{info.count} transaction{info.count !== 1 ? "s" : ""}</p>
                        {summary.totalExpenses > 0 && (
                          <p className="text-xs text-muted-foreground font-semibold">{Math.round((info.total / summary.totalExpenses) * 100)}% of expenses</p>
                        )}
                      </div>
                    </div>
                  ))}
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
                      return (
                        <div key={i} className={`flex items-start gap-3 bg-white border rounded-xl p-3 shadow-sm ${isDupe && isPro ? "border-amber-300 bg-amber-50/40" : "border-border"}`}>
                          <div className={`w-2 h-2 rounded-full shrink-0 mt-1.5 ${tx.type === "credit" ? "bg-green-500" : "bg-red-500"}`} />
                          <div className="flex-1 min-w-0">
                            <p className="text-sm font-semibold truncate">{tx.description}</p>
                            <div className="flex flex-wrap items-center gap-1.5 mt-1">
                              <span className="text-xs text-muted-foreground font-medium">{tx.date}</span>
                              <span className={`text-xs px-1.5 py-0.5 rounded-md ${tx.category.className}`}>{tx.category.name}</span>
                              {isDupe && isPro && (
                                <span className="text-[10px] font-bold bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded border border-amber-200">⚠ Possible Duplicate</span>
                              )}
                            </div>
                          </div>
                          <span className={`text-sm font-extrabold shrink-0 ${tx.type === "credit" ? "text-green-600" : "text-red-500"}`}>
                            {tx.type === "credit" ? "+" : "−"}{fmt(tx.amount)}
                          </span>
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
                            <span className="text-xs text-muted-foreground shrink-0">Budget ₦</span>
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
      <footer className="px-4 py-2 border-t border-border bg-white text-xs text-muted-foreground text-center font-semibold shrink-0">
        {appName} · Excel Add-in
      </footer>
    </div>
  );
}

import { useState, useEffect, useCallback } from "react";
import { getLicense, clearLicense, fetchSubscription, activateLicenseKey, type SubscriptionStatus } from "../lib/payment";
import { useAppConfig } from "../context/AppConfigContext";

interface Props {
  onStatusChange: (isPro: boolean) => void;
  onUpgrade: () => void;
}

export default function SubscriptionDashboard({ onStatusChange, onUpgrade }: Props) {
  const config = useAppConfig();
  const plans = config.plans;

  const [status, setStatus] = useState<SubscriptionStatus | null>(null);
  const [loading, setLoading] = useState(true);
  const [renewInput, setRenewInput] = useState("");
  const [renewing, setRenewing] = useState(false);
  const [renewError, setRenewError] = useState("");
  const [renewSuccess, setRenewSuccess] = useState(false);
  const [showRenewForm, setShowRenewForm] = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);

  const refresh = useCallback(async () => {
    setLoading(true);
    const key = getLicense();
    if (!key) {
      setStatus(null);
      setLoading(false);
      onStatusChange(false);
      return;
    }
    const s = await fetchSubscription(key);
    setStatus(s);
    setLoading(false);
    onStatusChange(s.valid);
  }, [onStatusChange]);

  useEffect(() => { refresh(); }, [refresh]);

  const handleRenew = useCallback(async () => {
    setRenewing(true);
    setRenewError("");
    const result = await activateLicenseKey(renewInput);
    setRenewing(false);
    if (result.success) {
      setRenewSuccess(true);
      setRenewInput("");
      setShowRenewForm(false);
      setTimeout(() => {
        setRenewSuccess(false);
        refresh();
      }, 1800);
    } else {
      setRenewError(result.error ?? "Invalid or expired key.");
    }
  }, [renewInput, refresh]);

  const handleClearLicense = () => {
    if (!confirmClear) { setConfirmClear(true); setTimeout(() => setConfirmClear(false), 4000); return; }
    clearLicense();
    setStatus(null);
    setConfirmClear(false);
    onStatusChange(false);
  };

  const planLabel = (planId: string | null) => {
    if (!planId) return "Pro";
    const plan = plans.find((p) => p.id === planId);
    return plan?.label ?? planId;
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center py-8">
        <div className="w-5 h-5 border-2 border-primary border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (!status || !status.valid) {
    return (
      <div className="flex flex-col items-center gap-4 py-6 px-4 text-center">
        <div className="w-12 h-12 rounded-full bg-muted flex items-center justify-center">
          <svg className="w-6 h-6 text-muted-foreground" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>
          </svg>
        </div>
        <div>
          <p className="font-semibold text-foreground text-sm">No Active Subscription</p>
          <p className="text-xs text-muted-foreground mt-1">
            {status?.reason === "expired" ? "Your license has expired. Renew to regain access." : "Unlock premium features to highlight cells and export reports."}
          </p>
        </div>

        {status?.reason === "expired" && (
          <div className="w-full bg-amber-50 border border-amber-200 rounded-xl p-3 text-left">
            <p className="text-[11px] font-semibold text-amber-700 mb-0.5">Expired plan</p>
            <p className="text-xs text-muted-foreground">Was: {planLabel(status.planId)}</p>
            {status.expiresAt && <p className="text-[11px] text-amber-600 mt-0.5">Expired {new Date(status.expiresAt).toLocaleDateString()}</p>}
          </div>
        )}

        <div className="w-full space-y-2">
          <button
            onClick={onUpgrade}
            className="w-full bg-primary text-primary-foreground text-sm font-semibold py-2.5 rounded-xl hover:bg-primary/90 transition-colors"
          >
            Choose a Plan
          </button>

          {!showRenewForm ? (
            <button
              onClick={() => setShowRenewForm(true)}
              className="w-full text-xs text-muted-foreground hover:text-primary transition-colors underline underline-offset-2 py-1"
            >
              Already have a license key?
            </button>
          ) : (
            <div className="space-y-2">
              <input
                type="text"
                value={renewInput}
                onChange={(e) => { setRenewInput(e.target.value.toUpperCase()); setRenewError(""); }}
                placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
                className="w-full rounded-lg border border-border bg-white px-3 py-2.5 text-sm font-mono focus:outline-none focus:ring-1 focus:ring-primary placeholder:text-muted-foreground/40 tracking-wider"
                spellCheck={false}
              />
              {renewError && <p className="text-[11px] text-destructive bg-destructive/10 rounded px-2.5 py-1.5">{renewError}</p>}
              <button
                onClick={handleRenew}
                disabled={renewing || !renewInput.trim()}
                className="w-full bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-xl hover:bg-blue-700 transition-colors disabled:opacity-60"
              >
                {renewing ? "Activating…" : "Activate Key"}
              </button>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ── Active subscription ────────────────────────────────────────────────────
  const daysLeft = status.daysLeft ?? 0;
  const daysTotal = status.daysTotal ?? 1;
  const pct = Math.max(0, Math.min(100, Math.round((daysLeft / daysTotal) * 100)));
  const isExpiringSoon = daysLeft > 0 && daysLeft <= 7;

  return (
    <div className="px-4 py-5 space-y-5">

      {renewSuccess && (
        <div className="flex items-center gap-2.5 bg-green-50 border border-green-200 rounded-xl px-3 py-2.5">
          <svg className="w-4 h-4 text-green-600 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
          <p className="text-sm font-semibold text-green-700">License updated successfully!</p>
        </div>
      )}

      {/* Plan card */}
      <div className="bg-gradient-to-br from-blue-600 to-blue-500 rounded-2xl p-4 text-white shadow-md">
        <div className="flex items-start justify-between mb-3">
          <div>
            <div className="flex items-center gap-1.5 mb-0.5">
              <svg className="w-3.5 h-3.5 text-amber-300" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2l2.4 7.4H22l-6.2 4.5 2.4 7.4L12 17l-6.2 4.3 2.4-7.4L2 9.4h7.6z"/></svg>
              <span className="text-xs font-bold text-amber-300 uppercase tracking-wider">Pro Active</span>
            </div>
            <p className="text-lg font-extrabold text-white">{planLabel(status.planId)}</p>
          </div>
          <span className="text-xs font-semibold bg-white/20 px-2 py-0.5 rounded-full">Active</span>
        </div>

        {/* Days progress */}
        <div className="space-y-1.5">
          <div className="flex items-center justify-between text-xs">
            <span className="text-blue-100">Time remaining</span>
            <span className={`font-bold ${isExpiringSoon ? "text-amber-300" : "text-white"}`}>
              {daysLeft === 0 ? "Expires today" : `${daysLeft} day${daysLeft !== 1 ? "s" : ""} left`}
            </span>
          </div>
          <div className="h-1.5 bg-white/20 rounded-full overflow-hidden">
            <div
              className={`h-full rounded-full transition-all ${pct > 30 ? "bg-white" : pct > 10 ? "bg-amber-300" : "bg-red-400"}`}
              style={{ width: `${pct}%` }}
            />
          </div>
          {status.expiresAt && (
            <p className="text-[10px] text-blue-200">
              Expires {new Date(status.expiresAt).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" })}
            </p>
          )}
        </div>
      </div>

      {/* Expiring soon warning */}
      {isExpiringSoon && (
        <div className="flex items-start gap-2.5 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5">
          <svg className="w-4 h-4 text-amber-500 shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <div>
            <p className="text-xs font-semibold text-amber-700">Expiring soon</p>
            <p className="text-[11px] text-amber-600 mt-0.5">Your subscription expires in {daysLeft} day{daysLeft !== 1 ? "s" : ""}. Renew to keep access.</p>
          </div>
        </div>
      )}

      {/* Active features */}
      <div className="bg-green-50 rounded-xl border border-green-100 p-3 space-y-2">
        <p className="text-[11px] font-semibold text-green-700 uppercase tracking-wider">Included in your plan</p>
        {["Analyze transactions", "Categorize spending", "Highlight cells by category", "Export summary sheet"].map((f) => (
          <div key={f} className="flex items-center gap-2">
            <svg className="w-3.5 h-3.5 text-green-500 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            <span className="text-xs text-green-800">{f}</span>
          </div>
        ))}
      </div>

      {/* Renew / extend */}
      <div className="border border-border rounded-xl overflow-hidden">
        <button
          onClick={() => setShowRenewForm((s) => !s)}
          className="w-full flex items-center justify-between px-4 py-3 text-left hover:bg-muted/40 transition-colors"
        >
          <div className="flex items-center gap-2">
            <svg className="w-4 h-4 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></svg>
            <span className="text-sm font-semibold text-foreground">Enter a new license key</span>
          </div>
          <svg className={`w-4 h-4 text-muted-foreground transition-transform ${showRenewForm ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="6 9 12 15 18 9"/></svg>
        </button>

        {showRenewForm && (
          <div className="px-4 pb-4 pt-1 border-t border-border bg-muted/20 space-y-3">
            <p className="text-xs text-muted-foreground">Enter a new key to extend or switch your plan.</p>
            <input
              type="text"
              value={renewInput}
              onChange={(e) => { setRenewInput(e.target.value.toUpperCase()); setRenewError(""); }}
              placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
              className="w-full rounded-lg border border-border bg-white px-3 py-2 text-sm font-mono focus:outline-none focus:ring-1 focus:ring-primary placeholder:text-muted-foreground/40 tracking-wider"
              spellCheck={false}
            />
            {renewError && <p className="text-[11px] text-destructive bg-destructive/10 rounded px-2.5 py-1.5">{renewError}</p>}
            <div className="flex gap-2">
              <button
                onClick={() => { setShowRenewForm(false); setRenewInput(""); setRenewError(""); }}
                className="flex-1 text-sm font-medium py-2 rounded-lg border border-border hover:bg-muted transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleRenew}
                disabled={renewing || !renewInput.trim()}
                className="flex-1 bg-primary text-primary-foreground text-sm font-semibold py-2 rounded-lg hover:bg-primary/90 transition-colors disabled:opacity-60"
              >
                {renewing ? "Activating…" : "Activate"}
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Upgrade plan */}
      <button
        onClick={onUpgrade}
        className="w-full flex items-center justify-between bg-amber-50 border border-amber-200 rounded-xl px-4 py-3 hover:bg-amber-100 transition-colors group"
      >
        <div className="text-left">
          <p className="text-sm font-semibold text-amber-800">Upgrade plan</p>
          <p className="text-[11px] text-amber-600 mt-0.5">Pay for a longer plan and save more</p>
        </div>
        <svg className="w-4 h-4 text-amber-500 group-hover:translate-x-0.5 transition-transform" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="9 18 15 12 9 6"/></svg>
      </button>

      {/* Remove license */}
      <div className="pt-1 border-t border-border">
        <button
          onClick={handleClearLicense}
          className="w-full text-xs text-muted-foreground hover:text-destructive transition-colors py-2"
        >
          {confirmClear ? "⚠ Tap again to confirm — this will remove your license" : "Remove license from this device"}
        </button>
      </div>
    </div>
  );
}

import { useState, useEffect, useCallback } from "react";
import { fetchPaymentConfig, verifyPayment, setLicense, activateLicenseKey } from "../lib/payment";
import { useAppConfig } from "../context/AppConfigContext";
import iconHighlight from "@assets/icons/icon-highlight.png";
import iconExport from "@assets/icons/icon-export.png";

interface Props {
  onUnlocked: () => void;
  onDismiss: () => void;
  initialMode?: "pay" | "key";
}

type PayMode = "pay" | "key";
type PayStep = "info" | "verifying" | "error" | "success";

export default function PaymentGate({ onUnlocked, onDismiss, initialMode = "pay" }: Props) {
  const config = useAppConfig();
  const plans = config.plans;

  const [mode, setMode] = useState<PayMode>(initialMode);

  // ── Pay mode state ──────────────────────────────────────────────────────────
  const [selectedPlanId, setSelectedPlanId] = useState<string>(() => plans[0]?.id ?? "monthly");
  const [wallet, setWallet] = useState<string>("");
  const [networkLabel, setNetworkLabel] = useState<string>("TRC-20 (Tron)");
  const [txHash, setTxHash] = useState("");
  const [userEmail, setUserEmail] = useState("");
  const [payStep, setPayStep] = useState<PayStep>("info");
  const [payError, setPayError] = useState("");
  const [copied, setCopied] = useState(false);
  const [generatedKey, setGeneratedKey] = useState("");
  const [keyCopied, setKeyCopied] = useState(false);

  // ── Key mode state ──────────────────────────────────────────────────────────
  const [licenseInput, setLicenseInput] = useState("");
  const [keyActivating, setKeyActivating] = useState(false);
  const [keyError, setKeyError] = useState("");
  const [keySuccess, setKeySuccess] = useState(false);

  useEffect(() => {
    fetchPaymentConfig().then((cfg) => {
      if (cfg) {
        setWallet(cfg.address);
        const label = cfg.network === "tron"
          ? "TRC-20 (Tron)"
          : cfg.network === "bsc"
          ? "BEP-20 (BSC)"
          : cfg.network === "eth"
          ? "ERC-20 (ETH)"
          : cfg.network;
        setNetworkLabel(label);
      }
    });
  }, []);

  // Keep selectedPlanId in sync when plans load
  useEffect(() => {
    if (plans.length > 0 && !plans.find((p) => p.id === selectedPlanId)) {
      setSelectedPlanId(plans[0].id);
    }
  }, [plans, selectedPlanId]);

  const selectedPlan = plans.find((p) => p.id === selectedPlanId) ?? plans[0];

  const copyWallet = useCallback(() => {
    navigator.clipboard.writeText(wallet).catch(() => {});
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }, [wallet]);

  const copyGeneratedKey = useCallback(() => {
    navigator.clipboard.writeText(generatedKey).catch(() => {});
    setKeyCopied(true);
    setTimeout(() => setKeyCopied(false), 2000);
  }, [generatedKey]);

  const handleVerifyPayment = useCallback(async () => {
    const hash = txHash.trim();
    if (!hash) { setPayError("Please paste your transaction hash."); return; }
    setPayStep("verifying");
    setPayError("");
    const email = userEmail.trim().includes("@") ? userEmail.trim() : undefined;
    const result = await verifyPayment(hash, selectedPlan?.id ?? "monthly", email);
    if (result.success && result.licenseKey) {
      setGeneratedKey(result.licenseKey);
      setLicense(result.licenseKey);
      setPayStep("success");
      setTimeout(onUnlocked, 4000);
    } else {
      setPayError(result.error ?? "Verification failed.");
      setPayStep("error");
    }
  }, [txHash, selectedPlan, onUnlocked]);

  const handleActivateKey = useCallback(async () => {
    setKeyActivating(true);
    setKeyError("");
    const result = await activateLicenseKey(licenseInput);
    setKeyActivating(false);
    if (result.success) {
      setKeySuccess(true);
      setTimeout(onUnlocked, 1500);
    } else {
      setKeyError(result.error ?? "Activation failed.");
    }
  }, [licenseInput, onUnlocked]);

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/30 backdrop-blur-sm">
      <div className="w-full bg-white rounded-t-2xl shadow-2xl overflow-hidden animate-in slide-in-from-bottom-4">

        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-blue-500 px-4 py-3 flex items-center justify-between">
          <div>
            <p className="text-white font-semibold text-sm">Unlock Premium Features</p>
            <p className="text-blue-100 text-[10px]">Choose a plan · Cancel anytime</p>
          </div>
          <button onClick={onDismiss} className="text-white/70 hover:text-white transition-colors p-1">
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
              <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* Mode tabs */}
        <div className="flex border-b border-border bg-gray-50">
          <button
            onClick={() => setMode("pay")}
            className={`flex-1 py-2.5 text-xs font-semibold transition-colors ${mode === "pay" ? "text-blue-600 border-b-2 border-blue-600 bg-white" : "text-muted-foreground hover:text-foreground"}`}
          >
            Pay with USDT
          </button>
          <button
            onClick={() => setMode("key")}
            className={`flex-1 py-2.5 text-xs font-semibold transition-colors ${mode === "key" ? "text-blue-600 border-b-2 border-blue-600 bg-white" : "text-muted-foreground hover:text-foreground"}`}
          >
            Enter License Key
          </button>
        </div>

        <div className="p-4 space-y-4 max-h-[75vh] overflow-y-auto">

          {/* ── PAY MODE ── */}
          {mode === "pay" && (
            <>
              {/* Premium feature cards */}
              <div className="grid grid-cols-2 gap-2">
                {[
                  { icon: iconHighlight, label: "Highlight Cells", desc: "Color-code by category" },
                  { icon: iconExport, label: "Export Sheet", desc: "Full summary report" },
                ].map((f) => (
                  <div key={f.label} className="bg-blue-50 rounded-lg p-2.5 border border-blue-100">
                    <img src={f.icon} alt={f.label} className="w-8 h-8 object-contain mb-1" />
                    <p className="text-xs font-semibold text-blue-800">{f.label}</p>
                    <p className="text-[10px] text-blue-600">{f.desc}</p>
                  </div>
                ))}
              </div>

              {payStep === "success" ? (
                <div className="flex flex-col items-center gap-3 py-3">
                  <div className="w-12 h-12 rounded-full bg-green-100 flex items-center justify-center">
                    <svg className="w-6 h-6 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="20 6 9 17 4 12" />
                    </svg>
                  </div>
                  <p className="text-sm font-bold text-green-700">Payment Verified!</p>

                  {generatedKey && (
                    <div className="w-full bg-amber-50 border border-amber-200 rounded-xl p-3">
                      <p className="text-[10px] font-bold text-amber-700 uppercase tracking-wider mb-1.5 flex items-center gap-1">
                        <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></svg>
                        Your License Key — Save this!
                      </p>
                      <div className="flex items-center gap-2 bg-white border border-amber-200 rounded-lg px-2.5 py-2">
                        <code className="flex-1 text-xs font-mono font-bold text-foreground tracking-wider select-all">{generatedKey}</code>
                        <button onClick={copyGeneratedKey}
                          className="shrink-0 text-[10px] font-semibold px-2 py-1 rounded bg-amber-500 text-white hover:bg-amber-600 transition-colors">
                          {keyCopied ? "✓ Copied" : "Copy"}
                        </button>
                      </div>
                      <p className="text-[10px] text-amber-600 mt-1.5">Keep this key safe — you'll need it to reactivate on other devices.</p>
                    </div>
                  )}

                  <p className="text-xs text-muted-foreground">Unlocking premium features…</p>
                </div>
              ) : (
                <>
                  {/* Plan selector */}
                  {plans.length > 0 && (
                    <div>
                      <p className="text-[11px] font-semibold text-foreground mb-2">Choose a plan</p>
                      <div className="grid grid-cols-2 gap-2">
                        {plans.map((plan) => {
                          const isSelected = selectedPlanId === plan.id;
                          return (
                            <button
                              key={plan.id}
                              onClick={() => setSelectedPlanId(plan.id)}
                              className={`relative rounded-xl border-2 p-2.5 text-left transition-all ${
                                isSelected
                                  ? "border-blue-500 bg-blue-50"
                                  : "border-border bg-white hover:border-blue-200"
                              }`}
                            >
                              {plan.id === "annual" && (
                                <span className="absolute -top-2 right-2 text-[9px] font-bold bg-green-500 text-white px-1.5 py-0.5 rounded-full">BEST VALUE</span>
                              )}
                              {plan.id === "quarterly" && (
                                <span className="absolute -top-2 right-2 text-[9px] font-bold bg-blue-500 text-white px-1.5 py-0.5 rounded-full">POPULAR</span>
                              )}
                              <p className={`text-xs font-bold ${isSelected ? "text-blue-700" : "text-foreground"}`}>{plan.label}</p>
                              <p className={`text-base font-extrabold mt-0.5 ${isSelected ? "text-blue-600" : "text-foreground"}`}>${plan.price} <span className="text-[10px] font-medium text-muted-foreground">USDT</span></p>
                              <p className="text-[10px] text-muted-foreground">{plan.days} days access</p>
                              {isSelected && (
                                <div className="absolute top-2 right-2 w-3.5 h-3.5 rounded-full bg-blue-500 flex items-center justify-center">
                                  <svg className="w-2 h-2 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>
                                </div>
                              )}
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* Step 1 — Send USDT */}
                  <div>
                    <p className="text-[11px] font-semibold text-foreground mb-1.5 flex items-center gap-1.5">
                      <span className="w-4 h-4 rounded-full bg-primary text-white text-[9px] flex items-center justify-center font-bold shrink-0">1</span>
                      Send exactly ${selectedPlan?.price ?? "—"} USDT to this wallet
                    </p>
                    <div className="flex items-center gap-2 bg-muted rounded-lg p-2 border border-border">
                      <p className="flex-1 text-[10px] font-mono text-foreground break-all leading-relaxed">
                        {wallet || "Loading…"}
                      </p>
                      {wallet && (
                        <button onClick={copyWallet}
                          className="shrink-0 text-[10px] font-medium px-2 py-1 rounded bg-primary/10 text-primary hover:bg-primary/20 transition-colors">
                          {copied ? "✓ Copied" : "Copy"}
                        </button>
                      )}
                    </div>
                    <p className="text-[10px] text-muted-foreground mt-1">
                      Network: <span className="font-medium text-foreground">{networkLabel}</span> · Use your crypto wallet or exchange
                    </p>
                  </div>

                  {/* Step 2 — Paste tx hash */}
                  <div>
                    <p className="text-[11px] font-semibold text-foreground mb-1.5 flex items-center gap-1.5">
                      <span className="w-4 h-4 rounded-full bg-primary text-white text-[9px] flex items-center justify-center font-bold shrink-0">2</span>
                      Paste your transaction hash and verify
                    </p>
                    <input
                      type="text"
                      value={txHash}
                      onChange={(e) => { setTxHash(e.target.value); setPayError(""); setPayStep("info"); }}
                      placeholder="e.g. 3f8a2b1c4d…"
                      className="w-full rounded-md border border-border bg-white px-2.5 py-2 text-xs font-mono focus:outline-none focus:ring-1 focus:ring-primary placeholder:text-muted-foreground/50"
                      data-testid="input-tx-hash"
                    />
                    {(payStep === "error" || payError) && (
                      <p className="text-[10px] text-destructive mt-1.5 bg-destructive/10 rounded px-2 py-1">{payError}</p>
                    )}
                  </div>

                  {/* Step 3 — Optional email for renewal reminders */}
                  <div>
                    <p className="text-[11px] font-semibold text-foreground mb-1 flex items-center gap-1.5">
                      <span className="w-4 h-4 rounded-full bg-muted text-muted-foreground text-[9px] flex items-center justify-center font-bold shrink-0 border border-border">3</span>
                      Your email <span className="text-muted-foreground font-normal">(optional — for renewal reminders)</span>
                    </p>
                    <input
                      type="email"
                      value={userEmail}
                      onChange={(e) => setUserEmail(e.target.value)}
                      placeholder="you@example.com"
                      className="w-full rounded-md border border-border bg-white px-2.5 py-2 text-xs focus:outline-none focus:ring-1 focus:ring-primary placeholder:text-muted-foreground/50"
                      data-testid="input-user-email"
                    />
                    <p className="text-[10px] text-muted-foreground mt-0.5">We'll remind you before your subscription expires. No spam.</p>
                  </div>

                  <button
                    onClick={payStep === "verifying" ? undefined : handleVerifyPayment}
                    disabled={payStep === "verifying" || !txHash.trim()}
                    className="w-full flex items-center justify-center gap-2 bg-primary text-primary-foreground text-sm font-semibold py-2.5 rounded-lg hover:bg-primary/90 transition-colors disabled:opacity-60"
                  >
                    {payStep === "verifying" ? (
                      <>
                        <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                        Verifying on blockchain…
                      </>
                    ) : `Verify Payment & Unlock (${selectedPlan ? `$${selectedPlan.price} USDT` : "—"})`}
                  </button>

                  <p className="text-[10px] text-center text-muted-foreground">
                    Already have a key?{" "}
                    <button onClick={() => setMode("key")} className="text-primary underline">
                      Enter license key instead
                    </button>
                  </p>
                </>
              )}
            </>
          )}

          {/* ── KEY MODE ── */}
          {mode === "key" && (
            <>
              {keySuccess ? (
                <div className="flex flex-col items-center gap-3 py-6">
                  <div className="w-12 h-12 rounded-full bg-green-100 flex items-center justify-center">
                    <svg className="w-6 h-6 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="20 6 9 17 4 12" />
                    </svg>
                  </div>
                  <p className="text-sm font-bold text-green-700">License Activated!</p>
                  <p className="text-xs text-muted-foreground">Unlocking premium features…</p>
                </div>
              ) : (
                <>
                  <div className="flex flex-col items-center gap-2 py-2 text-center">
                    <div className="w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center">
                      <svg className="w-5 h-5 text-blue-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>
                      </svg>
                    </div>
                    <p className="text-sm font-semibold text-foreground">Enter Your License Key</p>
                    <p className="text-xs text-muted-foreground max-w-[260px]">
                      Enter the license key you received after payment to activate Pro features.
                    </p>
                  </div>

                  <div className="space-y-2">
                    <input
                      type="text"
                      value={licenseInput}
                      onChange={(e) => { setLicenseInput(e.target.value.toUpperCase()); setKeyError(""); }}
                      placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
                      className="w-full rounded-md border border-border bg-white px-3 py-2.5 text-sm font-mono focus:outline-none focus:ring-1 focus:ring-primary placeholder:text-muted-foreground/40 tracking-wider"
                      spellCheck={false}
                      autoCapitalize="characters"
                    />
                    {keyError && (
                      <p className="text-[11px] text-destructive bg-destructive/10 rounded px-2.5 py-1.5">{keyError}</p>
                    )}
                  </div>

                  <button
                    onClick={handleActivateKey}
                    disabled={keyActivating || !licenseInput.trim()}
                    className="w-full flex items-center justify-center gap-2 bg-primary text-primary-foreground text-sm font-semibold py-2.5 rounded-lg hover:bg-primary/90 transition-colors disabled:opacity-60"
                  >
                    {keyActivating ? (
                      <>
                        <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                        Validating…
                      </>
                    ) : "Activate License Key"}
                  </button>

                  <p className="text-[10px] text-center text-muted-foreground">
                    Don't have a key yet?{" "}
                    <button onClick={() => setMode("pay")} className="text-primary underline">
                      Pay with USDT
                    </button>
                  </p>
                </>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

import { useState, useEffect, useCallback } from "react";
import { fetchAdminWallet, verifyPayment, setLicense } from "../lib/payment";

interface Props {
  onUnlocked: () => void;
  onDismiss: () => void;
}

type PayStep = "info" | "waiting" | "verifying" | "error" | "success";

export default function PaymentGate({ onUnlocked, onDismiss }: Props) {
  const [wallet, setWallet] = useState<string>("");
  const [price, setPrice] = useState<number>(5);
  const [network, setNetwork] = useState<string>("TRC-20 (Tron)");
  const [txHash, setTxHash] = useState("");
  const [payStep, setPayStep] = useState<PayStep>("info");
  const [errorMsg, setErrorMsg] = useState("");
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    fetchAdminWallet().then((cfg) => {
      if (cfg) {
        setWallet(cfg.address);
        setPrice(cfg.price);
        setNetwork(cfg.network);
      }
    });
  }, []);

  const copyWallet = useCallback(() => {
    navigator.clipboard.writeText(wallet).catch(() => {});
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }, [wallet]);

  const handleVerify = useCallback(async () => {
    const hash = txHash.trim();
    if (!hash) { setErrorMsg("Please paste your transaction hash."); return; }
    setPayStep("verifying");
    setErrorMsg("");
    const result = await verifyPayment(hash);
    if (result.success && result.licenseKey) {
      setLicense(result.licenseKey);
      setPayStep("success");
      setTimeout(onUnlocked, 1800);
    } else {
      setErrorMsg(result.error ?? "Verification failed.");
      setPayStep("error");
    }
  }, [txHash, onUnlocked]);

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/30 backdrop-blur-sm">
      <div className="w-full bg-white rounded-t-2xl shadow-2xl overflow-hidden animate-in slide-in-from-bottom-4">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-blue-500 px-4 py-3 flex items-center justify-between">
          <div>
            <p className="text-white font-semibold text-sm">Unlock Premium Features</p>
            <p className="text-blue-100 text-[10px]">One-time payment · No subscription</p>
          </div>
          <button onClick={onDismiss} className="text-white/70 hover:text-white transition-colors">
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
              <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        <div className="p-4 space-y-4 max-h-[80vh] overflow-y-auto">

          {/* Premium feature list */}
          <div className="grid grid-cols-2 gap-2">
            {[
              { icon: "✨", label: "Highlight Cells", desc: "Color-code by category" },
              { icon: "📊", label: "Export Sheet", desc: "Full summary report" },
            ].map((f) => (
              <div key={f.label} className="bg-blue-50 rounded-lg p-2.5 border border-blue-100">
                <p className="text-base mb-0.5">{f.icon}</p>
                <p className="text-xs font-semibold text-blue-800">{f.label}</p>
                <p className="text-[10px] text-blue-600">{f.desc}</p>
              </div>
            ))}
          </div>

          {/* Price badge */}
          <div className="flex items-center justify-center gap-2 bg-green-50 border border-green-200 rounded-lg py-2.5">
            <span className="text-lg font-bold text-green-700">${price} USDT</span>
            <span className="text-[10px] text-green-600 bg-green-100 px-1.5 py-0.5 rounded font-medium">{network}</span>
          </div>

          {payStep === "success" ? (
            <div className="flex flex-col items-center gap-2 py-4">
              <div className="w-12 h-12 rounded-full bg-green-100 flex items-center justify-center">
                <svg className="w-6 h-6 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="20 6 9 17 4 12" />
                </svg>
              </div>
              <p className="text-sm font-semibold text-green-700">Payment Verified!</p>
              <p className="text-xs text-muted-foreground">Unlocking premium features…</p>
            </div>
          ) : (
            <>
              {/* Step 1 */}
              <div>
                <p className="text-[11px] font-semibold text-foreground mb-1.5 flex items-center gap-1.5">
                  <span className="w-4 h-4 rounded-full bg-primary text-white text-[9px] flex items-center justify-center font-bold shrink-0">1</span>
                  Send exactly ${price} USDT to this wallet
                </p>
                <div className="flex items-center gap-2 bg-muted rounded-lg p-2 border border-border">
                  <p className="flex-1 text-[10px] font-mono text-foreground break-all leading-relaxed">
                    {wallet || "Loading…"}
                  </p>
                  {wallet && (
                    <button
                      onClick={copyWallet}
                      className="shrink-0 text-[10px] font-medium px-2 py-1 rounded bg-primary/10 text-primary hover:bg-primary/20 transition-colors"
                    >
                      {copied ? "✓ Copied" : "Copy"}
                    </button>
                  )}
                </div>
                <p className="text-[10px] text-muted-foreground mt-1">
                  Network: <span className="font-medium text-foreground">{network}</span> · Use your crypto wallet or exchange
                </p>
              </div>

              {/* Step 2 */}
              <div>
                <p className="text-[11px] font-semibold text-foreground mb-1.5 flex items-center gap-1.5">
                  <span className="w-4 h-4 rounded-full bg-primary text-white text-[9px] flex items-center justify-center font-bold shrink-0">2</span>
                  Paste your transaction hash and verify
                </p>
                <input
                  type="text"
                  value={txHash}
                  onChange={(e) => { setTxHash(e.target.value); setErrorMsg(""); }}
                  placeholder="e.g. 3f8a2b1c4d…"
                  className="w-full rounded-md border border-border bg-white px-2.5 py-2 text-xs font-mono focus:outline-none focus:ring-1 focus:ring-primary placeholder:text-muted-foreground/50"
                />
                {(payStep === "error" || errorMsg) && (
                  <p className="text-[10px] text-destructive mt-1.5 bg-destructive/10 rounded px-2 py-1">{errorMsg}</p>
                )}
              </div>

              <button
                onClick={payStep === "verifying" ? undefined : handleVerify}
                disabled={payStep === "verifying" || !txHash.trim()}
                className="w-full flex items-center justify-center gap-2 bg-primary text-primary-foreground text-sm font-semibold py-2.5 rounded-lg hover:bg-primary/90 transition-colors disabled:opacity-60"
              >
                {payStep === "verifying" ? (
                  <>
                    <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Verifying on Tron…
                  </>
                ) : (
                  "Verify Payment & Unlock"
                )}
              </button>

              <p className="text-[10px] text-center text-muted-foreground">
                Already paid?{" "}
                <button
                  onClick={() => setPayStep("waiting")}
                  className="text-primary underline"
                >
                  Enter your transaction hash above
                </button>
              </p>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

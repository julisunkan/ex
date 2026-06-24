import { useState, useEffect, useRef } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import {
  KeyRound, LogOut, RefreshCw, ShieldCheck, Copy, Check,
  CreditCard, Palette, Database, Download, Upload, Save, Eye, EyeOff,
  Rocket, ExternalLink, AlertCircle, CheckCircle2, Terminal, Bell, TrendingUp,
  Users, DollarSign, Activity, BarChart3, TicketCheck
} from "lucide-react";

const API_BASE = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");

// ── Types ─────────────────────────────────────────────────────────────────────
interface License {
  licenseKey: string; txHash: string; issuedAt: string;
  note?: string; expiresAt?: string; planId?: string;
  email?: string | null; reminderSent?: boolean;
}
interface Plan {
  id: string;
  label: string;
  price: number;
  days: number;
}

interface EmailNotifyCfg {
  enabled: boolean; to: string; smtpHost: string; smtpPort: number;
  smtpUser: string; smtpPass: string; from: string;
}

interface Settings {
  appearance: { name: string; tagline: string; primaryColor: string; accentColor: string; radius: string; };
  payment: { walletAddress: string; network: string; };
  plans: Plan[];
  notifications: { webhookUrl: string; remindersEnabled: boolean; reminderDays: number; email: EmailNotifyCfg; };
  features: { proEnabled: boolean; };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtDate(iso: string) {
  return new Date(iso).toLocaleString(undefined, { year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
}

function CopyBtn({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);
  return (
    <button onClick={() => { navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
      className="ml-1 text-muted-foreground hover:text-foreground transition-colors">
      {copied ? <Check className="w-3.5 h-3.5 text-green-500" /> : <Copy className="w-3.5 h-3.5" />}
    </button>
  );
}

// ── Login screen ──────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }: { onLogin: (pw: string) => void }) {
  const [pw, setPw] = useState("");
  const [show, setShow] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true); setError("");
    try {
      const res = await fetch(`${API_BASE}/api/admin/settings`, { headers: { "x-admin-password": pw } });
      setLoading(false);
      if (res.status === 401) { setError("Wrong password."); return; }
      if (res.status === 503) { setError("Backend not configured — set ADMIN_PASSWORD on the server."); return; }
      if (!res.ok) { setError(`Server error (${res.status}) — check that VITE_API_URL points to your backend.`); return; }
      sessionStorage.setItem("admin_pw", pw);
      onLogin(pw);
    } catch {
      setLoading(false);
      const hint = API_BASE ? `Cannot reach ${API_BASE}` : "VITE_API_URL is not set — rebuild the frontend with that env var.";
      setError(`Network error — ${hint}`);
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 p-4">
      <Card className="w-full max-w-sm">
        <CardHeader className="text-center">
          <div className="mx-auto mb-2 w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center">
            <ShieldCheck className="w-5 h-5 text-blue-600" />
          </div>
          <CardTitle className="text-lg">Admin Login</CardTitle>
          <p className="text-sm text-muted-foreground">Bank Statement Analyzer</p>
        </CardHeader>
        <CardContent>
          <form onSubmit={handleSubmit} className="space-y-3">
            <div className="relative">
              <Input type={show ? "text" : "password"} placeholder="Admin password" value={pw}
                onChange={e => setPw(e.target.value)} autoFocus className="pr-10" />
              <button type="button" onClick={() => setShow(s => !s)}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground">
                {show ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
            {error && <p className="text-sm text-red-500">{error}</p>}
            <Button type="submit" className="w-full" disabled={loading || !pw}>{loading ? "Checking…" : "Login"}</Button>
          </form>
        </CardContent>
      </Card>
    </div>
  );
}

// ── Licenses tab ──────────────────────────────────────────────────────────────
function LicensesTab({ pw }: { pw: string }) {
  const [data, setData] = useState<{ total: number; licenses: License[] } | null>(null);
  const [loading, setLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [genNote, setGenNote] = useState("");
  const [genExpiry, setGenExpiry] = useState("0");
  const [newKey, setNewKey] = useState<string | null>(null);
  const [newKeyCopied, setNewKeyCopied] = useState(false);
  const [revoking, setRevoking] = useState<string | null>(null);
  // Bulk generate state
  const [bulkCount, setBulkCount] = useState(5);
  const [bulkNote, setBulkNote] = useState("");
  const [bulkExpiry, setBulkExpiry] = useState("0");
  const [bulkGenerating, setBulkGenerating] = useState(false);
  const [bulkKeys, setBulkKeys] = useState<string[]>([]);
  const [bulkExpiresAt, setBulkExpiresAt] = useState<string | null>(null);

  async function load() {
    setLoading(true);
    const res = await fetch(`${API_BASE}/api/admin/licenses`, { headers: { "x-admin-password": pw } });
    if (res.ok) setData(await res.json());
    setLoading(false);
  }

  useEffect(() => { load(); }, []);

  async function generateKey() {
    setGenerating(true);
    setNewKey(null);
    const res = await fetch(`${API_BASE}/api/admin/licenses/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify({
        note: genNote || "Admin generated",
        expiryDays: parseInt(genExpiry) || 0,
      }),
    });
    setGenerating(false);
    if (res.ok) {
      const { licenseKey } = await res.json();
      setNewKey(licenseKey);
      setGenNote("");
      load();
    }
  }

  function copyNewKey() {
    if (!newKey) return;
    navigator.clipboard.writeText(newKey);
    setNewKeyCopied(true);
    setTimeout(() => setNewKeyCopied(false), 2000);
  }

  async function bulkGenerate() {
    setBulkGenerating(true);
    setBulkKeys([]);
    setBulkExpiresAt(null);
    const res = await fetch(`${API_BASE}/api/admin/licenses/bulk-generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify({
        count: bulkCount,
        note: bulkNote || "Bulk generated",
        expiryDays: parseInt(bulkExpiry) || 0,
      }),
    });
    setBulkGenerating(false);
    if (res.ok) {
      const { keys, expiresAt } = await res.json();
      setBulkKeys(keys);
      setBulkExpiresAt(expiresAt ?? null);
      load();
    }
  }

  function downloadCsv() {
    if (!bulkKeys.length) return;
    const note = bulkNote || "Bulk generated";
    const generatedAt = new Date().toISOString();
    const header = "License Key,Note,Generated At,Expires At";
    const rows = bulkKeys.map(k => `${k},"${note}",${generatedAt},${bulkExpiresAt ?? "Never"}`);
    const csv = [header, ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `bsa-licenses-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function revokeKey(key: string) {
    if (!confirm(`Revoke license key ${key}? This cannot be undone.`)) return;
    setRevoking(key);
    await fetch(`${API_BASE}/api/admin/licenses/${encodeURIComponent(key)}`, {
      method: "DELETE", headers: { "x-admin-password": pw },
    });
    setRevoking(null);
    load();
  }

  function exportSubscribersCsv() {
    if (!data || data.licenses.length === 0) return;
    const header = ["License Key", "Plan", "Subscriber Email", "Issued At", "Expires At", "Reminder Sent", "Source / TX Hash", "Note"];
    const rows = data.licenses.map(l => [
      l.licenseKey,
      l.planId ?? "",
      l.email ?? "",
      l.issuedAt ?? "",
      l.expiresAt ?? "",
      l.reminderSent ? "Yes" : "No",
      l.txHash === "MANUAL" ? "Manual" : (l.txHash ?? ""),
      l.note ?? "",
    ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(","));
    const csv = [header.join(","), ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `bsa-subscribers-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="space-y-4">
      {/* Generate key card */}
      <Card className="border-blue-100 bg-blue-50/40">
        <CardHeader className="pb-3">
          <CardTitle className="text-base flex items-center gap-2">
            <KeyRound className="w-4 h-4 text-blue-600" /> Generate License Key
          </CardTitle>
          <p className="text-sm text-muted-foreground">Create a key manually — useful for testing or granting access without payment.</p>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex gap-2 flex-wrap">
            <Input
              value={genNote}
              onChange={e => setGenNote(e.target.value)}
              placeholder="Note (optional, e.g. 'Test key')"
              className="flex-1 min-w-[140px]"
              onKeyDown={e => e.key === "Enter" && generateKey()}
            />
            <select
              value={genExpiry}
              onChange={e => setGenExpiry(e.target.value)}
              className="rounded-md border border-input bg-background px-2.5 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring shrink-0"
            >
              <option value="0">No expiry</option>
              <option value="7">7 days</option>
              <option value="30">30 days</option>
              <option value="90">90 days</option>
              <option value="365">1 year</option>
            </select>
            <Button onClick={generateKey} disabled={generating} className="gap-2 shrink-0">
              {generating
                ? <><RefreshCw className="w-4 h-4 animate-spin" /> Generating…</>
                : <><KeyRound className="w-4 h-4" /> Generate</>}
            </Button>
          </div>

          {newKey && (
            <div className="flex items-center gap-2 bg-white border border-blue-200 rounded-lg px-3 py-2.5">
              <code className="flex-1 text-sm font-mono font-bold text-foreground tracking-wider select-all">{newKey}</code>
              {genExpiry !== "0" && (
                <span className="text-[10px] text-amber-600 bg-amber-50 border border-amber-200 rounded px-1.5 py-0.5 font-medium shrink-0">
                  {genExpiry}d
                </span>
              )}
              <button onClick={copyNewKey}
                className="shrink-0 flex items-center gap-1 text-xs font-medium px-2.5 py-1.5 rounded bg-blue-600 text-white hover:bg-blue-700 transition-colors">
                {newKeyCopied ? <><Check className="w-3.5 h-3.5" /> Copied!</> : <><Copy className="w-3.5 h-3.5" /> Copy</>}
              </button>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Bulk generate card */}
      <Card className="border-purple-100 bg-purple-50/40">
        <CardHeader className="pb-3">
          <CardTitle className="text-base flex items-center gap-2">
            <Download className="w-4 h-4 text-purple-600" /> Bulk Generate Keys
          </CardTitle>
          <p className="text-sm text-muted-foreground">Generate up to 100 keys at once and download as a CSV file.</p>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex gap-2 flex-wrap">
            <div className="flex items-center gap-2 shrink-0">
              <label className="text-sm font-medium">Count</label>
              <Input
                type="number" min={1} max={100}
                value={bulkCount}
                onChange={e => setBulkCount(Math.max(1, Math.min(100, parseInt(e.target.value) || 1)))}
                className="w-20"
              />
            </div>
            <Input
              value={bulkNote}
              onChange={e => setBulkNote(e.target.value)}
              placeholder="Note (optional, e.g. 'Beta batch')"
              className="flex-1 min-w-[140px]"
            />
            <select
              value={bulkExpiry}
              onChange={e => setBulkExpiry(e.target.value)}
              className="rounded-md border border-input bg-background px-2.5 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring shrink-0"
            >
              <option value="0">No expiry</option>
              <option value="7">7 days</option>
              <option value="30">30 days</option>
              <option value="90">90 days</option>
              <option value="365">1 year</option>
            </select>
            <Button onClick={bulkGenerate} disabled={bulkGenerating} variant="outline" className="gap-2 border-purple-300 text-purple-700 hover:bg-purple-50 shrink-0">
              {bulkGenerating
                ? <><RefreshCw className="w-4 h-4 animate-spin" /> Generating…</>
                : <><KeyRound className="w-4 h-4" /> Generate {bulkCount}</>}
            </Button>
          </div>

          {bulkKeys.length > 0 && (
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-green-700">✓ {bulkKeys.length} keys generated</p>
                  {bulkExpiresAt && (
                    <p className="text-xs text-amber-600 mt-0.5">Expires {fmtDate(bulkExpiresAt)}</p>
                  )}
                </div>
                <Button onClick={downloadCsv} size="sm" className="gap-2 bg-purple-600 hover:bg-purple-700 text-white">
                  <Download className="w-3.5 h-3.5" /> Download CSV
                </Button>
              </div>
              <div className="max-h-36 overflow-y-auto rounded-lg border border-purple-200 bg-white">
                {bulkKeys.map((k, i) => (
                  <div key={k} className="flex items-center justify-between px-3 py-1.5 border-b border-purple-50 last:border-0 hover:bg-purple-50/50">
                    <span className="text-xs text-muted-foreground w-5 shrink-0">{i + 1}</span>
                    <code className="flex-1 text-xs font-mono font-medium text-foreground tracking-wider select-all">{k}</code>
                    <CopyBtn text={k} />
                  </div>
                ))}
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* List */}
      {/* Summary stats */}
      {data && data.licenses.length > 0 && (() => {
        const now = Date.now();
        const withEmail    = data.licenses.filter(l => l.email).length;
        const reminded     = data.licenses.filter(l => l.reminderSent).length;
        const active       = data.licenses.filter(l => !l.expiresAt || new Date(l.expiresAt).getTime() > now).length;
        const expiringSoon = data.licenses.filter(l => {
          if (!l.expiresAt) return false;
          const ms = new Date(l.expiresAt).getTime() - now;
          return ms > 0 && ms < 7 * 24 * 60 * 60 * 1000;
        }).length;
        return (
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            {[
              { label: "Total",          value: data.total,  color: "bg-gray-50 border-gray-200",     text: "text-gray-700"   },
              { label: "Active",         value: active,      color: "bg-green-50 border-green-200",   text: "text-green-700"  },
              { label: "Expiring soon",  value: expiringSoon,color: "bg-amber-50 border-amber-200",   text: "text-amber-700"  },
              { label: "With email",     value: withEmail,   color: "bg-blue-50 border-blue-200",     text: "text-blue-700"   },
            ].map(s => (
              <div key={s.label} className={`rounded-xl border ${s.color} px-4 py-3 text-center`}>
                <p className={`text-2xl font-bold ${s.text}`}>{s.value}</p>
                <p className="text-xs text-muted-foreground mt-0.5">{s.label}</p>
              </div>
            ))}
          </div>
        );
      })()}

      <div className="flex items-center justify-between">
        <Badge variant="secondary">{data?.total ?? 0} licenses issued</Badge>
        <div className="flex items-center gap-2">
          {data && data.licenses.length > 0 && (
            <Button variant="outline" size="sm" onClick={exportSubscribersCsv} className="gap-1.5 text-green-700 border-green-300 hover:bg-green-50" data-testid="button-export-subscribers-csv">
              <Download className="w-3.5 h-3.5" /> Export CSV
            </Button>
          )}
          <Button variant="outline" size="sm" onClick={load} disabled={loading}>
            <RefreshCw className={`w-4 h-4 mr-1 ${loading ? "animate-spin" : ""}`} /> Refresh
          </Button>
        </div>
      </div>

      <Card>
        <CardContent className="p-0">
          {!data || data.licenses.length === 0 ? (
            <div className="text-center py-12 text-muted-foreground text-sm">No licenses issued yet.</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b bg-gray-50/60">
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground w-8">#</th>
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground">License Key</th>
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground">Source / Note</th>
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground">Subscriber Email</th>
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground">Expiry</th>
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground">Reminder</th>
                    <th className="text-left px-4 py-3 font-medium text-muted-foreground">Issued At</th>
                    <th className="px-4 py-3 w-8" />
                  </tr>
                </thead>
                <tbody>
                  {data.licenses.map((l, i) => {
                    const now = Date.now();
                    const exp = l.expiresAt ? new Date(l.expiresAt).getTime() : null;
                    const expired = exp !== null && exp < now;
                    const expiringSoon = exp !== null && !expired && exp - now < 7 * 24 * 60 * 60 * 1000;
                    return (
                    <tr key={l.licenseKey} data-testid={`row-license-${l.licenseKey}`} className={`border-b last:border-0 hover:bg-gray-50 transition-colors ${expired ? "opacity-60" : ""}`}>
                      <td className="px-4 py-3 text-muted-foreground">{i + 1}</td>
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-1">
                          <code className="font-mono text-xs bg-gray-100 px-1.5 py-0.5 rounded" data-testid={`text-license-key-${i}`}>{l.licenseKey}</code>
                          <CopyBtn text={l.licenseKey} />
                        </div>
                        {l.planId && (
                          <span className="text-[10px] text-muted-foreground mt-0.5 block">{l.planId}</span>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        {l.txHash === "MANUAL" ? (
                          <div className="flex items-center gap-1.5">
                            <Badge variant="secondary" className="text-[10px] px-1.5 py-0 h-4 bg-blue-100 text-blue-700 border-0">Manual</Badge>
                            {l.note && <span className="text-xs text-muted-foreground truncate max-w-[120px]">{l.note}</span>}
                          </div>
                        ) : (
                          <div className="flex items-center gap-1 max-w-xs">
                            <span className="font-mono text-xs text-muted-foreground truncate">{l.txHash}</span>
                            <CopyBtn text={l.txHash} />
                          </div>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        {l.email ? (
                          <div className="flex items-center gap-1" data-testid={`text-subscriber-email-${i}`}>
                            <span className="text-xs text-foreground">{l.email}</span>
                            <CopyBtn text={l.email} />
                          </div>
                        ) : (
                          <span className="text-xs text-muted-foreground/50 italic">—</span>
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        {!l.expiresAt ? (
                          <span className="text-xs text-muted-foreground">Never</span>
                        ) : expired ? (
                          <span className="inline-flex items-center gap-1 text-[10px] font-semibold px-1.5 py-0.5 rounded bg-red-100 text-red-700">
                            Expired {fmtDate(l.expiresAt)}
                          </span>
                        ) : expiringSoon ? (
                          <span className="inline-flex items-center gap-1 text-[10px] font-semibold px-1.5 py-0.5 rounded bg-amber-100 text-amber-700">
                            ⚠ {fmtDate(l.expiresAt)}
                          </span>
                        ) : (
                          <span className="text-xs text-muted-foreground">{fmtDate(l.expiresAt)}</span>
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap" data-testid={`status-reminder-${i}`}>
                        {!l.email ? (
                          <span className="text-[10px] text-muted-foreground/40 italic">no email</span>
                        ) : l.reminderSent ? (
                          <span className="inline-flex items-center gap-1 text-[10px] font-semibold px-1.5 py-0.5 rounded bg-green-100 text-green-700">
                            <CheckCircle2 className="w-3 h-3" /> Sent
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-[10px] font-medium px-1.5 py-0.5 rounded bg-gray-100 text-gray-500">
                            Pending
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-muted-foreground whitespace-nowrap">{fmtDate(l.issuedAt)}</td>
                      <td className="px-4 py-3">
                        <button
                          onClick={() => revokeKey(l.licenseKey)}
                          disabled={revoking === l.licenseKey}
                          className="text-xs text-muted-foreground hover:text-destructive transition-colors disabled:opacity-40"
                          title="Revoke key"
                          data-testid={`button-revoke-${i}`}
                        >
                          {revoking === l.licenseKey ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : "✕"}
                        </button>
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

// ── Payment tab ───────────────────────────────────────────────────────────────
const DEFAULT_PLANS: Plan[] = [
  { id: "monthly",   label: "Monthly",  price: 5,  days: 30  },
  { id: "quarterly", label: "3-Month",  price: 12, days: 90  },
  { id: "biannual",  label: "6-Month",  price: 20, days: 180 },
  { id: "annual",    label: "1-Year",   price: 35, days: 365 },
];

function PaymentTab({ pw, settings, onSaved }: { pw: string; settings: Settings; onSaved: (s: Settings) => void }) {
  const [form, setForm] = useState(settings.payment);
  const [plans, setPlans] = useState<Plan[]>(settings.plans?.length ? settings.plans : DEFAULT_PLANS);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    setForm(settings.payment);
    setPlans(settings.plans?.length ? settings.plans : DEFAULT_PLANS);
  }, [settings]);

  function updatePlanPrice(id: string, price: number) {
    setPlans((prev) => prev.map((p) => p.id === id ? { ...p, price } : p));
  }

  async function save() {
    setSaving(true);
    const res = await fetch(`${API_BASE}/api/admin/settings`, {
      method: "PUT", headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify({ payment: form, plans }),
    });
    setSaving(false);
    if (res.ok) {
      const data = await res.json();
      onSaved(data.settings);
      setSaved(true);
      setTimeout(() => setSaved(false), 2000);
    }
  }

  return (
    <div className="space-y-4">
      {/* Wallet & network */}
      <Card>
        <CardHeader><CardTitle className="text-base">Wallet & Network</CardTitle></CardHeader>
        <CardContent className="space-y-4">
          <div>
            <label className="text-sm font-medium mb-1.5 block">USDT Wallet Address</label>
            <Input value={form.walletAddress} onChange={e => setForm(f => ({ ...f, walletAddress: e.target.value }))}
              placeholder="Your USDT wallet address" className="font-mono text-sm" />
            <p className="text-xs text-muted-foreground mt-1">Customers will send USDT to this address.</p>
          </div>
          <div>
            <label className="text-sm font-medium mb-1.5 block">Network</label>
            <select value={form.network} onChange={e => setForm(f => ({ ...f, network: e.target.value }))}
              className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring">
              <option value="tron">Tron (TRC-20) — Recommended, ~$0 fees</option>
              <option value="bsc">BNB Smart Chain (BEP-20) — Low fees</option>
              <option value="eth">Ethereum (ERC-20) — Higher fees</option>
            </select>
          </div>
          <div className="pt-2 border-t">
            <p className="text-xs font-medium text-muted-foreground mb-2">API Keys (optional — for higher rate limits)</p>
            <div className="space-y-2 text-xs text-muted-foreground bg-muted rounded-md p-3">
              <p>Add <code className="bg-background rounded px-1">TRONGRID_API_KEY</code>, <code className="bg-background rounded px-1">BSCSCAN_API_KEY</code>, or <code className="bg-background rounded px-1">ETHERSCAN_API_KEY</code> as Replit Secrets for higher blockchain API rate limits.</p>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Subscription plans */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Subscription Plans</CardTitle>
          <p className="text-sm text-muted-foreground">Set the price for each plan period. Users choose one before paying.</p>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="grid gap-3 sm:grid-cols-2">
            {plans.map((plan) => (
              <div key={plan.id} className="flex items-center gap-3 border border-border rounded-xl px-4 py-3 bg-muted/20">
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold text-foreground">{plan.label}</p>
                  <p className="text-xs text-muted-foreground">{plan.days} days access</p>
                </div>
                <div className="flex items-center gap-1.5 shrink-0">
                  <span className="text-xs text-muted-foreground">$</span>
                  <Input
                    type="number"
                    min="1"
                    step="0.5"
                    value={plan.price}
                    onChange={e => updatePlanPrice(plan.id, Number(e.target.value))}
                    className="w-20 text-right font-mono font-semibold"
                  />
                  <span className="text-xs text-muted-foreground">USDT</span>
                </div>
              </div>
            ))}
          </div>
          <p className="text-xs text-muted-foreground">Tip: Use tiered pricing (e.g. $5 → $12 → $20 → $35) to encourage longer commitments.</p>
        </CardContent>
      </Card>

      <Button onClick={save} disabled={saving} className="gap-2">
        {saved ? <><Check className="w-4 h-4" /> Saved!</> : saving ? "Saving…" : <><Save className="w-4 h-4" /> Save Payment Settings</>}
      </Button>
    </div>
  );
}

// ── Appearance tab ────────────────────────────────────────────────────────────
const PRESET_COLORS = [
  { label: "Blue", primary: "#3b82f6", accent: "#16a34a" },
  { label: "Purple", primary: "#8b5cf6", accent: "#0891b2" },
  { label: "Rose", primary: "#f43f5e", accent: "#0891b2" },
  { label: "Amber", primary: "#f59e0b", accent: "#10b981" },
  { label: "Slate", primary: "#475569", accent: "#0284c7" },
  { label: "Teal", primary: "#0d9488", accent: "#7c3aed" },
];

const RADIUS_OPTIONS = ["2px", "4px", "6px", "8px", "12px"];

function AppearanceTab({ pw, settings, onSaved }: { pw: string; settings: Settings; onSaved: (s: Settings) => void }) {
  const [form, setForm] = useState(settings.appearance);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);

  useEffect(() => { setForm(settings.appearance); }, [settings]);

  function applyPreview(a: typeof form) {
    function hexToHsl(hex: string) {
      const r = parseInt(hex.slice(1,3),16)/255, g = parseInt(hex.slice(3,5),16)/255, b = parseInt(hex.slice(5,7),16)/255;
      const max = Math.max(r,g,b), min = Math.min(r,g,b); let h=0,s=0; const l=(max+min)/2;
      if(max!==min){const d=max-min;s=l>0.5?d/(2-max-min):d/(max+min);switch(max){case r:h=((g-b)/d+(g<b?6:0))/6;break;case g:h=((b-r)/d+2)/6;break;case b:h=((r-g)/d+4)/6;break;}}
      return `${Math.round(h*360)} ${Math.round(s*100)}% ${Math.round(l*100)}%`;
    }
    document.documentElement.style.setProperty("--primary", hexToHsl(a.primaryColor));
    document.documentElement.style.setProperty("--ring", hexToHsl(a.primaryColor));
    document.documentElement.style.setProperty("--accent", hexToHsl(a.accentColor));
    document.documentElement.style.setProperty("--radius", a.radius);
  }

  function update(patch: Partial<typeof form>) {
    const next = { ...form, ...patch };
    setForm(next);
    applyPreview(next);
  }

  async function save() {
    setSaving(true);
    const res = await fetch(`${API_BASE}/api/admin/settings`, {
      method: "PUT", headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify({ appearance: form }),
    });
    setSaving(false);
    if (res.ok) {
      const data = await res.json();
      onSaved(data.settings);
      setSaved(true);
      setTimeout(() => setSaved(false), 2000);
    }
  }

  return (
    <Card>
      <CardHeader><CardTitle className="text-base">Appearance</CardTitle></CardHeader>
      <CardContent className="space-y-5">
        {/* Name & tagline */}
        <div className="grid gap-4 sm:grid-cols-2">
          <div>
            <label className="text-sm font-medium mb-1.5 block">Add-in Name</label>
            <Input value={form.name} onChange={e => update({ name: e.target.value })} placeholder="Bank Statement Analyzer" />
          </div>
          <div>
            <label className="text-sm font-medium mb-1.5 block">Tagline</label>
            <Input value={form.tagline} onChange={e => update({ tagline: e.target.value })} placeholder="Analyze transactions…" />
          </div>
        </div>

        {/* Color presets */}
        <div>
          <label className="text-sm font-medium mb-2 block">Color Presets</label>
          <div className="flex flex-wrap gap-2">
            {PRESET_COLORS.map(p => (
              <button key={p.label} onClick={() => update({ primaryColor: p.primary, accentColor: p.accent })}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-md border text-xs font-medium hover:border-primary transition-colors"
                style={{ borderColor: form.primaryColor === p.primary ? p.primary : undefined }}>
                <span className="w-3 h-3 rounded-full" style={{ background: p.primary }} />
                {p.label}
              </button>
            ))}
          </div>
        </div>

        {/* Custom colors */}
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="text-sm font-medium mb-1.5 block">Primary Color</label>
            <div className="flex items-center gap-2">
              <input type="color" value={form.primaryColor} onChange={e => update({ primaryColor: e.target.value })}
                className="w-9 h-9 rounded-md border border-input cursor-pointer p-0.5" />
              <Input value={form.primaryColor} onChange={e => /^#[0-9a-fA-F]{0,6}$/.test(e.target.value) && update({ primaryColor: e.target.value })}
                className="font-mono text-sm" maxLength={7} />
            </div>
          </div>
          <div>
            <label className="text-sm font-medium mb-1.5 block">Accent Color</label>
            <div className="flex items-center gap-2">
              <input type="color" value={form.accentColor} onChange={e => update({ accentColor: e.target.value })}
                className="w-9 h-9 rounded-md border border-input cursor-pointer p-0.5" />
              <Input value={form.accentColor} onChange={e => /^#[0-9a-fA-F]{0,6}$/.test(e.target.value) && update({ accentColor: e.target.value })}
                className="font-mono text-sm" maxLength={7} />
            </div>
          </div>
        </div>

        {/* Border radius */}
        <div>
          <label className="text-sm font-medium mb-2 block">Border Radius</label>
          <div className="flex gap-2 flex-wrap">
            {RADIUS_OPTIONS.map(r => (
              <button key={r} onClick={() => update({ radius: r })}
                className={`px-3 py-1.5 text-xs rounded-md border font-medium transition-colors ${form.radius === r ? "bg-primary text-primary-foreground border-primary" : "hover:border-primary"}`}>
                {r}
              </button>
            ))}
          </div>
        </div>

        {/* Live preview */}
        <div className="border rounded-lg p-4 bg-muted/30">
          <p className="text-xs font-medium text-muted-foreground mb-3">Live Preview</p>
          <div className="flex items-center gap-2 mb-3">
            <div className="w-7 h-7 rounded flex items-center justify-center" style={{ background: form.primaryColor }}>
              <svg className="w-4 h-4 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <rect x="2" y="3" width="20" height="18" rx="2"/><path d="M8 7h8M8 11h8M8 15h5"/>
              </svg>
            </div>
            <div>
              <div className="font-semibold text-sm">{form.name || "Add-in Name"}</div>
              <div className="text-[10px] text-muted-foreground">Excel Add-in</div>
            </div>
          </div>
          <p className="text-xs text-muted-foreground">{form.tagline || "Your tagline here"}</p>
          <div className="mt-3 flex gap-2">
            <span className="text-xs px-2.5 py-1 rounded-full text-white font-medium" style={{ background: form.primaryColor, borderRadius: form.radius }}>Primary Button</span>
            <span className="text-xs px-2.5 py-1 rounded-full text-white font-medium" style={{ background: form.accentColor, borderRadius: form.radius }}>Accent Button</span>
          </div>
        </div>

        <Button onClick={save} disabled={saving} className="gap-2">
          {saved ? <><Check className="w-4 h-4" /> Saved!</> : saving ? "Saving…" : <><Save className="w-4 h-4" /> Save Appearance</>}
        </Button>
      </CardContent>
    </Card>
  );
}

// ── Backup tab ────────────────────────────────────────────────────────────────
function BackupTab({ pw }: { pw: string }) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState("");

  async function exportData() {
    const res = await fetch(`${API_BASE}/api/admin/export`, { headers: { "x-admin-password": pw } });
    if (!res.ok) return;
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url;
    a.download = `bsa-backup-${new Date().toISOString().slice(0,10)}.json`;
    a.click(); URL.revokeObjectURL(url);
  }

  async function handleImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]; if (!file) return;
    setImporting(true); setImportMsg("");
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      const res = await fetch(`${API_BASE}/api/admin/import`, {
        method: "POST", headers: { "Content-Type": "application/json", "x-admin-password": pw },
        body: JSON.stringify(json),
      });
      setImportMsg(res.ok ? "✅ Import successful! Reload the page to see changes." : "❌ Import failed.");
    } catch {
      setImportMsg("❌ Invalid backup file.");
    }
    setImporting(false);
    if (fileRef.current) fileRef.current.value = "";
  }

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader><CardTitle className="text-base flex items-center gap-2"><Download className="w-4 h-4" /> Export Backup</CardTitle></CardHeader>
        <CardContent>
          <p className="text-sm text-muted-foreground mb-4">Download all settings and license records as a JSON file you can use to restore or migrate.</p>
          <Button onClick={exportData} variant="outline" className="gap-2"><Download className="w-4 h-4" /> Download Backup</Button>
        </CardContent>
      </Card>

      <Card>
        <CardHeader><CardTitle className="text-base flex items-center gap-2"><Upload className="w-4 h-4" /> Import Backup</CardTitle></CardHeader>
        <CardContent>
          <p className="text-sm text-muted-foreground mb-4">Restore settings and licenses from a previously exported backup file. <strong>This will overwrite current data.</strong></p>
          <input ref={fileRef} type="file" accept=".json" onChange={handleImport} className="hidden" />
          <Button onClick={() => fileRef.current?.click()} variant="outline" disabled={importing} className="gap-2">
            <Upload className="w-4 h-4" /> {importing ? "Importing…" : "Choose Backup File"}
          </Button>
          {importMsg && <p className="mt-3 text-sm">{importMsg}</p>}
        </CardContent>
      </Card>
    </div>
  );
}

// ── Setup Guide tab ───────────────────────────────────────────────────────────
function Step({ done, children }: { done?: boolean; children: React.ReactNode }) {
  return (
    <div className="flex gap-3 items-start">
      <div className={`mt-0.5 shrink-0 w-5 h-5 rounded-full flex items-center justify-center ${done ? "bg-green-100 text-green-600" : "bg-muted text-muted-foreground"}`}>
        {done ? <CheckCircle2 className="w-3.5 h-3.5" /> : <AlertCircle className="w-3.5 h-3.5" />}
      </div>
      <div className="text-sm leading-relaxed">{children}</div>
    </div>
  );
}

function CodeLine({ children }: { children: string }) {
  const [copied, setCopied] = useState(false);
  return (
    <div className="flex items-center gap-2 bg-gray-900 text-green-400 rounded-md px-3 py-2 font-mono text-xs my-1.5">
      <Terminal className="w-3.5 h-3.5 shrink-0 text-gray-500" />
      <span className="flex-1 select-all">{children}</span>
      <button onClick={() => { navigator.clipboard.writeText(children); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
        className="text-gray-500 hover:text-white transition-colors">
        {copied ? <Check className="w-3.5 h-3.5 text-green-400" /> : <Copy className="w-3.5 h-3.5" />}
      </button>
    </div>
  );
}

function EnvRow({ name, required, desc }: { name: string; required?: boolean; desc: string }) {
  return (
    <div className="flex items-start gap-3 py-2.5 border-b last:border-0">
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <code className="text-xs font-mono bg-gray-100 px-1.5 py-0.5 rounded text-gray-800">{name}</code>
          {required && <Badge variant="destructive" className="text-[10px] px-1.5 py-0 h-4">Required</Badge>}
        </div>
        <p className="text-xs text-muted-foreground mt-0.5">{desc}</p>
      </div>
    </div>
  );
}

function SetupTab() {
  const backendUrl = API_BASE || "https://your-api.onrender.com";

  return (
    <div className="space-y-5">

      {/* Render deployment */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base flex items-center gap-2">
            <Rocket className="w-4 h-4 text-primary" /> Render Deployment Guide
          </CardTitle>
          <p className="text-sm text-muted-foreground">Step-by-step setup for deploying both services to Render.</p>
        </CardHeader>
        <CardContent className="space-y-5">

          {/* Step 1 */}
          <div className="space-y-2">
            <p className="text-sm font-semibold text-foreground">1 · Deploy the Backend (Web Service)</p>
            <div className="space-y-2 pl-1">
              <Step>Create a new <strong>Web Service</strong> on <a href="https://render.com" target="_blank" rel="noreferrer" className="text-primary underline underline-offset-2 inline-flex items-center gap-0.5">render.com <ExternalLink className="w-3 h-3" /></a> and connect your Git repo.</Step>
              <Step>Set <strong>Root Directory</strong> to <code className="bg-muted px-1 rounded text-xs">server</code></Step>
              <Step>Use these build &amp; start commands:
                <CodeLine>npm install</CodeLine>
                <CodeLine>node index.js</CodeLine>
              </Step>
              <Step>Add the environment variables listed in the table below.</Step>
              <Step>Deploy — copy the service URL (e.g. <code className="bg-muted px-1 rounded text-xs">https://bank-analyzer-api.onrender.com</code>).</Step>
            </div>
          </div>

          <div className="border-t pt-4 space-y-2">
            <p className="text-sm font-semibold text-foreground">2 · Deploy the Frontend (Static Site)</p>
            <div className="space-y-2 pl-1">
              <Step>Create a new <strong>Static Site</strong> on Render from the same repo.</Step>
              <Step>Leave <strong>Root Directory</strong> blank (uses repo root).</Step>
              <Step>Build command &amp; publish directory:
                <CodeLine>npm install && npm run build</CodeLine>
                <span className="text-xs text-muted-foreground">Publish directory: <code className="bg-muted px-1 rounded">dist/public</code></span>
              </Step>
              <Step>Set <code className="bg-muted px-1 rounded text-xs">VITE_API_URL</code> to your backend URL from Step 1.</Step>
              <Step>Deploy — your add-in UI is now live.</Step>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Env vars table */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Backend Environment Variables</CardTitle>
          <p className="text-sm text-muted-foreground">Set these in Render → your Web Service → Environment.</p>
        </CardHeader>
        <CardContent className="p-0 px-4 pb-2">
          <EnvRow name="ADMIN_PASSWORD" required desc="Password to access this admin panel. Choose something strong." />
          <EnvRow name="USDT_WALLET_ADDRESS" required desc="Your USDT wallet address where customers send payments." />
          <EnvRow name="USDT_NETWORK" desc='Blockchain network: "tron" (default), "bsc", or "eth".' />
          <EnvRow name="USDT_PRICE" desc='Pro license price in USDT. Default: "5".' />
          <EnvRow name="TRONGRID_API_KEY" desc="TronGrid API key — raises rate limits for Tron payment verification." />
          <EnvRow name="BSCSCAN_API_KEY" desc="BscScan API key — raises rate limits for BSC payment verification." />
          <EnvRow name="ETHERSCAN_API_KEY" desc="Etherscan API key — raises rate limits for ETH payment verification." />
        </CardContent>
      </Card>

      {/* Frontend env var */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Frontend Environment Variable</CardTitle>
          <p className="text-sm text-muted-foreground">Set this in Render → your Static Site → Environment (before building).</p>
        </CardHeader>
        <CardContent className="p-0 px-4 pb-2">
          <EnvRow name="VITE_API_URL" required desc={`URL of your deployed backend, e.g. ${backendUrl}`} />
        </CardContent>
      </Card>

      {/* manifest.xml note */}
      <Card className="border-amber-200 bg-amber-50/50">
        <CardContent className="pt-4 pb-4">
          <div className="flex gap-3">
            <AlertCircle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
            <div className="text-sm">
              <p className="font-semibold text-amber-800 mb-1">Update manifest.xml before sideloading</p>
              <p className="text-amber-700">Set the <code className="bg-amber-100 px-1 rounded text-xs">SourceLocation</code> URL in <code className="bg-amber-100 px-1 rounded text-xs">manifest.xml</code> to your deployed frontend URL so Excel loads the correct add-in.</p>
            </div>
          </div>
        </CardContent>
      </Card>

    </div>
  );
}

// ── Notifications tab ─────────────────────────────────────────────────────────
const DEFAULT_EMAIL_CFG: EmailNotifyCfg = {
  enabled: false, to: "", smtpHost: "", smtpPort: 587, smtpUser: "", smtpPass: "", from: "",
};

function NotificationsTab({ pw, settings, onSaved }: { pw: string; settings: Settings; onSaved: (s: Settings) => void }) {
  const [webhookUrl, setWebhookUrl] = useState(settings.notifications?.webhookUrl ?? "");
  const [email, setEmail] = useState<EmailNotifyCfg>(settings.notifications?.email ?? DEFAULT_EMAIL_CFG);
  const [remindersEnabled, setRemindersEnabled] = useState(settings.notifications?.remindersEnabled ?? false);
  const [reminderDays, setReminderDays] = useState(settings.notifications?.reminderDays ?? 3);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [testing, setTesting] = useState(false);
  const [testMsg, setTestMsg] = useState<{ ok: boolean; text: string } | null>(null);

  useEffect(() => {
    setWebhookUrl(settings.notifications?.webhookUrl ?? "");
    setEmail(settings.notifications?.email ?? DEFAULT_EMAIL_CFG);
    setRemindersEnabled(settings.notifications?.remindersEnabled ?? false);
    setReminderDays(settings.notifications?.reminderDays ?? 3);
  }, [settings]);

  function setEmailField<K extends keyof EmailNotifyCfg>(k: K, v: EmailNotifyCfg[K]) {
    setEmail(e => ({ ...e, [k]: v }));
  }

  async function save() {
    setSaving(true);
    const res = await fetch(`${API_BASE}/api/admin/settings`, {
      method: "PUT",
      headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify({ notifications: { webhookUrl, email, remindersEnabled, reminderDays } }),
    });
    setSaving(false);
    if (res.ok) {
      const data = await res.json();
      onSaved(data.settings);
      setSaved(true);
      setTimeout(() => setSaved(false), 2000);
    }
  }

  async function sendTest() {
    setTesting(true);
    setTestMsg(null);
    const saveRes = await fetch(`${API_BASE}/api/admin/settings`, {
      method: "PUT",
      headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify({ notifications: { webhookUrl, email } }),
    });
    if (!saveRes.ok) { setTesting(false); setTestMsg({ ok: false, text: "Failed to save before testing." }); return; }
    const data = await saveRes.json();
    onSaved(data.settings);

    const res = await fetch(`${API_BASE}/api/admin/notify-test`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-admin-password": pw },
    });
    setTesting(false);
    const json = await res.json().catch(() => ({}));
    if (res.ok) setTestMsg({ ok: true, text: json.message || "Test notification sent!" });
    else setTestMsg({ ok: false, text: json.error || "Test failed." });
    setTimeout(() => setTestMsg(null), 6000);
  }

  return (
    <div className="space-y-4">
      {/* Webhook */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Webhook Notification</CardTitle>
          <p className="text-sm text-muted-foreground">Works with Discord, Slack, Make.com, Zapier, n8n — paste any incoming webhook URL.</p>
        </CardHeader>
        <CardContent className="space-y-3">
          <div>
            <label className="text-sm font-medium mb-1.5 block">Webhook URL</label>
            <Input
              value={webhookUrl}
              onChange={e => setWebhookUrl(e.target.value)}
              placeholder="https://discord.com/api/webhooks/… or https://hooks.slack.com/…"
              className="font-mono text-sm"
              data-testid="input-webhook-url"
            />
          </div>
          <p className="text-xs text-muted-foreground">
            A JSON payload with plan, license key, expiry, and TX hash is POSTed to this URL each time a new subscription activates.
          </p>
        </CardContent>
      </Card>

      {/* Email */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Email Notification (SMTP)</CardTitle>
          <p className="text-sm text-muted-foreground">Send yourself an email via your own SMTP server (Gmail, Mailgun, etc.) on each new activation.</p>
        </CardHeader>
        <CardContent className="space-y-3">
          <label className="flex items-center gap-2 cursor-pointer select-none">
            <input type="checkbox" checked={email.enabled} onChange={e => setEmailField("enabled", e.target.checked)}
              className="w-4 h-4 rounded accent-primary" data-testid="checkbox-email-enabled" />
            <span className="text-sm font-medium">Enable email notifications</span>
          </label>

          <div className={`space-y-3 transition-opacity ${email.enabled ? "opacity-100" : "opacity-40 pointer-events-none"}`}>
            <div className="grid gap-3 sm:grid-cols-2">
              <div>
                <label className="text-xs font-medium mb-1 block text-muted-foreground">Notify to (your email)</label>
                <Input value={email.to} onChange={e => setEmailField("to", e.target.value)}
                  placeholder="you@example.com" type="email" data-testid="input-email-to" />
              </div>
              <div>
                <label className="text-xs font-medium mb-1 block text-muted-foreground">From address (optional)</label>
                <Input value={email.from} onChange={e => setEmailField("from", e.target.value)}
                  placeholder="noreply@yourdomain.com" type="email" data-testid="input-email-from" />
              </div>
            </div>
            <div className="grid gap-3 sm:grid-cols-3">
              <div className="sm:col-span-2">
                <label className="text-xs font-medium mb-1 block text-muted-foreground">SMTP Host</label>
                <Input value={email.smtpHost} onChange={e => setEmailField("smtpHost", e.target.value)}
                  placeholder="smtp.gmail.com" data-testid="input-smtp-host" />
              </div>
              <div>
                <label className="text-xs font-medium mb-1 block text-muted-foreground">Port</label>
                <Input type="number" value={email.smtpPort} onChange={e => setEmailField("smtpPort", Number(e.target.value))}
                  placeholder="587" className="font-mono" data-testid="input-smtp-port" />
              </div>
            </div>
            <div className="grid gap-3 sm:grid-cols-2">
              <div>
                <label className="text-xs font-medium mb-1 block text-muted-foreground">SMTP Username</label>
                <Input value={email.smtpUser} onChange={e => setEmailField("smtpUser", e.target.value)}
                  placeholder="you@gmail.com" data-testid="input-smtp-user" />
              </div>
              <div>
                <label className="text-xs font-medium mb-1 block text-muted-foreground">SMTP Password / App Password</label>
                <Input type="password" value={email.smtpPass} onChange={e => setEmailField("smtpPass", e.target.value)}
                  placeholder="••••••••" data-testid="input-smtp-pass" />
              </div>
            </div>
            <p className="text-xs text-muted-foreground bg-muted rounded-md p-3">
              Gmail tip: Use port 587, your Gmail address as username, and an <strong>App Password</strong> (not your regular password). Generate one at myaccount.google.com → Security → App passwords.
            </p>
          </div>
        </CardContent>
      </Card>

      {/* Renewal reminders */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Renewal Reminders</CardTitle>
          <p className="text-sm text-muted-foreground">Automatically email subscribers before their plan expires (requires Email SMTP above).</p>
        </CardHeader>
        <CardContent className="space-y-3">
          <label className="flex items-center gap-2 cursor-pointer select-none">
            <input type="checkbox" checked={remindersEnabled} onChange={e => setRemindersEnabled(e.target.checked)}
              className="w-4 h-4 rounded accent-primary" data-testid="checkbox-reminders-enabled" />
            <span className="text-sm font-medium">Send renewal reminder emails to subscribers</span>
          </label>

          <div className={`space-y-1 transition-opacity ${remindersEnabled ? "opacity-100" : "opacity-40 pointer-events-none"}`}>
            <label className="text-xs font-medium text-muted-foreground block">Days before expiry to send reminder</label>
            <div className="flex items-center gap-2">
              <Input type="number" min="1" max="30" value={reminderDays}
                onChange={e => setReminderDays(Number(e.target.value))}
                className="w-24 font-mono" data-testid="input-reminder-days" />
              <span className="text-sm text-muted-foreground">days before expiry</span>
            </div>
          </div>

          <div className="text-xs text-muted-foreground bg-muted rounded-md p-3 space-y-1">
            <p>• Subscribers provide their email optionally during payment checkout.</p>
            <p>• The server checks every 6 hours and sends one reminder per license.</p>
            <p>• Requires <strong>Email SMTP</strong> to be configured and enabled above.</p>
          </div>
        </CardContent>
      </Card>

      {testMsg && (
        <div className={`flex items-center gap-2 rounded-lg px-4 py-3 text-sm ${testMsg.ok ? "bg-green-50 text-green-800 border border-green-200" : "bg-red-50 text-red-800 border border-red-200"}`}>
          {testMsg.ok ? <CheckCircle2 className="w-4 h-4 shrink-0" /> : <AlertCircle className="w-4 h-4 shrink-0" />}
          {testMsg.text}
        </div>
      )}

      <div className="flex flex-wrap gap-2">
        <Button onClick={save} disabled={saving} className="gap-2" data-testid="button-save-notifications">
          {saved ? <><Check className="w-4 h-4" /> Saved!</> : saving ? "Saving…" : <><Save className="w-4 h-4" /> Save Notifications</>}
        </Button>
        <Button variant="outline" onClick={sendTest} disabled={testing} className="gap-2" data-testid="button-test-notification">
          {testing ? "Sending…" : <><Bell className="w-4 h-4" /> Send Test</>}
        </Button>
      </div>
    </div>
  );
}

// ── Revenue tab ────────────────────────────────────────────────────────────────
interface RevenueData {
  totalRevenue: number;
  mrr: number;
  totalLicenses: number;
  activeLicenses: number;
  paidLicenses: number;
  withEmail: number;
  byPlan: { planId: string; count: number; revenue: number }[];
  monthly: { month: string; activations: number; revenue: number }[];
}

const PLAN_LABELS: Record<string, string> = {
  monthly: "Monthly", quarterly: "3-Month", biannual: "6-Month", annual: "Annual",
};
const PLAN_COLORS: Record<string, string> = {
  monthly: "bg-blue-500", quarterly: "bg-violet-500", biannual: "bg-amber-500", annual: "bg-emerald-500",
};

function RevenueTab({ pw }: { pw: string }) {
  const [data, setData] = useState<RevenueData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function load() {
    setLoading(true); setError("");
    try {
      const res = await fetch(`${API_BASE}/api/admin/revenue`, { headers: { "x-admin-password": pw } });
      if (!res.ok) throw new Error("Failed to load");
      setData(await res.json());
    } catch { setError("Could not load revenue data."); }
    finally { setLoading(false); }
  }

  useEffect(() => { load(); }, []);

  if (loading) return <div className="text-center py-20 text-muted-foreground text-sm">Loading revenue data…</div>;
  if (error || !data) return (
    <div className="text-center py-20 space-y-2">
      <p className="text-muted-foreground text-sm">{error || "No data"}</p>
      <Button variant="outline" size="sm" onClick={load}>Retry</Button>
    </div>
  );

  const maxMonthlyRev = Math.max(...data.monthly.map(m => m.revenue), 1);
  const maxMonthlyAct = Math.max(...data.monthly.map(m => m.activations), 1);
  const maxPlanRev    = Math.max(...data.byPlan.map(p => p.revenue), 1);
  const last13 = data.monthly.slice(-13);

  function fmt(n: number) { return `$${n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }
  function fmtMonth(s: string) {
    const [y, m] = s.split("-");
    return new Date(Number(y), Number(m) - 1).toLocaleDateString("en-US", { month: "short", year: "2-digit" });
  }

  return (
    <div className="space-y-6">
      {/* KPI cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <Card data-testid="card-total-revenue">
          <CardContent className="p-4">
            <div className="flex items-center justify-between mb-1">
              <span className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Total Revenue</span>
              <DollarSign className="w-4 h-4 text-emerald-500" />
            </div>
            <p className="text-2xl font-bold text-emerald-600" data-testid="text-total-revenue">{fmt(data.totalRevenue)}</p>
            <p className="text-xs text-muted-foreground mt-0.5">{data.paidLicenses} paid license{data.paidLicenses !== 1 ? "s" : ""}</p>
          </CardContent>
        </Card>
        <Card data-testid="card-mrr">
          <CardContent className="p-4">
            <div className="flex items-center justify-between mb-1">
              <span className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Est. MRR</span>
              <TrendingUp className="w-4 h-4 text-blue-500" />
            </div>
            <p className="text-2xl font-bold text-blue-600" data-testid="text-mrr">{fmt(data.mrr)}</p>
            <p className="text-xs text-muted-foreground mt-0.5">from active subscriptions</p>
          </CardContent>
        </Card>
        <Card data-testid="card-active-licenses">
          <CardContent className="p-4">
            <div className="flex items-center justify-between mb-1">
              <span className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Active</span>
              <Users className="w-4 h-4 text-violet-500" />
            </div>
            <p className="text-2xl font-bold text-violet-600" data-testid="text-active-licenses">{data.activeLicenses}</p>
            <p className="text-xs text-muted-foreground mt-0.5">of {data.totalLicenses} total issued</p>
          </CardContent>
        </Card>
        <Card data-testid="card-arpu">
          <CardContent className="p-4">
            <div className="flex items-center justify-between mb-1">
              <span className="text-xs font-medium text-muted-foreground uppercase tracking-wide">ARPU</span>
              <Activity className="w-4 h-4 text-amber-500" />
            </div>
            <p className="text-2xl font-bold text-amber-600" data-testid="text-arpu">
              {fmt(data.paidLicenses > 0 ? data.totalRevenue / data.paidLicenses : 0)}
            </p>
            <p className="text-xs text-muted-foreground mt-0.5">avg revenue per user</p>
          </CardContent>
        </Card>
      </div>

      {/* Monthly revenue chart */}
      <Card>
        <CardHeader className="pb-3">
          <div className="flex items-center justify-between">
            <CardTitle className="text-base flex items-center gap-2">
              <BarChart3 className="w-4 h-4 text-muted-foreground" /> Monthly Revenue (last 13 months)
            </CardTitle>
            <Button variant="ghost" size="sm" onClick={load} className="h-7 px-2" data-testid="button-refresh-revenue">
              <RefreshCw className="w-3.5 h-3.5" />
            </Button>
          </div>
        </CardHeader>
        <CardContent>
          {last13.every(m => m.revenue === 0) ? (
            <div className="text-center py-8 text-muted-foreground text-sm">No paid license revenue recorded yet.</div>
          ) : (
            <div className="space-y-2">
              <div className="flex items-end gap-1 h-40" data-testid="chart-monthly-revenue">
                {last13.map(m => (
                  <div key={m.month} className="flex-1 flex flex-col items-center gap-1 min-w-0">
                    <div
                      className="w-full bg-blue-500 rounded-t transition-all hover:bg-blue-600"
                      style={{ height: `${Math.max(2, (m.revenue / maxMonthlyRev) * 140)}px` }}
                      title={`${fmtMonth(m.month)}: ${fmt(m.revenue)}`}
                    />
                  </div>
                ))}
              </div>
              <div className="flex gap-1">
                {last13.map(m => (
                  <div key={m.month} className="flex-1 text-center min-w-0">
                    <span className="text-[9px] text-muted-foreground truncate block">{fmtMonth(m.month)}</span>
                  </div>
                ))}
              </div>
              <div className="flex gap-1 mt-1">
                {last13.map(m => (
                  <div key={m.month} className="flex-1 text-center min-w-0">
                    <span className="text-[9px] font-medium text-blue-700 truncate block">{m.revenue > 0 ? fmt(m.revenue) : ""}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Monthly activations chart */}
      <Card>
        <CardHeader className="pb-3">
          <CardTitle className="text-base flex items-center gap-2">
            <BarChart3 className="w-4 h-4 text-muted-foreground" /> Monthly Activations (last 13 months)
          </CardTitle>
        </CardHeader>
        <CardContent>
          {last13.every(m => m.activations === 0) ? (
            <div className="text-center py-8 text-muted-foreground text-sm">No activations recorded yet.</div>
          ) : (
            <div className="space-y-2">
              <div className="flex items-end gap-1 h-32" data-testid="chart-monthly-activations">
                {last13.map(m => (
                  <div key={m.month} className="flex-1 flex flex-col items-center min-w-0">
                    <div
                      className="w-full bg-violet-400 rounded-t transition-all hover:bg-violet-500"
                      style={{ height: `${Math.max(2, (m.activations / maxMonthlyAct) * 112)}px` }}
                      title={`${fmtMonth(m.month)}: ${m.activations} activation${m.activations !== 1 ? "s" : ""}`}
                    />
                  </div>
                ))}
              </div>
              <div className="flex gap-1">
                {last13.map(m => (
                  <div key={m.month} className="flex-1 text-center min-w-0">
                    <span className="text-[9px] text-muted-foreground truncate block">{fmtMonth(m.month)}</span>
                  </div>
                ))}
              </div>
              <div className="flex gap-1 mt-1">
                {last13.map(m => (
                  <div key={m.month} className="flex-1 text-center min-w-0">
                    <span className="text-[9px] font-medium text-violet-700 truncate block">{m.activations > 0 ? m.activations : ""}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Revenue by plan */}
      {data.byPlan.length > 0 && (
        <Card>
          <CardHeader className="pb-3">
            <CardTitle className="text-base">Revenue by Plan</CardTitle>
          </CardHeader>
          <CardContent className="space-y-3" data-testid="section-revenue-by-plan">
            {data.byPlan.map(p => (
              <div key={p.planId} className="space-y-1">
                <div className="flex items-center justify-between text-sm">
                  <span className="font-medium">{PLAN_LABELS[p.planId] ?? p.planId}</span>
                  <div className="flex items-center gap-3">
                    <span className="text-muted-foreground text-xs">{p.count} license{p.count !== 1 ? "s" : ""}</span>
                    <span className="font-semibold w-20 text-right">{fmt(p.revenue)}</span>
                  </div>
                </div>
                <div className="h-2 bg-muted rounded-full overflow-hidden">
                  <div
                    className={`h-full rounded-full transition-all ${PLAN_COLORS[p.planId] ?? "bg-gray-400"}`}
                    style={{ width: `${(p.revenue / maxPlanRev) * 100}%` }}
                  />
                </div>
              </div>
            ))}
            <div className="pt-2 border-t flex items-center justify-between text-sm font-semibold">
              <span>Total</span>
              <span className="text-emerald-600">{fmt(data.totalRevenue)}</span>
            </div>
          </CardContent>
        </Card>
      )}

      {data.byPlan.length === 0 && (
        <Card>
          <CardContent className="py-10 text-center text-muted-foreground text-sm">
            No paid license sales yet. Revenue will appear here once customers purchase subscriptions.
          </CardContent>
        </Card>
      )}
    </div>
  );
}

// ── Main admin page ───────────────────────────────────────────────────────────
type Tab = "licenses" | "revenue" | "payment" | "appearance" | "notifications" | "backup" | "setup" | "tickets";

// ── Tickets tab ────────────────────────────────────────────────────────────────
interface Ticket {
  id: string; name: string; email: string; licenseKey: string | null;
  category: string; subject: string; message: string;
  status: "open" | "resolved" | "closed"; createdAt: string;
  adminReply: string | null; repliedAt: string | null;
}

const STATUS_COLORS: Record<string, string> = {
  open:     "bg-amber-100 text-amber-700",
  resolved: "bg-green-100 text-green-700",
  closed:   "bg-gray-100 text-gray-600",
};

const CATEGORY_LABELS: Record<string, string> = {
  general: "General", billing: "Billing", license: "License",
  bug: "Bug", feature: "Feature", other: "Other",
};

function TicketsTab({ pw }: { pw: string }) {
  const [data, setData] = useState<{ total: number; tickets: Ticket[] } | null>(null);
  const [loading, setLoading] = useState(false);
  const [filter, setFilter] = useState<"" | "open" | "resolved" | "closed">("");
  const [selected, setSelected] = useState<Ticket | null>(null);
  const [reply, setReply] = useState("");
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState<string | null>(null);

  async function load() {
    setLoading(true);
    const url = filter ? `${API_BASE}/api/tickets?status=${filter}` : `${API_BASE}/api/tickets`;
    const res = await fetch(url, { headers: { "x-admin-password": pw } });
    if (res.ok) setData(await res.json());
    setLoading(false);
  }

  useEffect(() => { load(); }, [filter]);

  async function updateTicket(id: string, patch: { status?: string; adminReply?: string }) {
    setSaving(true);
    const res = await fetch(`${API_BASE}/api/tickets/${id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json", "x-admin-password": pw },
      body: JSON.stringify(patch),
    });
    setSaving(false);
    if (res.ok) { await load(); setSelected(null); setReply(""); }
  }

  async function deleteTicket(id: string) {
    setDeleting(id);
    await fetch(`${API_BASE}/api/tickets/${id}`, { method: "DELETE", headers: { "x-admin-password": pw } });
    setDeleting(null);
    if (selected?.id === id) setSelected(null);
    load();
  }

  const tickets = data?.tickets ?? [];
  const openCount = data ? (data.tickets.filter(t => t.status === "open").length) : 0;

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-lg font-bold">Support Tickets</h2>
          <p className="text-sm text-muted-foreground">{data?.total ?? 0} total · {openCount} open</p>
        </div>
        <div className="flex items-center gap-2">
          <select
            value={filter}
            onChange={e => setFilter(e.target.value as typeof filter)}
            className="text-sm border border-border rounded-lg px-2.5 py-1.5 bg-white"
          >
            <option value="">All</option>
            <option value="open">Open</option>
            <option value="resolved">Resolved</option>
            <option value="closed">Closed</option>
          </select>
          <Button variant="outline" size="sm" onClick={load} disabled={loading}>
            <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} />
          </Button>
        </div>
      </div>

      {loading && <div className="text-center py-12 text-muted-foreground text-sm">Loading…</div>}

      {!loading && tickets.length === 0 && (
        <div className="text-center py-12 text-muted-foreground text-sm">No tickets found.</div>
      )}

      {/* Detail view */}
      {selected && (
        <Card className="border-primary/30">
          <CardHeader className="pb-3">
            <div className="flex items-start justify-between gap-3">
              <div className="min-w-0">
                <div className="flex items-center gap-2 flex-wrap mb-1">
                  <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${STATUS_COLORS[selected.status]}`}>
                    {selected.status.toUpperCase()}
                  </span>
                  <span className="text-xs font-mono text-muted-foreground">{selected.id}</span>
                  <span className="text-xs bg-muted text-muted-foreground px-2 py-0.5 rounded-full">
                    {CATEGORY_LABELS[selected.category] ?? selected.category}
                  </span>
                </div>
                <CardTitle className="text-base leading-snug">{selected.subject}</CardTitle>
              </div>
              <button onClick={() => { setSelected(null); setReply(""); }}
                className="text-muted-foreground hover:text-foreground shrink-0">
                <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="text-xs space-y-0.5 text-muted-foreground">
              <p><span className="font-semibold text-foreground">From:</span> {selected.name} &lt;{selected.email}&gt;</p>
              {selected.licenseKey && <p><span className="font-semibold text-foreground">License:</span> <code className="font-mono">{selected.licenseKey}</code></p>}
              <p><span className="font-semibold text-foreground">Received:</span> {fmtDate(selected.createdAt)}</p>
            </div>
            <div className="bg-muted/50 rounded-lg p-3 text-sm leading-relaxed whitespace-pre-wrap border border-border">
              {selected.message}
            </div>

            {selected.adminReply && (
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 space-y-1">
                <p className="text-xs font-bold text-blue-700">Your Reply · {selected.repliedAt ? fmtDate(selected.repliedAt) : ""}</p>
                <p className="text-sm leading-relaxed whitespace-pre-wrap">{selected.adminReply}</p>
              </div>
            )}

            <div className="space-y-2">
              <label className="text-xs font-semibold">{selected.adminReply ? "Update Reply" : "Reply to Customer"}</label>
              <textarea
                value={reply}
                onChange={e => setReply(e.target.value)}
                placeholder="Type your reply…"
                rows={4}
                className="w-full rounded-lg border border-border p-3 text-sm resize-none focus:outline-none focus:border-primary"
              />
            </div>

            <div className="flex flex-wrap gap-2">
              <Button size="sm" onClick={() => updateTicket(selected.id, { adminReply: reply })} disabled={saving || !reply.trim()}>
                {saving ? "Saving…" : "Send Reply"}
              </Button>
              <Button size="sm" variant="outline" onClick={() => updateTicket(selected.id, { status: "resolved" })} disabled={saving}>
                Mark Resolved
              </Button>
              <Button size="sm" variant="outline" onClick={() => updateTicket(selected.id, { status: "closed" })} disabled={saving}>
                Close
              </Button>
              <Button size="sm" variant="outline" onClick={() => updateTicket(selected.id, { status: "open" })} disabled={saving}>
                Reopen
              </Button>
              <Button size="sm" variant="destructive" onClick={() => deleteTicket(selected.id)} disabled={!!deleting}>
                Delete
              </Button>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Ticket list */}
      {!loading && tickets.length > 0 && (
        <div className="space-y-2">
          {tickets.map(ticket => (
            <div key={ticket.id}
              onClick={() => { setSelected(ticket); setReply(ticket.adminReply ?? ""); }}
              className={`bg-white border rounded-xl p-4 cursor-pointer hover:border-primary/50 hover:shadow-sm transition-all ${selected?.id === ticket.id ? "border-primary/60 shadow-sm" : "border-border"}`}>
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2 flex-wrap mb-1">
                    <span className={`text-[11px] font-bold px-1.5 py-0.5 rounded-full ${STATUS_COLORS[ticket.status]}`}>
                      {ticket.status}
                    </span>
                    <span className="text-[11px] bg-muted text-muted-foreground px-1.5 py-0.5 rounded-full">
                      {CATEGORY_LABELS[ticket.category] ?? ticket.category}
                    </span>
                    <span className="text-[11px] font-mono text-muted-foreground">{ticket.id}</span>
                  </div>
                  <p className="text-sm font-semibold truncate">{ticket.subject}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">{ticket.name} · {ticket.email}</p>
                </div>
                <div className="text-right shrink-0">
                  <p className="text-[11px] text-muted-foreground">{fmtDate(ticket.createdAt)}</p>
                  {ticket.adminReply && (
                    <p className="text-[11px] text-green-600 font-semibold mt-0.5">Replied</p>
                  )}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const TABS: { id: Tab; label: string; icon: React.ReactNode }[] = [
  { id: "licenses",      label: "Licenses",      icon: <KeyRound className="w-4 h-4" /> },
  { id: "revenue",       label: "Revenue",        icon: <TrendingUp className="w-4 h-4" /> },
  { id: "tickets",       label: "Tickets",        icon: <TicketCheck className="w-4 h-4" /> },
  { id: "payment",       label: "Payment",        icon: <CreditCard className="w-4 h-4" /> },
  { id: "notifications", label: "Notifications",  icon: <Bell className="w-4 h-4" /> },
  { id: "appearance",    label: "Appearance",     icon: <Palette className="w-4 h-4" /> },
  { id: "backup",        label: "Backup",         icon: <Database className="w-4 h-4" /> },
  { id: "setup",         label: "Setup Guide",    icon: <Rocket className="w-4 h-4" /> },
];

export default function AdminPage() {
  const [pw, setPw] = useState(() => sessionStorage.getItem("admin_pw") ?? "");
  const [authed, setAuthed] = useState(false);
  const [tab, setTab] = useState<Tab>("licenses");
  const [settings, setSettings] = useState<Settings | null>(null);

  async function loadSettings(password: string) {
    const res = await fetch(`${API_BASE}/api/admin/settings`, { headers: { "x-admin-password": password } });
    if (res.ok) setSettings(await res.json());
  }

  function handleLogin(password: string) {
    setPw(password);
    setAuthed(true);
    loadSettings(password);
  }

  function logout() {
    sessionStorage.removeItem("admin_pw");
    setAuthed(false); setSettings(null); setPw("");
  }

  // Auto-login from session
  useEffect(() => {
    const saved = sessionStorage.getItem("admin_pw");
    if (saved) {
      fetch(`${API_BASE}/api/admin/settings`, { headers: { "x-admin-password": saved } })
        .then(r => { if (r.ok) { setAuthed(true); return r.json(); } throw new Error(); })
        .then(s => { setPw(saved); setSettings(s); })
        .catch(() => { sessionStorage.removeItem("admin_pw"); });
    }
  }, []);

  if (!authed) return <LoginScreen onLogin={handleLogin} />;

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Top bar */}
      <div className="bg-white border-b px-4 md:px-8 py-3 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <ShieldCheck className="w-5 h-5 text-primary" />
          <span className="font-semibold text-sm">Admin Panel</span>
          <span className="text-muted-foreground text-xs hidden sm:inline">· Bank Statement Analyzer</span>
        </div>
        <Button variant="ghost" size="sm" onClick={logout} className="gap-1.5">
          <LogOut className="w-4 h-4" /> Logout
        </Button>
      </div>

      <div className="max-w-4xl mx-auto p-4 md:p-8">
        {/* Tabs */}
        <div className="flex gap-1 bg-muted p-1 rounded-lg mb-6 w-fit">
          {TABS.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${tab === t.id ? "bg-white shadow-sm text-foreground" : "text-muted-foreground hover:text-foreground"}`}>
              {t.icon} <span className="hidden sm:inline">{t.label}</span>
            </button>
          ))}
        </div>

        {/* Tab content */}
        {tab === "licenses"      && <LicensesTab pw={pw} />}
        {tab === "revenue"       && <RevenueTab pw={pw} />}
        {tab === "tickets"       && <TicketsTab pw={pw} />}
        {tab === "payment"       && settings && <PaymentTab pw={pw} settings={settings} onSaved={setSettings} />}
        {tab === "notifications" && settings && <NotificationsTab pw={pw} settings={settings} onSaved={setSettings} />}
        {tab === "appearance"    && settings && <AppearanceTab pw={pw} settings={settings} onSaved={setSettings} />}
        {tab === "backup"        && <BackupTab pw={pw} />}
        {tab === "setup"         && <SetupTab />}
        {(tab === "payment" || tab === "appearance" || tab === "notifications") && !settings && (
          <div className="text-center py-12 text-muted-foreground text-sm">Loading settings…</div>
        )}
      </div>
    </div>
  );
}

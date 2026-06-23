import { useState, useEffect } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { KeyRound, LogOut, RefreshCw, ShieldCheck, Copy, Check } from "lucide-react";

const API_BASE = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");

interface License {
  licenseKey: string;
  txHash: string;
  issuedAt: string;
}

interface LicenseData {
  total: number;
  licenses: License[];
}

function formatDate(iso: string) {
  return new Date(iso).toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function CopyButton({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);
  const copy = () => {
    navigator.clipboard.writeText(text).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    });
  };
  return (
    <button onClick={copy} className="ml-1 text-muted-foreground hover:text-foreground transition-colors">
      {copied ? <Check className="w-3.5 h-3.5 text-green-500" /> : <Copy className="w-3.5 h-3.5" />}
    </button>
  );
}

export default function AdminPage() {
  const [password, setPassword] = useState(() => sessionStorage.getItem("admin_pw") ?? "");
  const [authed, setAuthed] = useState(false);
  const [data, setData] = useState<LicenseData | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function fetchLicenses(pw: string) {
    setLoading(true);
    setError("");
    try {
      const res = await fetch(`${API_BASE}/api/admin/licenses`, {
        headers: { "x-admin-password": pw },
      });
      if (res.status === 401) {
        setError("Wrong password.");
        setAuthed(false);
        return;
      }
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        setError(body.error ?? "Server error");
        return;
      }
      const json: LicenseData = await res.json();
      setData(json);
      setAuthed(true);
      sessionStorage.setItem("admin_pw", pw);
    } catch {
      setError("Could not reach the API. Make sure the backend is running.");
    } finally {
      setLoading(false);
    }
  }

  function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    fetchLicenses(password);
  }

  function logout() {
    sessionStorage.removeItem("admin_pw");
    setAuthed(false);
    setData(null);
    setPassword("");
  }

  // Auto-login if password saved in session
  useEffect(() => {
    const saved = sessionStorage.getItem("admin_pw");
    if (saved) fetchLicenses(saved);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!authed) {
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
            <form onSubmit={handleLogin} className="space-y-3">
              <Input
                type="password"
                placeholder="Admin password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                autoFocus
              />
              {error && <p className="text-sm text-red-500">{error}</p>}
              <Button type="submit" className="w-full" disabled={loading || !password}>
                {loading ? "Checking…" : "Login"}
              </Button>
            </form>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-8">
      <div className="max-w-5xl mx-auto space-y-6">
        {/* Header */}
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <KeyRound className="w-5 h-5 text-blue-600" />
            <h1 className="text-xl font-semibold">License Admin</h1>
            <Badge variant="secondary">{data?.total ?? 0} issued</Badge>
          </div>
          <div className="flex gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => fetchLicenses(password)}
              disabled={loading}
            >
              <RefreshCw className={`w-4 h-4 mr-1 ${loading ? "animate-spin" : ""}`} />
              Refresh
            </Button>
            <Button variant="ghost" size="sm" onClick={logout}>
              <LogOut className="w-4 h-4 mr-1" />
              Logout
            </Button>
          </div>
        </div>

        {/* Table */}
        <Card>
          <CardContent className="p-0">
            {!data || data.licenses.length === 0 ? (
              <div className="text-center py-16 text-muted-foreground text-sm">
                No licenses issued yet.
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b bg-gray-50/60">
                      <th className="text-left px-4 py-3 font-medium text-muted-foreground">#</th>
                      <th className="text-left px-4 py-3 font-medium text-muted-foreground">License Key</th>
                      <th className="text-left px-4 py-3 font-medium text-muted-foreground">Transaction Hash</th>
                      <th className="text-left px-4 py-3 font-medium text-muted-foreground">Issued At</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.licenses.map((lic, i) => (
                      <tr key={lic.licenseKey} className="border-b last:border-0 hover:bg-gray-50 transition-colors">
                        <td className="px-4 py-3 text-muted-foreground">{i + 1}</td>
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-1">
                            <code className="font-mono text-xs bg-gray-100 px-1.5 py-0.5 rounded">
                              {lic.licenseKey}
                            </code>
                            <CopyButton text={lic.licenseKey} />
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-1 max-w-xs">
                            <span className="font-mono text-xs text-muted-foreground truncate">
                              {lic.txHash}
                            </span>
                            <CopyButton text={lic.txHash} />
                          </div>
                        </td>
                        <td className="px-4 py-3 text-muted-foreground whitespace-nowrap">
                          {formatDate(lic.issuedAt)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

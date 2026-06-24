const LICENSE_KEY = "bsa_excel_pro_license";
const PRODUCT_ID  = "excel_addin_pro";
const API_BASE    = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");

function apiUrl(path: string) {
  return `${API_BASE}/api/payments${path}`;
}

// ── Local license storage ─────────────────────────────────────────────────────

export function getLicense(): string | null {
  try { return localStorage.getItem(LICENSE_KEY); } catch { return null; }
}

export function setLicense(key: string): void {
  try { localStorage.setItem(LICENSE_KEY, key); } catch { /* ignore */ }
}

export function clearLicense(): void {
  try { localStorage.removeItem(LICENSE_KEY); } catch { /* ignore */ }
}

// ── Types ─────────────────────────────────────────────────────────────────────

export interface Plan {
  id: string;
  label: string;
  price: number;
  days: number;
}

export interface PaymentConfig {
  address: string;
  network: string;
  plans: Plan[];
}

export interface SubscriptionStatus {
  valid: boolean;
  planId: string | null;
  planLabel: string | null;
  expiresAt: string | null;
  issuedAt: string | null;
  daysLeft: number | null;
  daysTotal: number | null;
  reason?: string;
}

// ── API helpers ───────────────────────────────────────────────────────────────

export async function fetchPaymentConfig(): Promise<PaymentConfig | null> {
  try {
    const res = await fetch(apiUrl("/config"));
    if (!res.ok) return null;
    return res.json();
  } catch {
    return null;
  }
}

export async function verifyPayment(
  txHash: string,
  planId: string,
  email?: string,
): Promise<{ success: boolean; licenseKey?: string; expiresAt?: string; planLabel?: string; error?: string }> {
  try {
    const res = await fetch(apiUrl("/verify"), {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ txHash, planId, productId: PRODUCT_ID, ...(email ? { email } : {}) }),
    });
    const data = await res.json();
    if (res.ok && data.licenseKey) {
      return { success: true, licenseKey: data.licenseKey, expiresAt: data.expiresAt, planLabel: data.planLabel };
    }
    return { success: false, error: data.error || "Verification failed" };
  } catch {
    return { success: false, error: "Network error — check your connection." };
  }
}

export async function fetchSubscription(licenseKey: string): Promise<SubscriptionStatus> {
  try {
    const res = await fetch(apiUrl(`/check/${encodeURIComponent(licenseKey)}`));
    if (!res.ok) return { valid: false, planId: null, planLabel: null, expiresAt: null, issuedAt: null, daysLeft: null, daysTotal: null };
    const data = await res.json();

    if (!data.valid) {
      return { valid: false, planId: data.planId ?? null, planLabel: null, expiresAt: data.expiresAt ?? null, issuedAt: null, daysLeft: null, daysTotal: null, reason: data.reason };
    }

    const expiresAt  = data.expiresAt  ?? null;
    const issuedAt   = data.issuedAt   ?? null;
    const daysLeft   = expiresAt ? Math.max(0, Math.ceil((new Date(expiresAt).getTime() - Date.now()) / 86400000)) : null;
    const daysTotal  = (expiresAt && issuedAt)
      ? Math.ceil((new Date(expiresAt).getTime() - new Date(issuedAt).getTime()) / 86400000)
      : null;

    return { valid: true, planId: data.planId, planLabel: data.planLabel, expiresAt, issuedAt, daysLeft, daysTotal };
  } catch {
    return { valid: false, planId: null, planLabel: null, expiresAt: null, issuedAt: null, daysLeft: null, daysTotal: null };
  }
}

export async function checkLicenseValid(licenseKey: string): Promise<boolean> {
  const sub = await fetchSubscription(licenseKey);
  return sub.valid;
}

export async function activateLicenseKey(key: string): Promise<{ success: boolean; error?: string }> {
  const trimmed = key.trim().toUpperCase();
  if (!trimmed) return { success: false, error: "Please enter a license key." };
  const sub = await fetchSubscription(trimmed);
  if (sub.valid) {
    setLicense(trimmed);
    return { success: true };
  }
  if (sub.reason === "expired") return { success: false, error: "This license key has expired. Please renew your subscription." };
  return { success: false, error: "Invalid license key. Check the key and try again." };
}

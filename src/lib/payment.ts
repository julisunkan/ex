const LICENSE_KEY = "bsa_excel_pro_license";
const PRODUCT_PRICE_USDT = 5;
const PRODUCT_ID = "excel_addin_pro";

export function getLicense(): string | null {
  try {
    return localStorage.getItem(LICENSE_KEY);
  } catch {
    return null;
  }
}

export function setLicense(key: string): void {
  try {
    localStorage.setItem(LICENSE_KEY, key);
  } catch {
    // ignore in environments without localStorage (Office desktop)
  }
}

export function clearLicense(): void {
  try {
    localStorage.removeItem(LICENSE_KEY);
  } catch {
    // ignore
  }
}

export const PRICE_USDT = PRODUCT_PRICE_USDT;

export async function fetchAdminWallet(): Promise<{ address: string; network: string; price: number } | null> {
  try {
    const res = await fetch("/api/payments/config");
    if (!res.ok) return null;
    return res.json();
  } catch {
    return null;
  }
}

export async function verifyPayment(txHash: string): Promise<{ success: boolean; licenseKey?: string; error?: string }> {
  try {
    const res = await fetch("/api/payments/verify", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ txHash, productId: PRODUCT_ID }),
    });
    const data = await res.json();
    if (res.ok && data.licenseKey) {
      return { success: true, licenseKey: data.licenseKey };
    }
    return { success: false, error: data.error || "Verification failed" };
  } catch {
    return { success: false, error: "Network error — check your connection." };
  }
}

export async function checkLicenseValid(licenseKey: string): Promise<boolean> {
  try {
    const res = await fetch(`/api/payments/check/${encodeURIComponent(licenseKey)}`);
    if (!res.ok) return false;
    const data = await res.json();
    return data.valid === true;
  } catch {
    return false;
  }
}

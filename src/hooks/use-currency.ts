import { useState, useCallback } from "react";

export const CURRENCIES = [
  { symbol: "$",   label: "USD – US Dollar" },
  { symbol: "€",   label: "EUR – Euro" },
  { symbol: "£",   label: "GBP – British Pound" },
  { symbol: "₦",   label: "NGN – Nigerian Naira" },
  { symbol: "₹",   label: "INR – Indian Rupee" },
  { symbol: "¥",   label: "JPY – Japanese Yen" },
  { symbol: "₩",   label: "KRW – South Korean Won" },
  { symbol: "R",   label: "ZAR – South African Rand" },
  { symbol: "KSh", label: "KES – Kenyan Shilling" },
  { symbol: "₵",   label: "GHS – Ghanaian Cedi" },
  { symbol: "₱",   label: "PHP – Philippine Peso" },
  { symbol: "฿",   label: "THB – Thai Baht" },
  { symbol: "د.إ", label: "AED – UAE Dirham" },
  { symbol: "CHF", label: "CHF – Swiss Franc" },
  { symbol: "C$",  label: "CAD – Canadian Dollar" },
  { symbol: "A$",  label: "AUD – Australian Dollar" },
];

const STORAGE_KEY = "bsa_currency_symbol";

function getStored(): string {
  try { return localStorage.getItem(STORAGE_KEY) || "₦"; } catch { return "₦"; }
}

export function useCurrency() {
  const [symbol, setSymbolState] = useState<string>(getStored);

  const setSymbol = useCallback((s: string) => {
    try { localStorage.setItem(STORAGE_KEY, s); } catch {}
    setSymbolState(s);
  }, []);

  const fmt = useCallback((n: number): string => {
    const abs = Math.abs(n);
    return `${symbol}${abs.toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
  }, [symbol]);

  const fmtShort = useCallback((n: number): string => {
    const abs = Math.abs(n);
    if (abs >= 1_000_000) return `${symbol}${(abs / 1_000_000).toFixed(1)}M`;
    if (abs >= 1_000) return `${symbol}${(abs / 1_000).toFixed(0)}K`;
    return `${symbol}${abs.toFixed(0)}`;
  }, [symbol]);

  return { symbol, setSymbol, fmt, fmtShort };
}

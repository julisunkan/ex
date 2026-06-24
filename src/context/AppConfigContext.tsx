import { createContext, useContext, useEffect, useState, ReactNode } from "react";

const API_BASE = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");

export interface Plan {
  id: string;
  label: string;
  price: number;
  days: number;
}

export interface AppConfig {
  appearance: {
    name: string;
    tagline: string;
    primaryColor: string;
    accentColor: string;
    radius: string;
  };
  features: { proEnabled: boolean };
  payment: { network: string };
  plans: Plan[];
}

const DEFAULT_PLANS: Plan[] = [
  { id: "monthly",   label: "Monthly",  price: 5,  days: 30  },
  { id: "quarterly", label: "3-Month",  price: 12, days: 90  },
  { id: "biannual",  label: "6-Month",  price: 20, days: 180 },
  { id: "annual",    label: "1-Year",   price: 35, days: 365 },
];

const DEFAULT_CONFIG: AppConfig = {
  appearance: {
    name: "Bank Statement Analyzer",
    tagline: "Analyze transactions, categorize spending, and export summary reports.",
    primaryColor: "#3b82f6",
    accentColor: "#16a34a",
    radius: "6px",
  },
  features: { proEnabled: true },
  payment: { network: "tron" },
  plans: DEFAULT_PLANS,
};

function hexToHsl(hex: string): string {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
      case g: h = ((b - r) / d + 2) / 6; break;
      case b: h = ((r - g) / d + 4) / 6; break;
    }
  }
  return `${Math.round(h * 360)} ${Math.round(s * 100)}% ${Math.round(l * 100)}%`;
}

function applyTheme(config: AppConfig) {
  const { primaryColor, accentColor, radius } = config.appearance;
  const root = document.documentElement;
  root.style.setProperty("--primary", hexToHsl(primaryColor));
  root.style.setProperty("--ring", hexToHsl(primaryColor));
  root.style.setProperty("--accent", hexToHsl(accentColor));
  root.style.setProperty("--success", hexToHsl(accentColor));
  root.style.setProperty("--radius", radius);
  document.title = config.appearance.name;
}

const AppConfigContext = createContext<AppConfig>(DEFAULT_CONFIG);

export function AppConfigProvider({ children }: { children: ReactNode }) {
  const [config, setConfig] = useState<AppConfig>(DEFAULT_CONFIG);

  useEffect(() => {
    fetch(`${API_BASE}/api/config`)
      .then((r) => r.ok ? r.json() : null)
      .then((data) => {
        if (!data) return;
        const merged: AppConfig = {
          appearance: { ...DEFAULT_CONFIG.appearance, ...data.appearance },
          features:   { ...DEFAULT_CONFIG.features,   ...data.features   },
          payment:    { ...DEFAULT_CONFIG.payment,    ...data.payment    },
          plans:      Array.isArray(data.plans) && data.plans.length ? data.plans : DEFAULT_PLANS,
        };
        setConfig(merged);
        applyTheme(merged);
      })
      .catch(() => {});
  }, []);

  return (
    <AppConfigContext.Provider value={config}>
      {children}
    </AppConfigContext.Provider>
  );
}

export function useAppConfig() {
  return useContext(AppConfigContext);
}

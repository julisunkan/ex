import { Router } from "express";
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { notifyNewLicense } from "../lib/notify.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SETTINGS_FILE = join(__dirname, "../data/settings.json");
const LICENSES_FILE = join(__dirname, "../data/licenses.json");

const router = Router();

const DEFAULT_PLANS = [
  { id: "monthly",   label: "Monthly",  price: 5,  days: 30  },
  { id: "quarterly", label: "3-Month",  price: 12, days: 90  },
  { id: "biannual",  label: "6-Month",  price: 20, days: 180 },
  { id: "annual",    label: "1-Year",   price: 35, days: 365 },
];

const DEFAULT_SETTINGS = {
  appearance: {
    name: "Bank Statement Analyzer",
    tagline: "Analyze transactions, categorize spending, and export summary reports.",
    primaryColor: "#3b82f6",
    accentColor: "#16a34a",
    radius: "6px",
  },
  payment: { walletAddress: "", network: "tron" },
  plans: DEFAULT_PLANS,
  notifications: {
    webhookUrl: "",
    remindersEnabled: false,
    reminderDays: 3,
    email: { enabled: false, to: "", smtpHost: "", smtpPort: 587, smtpUser: "", smtpPass: "", from: "" },
  },
  features: { proEnabled: true },
};

function loadSettings() {
  try {
    const raw = JSON.parse(readFileSync(SETTINGS_FILE, "utf8"));
    const rawEmail = raw.notifications?.email || {};
    return {
      ...DEFAULT_SETTINGS,
      ...raw,
      appearance:    { ...DEFAULT_SETTINGS.appearance,              ...(raw.appearance    || {}) },
      payment:       { ...DEFAULT_SETTINGS.payment,                 ...(raw.payment       || {}) },
      plans:         Array.isArray(raw.plans) && raw.plans.length === 4 ? raw.plans : DEFAULT_PLANS,
      notifications: {
        webhookUrl:       raw.notifications?.webhookUrl       ?? "",
        remindersEnabled: raw.notifications?.remindersEnabled ?? false,
        reminderDays:     raw.notifications?.reminderDays     ?? 3,
        email: { ...DEFAULT_SETTINGS.notifications.email, ...rawEmail },
      },
      features:      { ...DEFAULT_SETTINGS.features,                ...(raw.features      || {}) },
    };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

function saveSettings(s) {
  writeFileSync(SETTINGS_FILE, JSON.stringify(s, null, 2));
}

function loadLicenses() {
  try { return JSON.parse(readFileSync(LICENSES_FILE, "utf8")); } catch { return []; }
}

function saveLicenses(l) {
  writeFileSync(LICENSES_FILE, JSON.stringify(l, null, 2));
}

function requireAdmin(req, res, next) {
  const password = process.env.ADMIN_PASSWORD;
  if (!password) return res.status(503).json({ error: "ADMIN_PASSWORD not configured" });
  if (req.headers["x-admin-password"] !== password) return res.status(401).json({ error: "Unauthorized" });
  next();
}

// ── Public ────────────────────────────────────────────────────────────────────

// GET /api/config  (no auth — frontend loads this on startup)
router.get("/", (req, res) => {
  const s = loadSettings();
  res.json({
    appearance: s.appearance,
    features:   s.features,
    payment:    { network: s.payment.network },
    plans:      s.plans,
  });
});

// ── Admin ─────────────────────────────────────────────────────────────────────

// GET /api/admin/settings
router.get("/settings", requireAdmin, (req, res) => {
  res.json(loadSettings());
});

// PUT /api/admin/settings  (deep-merge patch)
router.put("/settings", requireAdmin, (req, res) => {
  const current = loadSettings();
  const patch = req.body || {};

  const patchEmail = patch.notifications?.email || {};
  const merged = {
    appearance:    { ...current.appearance,    ...(patch.appearance    || {}) },
    payment:       { ...current.payment,       ...(patch.payment       || {}) },
    plans:         Array.isArray(patch.plans) ? patch.plans : current.plans,
    notifications: {
      webhookUrl:       patch.notifications?.webhookUrl       ?? current.notifications?.webhookUrl       ?? "",
      remindersEnabled: patch.notifications?.remindersEnabled ?? current.notifications?.remindersEnabled ?? false,
      reminderDays:     patch.notifications?.reminderDays     ?? current.notifications?.reminderDays     ?? 3,
      email: { ...current.notifications?.email, ...patchEmail },
    },
    features:      { ...current.features,      ...(patch.features      || {}) },
  };

  if (!merged.payment.walletAddress) {
    merged.payment.walletAddress = process.env.USDT_WALLET_ADDRESS || "";
  }

  saveSettings(merged);
  res.json({ ok: true, settings: merged });
});

// POST /api/admin/notify-test  — fire a test notification
router.post("/notify-test", requireAdmin, async (req, res) => {
  const cfg = loadSettings().notifications ?? {};
  if (!cfg.webhookUrl && !cfg.email?.enabled) {
    return res.status(400).json({ error: "No notification channels configured. Add a webhook URL or enable email first." });
  }
  try {
    await notifyNewLicense({
      licenseKey:  "BSA-TEST-LICENSE",
      planLabel:   "Monthly",
      planId:      "monthly",
      expiresAt:   new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
      txHash:      "0xTEST_TRANSACTION_HASH",
      network:     "tron",
    });
    res.json({ ok: true, message: "Test notification sent! Check your webhook / inbox." });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET /api/admin/export  — full backup
router.get("/export", requireAdmin, (req, res) => {
  const backup = {
    exportedAt: new Date().toISOString(),
    settings: loadSettings(),
    licenses: loadLicenses(),
  };
  res.setHeader("Content-Disposition", `attachment; filename="bsa-backup-${Date.now()}.json"`);
  res.setHeader("Content-Type", "application/json");
  res.json(backup);
});

// POST /api/admin/import  — restore from backup
router.post("/import", requireAdmin, (req, res) => {
  const { settings, licenses } = req.body || {};
  if (settings) saveSettings(settings);
  if (Array.isArray(licenses)) saveLicenses(licenses);
  res.json({ ok: true });
});

export default router;

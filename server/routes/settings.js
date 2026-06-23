import { Router } from "express";
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SETTINGS_FILE = join(__dirname, "../data/settings.json");
const LICENSES_FILE = join(__dirname, "../data/licenses.json");

const router = Router();

function loadSettings() {
  try {
    return JSON.parse(readFileSync(SETTINGS_FILE, "utf8"));
  } catch {
    return {
      appearance: {
        name: "Bank Statement Analyzer",
        tagline: "Analyze transactions, categorize spending, and export summary reports.",
        primaryColor: "#3b82f6",
        accentColor: "#16a34a",
        radius: "6px",
      },
      payment: { walletAddress: "", network: "tron", price: 5 },
      features: { proEnabled: true },
    };
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

// GET /api/config  (no auth — used by frontend on every load)
router.get("/", (req, res) => {
  const s = loadSettings();
  res.json({
    appearance: s.appearance,
    features: s.features,
    payment: {
      price: s.payment.price,
      network: s.payment.network,
    },
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

  const merged = {
    appearance: { ...current.appearance, ...(patch.appearance || {}) },
    payment:    { ...current.payment,    ...(patch.payment    || {}) },
    features:   { ...current.features,  ...(patch.features   || {}) },
  };

  // If walletAddress is being cleared to empty, fall back to env var
  if (!merged.payment.walletAddress) {
    merged.payment.walletAddress = process.env.USDT_WALLET_ADDRESS || "";
  }

  saveSettings(merged);
  res.json({ ok: true, settings: merged });
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

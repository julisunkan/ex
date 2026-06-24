import { Router } from "express";
import { readFileSync, writeFileSync } from "fs";
import { randomBytes } from "crypto";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const LICENSES_FILE = join(__dirname, "../data/licenses.json");

const router = Router();

function requireAdmin(req, res, next) {
  const password = process.env.ADMIN_PASSWORD;
  if (!password) {
    return res.status(503).json({ error: "ADMIN_PASSWORD not configured on server" });
  }
  const provided = req.headers["x-admin-password"];
  if (provided !== password) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

function loadLicenses() {
  try {
    return JSON.parse(readFileSync(LICENSES_FILE, "utf8"));
  } catch {
    return [];
  }
}

function saveLicenses(licenses) {
  writeFileSync(LICENSES_FILE, JSON.stringify(licenses, null, 2));
}

function generateLicenseKey() {
  return "BSA-" + randomBytes(12).toString("hex").toUpperCase();
}

// GET /api/admin/licenses
router.get("/licenses", requireAdmin, (req, res) => {
  const licenses = loadLicenses();
  res.json({
    total: licenses.length,
    licenses: licenses.sort(
      (a, b) => new Date(b.issuedAt).getTime() - new Date(a.issuedAt).getTime()
    ),
  });
});

// POST /api/admin/licenses/generate  — manually issue a license key
router.post("/licenses/generate", requireAdmin, (req, res) => {
  const { note } = req.body || {};
  const licenseKey = generateLicenseKey();
  const licenses = loadLicenses();
  licenses.push({
    licenseKey,
    txHash: "MANUAL",
    note: (note || "Admin generated").slice(0, 100),
    issuedAt: new Date().toISOString(),
  });
  saveLicenses(licenses);
  console.log(`🔑 Manual license issued: ${licenseKey}${note ? ` (${note})` : ""}`);
  res.json({ licenseKey });
});

// POST /api/admin/licenses/bulk-generate  — generate multiple keys at once
router.post("/licenses/bulk-generate", requireAdmin, (req, res) => {
  const { count = 1, note } = req.body || {};
  const n = Math.max(1, Math.min(100, parseInt(count) || 1));
  const licenses = loadLicenses();
  const now = new Date().toISOString();
  const newKeys = [];
  for (let i = 0; i < n; i++) {
    const licenseKey = generateLicenseKey();
    licenses.push({
      licenseKey,
      txHash: "MANUAL",
      note: (note || "Bulk generated").slice(0, 100),
      issuedAt: now,
    });
    newKeys.push(licenseKey);
  }
  saveLicenses(licenses);
  console.log(`🔑 Bulk issued ${n} license(s)${note ? ` (${note})` : ""}`);
  res.json({ keys: newKeys, count: newKeys.length });
});

// DELETE /api/admin/licenses/:key  — revoke a license key
router.delete("/licenses/:key", requireAdmin, (req, res) => {
  const { key } = req.params;
  const licenses = loadLicenses();
  const updated = licenses.filter((l) => l.licenseKey !== key);
  if (updated.length === licenses.length) {
    return res.status(404).json({ error: "License not found" });
  }
  saveLicenses(updated);
  console.log(`🗑  License revoked: ${key}`);
  res.json({ ok: true });
});

export default router;

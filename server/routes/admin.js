import { Router } from "express";
import { readFileSync } from "fs";
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

export default router;

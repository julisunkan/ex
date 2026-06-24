import { Router } from "express";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { sendReportEmail } from "../lib/notify.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SETTINGS_FILE = join(__dirname, "../data/settings.json");

const router = Router();

function loadSmtp() {
  try {
    const s = JSON.parse(readFileSync(SETTINGS_FILE, "utf8"));
    return s.notifications?.email ?? null;
  } catch {
    return null;
  }
}

// POST /api/send-report
// Body: { email: string, html: string, appName: string }
router.post("/", async (req, res) => {
  const { email, html, appName } = req.body || {};

  if (!email || typeof email !== "string" || !email.includes("@")) {
    return res.status(400).json({ error: "A valid email address is required." });
  }
  if (!html || typeof html !== "string" || html.length < 100) {
    return res.status(400).json({ error: "Report HTML is missing or too short." });
  }

  const smtpCfg = loadSmtp();
  if (!smtpCfg || !smtpCfg.enabled || !smtpCfg.smtpHost) {
    return res.status(503).json({
      error: "Email delivery is not configured. Ask the admin to set up SMTP in the settings.",
    });
  }

  try {
    await sendReportEmail(smtpCfg, email, html, appName || "Bank Statement Analyzer");
    console.log(`[send-report] Report emailed to ${email}`);
    res.json({ ok: true });
  } catch (err) {
    console.warn(`[send-report] Failed to send to ${email}:`, err.message);
    res.status(500).json({ error: `Failed to send email: ${err.message}` });
  }
});

export default router;

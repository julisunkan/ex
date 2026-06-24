import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { notifyExpiryReminder } from "./notify.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const LICENSES_FILE = join(__dirname, "../data/licenses.json");
const SETTINGS_FILE = join(__dirname, "../data/settings.json");

const DEFAULT_PLANS = [
  { id: "monthly",   label: "Monthly",  price: 5,  days: 30  },
  { id: "quarterly", label: "3-Month",  price: 12, days: 90  },
  { id: "biannual",  label: "6-Month",  price: 20, days: 180 },
  { id: "annual",    label: "1-Year",   price: 35, days: 365 },
];

function loadSettings() {
  try { return JSON.parse(readFileSync(SETTINGS_FILE, "utf8")); } catch { return {}; }
}

function loadLicenses() {
  try { return JSON.parse(readFileSync(LICENSES_FILE, "utf8")); } catch { return []; }
}

function saveLicenses(licenses) {
  writeFileSync(LICENSES_FILE, JSON.stringify(licenses, null, 2));
}

async function runExpiryCheck() {
  const settings = loadSettings();
  const notifCfg = settings.notifications ?? {};
  const reminderDays = notifCfg.reminderDays ?? 3;
  const enabled = notifCfg.remindersEnabled ?? false;

  if (!enabled) return;

  const smtpCfg = notifCfg.email ?? {};
  if (!smtpCfg.enabled || !smtpCfg.smtpHost) {
    console.log("[expiry-checker] Skipping — SMTP not configured.");
    return;
  }

  const plans = Array.isArray(settings.plans) ? settings.plans : DEFAULT_PLANS;
  const licenses = loadLicenses();
  const now = Date.now();
  let updated = false;

  for (const license of licenses) {
    if (!license.expiresAt || !license.email) continue;
    if (license.reminderSent) continue;

    const msLeft = new Date(license.expiresAt).getTime() - now;
    const daysLeft = Math.ceil(msLeft / (1000 * 60 * 60 * 24));

    if (daysLeft > 0 && daysLeft <= reminderDays) {
      const plan = plans.find((p) => p.id === license.planId);
      const planLabel = plan?.label ?? (license.planId ?? "Pro");
      console.log(`[expiry-checker] Sending reminder to ${license.email} — ${daysLeft} days left (${license.licenseKey})`);
      try {
        await notifyExpiryReminder({
          licenseKey: license.licenseKey,
          planLabel,
          expiresAt: license.expiresAt,
          daysLeft,
          userEmail: license.email,
          smtpCfg,
        });
        license.reminderSent = true;
        updated = true;
      } catch (err) {
        console.warn("[expiry-checker] Failed to send reminder:", err.message);
      }
    }
  }

  if (updated) saveLicenses(licenses);
  console.log(`[expiry-checker] Check complete. ${licenses.length} licenses scanned.`);
}

export function startExpiryChecker() {
  const INTERVAL_MS = 6 * 60 * 60 * 1000; // every 6 hours
  console.log("[expiry-checker] Started — checks every 6 hours.");
  runExpiryCheck().catch((e) => console.warn("[expiry-checker] Initial check error:", e.message));
  setInterval(() => {
    runExpiryCheck().catch((e) => console.warn("[expiry-checker] Check error:", e.message));
  }, INTERVAL_MS);
}

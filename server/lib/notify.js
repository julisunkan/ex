import nodemailer from "nodemailer";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SETTINGS_FILE = join(__dirname, "../data/settings.json");

function getNotifyConfig() {
  try {
    const s = JSON.parse(readFileSync(SETTINGS_FILE, "utf8"));
    return s.notifications ?? {};
  } catch {
    return {};
  }
}

async function sendWebhook(url, payload) {
  if (!url) return;
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) console.warn(`[notify] Webhook responded ${res.status}`);
    else console.log("[notify] Webhook sent OK");
  } catch (err) {
    console.warn("[notify] Webhook failed:", err.message);
  }
}

async function sendEmail(cfg, subject, text) {
  if (!cfg?.enabled || !cfg?.smtpHost || !cfg?.to) return;
  try {
    const transporter = nodemailer.createTransport({
      host: cfg.smtpHost,
      port: cfg.smtpPort || 587,
      secure: (cfg.smtpPort || 587) === 465,
      auth: cfg.smtpUser ? { user: cfg.smtpUser, pass: cfg.smtpPass } : undefined,
    });
    await transporter.sendMail({
      from: cfg.from || cfg.smtpUser || "noreply@bankstatementanalyzer.app",
      to: cfg.to,
      subject,
      text,
    });
    console.log("[notify] Email sent to", cfg.to);
  } catch (err) {
    console.warn("[notify] Email failed:", err.message);
  }
}

export async function notifyNewLicense({ licenseKey, planLabel, planId, expiresAt, txHash, network }) {
  const cfg = getNotifyConfig();
  if (!cfg.webhookUrl && !cfg.email?.enabled) return;

  const lines = [
    `New license activated`,
    `Plan: ${planLabel} (${planId})`,
    `License: ${licenseKey}`,
    `Expires: ${expiresAt ? new Date(expiresAt).toLocaleString() : "Never"}`,
    `Network: ${network?.toUpperCase()}`,
    `TX Hash: ${txHash}`,
    `Time: ${new Date().toLocaleString()}`,
  ];
  const text = lines.join("\n");

  const webhookPayload = {
    text,
    embeds: [{
      title: "💰 New Subscription Activated",
      color: 0x16a34a,
      fields: [
        { name: "Plan",    value: `${planLabel} (${planId})`,                  inline: true },
        { name: "License", value: `\`${licenseKey}\``,                          inline: true },
        { name: "Expires", value: expiresAt ? new Date(expiresAt).toLocaleDateString() : "Never", inline: true },
        { name: "Network", value: network?.toUpperCase() || "-",                inline: true },
        { name: "TX Hash", value: `\`${txHash}\``,                              inline: false },
      ],
      timestamp: new Date().toISOString(),
    }],
  };

  await Promise.allSettled([
    sendWebhook(cfg.webhookUrl, webhookPayload),
    sendEmail(cfg.email, `New ${planLabel} subscription activated`, text),
  ]);
}

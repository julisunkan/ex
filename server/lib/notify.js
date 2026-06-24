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

function createTransporter(cfg) {
  return nodemailer.createTransport({
    host: cfg.smtpHost,
    port: cfg.smtpPort || 587,
    secure: (cfg.smtpPort || 587) === 465,
    auth: cfg.smtpUser ? { user: cfg.smtpUser, pass: cfg.smtpPass } : undefined,
  });
}

async function sendEmail(cfg, subject, text, overrideTo) {
  const recipient = overrideTo || cfg?.to;
  if (!cfg?.smtpHost || !recipient) return;
  try {
    const transporter = createTransporter(cfg);
    await transporter.sendMail({
      from: cfg.from || cfg.smtpUser || "noreply@bankstatementanalyzer.app",
      to: recipient,
      subject,
      text,
    });
    console.log("[notify] Email sent to", recipient);
  } catch (err) {
    console.warn("[notify] Email failed:", err.message);
  }
}

export async function sendReportEmail(cfg, to, html, appName) {
  if (!cfg?.smtpHost || !to) throw new Error("SMTP not configured or recipient missing.");
  const transporter = createTransporter(cfg);
  const from = cfg.from || cfg.smtpUser || "noreply@bankstatementanalyzer.app";
  const subject = `${appName} — Your Bank Statement Report`;
  await transporter.sendMail({
    from,
    to,
    subject,
    html,
    text: `Your bank statement analysis report from ${appName} is attached as HTML. Please open this email in a browser or HTML-capable client to view the full report.`,
  });
}

export async function notifyExpiryReminder({ licenseKey, planLabel, expiresAt, daysLeft, userEmail, smtpCfg }) {
  if (!userEmail || !smtpCfg?.enabled || !smtpCfg?.smtpHost) return;
  const expDate = expiresAt ? new Date(expiresAt).toLocaleDateString() : "soon";
  const subject = `Your ${planLabel} subscription expires in ${daysLeft} day${daysLeft === 1 ? "" : "s"}`;
  const text = [
    `Hi,`,
    ``,
    `Your Bank Statement Analyzer ${planLabel} subscription is expiring in ${daysLeft} day${daysLeft === 1 ? "" : "s"} (${expDate}).`,
    ``,
    `To continue using Pro features, open the add-in and renew by paying for a new subscription.`,
    `Your license key: ${licenseKey}`,
    ``,
    `Thank you for using Bank Statement Analyzer!`,
  ].join("\n");
  await sendEmail(smtpCfg, subject, text, userEmail);
}

export async function notifyNewTicket({ id, name, email, licenseKey, category, subject, message }) {
  const cfg = getNotifyConfig();
  if (!cfg.webhookUrl && !cfg.email?.enabled) return;

  const categoryLabel =
    { general: "General", billing: "Billing", license: "License", bug: "Bug", feature: "Feature", other: "Other" }[category] ?? category;

  const lines = [
    `New support ticket received`,
    `Ticket ID: ${id}`,
    `From: ${name} <${email}>`,
    ...(licenseKey ? [`License: ${licenseKey}`] : []),
    `Category: ${categoryLabel}`,
    `Subject: ${subject}`,
    ``,
    message,
    ``,
    `Time: ${new Date().toLocaleString()}`,
  ];
  const text = lines.join("\n");

  const webhookPayload = {
    text,
    embeds: [{
      title: `🎫 New Support Ticket: ${subject}`,
      color: 0xf59e0b,
      fields: [
        { name: "Ticket ID", value: id,           inline: true },
        { name: "Category",  value: categoryLabel, inline: true },
        { name: "From",      value: `${name}\n${email}`, inline: false },
        ...(licenseKey ? [{ name: "License", value: `\`${licenseKey}\``, inline: false }] : []),
        { name: "Message",   value: message.length > 300 ? message.slice(0, 300) + "…" : message, inline: false },
      ],
      timestamp: new Date().toISOString(),
    }],
  };

  await Promise.allSettled([
    sendWebhook(cfg.webhookUrl, webhookPayload),
    sendEmail(cfg.email, `[Support Ticket ${id}] ${subject}`, text),
  ]);
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

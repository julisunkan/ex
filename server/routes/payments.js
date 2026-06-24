import { Router } from "express";
import { randomBytes } from "crypto";
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { notifyNewLicense } from "../lib/notify.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const LICENSES_FILE = join(__dirname, "../data/licenses.json");
const SETTINGS_FILE = join(__dirname, "../data/settings.json");

const router = Router();

const DEFAULT_PLANS = [
  { id: "monthly",   label: "Monthly",  price: 5,  days: 30  },
  { id: "quarterly", label: "3-Month",  price: 12, days: 90  },
  { id: "biannual",  label: "6-Month",  price: 20, days: 180 },
  { id: "annual",    label: "1-Year",   price: 35, days: 365 },
];

const TRONGRID_API_KEY  = process.env.TRONGRID_API_KEY  || "";
const BSCSCAN_API_KEY   = process.env.BSCSCAN_API_KEY   || "";
const ETHERSCAN_API_KEY = process.env.ETHERSCAN_API_KEY || "";

const USDT_CONTRACTS = {
  tron: "TR7NHqjeKQxGTCi8q8ZY4pL8otSzgjLj6t",
  bsc:  "0x55d398326f99059fF775485246999027B3197955",
  eth:  "0xdAC17F958D2ee523a2206206994597C13D831ec7",
};

// ── Config ───────────────────────────────────────────────────────────────────

function getPaymentConfig() {
  try {
    const s = JSON.parse(readFileSync(SETTINGS_FILE, "utf8"));
    const plans = Array.isArray(s.plans) && s.plans.length ? s.plans : DEFAULT_PLANS;
    return {
      walletAddress: s.payment?.walletAddress || process.env.USDT_WALLET_ADDRESS || "",
      network: (s.payment?.network || process.env.USDT_NETWORK || "tron").toLowerCase(),
      plans,
    };
  } catch {
    return {
      walletAddress: process.env.USDT_WALLET_ADDRESS || "",
      network: (process.env.USDT_NETWORK || "tron").toLowerCase(),
      plans: DEFAULT_PLANS,
    };
  }
}

// ── License store ─────────────────────────────────────────────────────────────

function loadLicenses() {
  try { return JSON.parse(readFileSync(LICENSES_FILE, "utf8")); } catch { return []; }
}

function saveLicenses(licenses) {
  writeFileSync(LICENSES_FILE, JSON.stringify(licenses, null, 2));
}

function generateLicenseKey() {
  return "BSA-" + randomBytes(12).toString("hex").toUpperCase();
}

function txAlreadyUsed(txHash) {
  return loadLicenses().some((l) => l.txHash === txHash);
}

function saveLicense(licenseKey, txHash, planId, expiresAt, email) {
  const licenses = loadLicenses();
  licenses.push({
    licenseKey,
    txHash,
    planId:    planId    ?? null,
    expiresAt: expiresAt ?? null,
    issuedAt:  new Date().toISOString(),
    email:     email     ?? null,
  });
  saveLicenses(licenses);
}

// ── Blockchain verification ───────────────────────────────────────────────────

async function verifyTron(txHash, walletAddress, priceUsdt) {
  const headers = { Accept: "application/json" };
  if (TRONGRID_API_KEY) headers["TRON-PRO-API-KEY"] = TRONGRID_API_KEY;

  const res = await fetch(`https://api.trongrid.io/v1/transactions/${txHash}`, { headers });
  if (!res.ok) return { ok: false, reason: "TronGrid request failed" };

  const json = await res.json();
  const tx = json?.data?.[0];
  if (!tx) return { ok: false, reason: "Transaction not found" };

  const receipt = tx?.ret?.[0];
  if (receipt?.contractRet !== "SUCCESS") return { ok: false, reason: "Transaction not successful" };

  const data = tx?.raw_data?.contract?.[0]?.parameter?.value?.data || "";
  if (data.startsWith("a9059cbb")) {
    const to     = "41" + data.slice(32, 72);
    const amount = parseInt(data.slice(72, 136), 16) / 1e6;
    const ourHex = walletAddress;
    if (to.toLowerCase() !== ourHex.toLowerCase())
      return { ok: false, reason: "Wrong destination wallet" };
    if (amount < priceUsdt)
      return { ok: false, reason: `Amount too low: ${amount} USDT (need ${priceUsdt})` };
    return { ok: true };
  }
  return { ok: false, reason: "Not a TRC-20 USDT transfer" };
}

async function verifyBsc(txHash, walletAddress, priceUsdt) {
  if (!BSCSCAN_API_KEY) return { ok: false, reason: "BSCSCAN_API_KEY not configured" };
  const url = `https://api.bscscan.com/api?module=account&action=tokentx&contractaddress=${USDT_CONTRACTS.bsc}&address=${walletAddress}&apikey=${BSCSCAN_API_KEY}`;
  const res  = await fetch(url);
  const json = await res.json();
  if (json.status !== "1") return { ok: false, reason: "BSCScan query failed" };
  const tx = json.result?.find((t) => t.hash.toLowerCase() === txHash.toLowerCase());
  if (!tx) return { ok: false, reason: "Transaction not found" };
  const amount = Number(tx.value) / 10 ** Number(tx.tokenDecimal);
  if (amount < priceUsdt) return { ok: false, reason: `Amount too low: ${amount} USDT` };
  return { ok: true };
}

async function verifyEth(txHash, walletAddress, priceUsdt) {
  if (!ETHERSCAN_API_KEY) return { ok: false, reason: "ETHERSCAN_API_KEY not configured" };
  const url = `https://api.etherscan.io/api?module=account&action=tokentx&contractaddress=${USDT_CONTRACTS.eth}&address=${walletAddress}&apikey=${ETHERSCAN_API_KEY}`;
  const res  = await fetch(url);
  const json = await res.json();
  if (json.status !== "1") return { ok: false, reason: "Etherscan query failed" };
  const tx = json.result?.find((t) => t.hash.toLowerCase() === txHash.toLowerCase());
  if (!tx) return { ok: false, reason: "Transaction not found" };
  const amount = Number(tx.value) / 10 ** Number(tx.tokenDecimal);
  if (amount < priceUsdt) return { ok: false, reason: `Amount too low: ${amount} USDT` };
  return { ok: true };
}

// ── Routes ────────────────────────────────────────────────────────────────────

// GET /api/payments/config
router.get("/config", (req, res) => {
  const cfg = getPaymentConfig();
  if (!cfg.walletAddress) {
    return res.status(503).json({ error: "Wallet not configured. Set USDT_WALLET_ADDRESS in Admin > Payment." });
  }
  res.json({ address: cfg.walletAddress, network: cfg.network, plans: cfg.plans });
});

// POST /api/payments/verify
router.post("/verify", async (req, res) => {
  const { txHash, planId, email } = req.body || {};
  if (!txHash) return res.status(400).json({ error: "txHash is required" });

  if (txAlreadyUsed(txHash)) {
    return res.status(400).json({ error: "Transaction already redeemed" });
  }

  const cfg  = getPaymentConfig();
  const plan = cfg.plans.find((p) => p.id === planId) || cfg.plans[0];

  let result;
  try {
    if (cfg.network === "tron") result = await verifyTron(txHash, cfg.walletAddress, plan.price);
    else if (cfg.network === "bsc") result = await verifyBsc(txHash, cfg.walletAddress, plan.price);
    else if (cfg.network === "eth") result = await verifyEth(txHash, cfg.walletAddress, plan.price);
    else result = { ok: false, reason: `Unknown network: ${cfg.network}` };
  } catch (err) {
    console.error("Verification error:", err);
    return res.status(502).json({ error: "Blockchain lookup failed — try again shortly" });
  }

  if (!result.ok) {
    return res.status(402).json({ error: result.reason });
  }

  const licenseKey = generateLicenseKey();
  const expiresAt  = new Date(Date.now() + plan.days * 24 * 60 * 60 * 1000).toISOString();
  const cleanEmail = typeof email === "string" && email.includes("@") ? email.trim().toLowerCase() : null;
  saveLicense(licenseKey, txHash, plan.id, expiresAt, cleanEmail);
  console.log(`✅ License issued: ${licenseKey} | plan: ${plan.id} | expires: ${expiresAt}`);

  notifyNewLicense({ licenseKey, planLabel: plan.label, planId: plan.id, expiresAt, txHash, network: cfg.network })
    .catch((err) => console.warn("[notify] fire-and-forget error:", err.message));

  res.json({ licenseKey, expiresAt, planId: plan.id, planLabel: plan.label });
});

// GET /api/payments/check/:key
router.get("/check/:key", (req, res) => {
  const { key } = req.params;
  const licenses = loadLicenses();
  const license  = licenses.find((l) => l.licenseKey === key);
  if (!license) return res.json({ valid: false, reason: "not_found" });

  if (license.expiresAt && new Date(license.expiresAt) < new Date()) {
    return res.json({ valid: false, reason: "expired", expiresAt: license.expiresAt, planId: license.planId });
  }

  // Resolve plan label from current settings
  const cfg   = getPaymentConfig();
  const plan  = cfg.plans.find((p) => p.id === license.planId);
  const planLabel = plan?.label ?? (license.planId ? license.planId : "Pro");

  res.json({
    valid:      true,
    planId:     license.planId     ?? null,
    planLabel,
    expiresAt:  license.expiresAt  ?? null,
    issuedAt:   license.issuedAt   ?? null,
  });
});

export default router;

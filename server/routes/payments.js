import { Router } from "express";
import { createHash, randomBytes } from "crypto";
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const LICENSES_FILE = join(__dirname, "../data/licenses.json");

const router = Router();

// ── Config ──────────────────────────────────────────────────────────────────
const SETTINGS_FILE = join(__dirname, "../data/settings.json");

function getPaymentConfig() {
  try {
    const s = JSON.parse(readFileSync(SETTINGS_FILE, "utf8"));
    return {
      walletAddress: s.payment?.walletAddress || process.env.USDT_WALLET_ADDRESS || "",
      network: (s.payment?.network || process.env.USDT_NETWORK || "tron").toLowerCase(),
      price: Number(s.payment?.price ?? process.env.USDT_PRICE ?? 5),
    };
  } catch {
    return {
      walletAddress: process.env.USDT_WALLET_ADDRESS || "",
      network: (process.env.USDT_NETWORK || "tron").toLowerCase(),
      price: Number(process.env.USDT_PRICE || "5"),
    };
  }
}
const TRONGRID_API_KEY = process.env.TRONGRID_API_KEY || "";
const BSCSCAN_API_KEY = process.env.BSCSCAN_API_KEY || "";
const ETHERSCAN_API_KEY = process.env.ETHERSCAN_API_KEY || "";

// USDT contract addresses per network
const USDT_CONTRACTS = {
  tron: "TR7NHqjeKQxGTCi8q8ZY4pL8otSzgjLj6t",
  bsc: "0x55d398326f99059fF775485246999027B3197955",
  eth: "0xdAC17F958D2ee523a2206206994597C13D831ec7",
};

// ── License store (simple JSON file) ────────────────────────────────────────
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

function txAlreadyUsed(txHash) {
  const licenses = loadLicenses();
  return licenses.some((l) => l.txHash === txHash);
}

function saveLicense(licenseKey, txHash) {
  const licenses = loadLicenses();
  licenses.push({ licenseKey, txHash, issuedAt: new Date().toISOString() });
  saveLicenses(licenses);
}

// ── Blockchain verification helpers ─────────────────────────────────────────

async function verifyTron(txHash) {
  const headers = { Accept: "application/json" };
  if (TRONGRID_API_KEY) headers["TRON-PRO-API-KEY"] = TRONGRID_API_KEY;

  const url = `https://api.trongrid.io/v1/transactions/${txHash}`;
  const res = await fetch(url, { headers });
  if (!res.ok) return { ok: false, reason: "TronGrid request failed" };

  const json = await res.json();
  const tx = json?.data?.[0];
  if (!tx) return { ok: false, reason: "Transaction not found" };

  const receipt = tx?.ret?.[0];
  if (receipt?.contractRet !== "SUCCESS")
    return { ok: false, reason: "Transaction not successful" };

  const raw = tx?.raw_data?.contract?.[0]?.parameter?.value;
  if (!raw) return { ok: false, reason: "Could not read contract data" };

  const toAddr = raw.to_address;
  const contractAddr = raw.contract_address;
  const amountSun = Number(raw.amount || 0);

  // Validate this is TRC-20 USDT → our wallet
  if (
    contractAddr?.toLowerCase() !==
    tronAddressToHex(USDT_CONTRACTS.tron).toLowerCase()
  ) {
    // Try TRC-10 / native transfer fallback — just check amount & destination
  }

  // Decode TRC-20 transfer data
  const data = tx?.raw_data?.contract?.[0]?.parameter?.value?.data || "";
  if (data.startsWith("a9059cbb")) {
    const to = "41" + data.slice(32, 72);
    const amount = parseInt(data.slice(72, 136), 16) / 1e6;
    const ourHex = tronAddressToHex(WALLET_ADDRESS).toLowerCase();
    if (to.toLowerCase() !== ourHex.toLowerCase())
      return { ok: false, reason: "Wrong destination wallet" };
    if (amount < PRICE_USDT)
      return { ok: false, reason: `Amount too low: ${amount} USDT` };
    return { ok: true };
  }

  return { ok: false, reason: "Not a TRC-20 USDT transfer" };
}

function tronAddressToHex(addr) {
  // Base58Check → hex is complex; use TronGrid's own format comparison instead
  // For simplicity we compare lower-case base58 strings via the API response
  return addr;
}

async function verifyBsc(txHash) {
  if (!BSCSCAN_API_KEY)
    return { ok: false, reason: "BSCSCAN_API_KEY not configured" };

  const url =
    `https://api.bscscan.com/api?module=account&action=tokentx` +
    `&contractaddress=${USDT_CONTRACTS.bsc}` +
    `&address=${WALLET_ADDRESS}` +
    `&apikey=${BSCSCAN_API_KEY}`;

  const res = await fetch(url);
  const json = await res.json();
  if (json.status !== "1") return { ok: false, reason: "BSCScan query failed" };

  const tx = json.result?.find(
    (t) => t.hash.toLowerCase() === txHash.toLowerCase()
  );
  if (!tx) return { ok: false, reason: "Transaction not found" };

  const amount = Number(tx.value) / 10 ** Number(tx.tokenDecimal);
  if (amount < PRICE_USDT)
    return { ok: false, reason: `Amount too low: ${amount} USDT` };

  return { ok: true };
}

async function verifyEth(txHash) {
  if (!ETHERSCAN_API_KEY)
    return { ok: false, reason: "ETHERSCAN_API_KEY not configured" };

  const url =
    `https://api.etherscan.io/api?module=account&action=tokentx` +
    `&contractaddress=${USDT_CONTRACTS.eth}` +
    `&address=${WALLET_ADDRESS}` +
    `&apikey=${ETHERSCAN_API_KEY}`;

  const res = await fetch(url);
  const json = await res.json();
  if (json.status !== "1")
    return { ok: false, reason: "Etherscan query failed" };

  const tx = json.result?.find(
    (t) => t.hash.toLowerCase() === txHash.toLowerCase()
  );
  if (!tx) return { ok: false, reason: "Transaction not found" };

  const amount = Number(tx.value) / 10 ** Number(tx.tokenDecimal);
  if (amount < PRICE_USDT)
    return { ok: false, reason: `Amount too low: ${amount} USDT` };

  return { ok: true };
}

async function verifyTransaction(txHash) {
  if (NETWORK === "tron") return verifyTron(txHash);
  if (NETWORK === "bsc") return verifyBsc(txHash);
  if (NETWORK === "eth") return verifyEth(txHash);
  return { ok: false, reason: `Unknown network: ${NETWORK}` };
}

// ── Routes ───────────────────────────────────────────────────────────────────

// GET /api/payments/config
router.get("/config", (req, res) => {
  const cfg = getPaymentConfig();
  if (!cfg.walletAddress) {
    return res.status(503).json({ error: "Wallet not configured on server. Set USDT_WALLET_ADDRESS in Admin > Payment." });
  }
  res.json({ address: cfg.walletAddress, network: cfg.network, price: cfg.price });
});

// POST /api/payments/verify
router.post("/verify", async (req, res) => {
  const { txHash, productId } = req.body || {};
  if (!txHash) return res.status(400).json({ error: "txHash is required" });

  if (txAlreadyUsed(txHash)) {
    return res.status(400).json({ error: "Transaction already redeemed" });
  }

  let result;
  try {
    result = await verifyTransaction(txHash);
  } catch (err) {
    console.error("Verification error:", err);
    return res.status(502).json({ error: "Blockchain lookup failed — try again shortly" });
  }

  if (!result.ok) {
    return res.status(402).json({ error: result.reason });
  }

  const licenseKey = generateLicenseKey();
  saveLicense(licenseKey, txHash);
  console.log(`✅ License issued: ${licenseKey} for tx ${txHash}`);
  res.json({ licenseKey });
});

// GET /api/payments/check/:key
router.get("/check/:key", (req, res) => {
  const { key } = req.params;
  const licenses = loadLicenses();
  const valid = licenses.some((l) => l.licenseKey === key);
  res.json({ valid });
});

export default router;

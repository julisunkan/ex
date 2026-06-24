import express from "express";
import cors from "cors";
import paymentsRouter from "./routes/payments.js";
import adminRouter from "./routes/admin.js";
import settingsRouter from "./routes/settings.js";
import { startExpiryChecker } from "./lib/expiry-checker.js";

const PORT = Number(process.env.PORT || process.env.API_PORT || 3001);
const app = express();

app.use(cors({ origin: true }));
app.use(express.json());

app.use("/api/payments", paymentsRouter);
app.use("/api/admin", adminRouter);
app.use("/api/admin", settingsRouter);
app.use("/api/config", settingsRouter);

app.get("/", (_req, res) => res.send("Bank Statement Analyzer API is running. Use /api/health to check status."));
app.get("/api/health", (_req, res) => res.json({ ok: true }));

const HOST = "0.0.0.0";

app.listen(PORT, HOST, () => {
  console.log(`🚀 Payment API running on http://localhost:${PORT}`);
  startExpiryChecker();

  const wallet = process.env.USDT_WALLET_ADDRESS;
  const network = process.env.USDT_NETWORK || "tron";
  if (!wallet) {
    console.warn("⚠️  USDT_WALLET_ADDRESS is not set — /api/payments/config will return 503");
  } else {
    console.log(`💳 Wallet: ${wallet} (${network.toUpperCase()})`);
  }
});

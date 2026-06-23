import express from "express";
import cors from "cors";
import paymentsRouter from "./routes/payments.js";

const PORT = Number(process.env.API_PORT || 3001);
const app = express();

app.use(cors({ origin: true }));
app.use(express.json());

app.use("/api/payments", paymentsRouter);

app.get("/api/health", (_req, res) => res.json({ ok: true }));

app.listen(PORT, "127.0.0.1", () => {
  console.log(`🚀 Payment API running on http://localhost:${PORT}`);

  const wallet = process.env.USDT_WALLET_ADDRESS;
  const network = process.env.USDT_NETWORK || "tron";
  if (!wallet) {
    console.warn("⚠️  USDT_WALLET_ADDRESS is not set — /api/payments/config will return 503");
  } else {
    console.log(`💳 Wallet: ${wallet} (${network.toUpperCase()})`);
  }
});

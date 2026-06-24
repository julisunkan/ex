import { Router } from "express";
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { randomBytes } from "crypto";
import { notifyNewTicket } from "../lib/notify.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const TICKETS_FILE = join(__dirname, "../data/tickets.json");

const router = Router();

function loadTickets() {
  try { return JSON.parse(readFileSync(TICKETS_FILE, "utf8")); } catch { return []; }
}

function saveTickets(tickets) {
  writeFileSync(TICKETS_FILE, JSON.stringify(tickets, null, 2));
}

function requireAdmin(req, res, next) {
  const password = process.env.ADMIN_PASSWORD;
  if (!password) return res.status(503).json({ error: "ADMIN_PASSWORD not configured" });
  if (req.headers["x-admin-password"] !== password) return res.status(401).json({ error: "Unauthorized" });
  next();
}

// POST /api/tickets  — submit a new support ticket (public)
router.post("/", (req, res) => {
  const { name, email, licenseKey, category, subject, message } = req.body || {};
  if (!name || !email || !subject || !message) {
    return res.status(400).json({ error: "name, email, subject, and message are required." });
  }
  if (!email.includes("@")) {
    return res.status(400).json({ error: "A valid email address is required." });
  }
  if (message.trim().length < 20) {
    return res.status(400).json({ error: "Please provide more detail in your message (at least 20 characters)." });
  }

  const ticket = {
    id: "TKT-" + randomBytes(5).toString("hex").toUpperCase(),
    name: String(name).trim().slice(0, 100),
    email: String(email).trim().toLowerCase(),
    licenseKey: licenseKey ? String(licenseKey).trim() : null,
    category: category || "general",
    subject: String(subject).trim().slice(0, 200),
    message: String(message).trim().slice(0, 2000),
    status: "open",
    createdAt: new Date().toISOString(),
    adminReply: null,
    repliedAt: null,
  };

  const tickets = loadTickets();
  tickets.unshift(ticket);
  saveTickets(tickets);

  console.log(`🎫 New ticket: ${ticket.id} from ${ticket.email} — "${ticket.subject}"`);

  // Fire notification (webhook + email) without blocking the response
  notifyNewTicket(ticket).catch(err => console.warn("[tickets] Notify failed:", err.message));

  res.status(201).json({ ok: true, ticketId: ticket.id });
});

// GET /api/tickets  — list all tickets (admin only)
router.get("/", requireAdmin, (req, res) => {
  const tickets = loadTickets();
  const { status } = req.query;
  const filtered = status ? tickets.filter(t => t.status === status) : tickets;
  res.json({ total: tickets.length, tickets: filtered });
});

// PATCH /api/tickets/:id  — update status or add reply (admin only)
router.patch("/:id", requireAdmin, (req, res) => {
  const { id } = req.params;
  const { status, adminReply } = req.body || {};
  const tickets = loadTickets();
  const ticket = tickets.find(t => t.id === id);
  if (!ticket) return res.status(404).json({ error: "Ticket not found" });

  if (status) ticket.status = status;
  if (adminReply !== undefined) {
    ticket.adminReply = adminReply;
    ticket.repliedAt = new Date().toISOString();
    if (status !== "closed") ticket.status = "resolved";
  }

  saveTickets(tickets);
  res.json({ ok: true, ticket });
});

// DELETE /api/tickets/:id  — delete a ticket (admin only)
router.delete("/:id", requireAdmin, (req, res) => {
  const { id } = req.params;
  const tickets = loadTickets();
  const updated = tickets.filter(t => t.id !== id);
  if (updated.length === tickets.length) return res.status(404).json({ error: "Ticket not found" });
  saveTickets(updated);
  res.json({ ok: true });
});

export default router;

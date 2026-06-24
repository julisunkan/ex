import { useState } from "react";
import { getLicense } from "@/lib/payment";

const API_BASE = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");

const CATEGORIES = [
  { value: "general",      label: "General Question" },
  { value: "billing",      label: "Billing / Payment" },
  { value: "license",      label: "License Key Issue" },
  { value: "bug",          label: "Bug Report" },
  { value: "feature",      label: "Feature Request" },
  { value: "other",        label: "Other" },
];

export default function SupportPage() {
  const savedLicense = getLicense() ?? "";

  const [form, setForm] = useState({
    name: "",
    email: "",
    licenseKey: savedLicense,
    category: "general",
    subject: "",
    message: "",
  });
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState<string | null>(null);
  const [error, setError] = useState("");

  function update(field: string, value: string) {
    setForm((f) => ({ ...f, [field]: value }));
    setError("");
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setSubmitting(true);
    setError("");
    try {
      const res = await fetch(`${API_BASE}/api/tickets`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || `Error ${res.status}`);
      setSubmitted(data.ticketId);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to submit. Please try again.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col">
      <header className="flex items-center gap-3 px-4 py-3 bg-white border-b border-border shadow-sm shrink-0">
        <a
          href="/"
          className="flex items-center gap-1 text-sm font-semibold text-muted-foreground hover:text-foreground px-2.5 py-1.5 rounded-lg hover:bg-muted transition-colors"
        >
          <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
            <polyline points="15 18 9 12 15 6" />
          </svg>
          Back
        </a>
        <h1 className="text-sm font-bold text-foreground">Support</h1>
      </header>

      <div className="flex-1 overflow-y-auto px-4 py-5">
        {submitted ? (
          <div className="flex flex-col items-center justify-center text-center gap-4 pt-8 px-4">
            <div className="w-16 h-16 rounded-full bg-green-100 flex items-center justify-center">
              <svg className="w-8 h-8 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="20 6 9 17 4 12" />
              </svg>
            </div>
            <div>
              <p className="text-base font-bold text-foreground">Ticket Submitted!</p>
              <p className="text-xs text-muted-foreground mt-1">Your ticket ID is:</p>
              <p className="text-sm font-mono font-bold text-primary mt-1">{submitted}</p>
            </div>
            <p className="text-xs text-muted-foreground leading-relaxed max-w-xs">
              We'll review your request and respond to <span className="font-semibold text-foreground">{form.email}</span> as soon as possible. Please keep your ticket ID for reference.
            </p>
            <button
              onClick={() => { setSubmitted(null); setForm({ name: "", email: "", licenseKey: savedLicense, category: "general", subject: "", message: "" }); }}
              className="text-xs text-primary hover:underline mt-2"
            >
              Submit another ticket
            </button>
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="space-y-3.5">
            <div>
              <p className="text-xs text-muted-foreground leading-relaxed mb-4">
                Have a problem or question? Fill in the form below and our team will get back to you.
              </p>
            </div>

            {/* Name */}
            <div className="space-y-1">
              <label className="text-xs font-semibold text-foreground">Full Name <span className="text-destructive">*</span></label>
              <input
                type="text"
                value={form.name}
                onChange={(e) => update("name", e.target.value)}
                placeholder="Your name"
                required
                className="w-full rounded-xl border border-border bg-white px-3.5 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                data-testid="input-support-name"
              />
            </div>

            {/* Email */}
            <div className="space-y-1">
              <label className="text-xs font-semibold text-foreground">Email Address <span className="text-destructive">*</span></label>
              <input
                type="email"
                value={form.email}
                onChange={(e) => update("email", e.target.value)}
                placeholder="you@example.com"
                required
                className="w-full rounded-xl border border-border bg-white px-3.5 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                data-testid="input-support-email"
              />
            </div>

            {/* License key (optional) */}
            <div className="space-y-1">
              <label className="text-xs font-semibold text-foreground">License Key <span className="text-muted-foreground font-normal">(optional)</span></label>
              <input
                type="text"
                value={form.licenseKey}
                onChange={(e) => update("licenseKey", e.target.value)}
                placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
                className="w-full rounded-xl border border-border bg-white px-3.5 py-2.5 text-sm font-mono focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                data-testid="input-support-license"
              />
            </div>

            {/* Category */}
            <div className="space-y-1">
              <label className="text-xs font-semibold text-foreground">Category <span className="text-destructive">*</span></label>
              <select
                value={form.category}
                onChange={(e) => update("category", e.target.value)}
                className="w-full rounded-xl border border-border bg-white px-3.5 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors"
                data-testid="select-support-category"
              >
                {CATEGORIES.map((c) => (
                  <option key={c.value} value={c.value}>{c.label}</option>
                ))}
              </select>
            </div>

            {/* Subject */}
            <div className="space-y-1">
              <label className="text-xs font-semibold text-foreground">Subject <span className="text-destructive">*</span></label>
              <input
                type="text"
                value={form.subject}
                onChange={(e) => update("subject", e.target.value)}
                placeholder="Brief summary of your issue"
                required
                className="w-full rounded-xl border border-border bg-white px-3.5 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                data-testid="input-support-subject"
              />
            </div>

            {/* Message */}
            <div className="space-y-1">
              <label className="text-xs font-semibold text-foreground">Message <span className="text-destructive">*</span></label>
              <textarea
                value={form.message}
                onChange={(e) => update("message", e.target.value)}
                placeholder="Describe your issue in detail…"
                required
                rows={5}
                className="w-full rounded-xl border border-border bg-white px-3.5 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50 resize-none"
                data-testid="textarea-support-message"
              />
              <p className="text-[11px] text-muted-foreground">{form.message.length} / 2000</p>
            </div>

            {error && (
              <div className="flex items-start gap-2 bg-destructive/10 border border-destructive/20 rounded-xl px-3 py-2.5">
                <svg className="w-4 h-4 text-destructive shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <p className="text-xs text-destructive font-medium">{error}</p>
              </div>
            )}

            <button
              type="submit"
              disabled={submitting || !form.name || !form.email || !form.subject || !form.message}
              className="w-full flex items-center justify-center gap-2 bg-primary text-primary-foreground rounded-xl py-3 text-sm font-bold hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
              data-testid="button-support-submit"
            >
              {submitting ? (
                <><span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />Submitting…</>
              ) : (
                <><svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
                </svg>Submit Ticket</>
              )}
            </button>
          </form>
        )}
      </div>
    </div>
  );
}

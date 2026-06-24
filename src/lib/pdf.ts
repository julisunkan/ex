import type { Summary } from "./categorizer";

function makeFmt(symbol: string) {
  return (n: number) => {
    const abs = Math.abs(n);
    return `${symbol}${abs.toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
  };
}

function fmtPct(n: number): string {
  return `${n >= 0 ? "+" : ""}${n}%`;
}

function healthLabel(score: number): string {
  return score >= 75 ? "Excellent" : score >= 60 ? "Good" : score >= 45 ? "Fair" : "Needs Work";
}

function healthColor(score: number): string {
  return score >= 75 ? "#16a34a" : score >= 60 ? "#ca8a04" : score >= 45 ? "#d97706" : "#dc2626";
}

function catBadge(name: string, color: string): string {
  const bg = color + "22";
  return `<span style="background:${bg};color:${color};border:1px solid ${color}44;padding:1px 7px;border-radius:4px;font-size:10px;font-weight:700;white-space:nowrap">${name}</span>`;
}

export function buildReportHtml(summary: Summary, appName: string, currencySymbol = "₦", notes: Record<number, string> = {}, flaggedRows: Set<number> = new Set()): string {
  const fmtNum = makeFmt(currencySymbol);
  const sortedCats = Object.entries(summary.byCategory).sort((a, b) => b[1].total - a[1].total);
  const maxCat = sortedCats[0]?.[1].total || 1;
  const now = new Date().toLocaleString("en-US", { dateStyle: "full", timeStyle: "short" });

  // ── KPI cards ─────────────────────────────────────────────────────────────
  const kpiCards = [
    { label: "Total Income",    value: fmtNum(summary.totalIncome),    color: "#16a34a" },
    { label: "Total Expenses",  value: fmtNum(summary.totalExpenses),  color: "#dc2626" },
    { label: "Net Savings",     value: fmtNum(summary.net),            color: summary.net >= 0 ? "#2563eb" : "#dc2626" },
    { label: "Savings Rate",    value: `${summary.savingsRate}%`,      color: summary.savingsRate >= 20 ? "#16a34a" : summary.savingsRate >= 10 ? "#ca8a04" : "#dc2626" },
    { label: "Health Score",    value: `${summary.healthScore}/100`,   color: healthColor(summary.healthScore) },
    { label: "Transactions",    value: `${summary.transactions.length}`, color: "#4f46e5" },
  ]
    .map(
      (k) => `
      <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;flex:1;min-width:120px">
        <div style="font-size:10px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">${k.label}</div>
        <div style="font-size:18px;font-weight:800;color:${k.color};letter-spacing:-.02em">${k.value}</div>
      </div>`
    )
    .join("");

  // ── Health tips ────────────────────────────────────────────────────────────
  const tipsHtml = summary.healthTips.length
    ? `<div style="margin-top:10px;display:flex;flex-direction:column;gap:5px">
        ${summary.healthTips.map((t) => `<div style="font-size:11px;color:#475569;display:flex;gap:6px"><span>💡</span><span>${t}</span></div>`).join("")}
       </div>`
    : "";

  // ── Category bars ──────────────────────────────────────────────────────────
  const catRows = sortedCats
    .map(([name, info]) => {
      const barW = Math.round((info.total / maxCat) * 100);
      const pct = summary.totalExpenses > 0 ? Math.round((info.total / summary.totalExpenses) * 100) : 0;
      return `
      <tr>
        <td style="padding:6px 8px">${catBadge(name, info.color)}</td>
        <td style="padding:6px 8px;font-weight:700;text-align:right">${fmtNum(info.total)}</td>
        <td style="padding:6px 8px;text-align:center;color:#64748b;font-size:11px">${info.count}</td>
        <td style="padding:6px 8px;text-align:center;color:#64748b;font-size:11px">${pct}%</td>
        <td style="padding:6px 8px;width:120px">
          <div style="height:8px;background:#f1f5f9;border-radius:4px;overflow:hidden">
            <div style="height:8px;width:${barW}%;background:${info.color};border-radius:4px"></div>
          </div>
        </td>
      </tr>`;
    })
    .join("");

  // ── Monthly breakdown ──────────────────────────────────────────────────────
  const monthlyHtml =
    summary.monthly.length > 1
      ? `<div class="section">
          <h3>Monthly Breakdown</h3>
          <table>
            <thead><tr>
              <th>Month</th><th style="text-align:right">Income</th><th style="text-align:right">Expenses</th><th style="text-align:right">Net</th>
            </tr></thead>
            <tbody>
              ${summary.monthly
                .map(
                  (m) => `<tr>
                  <td style="font-weight:600">${m.month}</td>
                  <td style="text-align:right;color:#16a34a">${fmtNum(m.income)}</td>
                  <td style="text-align:right;color:#dc2626">${fmtNum(m.expenses)}</td>
                  <td style="text-align:right;color:${m.net >= 0 ? "#2563eb" : "#dc2626"};font-weight:700">${fmtNum(m.net)}</td>
                </tr>`
                )
                .join("")}
            </tbody>
          </table>
         </div>`
      : "";

  // ── Month-over-month comparison ────────────────────────────────────────────
  const momHtml =
    summary.momMonths && summary.momChanges.length > 0
      ? `<div class="section">
          <h3>Month-over-Month: <span style="color:#64748b;font-weight:500">${summary.momMonths[1]}</span> → <span style="color:#4f46e5">${summary.momMonths[0]}</span></h3>
          <table>
            <thead><tr>
              <th>Category</th>
              <th style="text-align:right">${summary.momMonths[1]}</th>
              <th style="text-align:right">${summary.momMonths[0]}</th>
              <th style="text-align:right">Change</th>
              <th style="text-align:right">%</th>
            </tr></thead>
            <tbody>
              ${summary.momChanges
                .map((c) => {
                  const badge = c.isNew
                    ? `<span style="font-size:9px;font-weight:700;background:#fef3c7;color:#92400e;padding:1px 5px;border-radius:3px">NEW</span>`
                    : c.isGone
                    ? `<span style="font-size:9px;font-weight:700;background:#f1f5f9;color:#64748b;padding:1px 5px;border-radius:3px">GONE</span>`
                    : c.delta > 0
                    ? `<span style="color:#dc2626;font-weight:700">▲ ${fmtNum(c.delta)}</span>`
                    : `<span style="color:#16a34a;font-weight:700">▼ ${fmtNum(Math.abs(c.delta))}</span>`;
                  const pctBadge =
                    !c.isNew && !c.isGone
                      ? `<span style="font-size:11px;font-weight:700;color:${c.delta > 0 ? "#dc2626" : "#16a34a"}">${fmtPct(c.pctChange)}</span>`
                      : "—";
                  return `<tr>
                    <td style="padding:5px 8px">${catBadge(c.category, c.color)}</td>
                    <td style="padding:5px 8px;text-align:right;color:#64748b">${c.previous > 0 ? fmtNum(c.previous) : "—"}</td>
                    <td style="padding:5px 8px;text-align:right;font-weight:600">${c.current > 0 ? fmtNum(c.current) : "—"}</td>
                    <td style="padding:5px 8px;text-align:right">${badge}</td>
                    <td style="padding:5px 8px;text-align:right">${pctBadge}</td>
                  </tr>`;
                })
                .join("")}
            </tbody>
          </table>
         </div>`
      : "";

  // ── Recurring ──────────────────────────────────────────────────────────────
  const recurringHtml =
    summary.recurring.length > 0
      ? `<div class="section">
          <h3>Recurring / Subscriptions Detected</h3>
          <table>
            <thead><tr><th>Description</th><th>Category</th><th style="text-align:center">Occurrences</th><th style="text-align:right">Avg Amount</th><th style="text-align:right">Total</th></tr></thead>
            <tbody>
              ${summary.recurring
                .map(
                  (r) => `<tr>
                  <td style="font-size:11px;max-width:180px;word-break:break-word">${r.description}</td>
                  <td>${catBadge(r.category.name, r.category.color)}</td>
                  <td style="text-align:center;font-weight:700">${r.count}×</td>
                  <td style="text-align:right">${fmtNum(Math.round(r.avgAmount))}</td>
                  <td style="text-align:right;font-weight:700;color:#dc2626">${fmtNum(Math.round(r.totalAmount))}</td>
                </tr>`
                )
                .join("")}
            </tbody>
          </table>
         </div>`
      : "";

  // ── Transactions ───────────────────────────────────────────────────────────
  const txRows = summary.transactions
    .map((tx) => {
      const isDupe = summary.duplicateRows.has(tx.row);
      const dupeTag = isDupe
        ? `<span style="font-size:9px;font-weight:700;background:#fef3c7;color:#92400e;padding:1px 4px;border-radius:3px;margin-left:4px">⚠ DUP</span>`
        : "";
      const note = notes[tx.row] ? notes[tx.row].trim() : "";
      const noteTag = note
        ? `<div style="font-size:9px;color:#64748b;font-style:italic;margin-top:2px">📝 ${note}</div>`
        : "";
      const isFlagged = flaggedRows.has(tx.row);
      const flagTag = isFlagged
        ? `<span style="font-size:9px;font-weight:700;background:#fee2e2;color:#b91c1c;padding:1px 4px;border-radius:3px;margin-left:4px">🚩</span>`
        : "";
      const rowBg = isFlagged ? "background:#fff5f5" : isDupe ? "background:#fffbeb" : "";
      return `<tr style="${rowBg}">
        <td style="font-size:10px;color:#64748b;white-space:nowrap">${tx.date}</td>
        <td style="font-size:11px;max-width:200px;word-break:break-word">${tx.description}${flagTag}${dupeTag}${noteTag}</td>
        <td>${catBadge(tx.category.name, tx.category.color)}</td>
        <td style="text-align:center">
          <span style="font-size:10px;font-weight:700;padding:1px 6px;border-radius:3px;${tx.type === "credit" ? "background:#dcfce7;color:#15803d" : "background:#fee2e2;color:#b91c1c"}">
            ${tx.type === "credit" ? "IN" : "OUT"}
          </span>
        </td>
        <td style="text-align:right;font-weight:700;color:${tx.type === "credit" ? "#16a34a" : "#dc2626"};white-space:nowrap">
          ${tx.type === "credit" ? "+" : "−"}${fmtNum(tx.amount)}
        </td>
      </tr>`;
    })
    .join("");

  // ── Full HTML document ─────────────────────────────────────────────────────
  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>${appName} — Statement Report</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      font-size: 12px;
      color: #1e293b;
      background: #fff;
      padding: 32px 36px;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    h2 { font-size: 20px; font-weight: 800; color: #1e3a8a; letter-spacing: -.02em; }
    h3 { font-size: 13px; font-weight: 700; color: #1e293b; margin-bottom: 10px; }
    .section { margin-bottom: 28px; }
    table { width: 100%; border-collapse: collapse; }
    th {
      text-align: left;
      font-size: 10px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: #64748b;
      padding: 6px 8px;
      border-bottom: 2px solid #e2e8f0;
    }
    td { padding: 6px 8px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
    tr:last-child td { border-bottom: none; }
    tr:nth-child(even) { background: #f8fafc; }
    .kpi-grid { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 10px; }
    .health-bar-bg { height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; width: 160px; display: inline-block; vertical-align: middle; }
    @media print {
      body { padding: 16px 20px; }
      .no-print { display: none !important; }
      tr { page-break-inside: avoid; }
      .section { page-break-inside: avoid; }
    }
  </style>
</head>
<body>

  <!-- Header -->
  <div class="section" style="border-bottom:3px solid #1e3a8a;padding-bottom:14px;margin-bottom:22px;display:flex;justify-content:space-between;align-items:flex-end">
    <div>
      <h2>${appName}</h2>
      <div style="font-size:11px;color:#64748b;margin-top:3px;font-weight:500">Bank Statement Analysis Report</div>
    </div>
    <div style="text-align:right;font-size:10px;color:#94a3b8">
      <div>Generated: ${now}</div>
      <div style="margin-top:2px">${summary.transactions.length} transactions</div>
    </div>
  </div>

  <!-- KPIs -->
  <div class="section">
    <h3>Overview</h3>
    <div class="kpi-grid">${kpiCards}</div>
    <!-- Health score bar -->
    <div style="display:flex;align-items:center;gap:10px;margin-top:8px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px 14px">
      <span style="font-size:11px;font-weight:600;color:#475569;min-width:90px">Financial Health</span>
      <div class="health-bar-bg">
        <div style="height:10px;width:${summary.healthScore}%;background:${healthColor(summary.healthScore)};border-radius:5px"></div>
      </div>
      <span style="font-size:13px;font-weight:800;color:${healthColor(summary.healthScore)}">${summary.healthScore}/100</span>
      <span style="font-size:11px;font-weight:600;color:${healthColor(summary.healthScore)}">${healthLabel(summary.healthScore)}</span>
    </div>
    ${tipsHtml}
  </div>

  <!-- Category Breakdown -->
  <div class="section">
    <h3>Spending by Category</h3>
    <table>
      <thead><tr>
        <th>Category</th>
        <th style="text-align:right">Total</th>
        <th style="text-align:center">Txns</th>
        <th style="text-align:center">% of Expenses</th>
        <th>Relative Spend</th>
      </tr></thead>
      <tbody>${catRows}</tbody>
    </table>
  </div>

  ${monthlyHtml}
  ${momHtml}
  ${recurringHtml}

  <!-- Transactions -->
  <div class="section">
    <h3>All Transactions
      ${summary.duplicateRows.size > 0 ? `<span style="font-size:10px;font-weight:500;color:#64748b;margin-left:8px">⚠ ${summary.duplicateRows.size} possible duplicate(s) flagged</span>` : ""}
    </h3>
    <table>
      <thead><tr>
        <th>Date</th>
        <th>Description</th>
        <th>Category</th>
        <th style="text-align:center">Type</th>
        <th style="text-align:right">Amount</th>
      </tr></thead>
      <tbody>${txRows}</tbody>
    </table>
  </div>

  <!-- Footer -->
  <div style="border-top:1px solid #e2e8f0;padding-top:10px;margin-top:8px;display:flex;justify-content:space-between;font-size:10px;color:#94a3b8">
    <span>${appName} · Excel Add-in</span>
    <span class="no-print" style="color:#4f46e5;font-weight:600;cursor:pointer" onclick="window.print()">🖨 Print / Save as PDF</span>
  </div>

  <script>
    window.addEventListener("load", () => { setTimeout(() => window.print(), 400); });
  </script>
</body>
</html>`;

  return html;
}

export function exportToPdf(summary: Summary, appName: string, currencySymbol = "₦", notes: Record<number, string> = {}, flaggedRows: Set<number> = new Set()): void {
  const html = buildReportHtml(summary, appName, currencySymbol, notes, flaggedRows);
  const win = window.open("", "_blank");
  if (!win) {
    alert("Pop-up blocked. Please allow pop-ups for this add-in to export PDF.");
    return;
  }
  win.document.write(html);
  win.document.close();
}

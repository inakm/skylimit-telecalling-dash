/* ============================================================
   Accounts Receivable — Telecaller Command Center.
   Customer-grouped table, live search, sorting, expandable bills,
   call status tracking, priority scoring, aging analysis,
   notes, promise-to-pay, follow-ups, call history, bulk actions.
   ============================================================ */

(() => {
  const $ = (sel) => document.querySelector(sel);
  const STORAGE_KEY = "receivables_call_status";
  const NOTES_KEY = "receivables_notes";
  const HISTORY_KEY = "receivables_history";

  const AGING_BUCKETS = [
    { label: "1–7 days", min: 1, max: 7 },
    { label: "8–14 days", min: 8, max: 14 },
    { label: "15–30 days", min: 15, max: 30 },
    { label: "31–60 days", min: 31, max: 60 },
    { label: "61–90 days", min: 61, max: 90 },
    { label: "90+ days", min: 91, max: Infinity },
  ];

  const AGING_COLORS = ["#27a644", "#34b856", "#5e6ad2", "#b45309", "#e07c00", "#d33b3b"];
  const SHEET = CONFIG.receivables;

  const headerMap = {
    reference: ["invoice", "inv no", "reference", "ref no", "receipt", "doc no", "invoice number", "inv no.", "bill no", "bill no."],
    customer: ["customer", "client", "company", "account", "customer name", "client name", "debtor"],
    phone: ["phone", "mobile", "contact", "phone number", "mobile number", "contact number", "tel"],
    issueDate: ["invoice date", "issue date", "date", "billing date", "transaction date", "bill date"],
    dueDate: ["due date", "payment due", "due", "due on"],
    amount: ["amount", "total", "balance", "outstanding", "amount due", "amount outstanding", "invoice amount", "balance due", "net amount", "bill value", "bill amount"],
    paid: ["paid", "paid amount", "payment", "collected", "amount paid"],
    status: ["status", "payment status", "state", "invoice status"],
    remarks: ["remarks", "remark", "notes", "comment"],
  };

  function matchHeader(headers) {
    const map = {};
    const normalized = headers.map((h) => String(h).toLowerCase().trim());
    for (const [key, candidates] of Object.entries(headerMap)) {
      for (const cand of candidates) {
        const idx = normalized.indexOf(cand);
        if (idx !== -1) { map[key] = headers[idx]; break; }
      }
    }
    return map;
  }

  function ageBucket(dueDate, today) {
    const days = Math.round((today - dueDate) / 86400000);
    const bucket = AGING_BUCKETS.find((b) => days >= b.min && days <= b.max) || AGING_BUCKETS[0];
    return { bucket, days };
  }

  function buildRows(records, cols) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const rows = [];
    for (const r of records) {
      const amount = Format.parseNumber(r[cols.amount]);
      if (!Number.isFinite(amount) || amount <= 0) continue;
      const paid = cols.paid ? Format.parseNumber(r[cols.paid]) : 0;
      const paidN = Number.isFinite(paid) ? paid : 0;
      const balance = Math.max(amount - paidN, 0);
      let dueDate = cols.dueDate ? Format.parseDate(r[cols.dueDate]) : null;
      if (!dueDate && cols.issueDate) {
        const issue = Format.parseDate(r[cols.issueDate]);
        if (issue) { dueDate = new Date(issue); dueDate.setDate(dueDate.getDate() + 30); }
      }
      const issueDate = cols.issueDate ? Format.parseDate(r[cols.issueDate]) : null;
      const daysOverdue = dueDate ? Math.round((today - dueDate) / 86400000) : null;
      const age = dueDate ? ageBucket(dueDate, today) : null;
      rows.push({
        reference: r[cols.reference] || "—", customer: r[cols.customer] || "—",
        phone: cols.phone ? (r[cols.phone] || "").trim() : "",
        issueDate, dueDate, amount, paid: paidN, balance,
        status: (r[cols.status] || "").trim(), daysOverdue,
        bucket: age ? age.bucket.label : null, bucketDays: age ? age.days : null,
        remarks: cols.remarks ? (r[cols.remarks] || "").trim() : "",
      });
    }
    return rows;
  }

  /* ---------- File Parsers ---------- */
  function parseExcel(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (!rows.length) throw new Error("The Excel file contains no data rows.");
    return rows;
  }

  function parseCSV(text) {
    const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter((l) => l.trim() !== "");
    if (lines.length < 2) throw new Error("CSV file must have a header row and at least one data row.");
    const rows = [];
    for (const line of lines) {
      const cells = []; let i = 0;
      while (i < line.length) {
        if (line[i] === '"') {
          let j = i + 1, val = "";
          while (j < line.length) {
            if (line[j] === '"') { if (j + 1 < line.length && line[j + 1] === '"') { val += '"'; j += 2; } else { j++; break; } }
            else { val += line[j]; j++; }
          }
          cells.push(val); i = j + 1; if (i < line.length && line[i] === ",") i++;
        } else {
          let j = line.indexOf(",", i); if (j === -1) j = line.length;
          cells.push(line.slice(i, j).trim()); i = j + 1;
        }
      }
      rows.push(cells);
    }
    const headers = rows[0].map((h) => h.replace(/^\uFEFF/, "").trim());
    return rows.slice(1).map((cells) => {
      const obj = {}; headers.forEach((h, i) => { obj[h] = cells[i] || ""; }); return obj;
    });
  }

  async function loadFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    let records;
    if (ext === "xlsx" || ext === "xls" || ext === "ods") { const buf = await file.arrayBuffer(); records = parseExcel(buf); }
    else { const text = await file.text(); records = parseCSV(text); }
    if (!records.length) throw new Error("The file contains no data rows.");
    const headers = Object.keys(records[0]);
    const cols = matchHeader(headers);
    if (!cols.amount || !cols.customer) throw new Error("Could not map columns. Expected at least customer and amount columns.");
    return { records, cols };
  }

  /* ---------- localStorage Helpers ---------- */
  function loadJSON(key) { try { return JSON.parse(localStorage.getItem(key)) || {}; } catch { return {}; } }
  function saveJSON(key, data) { localStorage.setItem(key, JSON.stringify(data)); }

  /* ---------- Call Status ---------- */
  function loadCallStatus() { return loadJSON(STORAGE_KEY); }
  function saveCallStatus(s) { saveJSON(STORAGE_KEY, s); }
  function getCallStatus(customer) { return loadCallStatus()[customer] || null; }
  function resetAllCallStatus() { localStorage.removeItem(STORAGE_KEY); localStorage.removeItem(HISTORY_KEY); }
  function purgeAllData() { localStorage.removeItem(STORAGE_KEY); localStorage.removeItem(HISTORY_KEY); localStorage.removeItem(NOTES_KEY); }

  /* ---------- History ---------- */
  function loadHistory() { return loadJSON(HISTORY_KEY); }
  function addHistoryEntry(customer, entry) {
    const hist = loadHistory();
    if (!hist[customer]) hist[customer] = [];
    hist[customer].push({ ...entry, date: new Date().toISOString() });
    saveJSON(HISTORY_KEY, hist);
  }
  function getHistory(customer) { return (loadHistory()[customer] || []).sort((a, b) => new Date(b.date) - new Date(a.date)); }

  /* ---------- Notes ---------- */
  function loadNotes() { return loadJSON(NOTES_KEY); }
  function saveNote(customer, text) {
    const notes = loadNotes();
    if (!notes[customer]) notes[customer] = [];
    notes[customer].push({ text, date: new Date().toISOString() });
    saveJSON(NOTES_KEY, notes);
  }
  function getNotes(customer) { return (loadNotes()[customer] || []).sort((a, b) => new Date(b.date) - new Date(a.date)); }
  function deleteNote(customer, idx) {
    const notes = loadNotes();
    if (notes[customer]) { notes[customer].splice(idx, 1); saveJSON(NOTES_KEY, notes); }
  }

  /* ---------- Mark Called (with history) ---------- */
  function markCalled(customer, action, extra) {
    const status = loadCallStatus();
    status[customer] = { action, date: new Date().toISOString(), ...extra };
    saveCallStatus(status);
    addHistoryEntry(customer, { action, ...extra });
  }

  /* ---------- Grouping ---------- */
  function groupByCustomer(rows) {
    const map = {};
    for (const r of rows) {
      const name = r.customer;
      if (!map[name]) {
        map[name] = { customer: name, phone: r.phone || "", bills: [], totalAmount: 0, totalPaid: 0, totalBalance: 0, maxDaysOverdue: null, latestDueDate: null };
      }
      const g = map[name];
      g.bills.push(r);
      g.totalAmount += r.amount; g.totalPaid += r.paid; g.totalBalance += r.balance;
      g.phone = g.phone || r.phone || "";
      if (r.daysOverdue !== null) { if (g.maxDaysOverdue === null || r.daysOverdue > g.maxDaysOverdue) g.maxDaysOverdue = r.daysOverdue; }
      if (r.dueDate && (!g.latestDueDate || r.dueDate < g.latestDueDate)) g.latestDueDate = r.dueDate;
    }
    return Object.values(map);
  }

  /* ---------- Priority Scoring ---------- */
  function computePriorityScore(g) {
    const overdueBills = g.bills.filter((b) => b.daysOverdue > 0);
    const overdueAmount = overdueBills.reduce((s, b) => s + b.balance, 0);
    const maxOverdue = g.maxDaysOverdue || 0;
    const billCount = g.bills.length;
    const overdueRatio = overdueBills.length / Math.max(billCount, 1);
    const score = Math.round((overdueAmount / 1000) * 2 + maxOverdue * 1.5 + billCount * 10 + overdueRatio * 50);
    const hist = getHistory(g.customer);
    const overduePromises = hist.filter((h) => h.action === "promised" && h.promisedDate && new Date(h.promisedDate) < new Date()).length;
    const adjustedScore = score + overduePromises * 20;
    return {
      overdueAmount, overdueBillCount: overdueBills.length,
      score: adjustedScore, urgency: adjustedScore >= 80 ? "critical" : adjustedScore >= 50 ? "high" : adjustedScore >= 25 ? "medium" : "low",
    };
  }

  /* ---------- Sorting ---------- */
  let currentSort = { key: "priority", dir: "desc" };
  function sortGroups(groups, sort) {
    const sorted = [...groups];
    switch (sort.key) {
      case "name": sorted.sort((a, b) => a.customer.localeCompare(b.customer)); break;
      case "amount": sorted.sort((a, b) => a.totalBalance - b.totalBalance); break;
      case "overdue": sorted.sort((a, b) => (b.maxDaysOverdue || 0) - (a.maxDaysOverdue || 0)); break;
      case "bills": sorted.sort((a, b) => b.bills.length - a.bills.length); break;
      case "priority": sorted.sort((a, b) => computePriorityScore(b).score - computePriorityScore(a).score); break;
    }
    if (sort.dir === "desc" && !["overdue", "bills", "priority"].includes(sort.key)) sorted.reverse();
    return sorted;
  }
  function toggleSort(key) {
    if (currentSort.key === key) currentSort.dir = currentSort.dir === "asc" ? "desc" : "asc";
    else currentSort = { key, dir: ["overdue", "bills", "priority"].includes(key) ? "desc" : "asc" };
    if (lastRows) renderDashboard(lastRows, lastSource);
  }

  /* ---------- Rendering Helpers ---------- */
  function statusBadge(status) {
    const s = String(status || "").toLowerCase();
    if (s.includes("paid") || s.includes("settled") || s.includes("closed")) return '<span class="badge badge-success badge-dot">Paid</span>';
    if (s.includes("partial")) return '<span class="badge badge-warning badge-dot">Partial</span>';
    if (s.includes("overdue") || s.includes("past due")) return '<span class="badge badge-danger badge-dot">Overdue</span>';
    if (s.includes("open") || s.includes("unpaid") || s.includes("pending") || s.includes("due")) return '<span class="badge badge-primary badge-dot">Open</span>';
    return `<span class="badge badge-neutral">${status || "—"}</span>`;
  }

  function overdueBadge(days) {
    if (days === null) return "—";
    if (days > 0) return `<span class="badge badge-danger badge-dot">${days}d overdue</span>`;
    return `<span class="badge badge-neutral">${Math.abs(days)}d to go</span>`;
  }

  function urgencyClass(days) {
    if (days === null) return "";
    if (days >= 90) return "row-critical";
    if (days >= 60) return "row-high";
    if (days >= 30) return "row-medium";
    if (days > 0) return "row-low";
    return "";
  }

  function urgencyLabel(score) {
    if (score >= 80) return '<span class="badge badge-danger badge-dot">Critical</span>';
    if (score >= 50) return '<span class="badge badge-warning badge-dot">High</span>';
    if (score >= 25) return '<span class="badge badge-primary badge-dot">Medium</span>';
    return '<span class="badge badge-neutral badge-dot">Low</span>';
  }

  function callStatusBadge(customer) {
    const st = getCallStatus(customer);
    if (!st) return "";
    const labels = { called: "Called", promised: "Promised", contacted: "Contacted", no_answer: "No answer", wrong_number: "Wrong #", follow_up: "Follow up", paid: "Paid" };
    const icons = { called: I.phone, promised: I.handshake, contacted: I.checkCircle, no_answer: I.phoneOff, wrong_number: I.xCircle, follow_up: I.refreshCw, paid: I.checkCircle };
    const label = labels[st.action] || st.action;
    const icon = icons[st.action] || "•";
    const date = new Date(st.date);
    const timeStr = date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    return `<span class="call-status-badge">${icon} ${label} <span class="call-status-time">${timeStr}</span></span>`;
  }

  function timeAgo(isoStr) {
    const diff = Date.now() - new Date(isoStr).getTime();
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return "just now";
    if (mins < 60) return `${mins}m ago`;
    const hrs = Math.floor(mins / 60);
    if (hrs < 24) return `${hrs}h ago`;
    const days = Math.floor(hrs / 24);
    return `${days}d ago`;
  }

  function statCard(label, value, opts = {}) {
    return Dom.el("div", { class: `stat-card ${opts.class || ""}`, html: `
      <div class="stat-label">${label}</div>
      <div class="stat-value">${value}</div>
      ${opts.delta ? `<div class="stat-delta">${opts.delta}</div>` : ""}
    `});
  }

  /* ---------- Lucide Icons (inline SVG) ---------- */
  const icon = (paths, sz = 14) => `<svg width="${sz}" height="${sz}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
  const I = {
    phone: icon('<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/>'),
    phoneOff: icon('<path d="M10.68 13.31a16 16 0 0 0 3.41 2.6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91"/><line x1="23" y1="1" x2="1" y2="23"/>'),
    phoneForwarded: icon('<polyline points="2 10 7 10 7 2"/><polyline points="22 10 17 10 17 2"/><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/>'),
    users: icon('<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>'),
    handshake: icon('<path d="M11 17a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9l4.39-4.39a2 2 0 0 1 2.82 0l2.83 2.83a2 2 0 0 0 2.82 0l2.83-2.83a2 2 0 0 1 2.82 0L23 9v8a1 1 0 0 1-1 1h-2a1 1 0 0 1-1-1v-2"/><path d="M13 13l-3-3"/><circle cx="7.5" cy="13.5" r="1.5"/>'),
    banknote: icon('<rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="2"/><path d="M6 12h.01M18 12h.01"/>'),
    alertTriangle: icon('<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" x2="12" y1="9" y2="13"/><line x1="12" x2="12.01" y1="17" y2="17"/>'),
    barChart: icon('<path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/>'),
    check: icon('<polyline points="20 6 9 17 4 12"/>'),
    checkCircle: icon('<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>'),
    xCircle: icon('<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'),
    refreshCw: icon('<polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>'),
    calendar: icon('<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>'),
    calendarCheck: icon('<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><path d="m9 16 2 2 4-4"/>'),
    clock: icon('<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>'),
    alertCircle: icon('<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>'),
    search: icon('<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>'),
    info: icon('<circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/>'),
    chevronDown: (sz) => `<svg width="${sz || 16}" height="${sz || 16}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>`,
    chevronRight: icon('<polyline points="9 18 15 12 9 6"/>'),
    download: icon('<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>'),
    fileDown: icon('<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><polyline points="9 15 12 18 15 15"/>'),
    fileText: icon('<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/>'),
    note: icon('<path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/>'),
    trash: icon('<polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>', 12),
    target: icon('<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>'),
    shieldAlert: icon('<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10"/><path d="M12 8v4"/><path d="M12 16h.01"/>'),
    history: icon('<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l4 2"/>'),
    whatsapp: icon('<path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347z"/><path d="M12 2C6.477 2 2 6.477 2 12c0 1.89.525 3.66 1.438 5.168L2 22l4.832-1.438A9.955 9.955 0 0 0 12 22c5.523 0 10-4.477 10-10S17.523 2 12 2z"/>'),
  };

  function phoneLink(phone) {
    const num = phone.replace(/[^0-9+]/g, "");
    if (!num) return null;
    return { call: `tel:${num}`, whatsapp: `https://wa.me/${num.replace("+", "")}` };
  }

  /* ---------- Section: Telecaller Guide ---------- */
  function renderGuide() {
    const wrap = Dom.el("div", { class: "telecaller-guide" });
    wrap.innerHTML = `
      <div class="guide-header" id="guideToggle">
        <div class="guide-title">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/></svg>
          Telecaller Quick Guide
        </div>
        <svg class="guide-chevron" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
      </div>
      <div class="guide-body" id="guideBody">
        <div class="guide-steps">
          <div class="guide-step"><div class="guide-step-num">1</div><div class="guide-step-content"><strong>Check Today's Call List</strong><span>Top section shows who to call first, sorted by urgency.</span></div></div>
          <div class="guide-step"><div class="guide-step-num">2</div><div class="guide-step-content"><strong>Call &amp; Update Status</strong><span>Click call buttons, mark status, add notes and set follow-ups.</span></div></div>
          <div class="guide-step"><div class="guide-step-num">3</div><div class="guide-step-content"><strong>Track Promises</strong><span>Record promised amounts and dates. Overdue promises boost priority.</span></div></div>
          <div class="guide-step"><div class="guide-step-num">4</div><div class="guide-step-content"><strong>Use Filters</strong><span>Filter by date range, urgency, call status. Use bulk actions for efficiency.</span></div></div>
        </div>
      </div>`;
    return wrap;
  }

  /* ---------- Section: Daily Summary ---------- */
  function renderDailySummary(rows, groups) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const hist = loadHistory();
    const status = loadCallStatus();
    const allEntries = Object.values(hist).flat();
    const todayEntries = allEntries.filter((e) => new Date(e.date) >= today);
    const todayCalls = todayEntries.filter((e) => ["called", "contacted", "no_answer", "wrong_number"].includes(e.action)).length;
    const todayPromises = todayEntries.filter((e) => e.action === "promised");
    const todayPromisedAmount = todayPromises.reduce((s, e) => s + (e.promisedAmount || 0), 0);
    const overduePromises = allEntries.filter((e) => e.action === "promised" && e.promisedDate && new Date(e.promisedDate) < new Date());
    const totalCustomers = groups.filter((g) => g.totalBalance > 0).length;
    const contactedToday = Object.keys(status).filter((c) => {
      const s = status[c]; return s && new Date(s.date) >= today;
    }).length;
    const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);

    const wrap = Dom.el("div", { class: "card daily-summary-card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div><div class="card-title">Daily Collection Summary</div>
      <div class="card-subtitle">${new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}</div></div>
    `}));

    const grid = Dom.el("div", { class: "daily-summary-grid" });
    const stats = [
      { label: "Calls today", value: todayCalls, icon: I.phone, color: "var(--primary)" },
      { label: "Customers reached", value: `${contactedToday}/${totalCustomers}`, icon: I.users, color: "var(--semantic-success)" },
      { label: "Promises today", value: todayPromises.length, icon: I.handshake, color: "var(--semantic-warning)" },
      { label: "Promised amount", value: Format.moneyWhole(todayPromisedAmount), icon: I.banknote, color: "var(--primary)" },
      { label: "Overdue promises", value: overduePromises.length, icon: I.alertTriangle, color: overduePromises.length ? "var(--semantic-danger)" : "var(--semantic-success)" },
      { label: "Outstanding", value: Format.moneyWhole(totalOutstanding), icon: I.barChart, color: "var(--ink)" },
    ];
    stats.forEach((s) => {
      grid.appendChild(Dom.el("div", { class: "daily-summary-stat", html: `
        <div class="daily-stat-icon">${s.icon}</div>
        <div class="daily-stat-info">
          <div class="daily-stat-label">${s.label}</div>
          <div class="daily-stat-value" style="color:${s.color}">${s.value}</div>
        </div>
      `}));
    });
    wrap.appendChild(grid);
    return wrap;
  }

  /* ---------- Section: Today's Call List ---------- */
  function renderTodayCallList(groups) {
    const pending = groups
      .filter((g) => g.totalBalance > 0)
      .map((g) => ({ ...g, priority: computePriorityScore(g) }))
      .sort((a, b) => b.priority.score - a.priority.score);
    const called = pending.filter((g) => getCallStatus(g.customer));
    const notCalled = pending.filter((g) => !getCallStatus(g.customer));
    const progressPct = pending.length ? Math.round((called.length / pending.length) * 100) : 0;

    const wrap = Dom.el("div", { class: "card today-call-card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div><div class="card-title">Today's Call List</div>
      <div class="card-subtitle">Top priority customers to call — ${called.length}/${pending.length} contacted</div></div>
      <div class="today-progress">
        <div class="today-progress-bar"><div class="today-progress-fill" style="width:${progressPct}%"></div></div>
        <span class="today-progress-label">${progressPct}%</span>
      </div>
    `}));

    const toShow = notCalled.slice(0, 8);
    if (!toShow.length) {
      const list = Dom.el("div", { class: "today-call-list" });
      list.appendChild(Dom.el("div", { class: "today-call-empty", html: `<div class="guide-step-num">${I.checkCircle}</div><span>All customers contacted today! Great work.</span>` }));
      wrap.appendChild(list);
    } else {
      const tWrap = Dom.el("div", { class: "table-wrap today-call-table-wrap" });
      const table = Dom.el("table", { class: "data-table today-call-table" });
      const thead = Dom.el("thead");
      const hr = Dom.el("tr");
      ["#", "Customer Name", "Amount", "Status", "Actions"].forEach((h) => {
        const th = Dom.el("th", { text: h });
        if (h === "Amount") th.classList.add("num");
        hr.appendChild(th);
      });
      thead.appendChild(hr);
      table.appendChild(thead);

      const tbody = Dom.el("tbody");
      toShow.forEach((g, i) => {
        const tr = Dom.el("tr", { class: `urgency-${g.priority.urgency}` });
        const links = phoneLink(g.phone);

        tr.appendChild(Dom.el("td", { class: "today-call-rank-cell", text: `${i + 1}` }));

        const nameTd = Dom.el("td", { class: "cell-title" });
        nameTd.appendChild(Dom.el("span", { text: g.customer }));
        if (g.phone) nameTd.appendChild(Dom.el("span", { class: "cell-phone", html: `&nbsp;· ${g.phone}` }));
        tr.appendChild(nameTd);

        tr.appendChild(Dom.el("td", { text: Format.moneyWhole(g.totalBalance), class: "num strong" }));

        const statusTd = Dom.el("td");
        const select = buildStatusSelect(g.customer);
        if (getCallStatus(g.customer)) select.value = getCallStatus(g.customer).action;
        statusTd.appendChild(select);
        tr.appendChild(statusTd);

        const actionsTd = Dom.el("td", { class: "today-call-actions-cell" });
        if (links) {
          actionsTd.appendChild(Dom.el("a", { class: "btn btn-call btn-sm", href: links.call, html: `${I.phone} Call` }));
          actionsTd.appendChild(Dom.el("a", { class: "btn btn-whatsapp btn-sm", href: links.whatsapp, target: "_blank", html: `${I.whatsapp} WhatsApp` }));
        }
        tr.appendChild(actionsTd);
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      tWrap.appendChild(table);
      wrap.appendChild(tWrap);
    }

    if (called.length > 0) {
      const resetBtn = Dom.el("button", { class: "btn btn-sm btn-secondary", text: "Reset all call status" });
      resetBtn.addEventListener("click", () => { resetAllCallStatus(); if (lastRows) renderDashboard(lastRows, lastSource); });
      wrap.appendChild(Dom.el("div", { class: "card-footer", style: "padding:0 20px 16px;" })).appendChild(resetBtn);
    }
    return wrap;
  }

  /* ---------- Status Select ---------- */
  function buildStatusSelect(customer, onChange) {
    const select = Dom.el("select", { class: "input input-sm status-select" });
    [
      { value: "", label: "Mark status..." },
      { value: "called", label: "Called" },
      { value: "promised", label: "Promised to pay" },
      { value: "contacted", label: "Contacted" },
      { value: "no_answer", label: "No answer" },
      { value: "wrong_number", label: "Wrong number" },
      { value: "follow_up", label: "Follow up needed" },
      { value: "paid", label: "Paid" },
    ].forEach((o) => { select.appendChild(Dom.el("option", { value: o.value, text: o.label })); });
    const st = getCallStatus(customer);
    if (st) select.value = st.action;
    select.addEventListener("change", (e) => {
      if (!e.target.value) return;
      if (e.target.value === "promised") {
        showPromiseModal(customer, () => { if (lastRows) renderDashboard(lastRows, lastSource); });
      } else if (e.target.value === "follow_up") {
        showFollowUpModal(customer, () => { if (lastRows) renderDashboard(lastRows, lastSource); });
      } else {
        markCalled(customer, e.target.value);
        if (onChange) onChange();
        else if (lastRows) renderDashboard(lastRows, lastSource);
      }
      e.target.value = getCallStatus(customer)?.action || "";
    });
    return select;
  }

  /* ---------- Promise Modal ---------- */
  function showPromiseModal(customer, onDone) {
    const overlay = Dom.el("div", { class: "modal-overlay" });
    const modal = Dom.el("div", { class: "modal" });
    modal.appendChild(Dom.el("div", { class: "modal-header", html: `<div class="modal-title">Promise to Pay</div><div class="modal-subtitle">${customer}</div>` }));
    const body = Dom.el("div", { class: "modal-body" });

    const amtRow = Dom.el("div", { class: "form-row" });
    amtRow.appendChild(Dom.el("label", { class: "form-label", text: "Promised Amount" }));
    const amtInput = Dom.el("input", { class: "input", type: "number", placeholder: "Enter amount" });
    amtRow.appendChild(amtInput);
    body.appendChild(amtRow);

    const dateRow = Dom.el("div", { class: "form-row" });
    dateRow.appendChild(Dom.el("label", { class: "form-label", text: "Promised Date" }));
    const dateInput = Dom.el("input", { class: "input", type: "date" });
    dateInput.value = new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10);
    dateRow.appendChild(dateInput);
    body.appendChild(dateRow);

    const noteRow = Dom.el("div", { class: "form-row" });
    noteRow.appendChild(Dom.el("label", { class: "form-label", text: "Note (optional)" }));
    const noteInput = Dom.el("input", { class: "input", type: "text", placeholder: "e.g. Will pay by cheque" });
    noteRow.appendChild(noteInput);
    body.appendChild(noteRow);

    modal.appendChild(body);

    const footer = Dom.el("div", { class: "modal-footer" });
    const cancelBtn = Dom.el("button", { class: "btn btn-secondary", text: "Cancel" });
    cancelBtn.addEventListener("click", () => { document.body.removeChild(overlay); });
    footer.appendChild(cancelBtn);
    const saveBtn = Dom.el("button", { class: "btn btn-primary", text: "Save Promise" });
    saveBtn.addEventListener("click", () => {
      markCalled(customer, "promised", {
        promisedAmount: parseFloat(amtInput.value) || 0,
        promisedDate: dateInput.value,
        promiseNote: noteInput.value.trim(),
      });
      document.body.removeChild(overlay);
      onDone();
    });
    footer.appendChild(saveBtn);
    modal.appendChild(footer);
    overlay.appendChild(modal);
    overlay.addEventListener("click", (e) => { if (e.target === overlay) document.body.removeChild(overlay); });
    document.body.appendChild(overlay);
  }

  /* ---------- Follow-Up Modal ---------- */
  function showFollowUpModal(customer, onDone) {
    const overlay = Dom.el("div", { class: "modal-overlay" });
    const modal = Dom.el("div", { class: "modal" });
    modal.appendChild(Dom.el("div", { class: "modal-header", html: `<div class="modal-title">Schedule Follow-Up</div><div class="modal-subtitle">${customer}</div>` }));
    const body = Dom.el("div", { class: "modal-body" });

    const dateRow = Dom.el("div", { class: "form-row" });
    dateRow.appendChild(Dom.el("label", { class: "form-label", text: "Follow-up Date" }));
    const dateInput = Dom.el("input", { class: "input", type: "date" });
    dateInput.value = new Date(Date.now() + 3 * 86400000).toISOString().slice(0, 10);
    dateRow.appendChild(dateInput);
    body.appendChild(dateRow);

    const noteRow = Dom.el("div", { class: "form-row" });
    noteRow.appendChild(Dom.el("label", { class: "form-label", text: "Reason" }));
    const noteInput = Dom.el("input", { class: "input", type: "text", placeholder: "e.g. Customer asked to call back" });
    noteRow.appendChild(noteInput);
    body.appendChild(noteRow);

    modal.appendChild(body);
    const footer = Dom.el("div", { class: "modal-footer" });
    const cancelBtn = Dom.el("button", { class: "btn btn-secondary", text: "Cancel" });
    cancelBtn.addEventListener("click", () => { document.body.removeChild(overlay); });
    footer.appendChild(cancelBtn);
    const saveBtn = Dom.el("button", { class: "btn btn-primary", text: "Schedule" });
    saveBtn.addEventListener("click", () => {
      markCalled(customer, "follow_up", { followUpDate: dateInput.value, followUpNote: noteInput.value.trim() });
      document.body.removeChild(overlay);
      onDone();
    });
    footer.appendChild(saveBtn);
    modal.appendChild(footer);
    overlay.appendChild(modal);
    overlay.addEventListener("click", (e) => { if (e.target === overlay) document.body.removeChild(overlay); });
    document.body.appendChild(overlay);
  }

  /* ---------- Customer Detail Panel ---------- */
  function renderCustomerDetailPanel(g) {
    const panel = Dom.el("div", { class: "detail-panel" });
    const links = phoneLink(g.phone);
    const callSt = getCallStatus(g.customer);
    const hist = getHistory(g.customer);
    const notes = getNotes(g.customer);

    /* --- Header --- */
    const hdr = Dom.el("div", { class: "detail-panel-header" });
    hdr.appendChild(Dom.el("div", { class: "detail-panel-name", text: g.customer }));
    if (g.phone) hdr.appendChild(Dom.el("div", { class: "detail-panel-phone", text: g.phone }));
    hdr.appendChild(Dom.el("div", { class: "detail-panel-balance", html: `<span class="detail-label">Balance</span> ${Format.moneyWhole(g.totalBalance)}` }));
    panel.appendChild(hdr);

    /* --- Actions row --- */
    const actRow = Dom.el("div", { class: "detail-panel-actions" });
    if (links) {
      actRow.appendChild(Dom.el("a", { class: "btn btn-call btn-sm", href: links.call, html: `${I.phone} Call` }));
      actRow.appendChild(Dom.el("a", { class: "btn btn-whatsapp btn-sm", href: links.whatsapp, target: "_blank", html: `${I.whatsapp} WhatsApp` }));
    }
    actRow.appendChild(buildStatusSelect(g.customer));
    panel.appendChild(actRow);

    /* --- Promised Info --- */
    const overduePromises = hist.filter((h) => h.action === "promised" && h.promisedDate && new Date(h.promisedDate) < new Date());
    const activePromises = hist.filter((h) => h.action === "promised" && h.promisedDate && new Date(h.promisedDate) >= new Date());
    if (activePromises.length || overduePromises.length) {
      const promSection = Dom.el("div", { class: "detail-section" });
      promSection.appendChild(Dom.el("div", { class: "detail-section-title", text: "Promises" }));
      activePromises.forEach((p) => {
        promSection.appendChild(Dom.el("div", { class: "promise-item promise-active", html: `
          <div class="promise-info"><strong>${Format.moneyWhole(p.promisedAmount || 0)}</strong> by ${p.promisedDate || "—"}${p.promiseNote ? ` · ${p.promiseNote}` : ""}</div>
          <div class="promise-date">${timeAgo(p.date)}</div>
        `}));
      });
      overduePromises.forEach((p) => {
        promSection.appendChild(Dom.el("div", { class: "promise-item promise-overdue", html: `
          <div class="promise-info"><strong>${Format.moneyWhole(p.promisedAmount || 0)}</strong> was due ${p.promisedDate || "—"} — <span style="color:var(--semantic-danger)">OVERDUE</span>${p.promiseNote ? ` · ${p.promiseNote}` : ""}</div>
          <div class="promise-date">${timeAgo(p.date)}</div>
        `}));
      });
      panel.appendChild(promSection);
    }

    /* --- Follow-Up Info --- */
    const followUps = hist.filter((h) => h.action === "follow_up");
    if (followUps.length) {
      const fuSection = Dom.el("div", { class: "detail-section" });
      fuSection.appendChild(Dom.el("div", { class: "detail-section-title", text: "Follow-ups" }));
      followUps.slice(0, 3).forEach((f) => {
        const isOverdue = f.followUpDate && new Date(f.followUpDate) < new Date();
        fuSection.appendChild(Dom.el("div", { class: `promise-item ${isOverdue ? "promise-overdue" : "promise-active"}`, html: `
          <div class="promise-info">${isOverdue ? `${I.alertTriangle} Overdue: ` : `${I.calendar} `} Follow-up ${f.followUpDate || "—"}${f.followUpNote ? ` · ${f.followUpNote}` : ""}</div>
          <div class="promise-date">${timeAgo(f.date)}</div>
        `}));
      });
      panel.appendChild(fuSection);
    }

    /* --- Call History Timeline --- */
    if (hist.length) {
      const tlSection = Dom.el("div", { class: "detail-section" });
      tlSection.appendChild(Dom.el("div", { class: "detail-section-title", text: `Call History (${hist.length})` }));
      const tl = Dom.el("div", { class: "call-timeline" });
      hist.slice(0, 10).forEach((h) => {
        const actionIcons = { called: I.phone, promised: I.handshake, contacted: I.checkCircle, no_answer: I.phoneOff, wrong_number: I.xCircle, follow_up: I.refreshCw };
        const actionLabels = { called: "Called", promised: "Promised to pay", contacted: "Contacted", no_answer: "No answer", wrong_number: "Wrong number", follow_up: "Follow-up scheduled" };
        tl.appendChild(Dom.el("div", { class: "timeline-item", html: `
          <div class="timeline-icon">${actionIcons[h.action] || "•"}</div>
          <div class="timeline-content">
            <div class="timeline-label">${actionLabels[h.action] || h.action}${h.promisedAmount ? ` — ${Format.moneyWhole(h.promisedAmount)} by ${h.promisedDate || "—"}` : ""}${h.followUpDate ? ` — ${h.followUpDate}` : ""}${h.promiseNote || h.followUpNote ? ` · ${h.promiseNote || h.followUpNote}` : ""}</div>
            <div class="timeline-time">${timeAgo(h.date)}</div>
          </div>
        `}));
      });
      tlSection.appendChild(tl);
      panel.appendChild(tlSection);
    }

    /* --- Remark --- */
    const notesSection = Dom.el("div", { class: "detail-section" });
    notesSection.appendChild(Dom.el("div", { class: "detail-section-title", text: `Remark${notes.length ? ` (${notes.length})` : ""}` }));
    const noteInput = Dom.el("div", { class: "note-input-wrap" });
    const noteText = Dom.el("textarea", { class: "input note-textarea", placeholder: "Add a remark...", rows: "2" });
    noteInput.appendChild(noteText);
    const noteSaveBtn = Dom.el("button", { class: "btn btn-sm btn-primary", text: "Add Remark" });
    noteSaveBtn.addEventListener("click", () => {
      const txt = noteText.value.trim();
      if (!txt) return;
      saveNote(g.customer, txt);
      noteText.value = "";
      if (lastRows) renderDashboard(lastRows, lastSource);
    });
    noteInput.appendChild(noteSaveBtn);
    notesSection.appendChild(noteInput);

    if (notes.length) {
      const noteList = Dom.el("div", { class: "note-list" });
      notes.forEach((n, idx) => {
        const noteItem = Dom.el("div", { class: "note-item" });
        noteItem.appendChild(Dom.el("div", { class: "note-text", text: n.text }));
        noteItem.appendChild(Dom.el("div", { class: "note-meta", html: `${timeAgo(n.date)}` }));
        const delBtn = Dom.el("button", { class: "note-delete", html: "×" });
        delBtn.addEventListener("click", () => { deleteNote(g.customer, idx); if (lastRows) renderDashboard(lastRows, lastSource); });
        noteItem.appendChild(delBtn);
        noteList.appendChild(noteItem);
      });
      notesSection.appendChild(noteList);
    }
    panel.appendChild(notesSection);

    return panel;
  }

  /* ---------- Section: Customer Table ---------- */
  let selectedCustomers = new Set();

  function renderCustomerTable(groups) {
    const pending = groups.filter((g) => g.totalBalance > 0);
    const sorted = sortGroups(pending, currentSort);
    const wrap = Dom.el("div", { class: "card customer-table-card", style: "padding:0;overflow:hidden;" });

    const head = Dom.el("div", { class: "card-header", style: "padding:16px 20px;margin:0;border-bottom:1px solid var(--hairline);" });
    head.appendChild(Dom.el("div", { html: `
      <div class="card-title">Customer Outstanding</div>
      <div class="card-subtitle">${sorted.length} customers · click row to expand bills</div>
    `}));
    wrap.appendChild(head);

    const toolbar = Dom.el("div", { class: "table-toolbar" });
    const searchBox = Dom.el("div", { class: "search-box", html: `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      <input class="input" id="searchInput" type="text" placeholder="Search customer or bill number..." />
    `});
    toolbar.appendChild(searchBox);

    const filterBar = Dom.el("div", { class: "filter-bar" });

    /* --- Date Filter --- */
    const dateFilter = Dom.el("div", { class: "date-filter" });
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let activeDateRange = null;
    function startOfWeek(d) { const dt = new Date(d); const day = dt.getDay(); dt.setDate(dt.getDate() - (day === 0 ? 6 : day - 1)); dt.setHours(0, 0, 0, 0); return dt; }
    function startOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
    function startOfQuarter(d) { return new Date(d.getFullYear(), Math.floor(d.getMonth() / 3) * 3, 1); }
    function fmtDate(d) { return d.toISOString().slice(0, 10); }

    const datePresets = [
      { key: "overdue", label: "Overdue" },
      { key: "this_week", label: "Due this week" },
      { key: "this_month", label: "Due this month" },
      { key: "next_30", label: "Due in 30d" },
      { key: "last_month", label: "Due last month" },
      { key: "this_quarter", label: "Due this quarter" },
      { key: "custom", label: "Custom range" },
    ];
    const dateChips = Dom.el("div", { class: "date-chips" });
    let customRow = null;
    datePresets.forEach((p) => {
      const chip = Dom.el("button", { class: "filter-chip date-chip", text: p.label });
      chip.addEventListener("click", () => {
        dateChips.querySelectorAll(".date-chip").forEach((c) => c.classList.remove("active"));
        chip.classList.add("active");
        if (p.key === "custom") { if (customRow) customRow.style.display = "flex"; activeDateRange = null; }
        else {
          if (customRow) customRow.style.display = "none";
          const ranges = {
            overdue: { from: null, to: today },
            this_week: { from: startOfWeek(today), to: new Date(today.getTime() + 7 * 86400000) },
            this_month: { from: startOfMonth(today), to: new Date(today.getFullYear(), today.getMonth() + 1, 0) },
            next_30: { from: today, to: new Date(today.getTime() + 30 * 86400000) },
            last_month: { from: new Date(today.getFullYear(), today.getMonth() - 1, 1), to: new Date(today.getFullYear(), today.getMonth(), 0) },
            this_quarter: { from: startOfQuarter(today), to: new Date(today.getFullYear(), Math.floor(today.getMonth() / 3) * 3 + 3, 0) },
          };
          activeDateRange = ranges[p.key] || null;
        }
        applyFilters();
      });
      dateChips.appendChild(chip);
    });
    customRow = Dom.el("div", { class: "date-custom-row", style: "display:none;" });
    const fromInput = Dom.el("input", { class: "input input-sm", type: "date" });
    const toInput = Dom.el("input", { class: "input input-sm", type: "date" });
    fromInput.value = fmtDate(new Date(today.getFullYear(), today.getMonth(), 1));
    toInput.value = fmtDate(today);
    customRow.appendChild(Dom.el("span", { class: "date-custom-label", text: "From" }));
    customRow.appendChild(fromInput);
    customRow.appendChild(Dom.el("span", { class: "date-custom-label", text: "To" }));
    customRow.appendChild(toInput);
    const applyCustom = Dom.el("button", { class: "btn btn-sm btn-primary", text: "Apply" });
    applyCustom.addEventListener("click", () => {
      const fromDate = fromInput.value ? new Date(fromInput.value + "T00:00:00") : null;
      const toDate = toInput.value ? new Date(toInput.value + "T23:59:59") : null;
      activeDateRange = { from: fromDate, to: toDate };
      applyFilters();
    });
    customRow.appendChild(applyCustom);
    const clearCustom = Dom.el("button", { class: "btn btn-sm btn-secondary", text: "Clear" });
    clearCustom.addEventListener("click", () => { activeDateRange = null; fromInput.value = ""; toInput.value = ""; dateChips.querySelectorAll(".date-chip").forEach((c) => c.classList.remove("active")); applyFilters(); });
    customRow.appendChild(clearCustom);
    dateFilter.appendChild(dateChips);
    dateFilter.appendChild(customRow);
    filterBar.appendChild(dateFilter);

    /* --- Status Chips --- */
    const filterChips = Dom.el("div", { class: "filter-chips" });
    const activeFilters = new Set();
    const chipDefs = [
      { key: "critical", label: "Critical (90d+)", color: "#d33b3b" },
      { key: "high", label: "High (60d+)", color: "#e07c00" },
      { key: "medium", label: "Medium (30d+)", color: "#b45309" },
      { key: "called", label: "Already Called", color: "#27a644" },
      { key: "not_called", label: "Not Called Yet", color: "#5e6ad2" },
      { key: "follow_up", label: "Follow-up Due", color: "#8b5cf6" },
      { key: "promised", label: "Promised to Pay", color: "#0891b2" },
    ];
    chipDefs.forEach((c) => {
      const chip = Dom.el("button", { class: "filter-chip", html: `<span class="filter-chip-dot" style="background:${c.color}"></span>${c.label}` });
      chip.addEventListener("click", () => {
        if (activeFilters.has(c.key)) { activeFilters.delete(c.key); chip.classList.remove("active"); }
        else { activeFilters.add(c.key); chip.classList.add("active"); }
        applyFilters();
      });
      filterChips.appendChild(chip);
    });
    filterBar.appendChild(filterChips);

    /* --- Sort Buttons --- */
    const sortBtns = Dom.el("div", { class: "sort-btns" });
    [
      { key: "priority", label: "Priority" },
      { key: "overdue", label: "Most overdue" },
      { key: "amount", label: "Amount" },
      { key: "name", label: "Name" },
      { key: "bills", label: "Bill count" },
    ].forEach((s) => {
      const btn = Dom.el("button", {
        class: `btn btn-sm ${currentSort.key === s.key ? "btn-primary" : "btn-secondary"}`,
        text: s.label + (currentSort.key === s.key ? (currentSort.dir === "asc" ? " ↑" : " ↓") : ""),
      });
      btn.addEventListener("click", () => toggleSort(s.key));
      sortBtns.appendChild(btn);
    });
    filterBar.appendChild(sortBtns);

    /* --- Reset Filters --- */
    const resetBtn = Dom.el("button", { class: "btn btn-sm btn-secondary reset-filters-btn", text: "Reset filters" });
    resetBtn.addEventListener("click", () => {
      activeFilters.clear();
      filterChips.querySelectorAll(".filter-chip").forEach((c) => c.classList.remove("active"));
      activeDateRange = null;
      dateChips.querySelectorAll(".date-chip").forEach((c) => c.classList.remove("active"));
      if (customRow) customRow.style.display = "none";
      const searchEl = $("#searchInput");
      if (searchEl) searchEl.value = "";
      currentSort = { key: "priority", dir: "desc" };
      if (lastRows) renderDashboard(lastRows, lastSource);
    });
    filterBar.appendChild(resetBtn);

    toolbar.appendChild(filterBar);

    /* --- Bulk Actions --- */
    const bulkBar = Dom.el("div", { class: "bulk-bar", style: "display:none;" });
    const bulkCount = Dom.el("span", { class: "bulk-count", text: "0 selected" });
    bulkBar.appendChild(bulkCount);
    const bulkActions = Dom.el("div", { class: "bulk-actions" });
    const bulkSelect = Dom.el("select", { class: "input input-sm" });
    [{ value: "", label: "Bulk action..." }, { value: "called", label: "Mark Called" }, { value: "contacted", label: "Mark Contacted" }, { value: "no_answer", label: "Mark No Answer" }, { value: "follow_up", label: "Schedule Follow-up" }].forEach((o) => {
      bulkSelect.appendChild(Dom.el("option", { value: o.value, text: o.label }));
    });
    bulkActions.appendChild(bulkSelect);
    const bulkApplyBtn = Dom.el("button", { class: "btn btn-sm btn-primary", text: "Apply" });
    bulkApplyBtn.addEventListener("click", () => {
      const action = bulkSelect.value;
      if (!action || !selectedCustomers.size) return;
      if (action === "follow_up") {
        showFollowUpModal([...selectedCustomers][0], () => { selectedCustomers.clear(); updateBulkBar(); if (lastRows) renderDashboard(lastRows, lastSource); });
      } else {
        selectedCustomers.forEach((c) => { markCalled(c, action); });
        selectedCustomers.clear();
        updateBulkBar();
        if (lastRows) renderDashboard(lastRows, lastSource);
      }
    });
    bulkActions.appendChild(bulkApplyBtn);
    const bulkClearBtn = Dom.el("button", { class: "btn btn-sm btn-secondary", text: "Clear selection" });
    bulkClearBtn.addEventListener("click", () => { selectedCustomers.clear(); document.querySelectorAll(".row-select:checked").forEach((cb) => { cb.checked = false; }); updateBulkBar(); });
    bulkActions.appendChild(bulkClearBtn);
    bulkBar.appendChild(bulkActions);
    toolbar.appendChild(bulkBar);

    /* --- Export --- */
    const exportBtn = Dom.el("button", { class: "btn btn-sm btn-secondary export-btn", html: `${I.download} Full Report` });
    exportBtn.addEventListener("click", () => exportFilteredCSV(groups));
    toolbar.appendChild(exportBtn);

    /* --- Import --- */
    const importInput = Dom.el("input", { type: "file", accept: ".csv,.txt", hidden: "" });
    const importBtn = Dom.el("button", { class: "btn btn-sm btn-secondary", html: `${I.fileDown} Import Report` });
    importBtn.addEventListener("click", () => {
      if (!confirm("This will replace all existing call status, history, and remarks on this dashboard. Continue?")) return;
      importInput.click();
    });
    importInput.addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const count = importReportCSV(ev.target.result);
          alert(`Imported ${count} entries (call status, history, notes).`);
          if (lastRows) renderDashboard(lastRows, lastSource);
        } catch (err) { alert("Import failed: " + err.message); }
      };
      reader.readAsText(file);
      importInput.value = "";
    });
    toolbar.appendChild(importBtn);

    wrap.appendChild(toolbar);

    /* --- Table --- */
    const tWrap = Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" });
    const table = Dom.el("table", { class: "data-table customer-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["", "", "Customer", "Pending bills", "Total balance", "Max overdue", "Priority", "Status"].forEach((h) => {
      const th = Dom.el("th", { text: h });
      if (h === "Total balance" || h === "Pending bills") th.classList.add("num");
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = Dom.el("tbody");
    sorted.forEach((g) => {
      const row = buildCustomerRow(g);
      tbody.appendChild(row.main);
      tbody.appendChild(row.detail);
    });
    table.appendChild(tbody);
    tWrap.appendChild(table);
    wrap.appendChild(tWrap);

    function updateBulkBar() {
      const bar = toolbar.querySelector(".bulk-bar");
      const count = selectedCustomers.size;
      bar.style.display = count ? "flex" : "none";
      bulkCount.textContent = `${count} selected`;
    }

    function applyFilters() {
      document.querySelectorAll("tr.customer-main").forEach((tr) => {
        const name = tr.dataset.customer || "";
        const bills = tr.dataset.bills || "";
        const maxDays = parseInt(tr.dataset.maxDays || "0", 10);
        const isCalled = tr.dataset.called === "1";
        const hasFollowUp = tr.dataset.followUp === "1";
        const hasPromise = tr.dataset.hasPromise === "1";
        const dueDates = (tr.dataset.dueDates || "").split("|").filter(Boolean);
        const q = ($("#searchInput") || {}).value || "";
        const matchSearch = !q || name.includes(q.toLowerCase()) || bills.includes(q.toLowerCase());

        let matchFilter = activeFilters.size === 0;
        if (activeFilters.has("critical")) matchFilter = matchFilter || maxDays >= 90;
        if (activeFilters.has("high")) matchFilter = matchFilter || (maxDays >= 60 && maxDays < 90);
        if (activeFilters.has("medium")) matchFilter = matchFilter || (maxDays >= 30 && maxDays < 60);
        if (activeFilters.has("called")) matchFilter = matchFilter || isCalled;
        if (activeFilters.has("not_called")) matchFilter = matchFilter || !isCalled;
        if (activeFilters.has("follow_up")) matchFilter = matchFilter || hasFollowUp;
        if (activeFilters.has("promised")) matchFilter = matchFilter || hasPromise;

        let matchDate = true;
        if (activeDateRange) {
          matchDate = false;
          for (const ds of dueDates) {
            const d = new Date(ds + "T00:00:00");
            if (activeDateRange.from && d < activeDateRange.from) continue;
            if (activeDateRange.to && d > activeDateRange.to) continue;
            matchDate = true; break;
          }
        }

        const show = matchSearch && matchFilter && matchDate;
        tr.style.display = show ? "" : "none";
        const detail = tr.nextElementSibling;
        if (detail && detail.classList.contains("customer-detail")) detail.style.display = show ? "" : "none";
      });
    }

    const searchInput = searchBox.querySelector("#searchInput");
    searchInput.addEventListener("input", Dom.debounce(() => applyFilters(), 200));

    return wrap;
  }

  function buildCustomerRow(g) {
    const isOpen = expandedCustomers.has(g.customer);
    const cls = urgencyClass(g.maxDaysOverdue);
    const priority = computePriorityScore(g);
    const callSt = getCallStatus(g.customer);
    const isCalled = !!callSt;
    const hist = getHistory(g.customer);
    const hasFollowUp = hist.some((h) => h.action === "follow_up" && h.followUpDate && new Date(h.followUpDate) >= new Date());
    const hasPromise = hist.some((h) => h.action === "promised");
    const dueDateStrs = g.bills.filter((b) => b.dueDate).map((b) => b.dueDate.toISOString().slice(0, 10));

    const mainTr = Dom.el("tr", {
      class: `customer-main ${cls}`,
      "data-customer": g.customer.toLowerCase(),
      "data-bills": g.bills.map((b) => b.reference.toLowerCase()).join(" "),
      "data-max-days": String(g.maxDaysOverdue || 0),
      "data-called": isCalled ? "1" : "0",
      "data-due-dates": dueDateStrs.join("|"),
      "data-follow-up": hasFollowUp ? "1" : "0",
      "data-has-promise": hasPromise ? "1" : "0",
    });

    /* Checkbox */
    const cb = Dom.el("input", { class: "row-select", type: "checkbox" });
    cb.checked = selectedCustomers.has(g.customer);
    cb.addEventListener("change", (e) => {
      e.stopPropagation();
      if (e.target.checked) selectedCustomers.add(g.customer); else selectedCustomers.delete(g.customer);
    });
    const cbTd = Dom.el("td", { class: "cell-checkbox" });
    cbTd.appendChild(cb);
    mainTr.appendChild(cbTd);

    const chevron = Dom.el("td", { class: "cell-chevron", html: `<span class="chevron ${isOpen ? "open" : ""}">▶</span>` });
    mainTr.appendChild(chevron);

    const custCell = Dom.el("td", { class: "cell-title" });
    custCell.appendChild(Dom.el("span", { text: g.customer }));
    if (g.phone) custCell.appendChild(Dom.el("span", { class: "cell-phone", html: `&nbsp;· ${g.phone}` }));
    if (callSt) custCell.appendChild(Dom.el("span", { html: callStatusBadge(g.customer) }));
    if (hasFollowUp) custCell.appendChild(Dom.el("span", { class: "badge badge-info badge-dot", style: "margin-left:4px;font-size:10px;", text: "Follow-up" }));
    if (hasPromise) custCell.appendChild(Dom.el("span", { class: "badge badge-info badge-dot", style: "margin-left:4px;font-size:10px;", text: "Promised" }));
    mainTr.appendChild(custCell);

    mainTr.appendChild(Dom.el("td", { text: `${g.bills.length} bill${g.bills.length > 1 ? "s" : ""}`, class: "num" }));
    mainTr.appendChild(Dom.el("td", { text: Format.money(g.totalBalance), class: "num strong" }));
    mainTr.appendChild(Dom.el("td", { html: overdueBadge(g.maxDaysOverdue), class: "num" }));
    mainTr.appendChild(Dom.el("td", { html: urgencyLabel(priority.score), class: "num" }));

    mainTr.appendChild(Dom.el("td", {
      html: g.bills.every((b) => { const s = String(b.status || "").toLowerCase(); return s.includes("paid") || s.includes("settled"); })
        ? '<span class="badge badge-success badge-dot">All paid</span>'
        : g.maxDaysOverdue > 0
          ? '<span class="badge badge-danger badge-dot">Needs follow-up</span>'
          : '<span class="badge badge-primary badge-dot">Pending</span>',
    }));

    /* Detail row */
    const detailTr = Dom.el("tr", { class: `customer-detail ${isOpen ? "open" : ""}` });
    const detailTd = Dom.el("td", { colspan: "8" });

    const detailContent = Dom.el("div", { class: "detail-content" });

    /* Bill sub-table */
    const subTable = Dom.el("div", { class: "bill-detail-wrap" });
    const subHead = Dom.el("div", { class: "bill-detail-header" });
    ["Bill No", "Bill date", "Due date", "Bill value", "Paid", "Balance", "Overdue", "Status", "Remarks"].forEach((h) => {
      subHead.appendChild(Dom.el("div", { class: `bill-detail-cell ${["Bill value", "Paid", "Balance"].includes(h) ? "num" : ""}`, text: h }));
    });
    subTable.appendChild(subHead);

    [...g.bills].sort((a, b) => (b.daysOverdue || 0) - (a.daysOverdue || 0)).forEach((b) => {
      const r = Dom.el("div", { class: `bill-detail-row ${urgencyClass(b.daysOverdue)}` });
      r.appendChild(Dom.el("div", { class: "bill-detail-cell cell-title", text: b.reference }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", text: Format.dateShort(b.issueDate) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", text: Format.dateShort(b.dueDate) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell num", text: Format.money(b.amount) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell num", text: Format.money(b.paid) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell num strong", text: Format.money(b.balance) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", html: overdueBadge(b.daysOverdue) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", html: statusBadge(b.status) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", text: b.remarks || "—" }));
      subTable.appendChild(r);
    });
    detailContent.appendChild(subTable);

    /* Customer detail panel (notes, history, promises) */
    detailContent.appendChild(renderCustomerDetailPanel(g));

    detailTd.appendChild(detailContent);
    detailTr.appendChild(detailTd);

    mainTr.addEventListener("click", (e) => {
      if (e.target.type === "checkbox" || e.target.closest(".row-select")) return;
      const chev = mainTr.querySelector(".chevron");
      const isOpenNow = detailTr.classList.toggle("open");
      chev.classList.toggle("open", isOpenNow);
      if (isOpenNow) expandedCustomers.add(g.customer); else expandedCustomers.delete(g.customer);
    });

    return { main: mainTr, detail: detailTr };
  }

  /* ---------- Paid Customers Table ---------- */
  function renderPaidTable(paidGroups) {
    if (!paidGroups.length) return null;
    const totalPaid = paidGroups.reduce((s, g) => s + g.totalAmount, 0);
    const wrap = Dom.el("div", { class: "card paid-table-card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div><div class="card-title">${I.checkCircle} Paid / Settled Customers</div>
      <div class="card-subtitle">${paidGroups.length} customer${paidGroups.length > 1 ? "s" : ""} · ${Format.moneyWhole(totalPaid)} total settled</div></div>
    `}));
    const tWrap = Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" });
    const table = Dom.el("table", { class: "data-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["", "Customer", "Phone", "Bills", "Total Amount", "Settled Status"].forEach((h) => {
      const th = Dom.el("th", { text: h });
      if (h === "Total Amount" || h === "Bills") th.classList.add("num");
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    table.appendChild(thead);
    const tbody = Dom.el("tbody");
    paidGroups.forEach((g) => {
      const st = getCallStatus(g.customer);
      const tr = Dom.el("tr");
      tr.appendChild(Dom.el("td", { class: "cell-chevron", html: `<span style="color:var(--semantic-success);font-size:12px;">${I.checkCircle}</span>` }));
      const custCell = Dom.el("td", { class: "cell-title" });
      custCell.appendChild(Dom.el("span", { text: g.customer }));
      if (g.phone) custCell.appendChild(Dom.el("span", { class: "cell-phone", html: `&nbsp;· ${g.phone}` }));
      tr.appendChild(custCell);
      tr.appendChild(Dom.el("td", { text: g.phone || "—", class: "num", style: "font-family:var(--font-mono);font-size:12px;color:var(--ink-subtle);" }));
      tr.appendChild(Dom.el("td", { text: `${g.bills.length}`, class: "num" }));
      tr.appendChild(Dom.el("td", { text: Format.money(g.totalAmount), class: "num strong" }));
      tr.appendChild(Dom.el("td", { html: st ? callStatusBadge(g.customer) : '<span class="badge badge-success badge-dot">All paid</span>' }));
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    tWrap.appendChild(table);
    wrap.appendChild(tWrap);
    return wrap;
  }

  /* ---------- Purge Button ---------- */
  function renderPurgeBtn() {
    const wrap = Dom.el("div", { class: "purge-wrap" });
    const purgeBtn = Dom.el("button", { class: "btn btn-sm btn-danger", html: `${I.trash} Purge All Data` });
    purgeBtn.addEventListener("click", () => {
      if (!confirm("This will permanently delete ALL call status, history, and remarks for every customer. This cannot be undone. Are you sure?")) return;
      if (!confirm("Last chance — all telecaller data will be erased. Continue?")) return;
      purgeAllData();
      if (lastRows) renderDashboard(lastRows, lastSource);
    });
    wrap.appendChild(purgeBtn);
    return wrap;
  }

  /* ---------- Export ---------- */
  function csvEscape(val) {
    const s = String(val == null ? "" : val);
    if (s.includes(",") || s.includes('"') || s.includes("\n")) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function exportFilteredCSV(groups) {
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    const allNotes = loadNotes();
    const allHistory = loadHistory();
    const allStatus = loadCallStatus();

    /* --- Sheet 1: Customer Summary --- */
    const header1 = [
      "Rank", "Customer", "Phone", "Total Balance", "Bills", "Max Overdue Days",
      "Priority Score", "Urgency", "Current Status", "Last Action", "Last Action Date",
      "Promise Amount", "Promise Due Date", "Follow-Up Date", "Notes Count", "All Notes",
    ].join(",");

    const rows1 = groups.map((g, i) => {
      const p = computePriorityScore(g);
      const st = getCallStatus(g.customer);
      const notes = getNotes(g.customer);
      const hist = getHistory(g.customer);
      const lastAction = hist.length ? hist[0] : null;
      const promises = hist.filter((h) => h.action === "promised");
      const lastPromise = promises.length ? promises[0] : null;
      const followUps = hist.filter((h) => h.action === "follow_up");
      const lastFollowUp = followUps.length ? followUps[0] : null;
      const notesText = notes.map((n) => `${n.text} (${n.date})`).join(" | ");

      return [
        i + 1,
        csvEscape(g.customer),
        csvEscape(g.phone || ""),
        g.totalBalance.toFixed(2),
        g.bills.length,
        g.maxDaysOverdue || 0,
        p.score,
        p.urgency,
        st ? st.action : "not_called",
        lastAction ? lastAction.action : "",
        lastAction ? lastAction.date : "",
        lastPromise ? (lastPromise.promisedAmount || 0).toFixed(2) : "",
        lastPromise ? lastPromise.promisedDate || "" : "",
        lastFollowUp ? lastFollowUp.followUpDate || "" : "",
        notes.length,
        csvEscape(notesText),
      ].join(",");
    });

    /* --- Sheet 2: Full Activity Log (all calls, promises, follow-ups, notes) --- */
    const header2 = [
      "Customer", "Phone", "Activity Type", "Action", "Amount", "Date",
      "Promised Amount", "Promised Date", "Follow-Up Date", "Note", "Timestamp",
    ].join(",");

    const rows2 = [];
    groups.forEach((g) => {
      const hist = getHistory(g.customer);
      hist.forEach((h) => {
        rows2.push([
          csvEscape(g.customer),
          csvEscape(g.phone || ""),
          "Call Status",
          h.action || "",
          "",
          "",
          h.promisedAmount ? h.promisedAmount.toFixed(2) : "",
          h.promisedDate || "",
          h.followUpDate || "",
          csvEscape(h.promiseNote || h.followUpNote || ""),
          h.date || "",
        ].join(","));
      });

      const notes = getNotes(g.customer);
      notes.forEach((n) => {
        rows2.push([
          csvEscape(g.customer),
          csvEscape(g.phone || ""),
          "Note",
          "",
          "",
          "",
          "",
          "",
          "",
          csvEscape(n.text),
          n.date || "",
        ].join(","));
      });
    });

    /* --- Sheet 3: Bill Details (with Excel remarks) --- */
    const header3 = [
      "Customer", "Phone", "Bill No", "Bill Date", "Due Date", "Amount",
      "Paid", "Balance", "Days Overdue", "Aging Bucket", "Status", "Call Status", "Bill Remarks", "Notes",
    ].join(",");

    const rows3 = [];
    groups.forEach((g) => {
      const st = getCallStatus(g.customer);
      const notes = getNotes(g.customer).map((n) => n.text).join("; ");
      g.bills.forEach((b) => {
        rows3.push([
          csvEscape(g.customer),
          csvEscape(g.phone || ""),
          csvEscape(b.reference),
          b.issueDate ? b.issueDate.toISOString().slice(0, 10) : "",
          b.dueDate ? b.dueDate.toISOString().slice(0, 10) : "",
          b.amount.toFixed(2),
          b.paid.toFixed(2),
          b.balance.toFixed(2),
          b.daysOverdue != null ? b.daysOverdue : "",
          b.bucket || "",
          b.status || "",
          st ? st.action : "",
          csvEscape(b.remarks || ""),
          csvEscape(notes),
        ].join(","));
      });
    });

    /* --- Build multi-sheet CSV as a single file with markers --- */
    let csv = `--- CUSTOMER SUMMARY ---\n`;
    csv += header1 + "\n";
    csv += rows1.join("\n") + "\n\n";
    csv += `--- FULL ACTIVITY LOG (Calls + Notes) ---\n`;
    csv += header2 + "\n";
    csv += rows2.join("\n") + "\n\n";
    csv += `--- BILL DETAILS ---\n`;
    csv += header3 + "\n";
    csv += rows3.join("\n");

    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `receivables-full-report-${dateStr}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function importReportCSV(text) {
    const sections = text.split(/^--- .+ ---$/m).map((s) => s.trim()).filter(Boolean);
    if (sections.length < 2) throw new Error("Invalid report format. Expected multiple sections.");
    const activitySection = sections[1];
    const lines = activitySection.split("\n").filter((l) => l.trim());
    if (lines.length < 2) throw new Error("Activity log section is empty.");
    const headers = parseCSVLine(lines[0]);
    const status = {};
    const hist = {};
    const notes = {};
    let imported = 0;
    for (let i = 1; i < lines.length; i++) {
      const cells = parseCSVLine(lines[i]);
      if (cells.length < 6) continue;
      const customer = cells[0];
      if (!customer) continue;
      const activityType = cells[2];
      const action = cells[3];
      const timestamp = cells[10] || "";
      if (activityType === "Call Status" && action) {
        if (!status[customer] || (timestamp && new Date(timestamp) > new Date(status[customer].date || 0))) {
          status[customer] = {
            action,
            date: timestamp,
            ...(cells[6] ? { promisedAmount: parseFloat(cells[6]) || 0 } : {}),
            ...(cells[7] ? { promisedDate: cells[7] } : {}),
            ...(cells[8] ? { followUpDate: cells[8] } : {}),
            ...(cells[9] && action === "promised" ? { promiseNote: cells[9] } : {}),
            ...(cells[9] && action === "follow_up" ? { followUpNote: cells[9] } : {}),
          };
        }
        if (!hist[customer]) hist[customer] = [];
        const entry = { action, date: timestamp };
        if (cells[6]) entry.promisedAmount = parseFloat(cells[6]) || 0;
        if (cells[7]) entry.promisedDate = cells[7];
        if (cells[8]) entry.followUpDate = cells[8];
        if (cells[9]) { if (action === "promised") entry.promiseNote = cells[9]; else entry.followUpNote = cells[9]; }
        hist[customer].push(entry);
        imported++;
      } else if (activityType === "Note" && cells[9]) {
        if (!notes[customer]) notes[customer] = [];
        notes[customer].push({ text: cells[9], date: timestamp });
        imported++;
      }
    }
    saveCallStatus(status);
    saveJSON(HISTORY_KEY, hist);
    saveJSON(NOTES_KEY, notes);
    return imported;
  }

  function parseCSVLine(line) {
    const cells = []; let i = 0;
    while (i < line.length) {
      if (line[i] === '"') {
        let j = i + 1, val = "";
        while (j < line.length) {
          if (line[j] === '"') { if (j + 1 < line.length && line[j + 1] === '"') { val += '"'; j += 2; } else { j++; break; } }
          else { val += line[j]; j++; }
        }
        cells.push(val); i = j + 1; if (i < line.length && line[i] === ",") i++;
      } else {
        let j = line.indexOf(",", i); if (j === -1) j = line.length;
        cells.push(line.slice(i, j).trim()); i = j + 1;
      }
    }
    return cells;
  }

  /* ---------- Aging Analysis ---------- */
  function renderAgingChart(rows) {
    const buckets = AGING_BUCKETS.map((b) => ({ label: b.label, value: rows.filter((r) => r.bucket === b.label).reduce((s, r) => s + r.balance, 0) }));
    const segs = buckets.map((b, i) => ({ ...b, color: AGING_COLORS[i] }));
    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `<div><div class="card-title">Aging summary</div><div class="card-subtitle">Outstanding balance by age bucket</div></div>` }));
    const body = Dom.el("div", { class: "grid-2", style: "margin-bottom:0;align-items:center;" });
    body.appendChild(Dom.el("div", { class: "chart-wrap", style: "display:flex;justify-content:center;" }));
    body.lastChild.appendChild(Charts.donutChart({ segments: segs, size: 240, thickness: 30, format: (v) => Format.moneyWhole(v) }));
    const legend = Dom.el("div", { class: "legend" });
    segs.forEach((s, i) => {
      legend.appendChild(Dom.el("div", { class: "legend-item", html: `<span class="legend-swatch" style="background:${AGING_COLORS[i]}"></span>${s.label} · ${Format.moneyWhole(s.value)}` }));
    });
    body.appendChild(legend); wrap.appendChild(body); return wrap;
  }

  function renderBarChart(rows) {
    const topCustomers = [...rows].filter((r) => r.balance > 0).reduce((acc, r) => {
      const k = r.customer; acc[k] = acc[k] || { label: k, value: 0 }; acc[k].value += r.balance; return acc;
    }, {});
    const topSorted = Object.values(topCustomers).sort((a, b) => b.value - a.value).slice(0, 8);
    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `<div><div class="card-title">Top customers by balance</div><div class="card-subtitle">Outstanding amount by customer</div></div>` }));
    wrap.appendChild(Dom.el("div", { class: "chart-wrap" }));
    wrap.lastChild.appendChild(Charts.barChart({ items: topSorted, height: 280, horizontal: true, format: (v) => Format.moneyWhole(v) }));
    return wrap;
  }

  function renderAgingAnalysis(rows) {
    const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);
    const overdueTotal = rows.filter((r) => r.daysOverdue > 0).reduce((s, r) => s + r.balance, 0);
    const data = AGING_BUCKETS.map((b, i) => {
      const matching = rows.filter((r) => r.bucket === b.label);
      const total = matching.reduce((s, r) => s + r.balance, 0);
      const count = matching.length;
      const pct = totalOutstanding ? (total / totalOutstanding) * 100 : 0;
      const overduePct = overdueTotal ? (total / overdueTotal) * 100 : 0;
      const avgDays = matching.length ? Math.round(matching.reduce((s, r) => s + (r.daysOverdue || 0), 0) / matching.length) : 0;
      return { ...b, color: AGING_COLORS[i], total, count, pct, overduePct, avgDays };
    });
    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `<div><div class="card-title">Advanced Aging Analysis</div><div class="card-subtitle">Detailed breakdown by age bucket</div></div>` }));
    const summaryRow = Dom.el("div", { class: "aging-summary-row" });
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `<span class="aging-summary-label">Total Outstanding</span><span class="aging-summary-value">${Format.money(totalOutstanding)}</span>` }));
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `<span class="aging-summary-label">Overdue Amount</span><span class="aging-summary-value danger">${Format.money(overdueTotal)}</span>` }));
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `<span class="aging-summary-label">Total Bills</span><span class="aging-summary-value">${rows.length}</span>` }));
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `<span class="aging-summary-label">Overdue Bills</span><span class="aging-summary-value danger">${rows.filter((r) => r.daysOverdue > 0).length}</span>` }));
    wrap.appendChild(summaryRow);
    const stackedBar = Dom.el("div", { class: "aging-stacked-bar" });
    data.forEach((d) => { if (d.total > 0) stackedBar.appendChild(Dom.el("div", { class: "aging-stacked-seg", style: `width:${d.pct}%;background:${d.color};`, title: `${d.label}: ${Format.money(d.total)} (${d.pct.toFixed(1)}%)` })); });
    wrap.appendChild(stackedBar);
    const table = Dom.el("table", { class: "data-table aging-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["Aging Bucket", "Bills", "Amount", "% of Total", "% of Overdue", "Avg Days", "Distribution"].forEach((h) => {
      const th = Dom.el("th", { text: h }); if (h !== "Aging Bucket") th.classList.add("num"); hr.appendChild(th);
    });
    thead.appendChild(hr); table.appendChild(thead);
    const tbody = Dom.el("tbody");
    data.forEach((d) => {
      const tr = Dom.el("tr");
      const bucketLabel = Dom.el("td", { class: "cell-title" });
      bucketLabel.appendChild(Dom.el("span", { class: "aging-dot", style: `background:${d.color};` }));
      bucketLabel.appendChild(document.createTextNode(` ${d.label}`));
      tr.appendChild(bucketLabel);
      tr.appendChild(Dom.el("td", { text: `${d.count}`, class: "num" }));
      tr.appendChild(Dom.el("td", { text: Format.money(d.total), class: "num strong" }));
      tr.appendChild(Dom.el("td", { text: `${d.pct.toFixed(1)}%`, class: "num" }));
      tr.appendChild(Dom.el("td", { text: `${d.overduePct.toFixed(1)}%`, class: "num" }));
      tr.appendChild(Dom.el("td", { text: d.avgDays > 0 ? `${d.avgDays}d` : "—", class: "num" }));
      const barCell = Dom.el("td", { class: "aging-bar-cell" });
      const barWrap = Dom.el("div", { class: "aging-bar-wrap" });
      barWrap.appendChild(Dom.el("div", { class: "aging-bar-fill", style: `width:${d.pct}%;background:${d.color};` }));
      barCell.appendChild(barWrap); tr.appendChild(barCell);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" })).appendChild(table);
    return wrap;
  }

  function renderCustomerAgingHeatmap(rows, groups) {
    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `<div><div class="card-title">Customer Aging Heatmap</div><div class="card-subtitle">Top 15 customers by balance across aging buckets</div></div>` }));
    const topGroups = [...groups].filter((g) => g.totalBalance > 0).sort((a, b) => b.totalBalance - a.totalBalance).slice(0, 15);
    const table = Dom.el("table", { class: "data-table heatmap-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    hr.appendChild(Dom.el("th", { text: "Customer" }));
    hr.appendChild(Dom.el("th", { text: "Total", class: "num" }));
    AGING_BUCKETS.forEach((b, i) => { hr.appendChild(Dom.el("th", { text: b.label, class: "num", style: `color:${AGING_COLORS[i]};` })); });
    thead.appendChild(hr); table.appendChild(thead);
    const tbody = Dom.el("tbody");
    topGroups.forEach((g) => {
      const tr = Dom.el("tr");
      const nameCell = Dom.el("td", { class: "cell-title" });
      nameCell.textContent = g.customer.length > 28 ? g.customer.slice(0, 26) + "…" : g.customer;
      tr.appendChild(nameCell);
      tr.appendChild(Dom.el("td", { text: Format.moneyWhole(g.totalBalance), class: "num strong" }));
      AGING_BUCKETS.forEach((b, i) => {
        const bucketBalance = g.bills.filter((r) => r.bucket === b.label).reduce((s, r) => s + r.balance, 0);
        const cell = Dom.el("td", { class: "num heatmap-cell" });
        if (bucketBalance > 0) {
          const intensity = g.totalBalance ? bucketBalance / g.totalBalance : 0;
          cell.appendChild(Dom.el("span", {
            class: "heatmap-value",
            style: `background:${AGING_COLORS[i]};opacity:${0.15 + intensity * 0.85};color:${intensity > 0.4 ? "#fff" : AGING_COLORS[i]};`,
            text: Format.moneyWhole(bucketBalance),
          }));
        } else cell.textContent = "—";
        tr.appendChild(cell);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" })).appendChild(table);
    return wrap;
  }

  /* ---------- Dashboard ---------- */
  let lastRows = null;
  let lastSource = null;
  const expandedCustomers = new Set();

  function renderDashboard(rows, sourceLabel) {
    lastRows = rows; lastSource = sourceLabel;
    const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);
    const groups = groupByCustomer(rows);
    const pendingGroups = groups.filter((g) => g.totalBalance > 0);
    const paidGroups = groups.filter((g) => g.totalBalance <= 0);
    const dash = $("#dashboard"); Dom.clear(dash);
    if (sourceLabel) dash.appendChild(Dom.el("div", { class: "source-badge", html: `<span class="badge badge-primary">${sourceLabel}</span>` }));
    dash.appendChild(renderGuide());
    dash.appendChild(renderKpis(rows, totalOutstanding, groups));
    dash.appendChild(renderCallingStatusKpis(groups));
    dash.appendChild(renderDailySummary(rows, groups));
    dash.appendChild(renderTodayCallList(groups));
    dash.appendChild(renderCustomerTable(groups));
    const paidTable = renderPaidTable(paidGroups);
    if (paidTable) dash.appendChild(paidTable);
    dash.appendChild(renderPurgeBtn());
    const charts = Dom.el("div", { class: "grid-2" });
    charts.appendChild(renderAgingChart(rows));
    charts.appendChild(renderBarChart(rows));
    dash.appendChild(charts);
    dash.appendChild(renderAgingAnalysis(rows));
    dash.appendChild(renderCustomerAgingHeatmap(rows, groups));
    const now = new Date();
    $("#lastUpdated").textContent = `Updated ${now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
  }

  function renderKpis(rows, totalOutstanding, groups) {
    const overdueRows = rows.filter((r) => r.daysOverdue !== null && r.daysOverdue > 0);
    const overdueTotal = overdueRows.reduce((s, r) => s + r.balance, 0);
    const days90 = rows.filter((r) => r.bucketDays !== null && r.bucketDays >= 91).reduce((s, r) => s + r.balance, 0);
    const customersWithDue = groups.filter((g) => g.totalBalance > 0).length;
    const pctOverdue = totalOutstanding ? (overdueTotal / totalOutstanding) * 100 : 0;
    const grid = Dom.el("div", { class: "stat-grid" });
    grid.appendChild(statCard("Total outstanding", Format.moneyWhole(totalOutstanding)));
    grid.appendChild(statCard("Customers pending", `<span class="mono">${customersWithDue}</span>`, { delta: `<span>${rows.length} total bills</span>` }));
    grid.appendChild(statCard("Total overdue", Format.moneyWhole(overdueTotal), { delta: `<span class="down">${pctOverdue.toFixed(1)}% of outstanding</span>` }));
    grid.appendChild(statCard("High risk (60d+)", Format.moneyWhole(days90)));
    return grid;
  }

  function renderCallingStatusKpis(groups) {
    const status = loadCallStatus();
    const hist = loadHistory();
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const allEntries = Object.values(hist).flat();
    const todayEntries = allEntries.filter((e) => new Date(e.date) >= today);

    const statusCounts = { called: 0, promised: 0, contacted: 0, no_answer: 0, wrong_number: 0, follow_up: 0, paid: 0, not_called: 0 };
    const statusAmounts = { called: 0, promised: 0, contacted: 0, no_answer: 0, wrong_number: 0, follow_up: 0, paid: 0, not_called: 0 };
    const pending = groups.filter((g) => g.totalBalance > 0);

    pending.forEach((g) => {
      const st = status[g.customer];
      const key = st ? st.action : "not_called";
      if (statusCounts.hasOwnProperty(key)) {
        statusCounts[key]++;
        statusAmounts[key] += g.totalBalance;
      } else {
        statusCounts.not_called++;
        statusAmounts.not_called += g.totalBalance;
      }
    });

    const todayPromises = todayEntries.filter((e) => e.action === "promised");
    const todayPromisedAmt = todayPromises.reduce((s, e) => s + (e.promisedAmount || 0), 0);
    const overduePromises = allEntries.filter((e) => e.action === "promised" && e.promisedDate && new Date(e.promisedDate) < today);
    const dueFollowUps = allEntries.filter((e) => e.action === "follow_up" && e.followUpDate && new Date(e.followUpDate) <= today);

    const grid = Dom.el("div", { class: "stat-grid calling-status-grid" });
    const items = [
      { label: "Not called yet", value: statusCounts.not_called, amount: statusAmounts.not_called, color: "var(--primary)", icon: I.phone },
      { label: "Called", value: statusCounts.called, amount: statusAmounts.called, color: "var(--semantic-success)", icon: I.checkCircle },
      { label: "Contacted", value: statusCounts.contacted, amount: statusAmounts.contacted, color: "#0891b2", icon: I.users },
      { label: "Promised to pay", value: statusCounts.promised, amount: statusAmounts.promised, color: "var(--semantic-warning)", icon: I.handshake },
      { label: "Follow-up due", value: dueFollowUps.length, amount: 0, color: "#8b5cf6", icon: I.calendar },
      { label: "Paid", value: statusCounts.paid, amount: statusAmounts.paid, color: "var(--semantic-success)", icon: I.banknote },
      { label: "No answer", value: statusCounts.no_answer, amount: statusAmounts.no_answer, color: "#a8abb3", icon: I.phoneOff },
      { label: "Wrong number", value: statusCounts.wrong_number, amount: statusAmounts.wrong_number, color: "#a8abb3", icon: I.xCircle },
    ];
    items.forEach((item) => {
      grid.appendChild(Dom.el("div", { class: "stat-card calling-status-card", html: `
        <div class="stat-label" style="color:${item.color}">${item.icon} ${item.label}</div>
        <div class="stat-value">${item.value}</div>
        ${item.amount > 0 ? `<div class="stat-delta">${Format.moneyWhole(item.amount)}</div>` : ""}
      `}));
    });
    return grid;
  }

  function renderPlaceholder() {
    const banner = Dom.el("div", { class: "status-banner warning", html: `<strong>Data source not configured</strong><span>&nbsp;—&nbsp;Configure the Google Sheet URL in <code>assets/js/config.js</code> or upload a file.</span>` });
    $("#statusBanner").appendChild(banner);
    const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
    box.appendChild(Dom.el("div", { class: "state-icon", html: I.phone }));
    box.appendChild(Dom.el("h3", { text: "No receivables data loaded" }));
    box.appendChild(Dom.el("p", { text: "Configure the Google Sheet or upload a file to start tracking outstanding payments." }));
    box.appendChild(Dom.el("a", { class: "btn btn-secondary", href: "index.html", text: "Back to home" }));
    $("#dashboard").appendChild(box);
  }

  function renderError(err) {
    $("#dashboard").innerHTML = "";
    const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
    box.appendChild(Dom.el("div", { class: "state-icon", html: I.alertTriangle }));
    box.appendChild(Dom.el("h3", { text: "Could not load receivables data" }));
    box.appendChild(Dom.el("p", { text: err.message }));
    const retry = Dom.el("button", { class: "btn btn-secondary", text: "Try again" });
    retry.addEventListener("click", loadFromSheet);
    box.appendChild(retry);
    $("#dashboard").appendChild(box);
  }

  async function loadFromSheet() {
    const refreshBtn = $("#refreshBtn"); refreshBtn.classList.add("loading");
    try {
      if (!SHEET.sheetUrl) { renderPlaceholder(); return; }
      $("#statusBanner").innerHTML = "";
      const records = await Sheets.getRows(SHEET.sheetUrl, SHEET.gid);
      if (!records.length) throw new Error("The sheet contains no data rows.");
      const headers = Object.keys(records[0]); const cols = matchHeader(headers);
      if (!cols.amount || !cols.customer) throw new Error("Could not map columns. Expected at least customer and amount columns.");
      const rows = buildRows(records, cols);
      renderDashboard(rows, "Google Sheets");
    } catch (err) { renderError(err); } finally { refreshBtn.classList.remove("loading"); }
  }

  async function loadFromCSV(file) {
    try {
      $("#statusBanner").innerHTML = "";
      const { records, cols } = await loadFile(file);
      const rows = buildRows(records, cols);
      renderDashboard(rows, `Uploaded: ${file.name}`);
    } catch (err) { renderError(err); }
  }

  /* ---------- Wire up ---------- */
  const fileInput = $("#fileInput");
  const uploadBtn = $("#uploadBtn");
  uploadBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => { const file = e.target.files && e.target.files[0]; if (file) loadFromCSV(file); fileInput.value = ""; });
  $("#refreshBtn").addEventListener("click", loadFromSheet);
  document.addEventListener("click", (e) => {
    const toggle = e.target.closest("#guideToggle");
    if (toggle) { const body = document.getElementById("guideBody"); const chev = toggle.querySelector(".guide-chevron"); body.classList.toggle("open"); chev.classList.toggle("open"); }
  });
  loadFromSheet();
})();

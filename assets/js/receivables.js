/* ============================================================
   Accounts Receivable — Telecaller Command Center.
   Customer-grouped table, live search, sorting, expandable bills,
   call status tracking, priority scoring, aging analysis.
   ============================================================ */

(() => {
  const $ = (sel) => document.querySelector(sel);
  const STORAGE_KEY = "receivables_call_status";

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
  };

  function matchHeader(headers) {
    const map = {};
    const normalized = headers.map((h) => String(h).toLowerCase().trim());
    for (const [key, candidates] of Object.entries(headerMap)) {
      for (const cand of candidates) {
        const idx = normalized.indexOf(cand);
        if (idx !== -1) {
          map[key] = headers[idx];
          break;
        }
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
    const today = new Date();
    today.setHours(0, 0, 0, 0);
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
        if (issue) {
          dueDate = new Date(issue);
          dueDate.setDate(dueDate.getDate() + 30);
        }
      }

      const issueDate = cols.issueDate ? Format.parseDate(r[cols.issueDate]) : null;
      const daysOverdue = dueDate ? Math.round((today - dueDate) / 86400000) : null;
      const age = dueDate ? ageBucket(dueDate, today) : null;

      rows.push({
        reference: r[cols.reference] || "—",
        customer: r[cols.customer] || "—",
        phone: cols.phone ? (r[cols.phone] || "").trim() : "",
        issueDate,
        dueDate,
        amount,
        paid: paidN,
        balance,
        status: (r[cols.status] || "").trim(),
        daysOverdue,
        bucket: age ? age.bucket.label : null,
        bucketDays: age ? age.days : null,
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
      const cells = [];
      let i = 0;
      while (i < line.length) {
        if (line[i] === '"') {
          let j = i + 1;
          let val = "";
          while (j < line.length) {
            if (line[j] === '"') {
              if (j + 1 < line.length && line[j + 1] === '"') {
                val += '"';
                j += 2;
              } else {
                j++;
                break;
              }
            } else {
              val += line[j];
              j++;
            }
          }
          cells.push(val);
          i = j + 1;
          if (i < line.length && line[i] === ",") i++;
        } else {
          let j = line.indexOf(",", i);
          if (j === -1) j = line.length;
          cells.push(line.slice(i, j).trim());
          i = j + 1;
        }
      }
      rows.push(cells);
    }

    const headers = rows[0].map((h) => h.replace(/^\uFEFF/, "").trim());
    return rows.slice(1).map((cells) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = cells[i] || ""; });
      return obj;
    });
  }

  async function loadFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    let records;

    if (ext === "xlsx" || ext === "xls" || ext === "ods") {
      const buf = await file.arrayBuffer();
      records = parseExcel(buf);
    } else {
      const text = await file.text();
      records = parseCSV(text);
    }

    if (!records.length) throw new Error("The file contains no data rows.");

    const headers = Object.keys(records[0]);
    const cols = matchHeader(headers);
    if (!cols.amount || !cols.customer) {
      throw new Error("Could not map columns. Expected at least customer and amount columns (e.g. Client Name, Amount).");
    }

    return { records, cols };
  }

  /* ---------- Call Status (localStorage) ---------- */

  function loadCallStatus() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {};
    } catch {
      return {};
    }
  }

  function saveCallStatus(status) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(status));
  }

  function markCalled(customer, action) {
    const status = loadCallStatus();
    status[customer] = {
      action,
      date: new Date().toISOString(),
    };
    saveCallStatus(status);
  }

  function getCallStatus(customer) {
    const status = loadCallStatus();
    return status[customer] || null;
  }

  function resetAllCallStatus() {
    localStorage.removeItem(STORAGE_KEY);
  }

  /* ---------- Grouping ---------- */

  function groupByCustomer(rows) {
    const map = {};
    for (const r of rows) {
      const name = r.customer;
      if (!map[name]) {
        map[name] = {
          customer: name,
          phone: r.phone || "",
          bills: [],
          totalAmount: 0,
          totalPaid: 0,
          totalBalance: 0,
          maxDaysOverdue: null,
          latestDueDate: null,
        };
      }
      const g = map[name];
      g.bills.push(r);
      g.totalAmount += r.amount;
      g.totalPaid += r.paid;
      g.totalBalance += r.balance;
      g.phone = g.phone || r.phone || "";
      if (r.daysOverdue !== null) {
        if (g.maxDaysOverdue === null || r.daysOverdue > g.maxDaysOverdue) {
          g.maxDaysOverdue = r.daysOverdue;
        }
      }
      if (r.dueDate && (!g.latestDueDate || r.dueDate < g.latestDueDate)) {
        g.latestDueDate = r.dueDate;
      }
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

    const score = Math.round(
      (overdueAmount / 1000) * 2 +
      maxOverdue * 1.5 +
      billCount * 10 +
      overdueRatio * 50
    );

    return {
      overdueAmount,
      overdueBillCount: overdueBills.length,
      score,
      urgency: score >= 80 ? "critical" : score >= 50 ? "high" : score >= 25 ? "medium" : "low",
    };
  }

  /* ---------- Sorting ---------- */

  let currentSort = { key: "priority", dir: "desc" };

  function sortGroups(groups, sort) {
    const sorted = [...groups];
    switch (sort.key) {
      case "name":
        sorted.sort((a, b) => a.customer.localeCompare(b.customer));
        break;
      case "amount":
        sorted.sort((a, b) => a.totalBalance - b.totalBalance);
        break;
      case "overdue":
        sorted.sort((a, b) => (b.maxDaysOverdue || 0) - (a.maxDaysOverdue || 0));
        break;
      case "bills":
        sorted.sort((a, b) => b.bills.length - a.bills.length);
        break;
      case "priority":
        sorted.sort((a, b) => computePriorityScore(b).score - computePriorityScore(a).score);
        break;
    }
    if (sort.dir === "desc" && sort.key !== "overdue" && sort.key !== "bills" && sort.key !== "priority") {
      sorted.reverse();
    }
    return sorted;
  }

  function toggleSort(key) {
    if (currentSort.key === key) {
      currentSort.dir = currentSort.dir === "asc" ? "desc" : "asc";
    } else {
      currentSort = { key, dir: key === "overdue" || key === "bills" || key === "priority" ? "desc" : "asc" };
    }
    if (lastRows) renderDashboard(lastRows, lastSource);
  }

  /* ---------- Rendering Helpers ---------- */

  function statusBadge(status) {
    const s = String(status || "").toLowerCase();
    if (s.includes("paid") || s.includes("settled") || s.includes("closed")) {
      return '<span class="badge badge-success badge-dot">Paid</span>';
    }
    if (s.includes("partial")) {
      return '<span class="badge badge-warning badge-dot">Partial</span>';
    }
    if (s.includes("overdue") || s.includes("past due")) {
      return '<span class="badge badge-danger badge-dot">Overdue</span>';
    }
    if (s.includes("open") || s.includes("unpaid") || s.includes("pending") || s.includes("due")) {
      return '<span class="badge badge-primary badge-dot">Open</span>';
    }
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
    const actionLabels = {
      called: "📞 Called",
      promised: "🤝 Promised to pay",
      contacted: "✅ Contacted",
      no_answer: "📵 No answer",
      wrong_number: "❌ Wrong number",
      follow_up: "🔄 Follow up",
    };
    const label = actionLabels[st.action] || st.action;
    const date = new Date(st.date);
    const timeStr = date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    return `<span class="call-status-badge">${label} <span class="call-status-time">${timeStr}</span></span>`;
  }

  function statCard(label, value, opts = {}) {
    return Dom.el("div", { class: "stat-card", html: `
      <div class="stat-label">${label}</div>
      <div class="stat-value">${value}</div>
      ${opts.delta ? `<div class="stat-delta">${opts.delta}</div>` : ""}
    `});
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
          <div class="guide-step">
            <div class="guide-step-num">1</div>
            <div class="guide-step-content">
              <strong>Check Today's Call List</strong>
              <span>Top section shows who to call first, sorted by urgency.</span>
            </div>
          </div>
          <div class="guide-step">
            <div class="guide-step-num">2</div>
            <div class="guide-step-content">
              <strong>Call &amp; Update Status</strong>
              <span>Click the action buttons to mark customers as called, promised to pay, or follow-up needed.</span>
            </div>
          </div>
          <div class="guide-step">
            <div class="guide-step-num">3</div>
            <div class="guide-step-content">
              <strong>Track Progress</strong>
              <span>The progress bar shows how many customers you've contacted today.</span>
            </div>
          </div>
          <div class="guide-step">
            <div class="guide-step-num">4</div>
            <div class="guide-step-content">
              <strong>Export &amp; Share</strong>
              <span>Download your priority list or filtered results as a file.</span>
            </div>
          </div>
        </div>
      </div>
    `;
    return wrap;
  }

  /* ---------- Section: KPI Cards ---------- */

  function renderKpis(rows, totalOutstanding, groups) {
    const overdueRows = rows.filter((r) => r.daysOverdue !== null && r.daysOverdue > 0);
    const overdueTotal = overdueRows.reduce((s, r) => s + r.balance, 0);
    const currentTotal = rows.filter((r) => r.daysOverdue !== null && r.daysOverdue <= 0).reduce((s, r) => s + r.balance, 0);
    const days90 = rows.filter((r) => r.bucketDays !== null && r.bucketDays >= 91).reduce((s, r) => s + r.balance, 0);
    const customersWithDue = groups.filter((g) => g.totalBalance > 0).length;

    const pctOverdue = totalOutstanding ? (overdueTotal / totalOutstanding) * 100 : 0;

    const grid = Dom.el("div", { class: "stat-grid" });
    grid.appendChild(statCard("Total outstanding", Format.moneyWhole(totalOutstanding)));
    grid.appendChild(statCard("Customers pending", `<span class="mono">${customersWithDue}</span>`, {
      delta: `<span>${rows.length} total bills</span>`,
    }));
    grid.appendChild(statCard("Total overdue", Format.moneyWhole(overdueTotal), {
      delta: `<span class="down">${pctOverdue.toFixed(1)}% of outstanding</span>`,
    }));
    grid.appendChild(statCard("High risk (60d+)", Format.moneyWhole(days90)));
    return grid;
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
      <div>
        <div class="card-title">Today's Call List</div>
        <div class="card-subtitle">Top priority customers to call — ${called.length}/${pending.length} contacted</div>
      </div>
      <div class="today-progress">
        <div class="today-progress-bar">
          <div class="today-progress-fill" style="width:${progressPct}%"></div>
        </div>
        <span class="today-progress-label">${progressPct}%</span>
      </div>
    `}));

    const list = Dom.el("div", { class: "today-call-list" });
    const toShow = notCalled.slice(0, 8);
    if (!toShow.length) {
      list.appendChild(Dom.el("div", { class: "today-call-empty", html: `
        <div class="guide-step-num">✓</div>
        <span>All customers contacted today! Great work.</span>
      `}));
    } else {
      toShow.forEach((g, i) => {
        const item = Dom.el("div", { class: `today-call-item urgency-${g.priority.urgency}` });
        const phoneNum = g.phone.replace(/[^0-9+]/g, "");

        const rankEl = Dom.el("div", { class: "today-call-rank", text: `${i + 1}` });

        const info = Dom.el("div", { class: "today-call-info" });
        info.appendChild(Dom.el("div", { class: "today-call-name", text: g.customer }));
        const metaParts = [];
        if (g.phone) metaParts.push(g.phone);
        metaParts.push(Format.moneyWhole(g.totalBalance));
        if (g.maxDaysOverdue > 0) metaParts.push(`${g.maxDaysOverdue}d overdue`);
        metaParts.push(`${g.bills.length} bill${g.bills.length > 1 ? "s" : ""}`);
        info.appendChild(Dom.el("div", { class: "today-call-meta", text: metaParts.join(" · ") }));

        const actions = Dom.el("div", { class: "today-call-actions" });
        if (phoneNum) {
          const callBtn = Dom.el("a", {
            class: "btn btn-call",
            href: `tel:${phoneNum}`,
            html: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07 19.5 19.5 0 01-6-6 19.79 19.79 0 01-3.07-8.67A2 2 0 014.11 2h3a2 2 0 012 1.72c.127.96.361 1.903.7 2.81a2 2 0 01-.45 2.11L8.09 9.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0122 16.92z"/></svg> Call',
          });
          actions.appendChild(callBtn);

          const whatsBtn = Dom.el("a", {
            class: "btn btn-whatsapp",
            href: `https://wa.me/${phoneNum.replace("+", "")}`,
            target: "_blank",
            html: '<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg> WhatsApp',
          });
          actions.appendChild(whatsBtn);
        }

        const selectWrap = Dom.el("div", { class: "today-call-select" });
        const select = Dom.el("select", { class: "input input-sm" });
        const opts = [
          { value: "", label: "Mark status..." },
          { value: "called", label: "Called" },
          { value: "promised", label: "Promised to pay" },
          { value: "contacted", label: "Contacted" },
          { value: "no_answer", label: "No answer" },
          { value: "wrong_number", label: "Wrong number" },
          { value: "follow_up", label: "Follow up needed" },
        ];
        opts.forEach((o) => {
          const opt = Dom.el("option", { value: o.value, text: o.label });
          select.appendChild(opt);
        });
        const st = getCallStatus(g.customer);
        if (st) select.value = st.action;
        select.addEventListener("change", (e) => {
          if (e.target.value) markCalled(g.customer, e.target.value);
          if (lastRows) renderDashboard(lastRows, lastSource);
        });
        selectWrap.appendChild(select);
        actions.appendChild(selectWrap);

        item.appendChild(rankEl);
        item.appendChild(info);
        item.appendChild(actions);
        list.appendChild(item);
      });
    }
    wrap.appendChild(list);

    if (called.length > 0) {
      const resetBtn = Dom.el("button", { class: "btn btn-sm btn-secondary", text: "Reset all call status" });
      resetBtn.addEventListener("click", () => {
        resetAllCallStatus();
        if (lastRows) renderDashboard(lastRows, lastSource);
      });
      wrap.appendChild(Dom.el("div", { class: "card-footer", style: "padding:0 20px 16px;" })).appendChild(resetBtn);
    }

    return wrap;
  }

  /* ---------- Section: Customer Table ---------- */

  function renderCustomerTable(groups) {
    const sorted = sortGroups(groups, currentSort);

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
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let activeDateRange = null;

    function startOfWeek(d) { const dt = new Date(d); const day = dt.getDay(); dt.setDate(dt.getDate() - (day === 0 ? 6 : day - 1)); dt.setHours(0, 0, 0, 0); return dt; }
    function startOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
    function startOfQuarter(d) { return new Date(d.getFullYear(), Math.floor(d.getMonth() / 3) * 3, 1); }
    function fmtDate(d) { return d.toISOString().slice(0, 10); }

    const datePresets = [
      { key: "overdue", label: "Overdue only" },
      { key: "this_week", label: "Due this week" },
      { key: "this_month", label: "Due this month" },
      { key: "next_30", label: "Due in 30 days" },
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

        if (p.key === "custom") {
          if (customRow) customRow.style.display = "flex";
          activeDateRange = null;
        } else {
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
      const from = fromInput.value ? new Date(fromInput.value + "T00:00:00") : null;
      const to = toInput.value ? new Date(toInput.value + "T23:59:59") : null;
      activeDateRange = { from, to };
      applyFilters();
    });
    customRow.appendChild(applyCustom);
    const clearCustom = Dom.el("button", { class: "btn btn-sm btn-secondary", text: "Clear" });
    clearCustom.addEventListener("click", () => {
      activeDateRange = null;
      fromInput.value = "";
      toInput.value = "";
      dateChips.querySelectorAll(".date-chip").forEach((c) => c.classList.remove("active"));
      applyFilters();
    });
    customRow.appendChild(clearCustom);

    dateFilter.appendChild(dateChips);
    dateFilter.appendChild(customRow);
    filterBar.appendChild(dateFilter);

    /* --- Status / Urgency Chips --- */
    const filterChips = Dom.el("div", { class: "filter-chips", id: "filterChips" });
    const activeFilters = new Set();

    const chipDefs = [
      { key: "critical", label: "Critical (90d+)", color: "#d33b3b", count: 0 },
      { key: "high", label: "High (60d+)", color: "#e07c00", count: 0 },
      { key: "medium", label: "Medium (30d+)", color: "#b45309", count: 0 },
      { key: "called", label: "Already Called", color: "#27a644", count: 0 },
      { key: "not_called", label: "Not Called Yet", color: "#5e6ad2", count: 0 },
    ];

    chipDefs.forEach((c) => {
      if (c.key === "critical") c.count = groups.filter((g) => (g.maxDaysOverdue || 0) >= 90 && g.totalBalance > 0).length;
      if (c.key === "high") c.count = groups.filter((g) => (g.maxDaysOverdue || 0) >= 60 && (g.maxDaysOverdue || 0) < 90 && g.totalBalance > 0).length;
      if (c.key === "medium") c.count = groups.filter((g) => (g.maxDaysOverdue || 0) >= 30 && (g.maxDaysOverdue || 0) < 60 && g.totalBalance > 0).length;
      if (c.key === "called") c.count = groups.filter((g) => g.totalBalance > 0 && getCallStatus(g.customer)).length;
      if (c.key === "not_called") c.count = groups.filter((g) => g.totalBalance > 0 && !getCallStatus(g.customer)).length;
    });

    function applyFilters() {
      const allMain = document.querySelectorAll("tr.customer-main");
      allMain.forEach((tr) => {
        const name = tr.dataset.customer || "";
        const bills = tr.dataset.bills || "";
        const maxDays = parseInt(tr.dataset.maxDays || "0", 10);
        const isCalled = tr.dataset.called === "1";
        const dueDates = (tr.dataset.dueDates || "").split("|").filter(Boolean);
        const q = ($("#searchInput") || {}).value || "";
        const matchSearch = !q || name.includes(q.toLowerCase()) || bills.includes(q.toLowerCase());

        let matchFilter = activeFilters.size === 0;
        if (activeFilters.has("critical")) matchFilter = matchFilter || maxDays >= 90;
        if (activeFilters.has("high")) matchFilter = matchFilter || (maxDays >= 60 && maxDays < 90);
        if (activeFilters.has("medium")) matchFilter = matchFilter || (maxDays >= 30 && maxDays < 60);
        if (activeFilters.has("called")) matchFilter = matchFilter || isCalled;
        if (activeFilters.has("not_called")) matchFilter = matchFilter || !isCalled;

        let matchDate = true;
        if (activeDateRange) {
          matchDate = false;
          for (const ds of dueDates) {
            const d = new Date(ds + "T00:00:00");
            if (activeDateRange.from && d < activeDateRange.from) continue;
            if (activeDateRange.to && d > activeDateRange.to) continue;
            matchDate = true;
            break;
          }
        }

        const show = matchSearch && matchFilter && matchDate;
        tr.style.display = show ? "" : "none";
        const detail = tr.nextElementSibling;
        if (detail && detail.classList.contains("customer-detail")) {
          detail.style.display = show ? "" : "none";
        }
      });
    }

    chipDefs.forEach((c) => {
      const chip = Dom.el("button", {
        class: "filter-chip",
        html: `<span class="filter-chip-dot" style="background:${c.color}"></span>${c.label} <span class="filter-chip-count">${c.count}</span>`,
      });
      chip.addEventListener("click", () => {
        if (activeFilters.has(c.key)) {
          activeFilters.delete(c.key);
          chip.classList.remove("active");
        } else {
          activeFilters.add(c.key);
          chip.classList.add("active");
        }
        applyFilters();
      });
      filterChips.appendChild(chip);
    });

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

    filterBar.appendChild(filterChips);
    filterBar.appendChild(sortBtns);
    toolbar.appendChild(filterBar);
    wrap.appendChild(toolbar);

    const exportBtn = Dom.el("button", { class: "btn btn-sm btn-secondary export-btn", html: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Export CSV' });
    exportBtn.addEventListener("click", () => exportFilteredCSV(sorted));
    toolbar.appendChild(exportBtn);

    const tWrap = Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" });
    const table = Dom.el("table", { class: "data-table customer-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["", "Customer", "Pending bills", "Total balance", "Max overdue", "Priority", "Status"].forEach((h) => {
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

    const dueDateStrs = g.bills.filter((b) => b.dueDate).map((b) => b.dueDate.toISOString().slice(0, 10));

    const mainTr = Dom.el("tr", {
      class: `customer-main ${cls}`,
      "data-customer": g.customer.toLowerCase(),
      "data-bills": g.bills.map((b) => b.reference.toLowerCase()).join(" "),
      "data-max-days": String(g.maxDaysOverdue || 0),
      "data-called": isCalled ? "1" : "0",
      "data-due-dates": dueDateStrs.join("|"),
    });

    const chevron = Dom.el("td", { class: "cell-chevron", html: `<span class="chevron ${isOpen ? "open" : ""}">▶</span>` });
    mainTr.appendChild(chevron);

    const custCell = Dom.el("td", { class: "cell-title" });
    custCell.appendChild(Dom.el("span", { text: g.customer }));
    if (g.phone) {
      custCell.appendChild(Dom.el("span", { class: "cell-phone", html: `&nbsp;· ${g.phone}` }));
    }
    if (callSt) {
      custCell.appendChild(Dom.el("span", { html: callStatusBadge(g.customer) }));
    }
    mainTr.appendChild(custCell);

    mainTr.appendChild(Dom.el("td", { text: `${g.bills.length} bill${g.bills.length > 1 ? "s" : ""}`, class: "num" }));
    mainTr.appendChild(Dom.el("td", { text: Format.money(g.totalBalance), class: "num strong" }));
    mainTr.appendChild(Dom.el("td", { html: overdueBadge(g.maxDaysOverdue), class: "num" }));
    mainTr.appendChild(Dom.el("td", { html: urgencyLabel(priority.score), class: "num" }));

    mainTr.appendChild(Dom.el("td", {
      html: g.bills.every((b) => {
        const s = String(b.status || "").toLowerCase();
        return s.includes("paid") || s.includes("settled");
      })
        ? '<span class="badge badge-success badge-dot">All paid</span>'
        : g.maxDaysOverdue > 0
          ? '<span class="badge badge-danger badge-dot">Needs follow-up</span>'
          : '<span class="badge badge-primary badge-dot">Pending</span>',
    }));

    const detailTr = Dom.el("tr", { class: `customer-detail ${isOpen ? "open" : ""}` });
    const detailTd = Dom.el("td", { colspan: "7" });

    const subTable = Dom.el("div", { class: "bill-detail-wrap" });
    const subHead = Dom.el("div", { class: "bill-detail-header" });
    ["Bill No", "Bill date", "Due date", "Bill value", "Paid", "Balance", "Overdue", "Status"].forEach((h) => {
      subHead.appendChild(Dom.el("div", { class: `bill-detail-cell ${h === "Bill value" || h === "Paid" || h === "Balance" ? "num" : ""}`, text: h }));
    });
    subTable.appendChild(subHead);

    const sortedBills = [...g.bills].sort((a, b) => (b.daysOverdue || 0) - (a.daysOverdue || 0));
    sortedBills.forEach((b) => {
      const r = Dom.el("div", { class: `bill-detail-row ${urgencyClass(b.daysOverdue)}` });
      r.appendChild(Dom.el("div", { class: "bill-detail-cell cell-title", text: b.reference }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", text: Format.dateShort(b.issueDate) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", text: Format.dateShort(b.dueDate) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell num", text: Format.money(b.amount) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell num", text: Format.money(b.paid) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell num strong", text: Format.money(b.balance) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", html: overdueBadge(b.daysOverdue) }));
      r.appendChild(Dom.el("div", { class: "bill-detail-cell", html: statusBadge(b.status) }));
      subTable.appendChild(r);
    });

    const detailActions = Dom.el("div", { class: "detail-actions" });
    const phoneNum = g.phone.replace(/[^0-9+]/g, "");
    if (phoneNum) {
      const callBtn = Dom.el("a", {
        class: "btn btn-call btn-sm",
        href: `tel:${phoneNum}`,
        html: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07 19.5 19.5 0 01-6-6 19.79 19.79 0 01-3.07-8.67A2 2 0 014.11 2h3a2 2 0 012 1.72c.127.96.361 1.903.7 2.81a2 2 0 01-.45 2.11L8.09 9.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0122 16.92z"/></svg> Call',
      });
      detailActions.appendChild(callBtn);
      const whatsBtn = Dom.el("a", {
        class: "btn btn-whatsapp btn-sm",
        href: `https://wa.me/${phoneNum.replace("+", "")}`,
        target: "_blank",
        html: '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg> WhatsApp',
      });
      detailActions.appendChild(whatsBtn);
    }

    const selectWrap = Dom.el("div", { class: "detail-select-wrap" });
    const select = Dom.el("select", { class: "input input-sm" });
    [
      { value: "", label: "Mark status..." },
      { value: "called", label: "Called" },
      { value: "promised", label: "Promised to pay" },
      { value: "contacted", label: "Contacted" },
      { value: "no_answer", label: "No answer" },
      { value: "wrong_number", label: "Wrong number" },
      { value: "follow_up", label: "Follow up needed" },
    ].forEach((o) => {
      select.appendChild(Dom.el("option", { value: o.value, text: o.label }));
    });
    if (callSt) select.value = callSt.action;
    select.addEventListener("change", (e) => {
      if (e.target.value) markCalled(g.customer, e.target.value);
      if (lastRows) renderDashboard(lastRows, lastSource);
    });
    selectWrap.appendChild(select);
    detailActions.appendChild(selectWrap);

    subTable.appendChild(detailActions);
    detailTd.appendChild(subTable);
    detailTr.appendChild(detailTd);

    mainTr.addEventListener("click", () => {
      const chev = mainTr.querySelector(".chevron");
      const isOpenNow = detailTr.classList.toggle("open");
      chev.classList.toggle("open", isOpenNow);
      if (isOpenNow) expandedCustomers.add(g.customer);
      else expandedCustomers.delete(g.customer);
    });

    return { main: mainTr, detail: detailTr };
  }

  /* ---------- Export ---------- */

  function exportFilteredCSV(groups) {
    const header = "Rank,Customer,Phone,Total Balance,Bills,Max Overdue Days,Priority Score,Urgency,Call Status\n";
    const rows = groups.map((g, i) => {
      const p = computePriorityScore(g);
      const st = getCallStatus(g.customer);
      return [
        i + 1,
        `"${g.customer}"`,
        g.phone || "",
        g.totalBalance.toFixed(2),
        g.bills.length,
        g.maxDaysOverdue || 0,
        p.score,
        p.urgency,
        st ? st.action : "not_called",
      ].join(",");
    });
    const blob = new Blob([header + rows.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `receivables-priority-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  /* ---------- Aging Analysis ---------- */

  function renderAgingChart(rows) {
    const buckets = AGING_BUCKETS.map((b) => ({
      label: b.label,
      value: rows.filter((r) => r.bucket === b.label).reduce((s, r) => s + r.balance, 0),
    }));

    const segs = buckets.map((b, i) => ({ ...b, color: AGING_COLORS[i] }));

    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Aging summary</div>
        <div class="card-subtitle">Outstanding balance by age bucket</div>
      </div>
    `}));

    const body = Dom.el("div", { class: "grid-2", style: "margin-bottom:0;align-items:center;" });
    body.appendChild(Dom.el("div", { class: "chart-wrap", style: "display:flex;justify-content:center;" }));
    body.lastChild.appendChild(Charts.donutChart({
      segments: segs,
      size: 240,
      thickness: 30,
      format: (v) => Format.moneyWhole(v),
    }));

    const legend = Dom.el("div", { class: "legend" });
    segs.forEach((s, i) => {
      legend.appendChild(Dom.el("div", { class: "legend-item", html: `
        <span class="legend-swatch" style="background:${AGING_COLORS[i]}"></span>
        ${s.label} · ${Format.moneyWhole(s.value)}
      `}));
    });
    body.appendChild(legend);
    wrap.appendChild(body);
    return wrap;
  }

  function renderBarChart(rows) {
    const topCustomers = [...rows]
      .filter((r) => r.balance > 0)
      .reduce((acc, r) => {
        const k = r.customer;
        acc[k] = acc[k] || { label: k, value: 0 };
        acc[k].value += r.balance;
        return acc;
      }, {});
    const topSorted = Object.values(topCustomers).sort((a, b) => b.value - a.value).slice(0, 8);

    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Top customers by balance</div>
        <div class="card-subtitle">Outstanding amount by customer</div>
      </div>
    `}));
    wrap.appendChild(Dom.el("div", { class: "chart-wrap" }));
    wrap.lastChild.appendChild(Charts.barChart({
      items: topSorted,
      height: 280,
      horizontal: true,
      format: (v) => Format.moneyWhole(v),
    }));
    return wrap;
  }

  function renderAgingAnalysis(rows) {
    const data = AGING_BUCKETS.map((b, i) => {
      const matching = rows.filter((r) => r.bucket === b.label);
      const total = matching.reduce((s, r) => s + r.balance, 0);
      const count = matching.length;
      const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);
      const overdueTotal = rows.filter((r) => r.daysOverdue > 0).reduce((s, r) => s + r.balance, 0);
      const pct = totalOutstanding ? (total / totalOutstanding) * 100 : 0;
      const overduePct = overdueTotal ? (total / overdueTotal) * 100 : 0;
      const avgDays = matching.length
        ? Math.round(matching.reduce((s, r) => s + (r.daysOverdue || 0), 0) / matching.length)
        : 0;
      return { ...b, color: AGING_COLORS[i], total, count, pct, overduePct, avgDays };
    });

    const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);
    const overdueTotal = rows.filter((r) => r.daysOverdue > 0).reduce((s, r) => s + r.balance, 0);
    const totalBills = rows.length;
    const overdueBills = rows.filter((r) => r.daysOverdue > 0).length;

    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Advanced Aging Analysis</div>
        <div class="card-subtitle">Detailed breakdown of outstanding balances by age bucket</div>
      </div>
    `}));

    const summaryRow = Dom.el("div", { class: "aging-summary-row" });
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `
      <span class="aging-summary-label">Total Outstanding</span>
      <span class="aging-summary-value">${Format.money(totalOutstanding)}</span>
    `}));
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `
      <span class="aging-summary-label">Overdue Amount</span>
      <span class="aging-summary-value danger">${Format.money(overdueTotal)}</span>
    `}));
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `
      <span class="aging-summary-label">Total Bills</span>
      <span class="aging-summary-value">${totalBills}</span>
    `}));
    summaryRow.appendChild(Dom.el("div", { class: "aging-summary-stat", html: `
      <span class="aging-summary-label">Overdue Bills</span>
      <span class="aging-summary-value danger">${overdueBills}</span>
    `}));
    wrap.appendChild(summaryRow);

    const stackedBar = Dom.el("div", { class: "aging-stacked-bar" });
    data.forEach((d) => {
      if (d.total > 0) {
        stackedBar.appendChild(Dom.el("div", {
          class: "aging-stacked-seg",
          style: `width:${d.pct}%;background:${d.color};`,
          title: `${d.label}: ${Format.money(d.total)} (${d.pct.toFixed(1)}%)`,
        }));
      }
    });
    wrap.appendChild(stackedBar);

    const table = Dom.el("table", { class: "data-table aging-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["Aging Bucket", "Bills", "Amount", "% of Total", "% of Overdue", "Avg Days Overdue", "Distribution"].forEach((h) => {
      const th = Dom.el("th", { text: h });
      if (h !== "Aging Bucket") th.classList.add("num");
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    table.appendChild(thead);

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
      barCell.appendChild(barWrap);
      tr.appendChild(barCell);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" })).appendChild(table);
    return wrap;
  }

  function renderCustomerAgingHeatmap(rows, groups) {
    const wrap = Dom.el("div", { class: "card" });
    wrap.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Customer Aging Heatmap</div>
        <div class="card-subtitle">Which aging buckets each customer falls into (top 15 by balance)</div>
      </div>
    `}));

    const topGroups = [...groups]
      .filter((g) => g.totalBalance > 0)
      .sort((a, b) => b.totalBalance - a.totalBalance)
      .slice(0, 15);

    const table = Dom.el("table", { class: "data-table heatmap-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    hr.appendChild(Dom.el("th", { text: "Customer" }));
    hr.appendChild(Dom.el("th", { text: "Total", class: "num" }));
    AGING_BUCKETS.forEach((b, i) => {
      hr.appendChild(Dom.el("th", { text: b.label, class: "num", style: `color:${AGING_COLORS[i]};` }));
    });
    thead.appendChild(hr);
    table.appendChild(thead);

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
        } else {
          cell.textContent = "—";
        }
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
    lastRows = rows;
    lastSource = sourceLabel;

    const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);
    const groups = groupByCustomer(rows);
    const dash = $("#dashboard");
    Dom.clear(dash);

    if (sourceLabel) {
      dash.appendChild(Dom.el("div", { class: "source-badge", html: `<span class="badge badge-primary">${sourceLabel}</span>` }));
    }

    dash.appendChild(renderGuide());
    dash.appendChild(renderKpis(rows, totalOutstanding, groups));
    dash.appendChild(renderTodayCallList(groups));
    dash.appendChild(renderCustomerTable(groups));

    const charts = Dom.el("div", { class: "grid-2" });
    charts.appendChild(renderAgingChart(rows));
    charts.appendChild(renderBarChart(rows));
    dash.appendChild(charts);

    dash.appendChild(renderAgingAnalysis(rows));
    dash.appendChild(renderCustomerAgingHeatmap(rows, groups));

    const now = new Date();
    $("#lastUpdated").textContent = `Updated ${now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
  }

  function renderPlaceholder() {
    const banner = Dom.el("div", { class: "status-banner warning", html: `
      <strong>Data source not configured</strong>
      <span>&nbsp;—&nbsp;The Google Sheet URL is not set. Configure it in
      <code>assets/js/config.js</code> or upload a CSV/Excel file using the button above.</span>
    `});
    $("#statusBanner").appendChild(banner);

    const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
    box.appendChild(Dom.el("div", { class: "state-icon", html: "📞" }));
    box.appendChild(Dom.el("h3", { text: "No receivables data loaded" }));
    box.appendChild(Dom.el("p", { text: "Configure the Google Sheet or upload a file to start tracking outstanding payments." }));
    box.appendChild(Dom.el("a", { class: "btn btn-secondary", href: "index.html", text: "Back to home" }));
    $("#dashboard").appendChild(box);
  }

  function renderError(err) {
    $("#dashboard").innerHTML = "";
    const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
    box.appendChild(Dom.el("div", { class: "state-icon", html: "⚠️" }));
    box.appendChild(Dom.el("h3", { text: "Could not load receivables data" }));
    box.appendChild(Dom.el("p", { text: err.message }));
    const retry = Dom.el("button", { class: "btn btn-secondary", text: "Try again" });
    retry.addEventListener("click", loadFromSheet);
    box.appendChild(retry);
    $("#dashboard").appendChild(box);
  }

  /* ---------- Load ---------- */

  async function loadFromSheet() {
    const refreshBtn = $("#refreshBtn");
    refreshBtn.classList.add("loading");
    try {
      if (!SHEET.sheetUrl) {
        renderPlaceholder();
        return;
      }
      $("#statusBanner").innerHTML = "";
      const records = await Sheets.getRows(SHEET.sheetUrl, SHEET.gid);
      if (!records.length) throw new Error("The sheet contains no data rows.");

      const headers = Object.keys(records[0]);
      const cols = matchHeader(headers);
      if (!cols.amount || !cols.customer) {
        throw new Error("Could not map columns. Expected at least customer and amount columns (e.g. Client Name, Amount/Balance).");
      }

      const rows = buildRows(records, cols);
      renderDashboard(rows, "Google Sheets");
    } catch (err) {
      renderError(err);
    } finally {
      refreshBtn.classList.remove("loading");
    }
  }

  async function loadFromCSV(file) {
    try {
      $("#statusBanner").innerHTML = "";
      const { records, cols } = await loadFile(file);
      const rows = buildRows(records, cols);
      renderDashboard(rows, `Uploaded: ${file.name}`);
    } catch (err) {
      renderError(err);
    }
  }

  /* ---------- Wire up ---------- */

  const fileInput = $("#fileInput");
  const uploadBtn = $("#uploadBtn");

  uploadBtn.addEventListener("click", () => fileInput.click());

  fileInput.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) loadFromCSV(file);
    fileInput.value = "";
  });

  $("#refreshBtn").addEventListener("click", loadFromSheet);

  document.addEventListener("click", (e) => {
    const toggle = e.target.closest("#guideToggle");
    if (toggle) {
      const body = document.getElementById("guideBody");
      const chev = toggle.querySelector(".guide-chevron");
      body.classList.toggle("open");
      chev.classList.toggle("open");
    }
  });

  loadFromSheet();
})();

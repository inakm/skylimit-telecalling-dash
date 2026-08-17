/* ============================================================
   Accounts Receivable — Telecaller Dashboard.
   Customer-grouped table, live search, sorting, expandable bills.
   ============================================================ */

(() => {
  const $ = (sel) => document.querySelector(sel);

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

  /* ---------- Sorting ---------- */

  let currentSort = { key: "overdue", dir: "desc" };

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
    }
    if (sort.dir === "desc" && sort.key !== "overdue" && sort.key !== "bills") {
      sorted.reverse();
    }
    return sorted;
  }

  function toggleSort(key) {
    if (currentSort.key === key) {
      currentSort.dir = currentSort.dir === "asc" ? "desc" : "asc";
    } else {
      currentSort = { key, dir: key === "overdue" || key === "bills" ? "desc" : "asc" };
    }
    if (lastRows) renderDashboard(lastRows, lastSource);
  }

  /* ---------- Rendering ---------- */

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

  function statCard(label, value, opts = {}) {
    return Dom.el("div", { class: "stat-card", html: `
      <div class="stat-label">${label}</div>
      <div class="stat-value">${value}</div>
      ${opts.delta ? `<div class="stat-delta">${opts.delta}</div>` : ""}
    `});
  }

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

    const sortBtns = Dom.el("div", { class: "sort-btns" });
    [
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
    toolbar.appendChild(sortBtns);
    wrap.appendChild(toolbar);

    const tWrap = Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" });
    const table = Dom.el("table", { class: "data-table customer-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["", "Customer", "Pending bills", "Total balance", "Max overdue", "Status"].forEach((h) => {
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
    searchInput.addEventListener("input", Dom.debounce((e) => {
      const q = e.target.value.toLowerCase().trim();
      const allMain = tbody.querySelectorAll("tr.customer-main");
      allMain.forEach((tr) => {
        const name = tr.dataset.customer || "";
        const bills = tr.dataset.bills || "";
        const match = !q || name.includes(q) || bills.includes(q);
        tr.style.display = match ? "" : "none";
        const detail = tr.nextElementSibling;
        if (detail && detail.classList.contains("customer-detail")) {
          detail.style.display = match ? "" : "none";
          if (match && q) detail.classList.add("open");
        }
      });
    }, 200));

    return wrap;
  }

  function buildCustomerRow(g) {
    const isOpen = expandedCustomers.has(g.customer);
    const cls = urgencyClass(g.maxDaysOverdue);

    const mainTr = Dom.el("tr", { class: `customer-main ${cls}`, "data-customer": g.customer.toLowerCase(), "data-bills": g.bills.map((b) => b.reference.toLowerCase()).join(" ") });

    const chevron = Dom.el("td", { class: "cell-chevron", html: `<span class="chevron ${isOpen ? "open" : ""}">▶</span>` });
    mainTr.appendChild(chevron);

    const custCell = Dom.el("td", { class: "cell-title" });
    custCell.appendChild(Dom.el("span", { text: g.customer }));
    if (g.phone) {
      custCell.appendChild(Dom.el("span", { class: "cell-phone", html: `&nbsp;· ${g.phone}` }));
    }
    mainTr.appendChild(custCell);

    mainTr.appendChild(Dom.el("td", { text: `${g.bills.length} bill${g.bills.length > 1 ? "s" : ""}`, class: "num" }));
    mainTr.appendChild(Dom.el("td", { text: Format.money(g.totalBalance), class: "num strong" }));
    mainTr.appendChild(Dom.el("td", { html: overdueBadge(g.maxDaysOverdue), class: "num" }));

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
    const detailTd = Dom.el("td", { colspan: "6" });

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

    dash.appendChild(renderKpis(rows, totalOutstanding, groups));
    dash.appendChild(renderCustomerTable(groups));

    const charts = Dom.el("div", { class: "grid-2" });
    charts.appendChild(renderAgingChart(rows));
    charts.appendChild(renderBarChart(rows));
    dash.appendChild(charts);

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
  loadFromSheet();
})();

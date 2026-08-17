/* ============================================================
   Accounts Receivable dashboard.
   Data source: Google Sheets (configured in config.js) or CSV upload.
   ============================================================ */

(() => {
  const $ = (sel) => document.querySelector(sel);

  const AGING_BUCKETS = [
    { label: "Current", min: 0, max: 30 },
    { label: "31–60 days", min: 31, max: 60 },
    { label: "61–90 days", min: 61, max: 90 },
    { label: "90+ days", min: 91, max: Infinity },
  ];

  const SHEET = CONFIG.receivables;

  const headerMap = {
    reference: ["invoice", "inv no", "reference", "ref no", "receipt", "doc no", "invoice number", "inv no.", "bill no", "bill no."],
    customer: ["customer", "client", "company", "account", "customer name", "client name", "debtor"],
    issueDate: ["invoice date", "issue date", "date", "billing date", "transaction date", "bill date"],
    dueDate: ["due date", "payment due", "due", "due on"],
    amount: ["amount", "total", "balance", "outstanding", "amount due", "amount outstanding", "invoice amount", "balance due", "net amount"],
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
    const bucket = AGING_BUCKETS.find((b) => days >= b.min && days <= b.max);
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

  /* ---------- Rendering ---------- */

  function statusBadge(status) {
    const s = String(status || "").toLowerCase();
    if (s.includes("paid") || s.includes("settled") || s.includes("closed")) {
      return '<span class="badge badge-success badge-dot">Paid</span>';
    }
    if (s.includes("partial")) {
      return '<span class="badge badge-warning badge-dot">Partial</span>';
    }
    if (s.includes("overdue") || s.includes("past due") || s.includes("due")) {
      return '<span class="badge badge-danger badge-dot">Overdue</span>';
    }
    if (s.includes("open") || s.includes("unpaid") || s.includes("pending") || s.includes("due now")) {
      return '<span class="badge badge-primary badge-dot">Open</span>';
    }
    return `<span class="badge badge-neutral">${status || "—"}</span>`;
  }

  function statCard(label, value, opts = {}) {
    const cls = ["stat-card"];
    if (opts.mono) cls.push("stat-mono");
    const el = Dom.el("div", { class: cls.join(" "), html: `
      <div class="stat-label">${label}</div>
      <div class="stat-value ${opts.mono ? "mono" : ""}">${value}</div>
      ${opts.delta ? `<div class="stat-delta">${opts.delta}</div>` : ""}
    `});
    return el;
  }

  function renderKpis(rows, totalOutstanding) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const overdue = rows.filter((r) => r.daysOverdue !== null && r.daysOverdue > 0);
    const overdueTotal = overdue.reduce((s, r) => s + r.balance, 0);
    const currentTotal = rows.filter((r) => r.daysOverdue !== null && r.daysOverdue <= 0).reduce((s, r) => s + r.balance, 0);
    const days90 = rows.filter((r) => r.bucketDays !== null && r.bucketDays >= 91).reduce((s, r) => s + r.balance, 0);

    const pctOverdue = totalOutstanding ? (overdueTotal / totalOutstanding) * 100 : 0;
    const pct90 = totalOutstanding ? (days90 / totalOutstanding) * 100 : 0;

    const grid = Dom.el("div", { class: "stat-grid" });
    grid.appendChild(statCard("Total outstanding", Format.moneyWhole(totalOutstanding)));
    grid.appendChild(statCard("Current (0–30 days)", Format.moneyWhole(currentTotal)));
    grid.appendChild(statCard("Total overdue", Format.moneyWhole(overdueTotal), {
      delta: `<span class="down">${pctOverdue.toFixed(1)}% of outstanding</span>`,
    }));
    grid.appendChild(statCard("90+ days risk", Format.moneyWhole(days90), {
      delta: pct90 > 20 ? `<span class="down">${pct90.toFixed(1)}% of outstanding</span>` : `<span>${pct90.toFixed(1)}% of outstanding</span>`,
    }));

    return grid;
  }

  function renderAgingChart(rows) {
    const buckets = AGING_BUCKETS.map((b) => ({
      label: b.label,
      value: rows.filter((r) => r.bucket === b.label).reduce((s, r) => s + r.balance, 0),
    }));

    const colors = ["#27a644", "#5e6ad2", "#b45309", "#d33b3b"];
    const segs = buckets.map((b, i) => ({ ...b, color: colors[i] }));

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
        <span class="legend-swatch" style="background:${colors[i]}"></span>
        ${s.label} · ${Format.moneyWhole(s.value)}
      `}));
    });
    body.appendChild(legend);
    wrap.appendChild(body);
    return wrap;
  }

  function renderTable(rows) {
    const sorted = [...rows].sort((a, b) => (b.daysOverdue || 0) - (a.daysOverdue || 0));

    const wrap = Dom.el("div", { class: "card", style: "padding:0;overflow:hidden;" });
    const head = Dom.el("div", { class: "card-header", style: "padding:16px 20px;margin:0;border-bottom:1px solid var(--hairline);" });
    head.appendChild(Dom.el("div", { html: `
      <div class="card-title">Receivable detail</div>
      <div class="card-subtitle">${sorted.length} records · sorted by most overdue</div>
    `}));
    wrap.appendChild(head);

    const tWrap = Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" });
    const table = Dom.el("table", { class: "data-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["Reference", "Customer", "Invoice date", "Due date", "Overdue", "Balance", "Status"].forEach((h) => {
      hr.appendChild(Dom.el("th", { text: h, class: h === "Balance" ? "num" : "" }));
    });
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = Dom.el("tbody");
    sorted.forEach((r) => {
      const tr = Dom.el("tr");
      tr.appendChild(Dom.el("td", { text: r.reference, class: "cell-title" }));
      tr.appendChild(Dom.el("td", { text: r.customer }));
      tr.appendChild(Dom.el("td", { text: Format.dateShort(r.issueDate) }));
      tr.appendChild(Dom.el("td", { text: Format.dateShort(r.dueDate) }));
      tr.appendChild(Dom.el("td", {
        html: r.daysOverdue > 0
          ? `<span class="badge badge-danger badge-dot">${r.daysOverdue}d overdue</span>`
          : r.daysOverdue <= 0 && r.daysOverdue !== null
            ? `<span class="badge badge-neutral">${Math.abs(r.daysOverdue)}d to go</span>`
            : "—",
      }));
      tr.appendChild(Dom.el("td", { text: Format.money(r.balance), class: "num strong" }));
      tr.appendChild(Dom.el("td", { html: statusBadge(r.status) }));
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    tWrap.appendChild(table);
    wrap.appendChild(tWrap);
    return wrap;
  }

  function renderDashboard(rows, sourceLabel) {
    const totalOutstanding = rows.reduce((s, r) => s + r.balance, 0);
    const dash = $("#dashboard");
    Dom.clear(dash);

    if (sourceLabel) {
      const badge = Dom.el("div", { class: "source-badge", html: `<span class="badge badge-primary">${sourceLabel}</span>` });
      dash.appendChild(badge);
    }

    dash.appendChild(renderKpis(rows, totalOutstanding));

    const charts = Dom.el("div", { class: "grid-2" });
    charts.appendChild(renderAgingChart(rows));
    charts.appendChild(Dom.el("div", { class: "card", html: `
      <div class="card-header">
        <div>
          <div class="card-title">Overdue concentration</div>
          <div class="card-subtitle">Outstanding balance by customer, most overdue first</div>
        </div>
      </div>
    `}));
    const topCustomers = [...rows]
      .filter((r) => r.balance > 0)
      .reduce((acc, r) => {
        const k = r.customer;
        acc[k] = acc[k] || { label: k, value: 0, count: 0 };
        acc[k].value += r.balance;
        acc[k].count += 1;
        return acc;
      }, {});
    const topSorted = Object.values(topCustomers).sort((a, b) => b.value - a.value).slice(0, 8);
    charts.lastChild.appendChild(Dom.el("div", { class: "chart-wrap" }));
    charts.lastChild.lastChild.appendChild(Charts.barChart({
      items: topSorted,
      height: 280,
      horizontal: true,
      format: (v) => Format.moneyWhole(v),
    }));
    dash.appendChild(charts);

    dash.appendChild(renderTable(rows));

    const now = new Date();
    $("#lastUpdated").textContent = `Updated ${now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
  }

  function renderPlaceholder() {
    const banner = Dom.el("div", { class: "status-banner warning", html: `
      <strong>Data source not configured</strong>
      <span>&nbsp;—&nbsp;The Google Sheet URL is not set, and no file has been uploaded. Configure the URL in
      <code>assets/js/config.js</code> or upload a CSV/Excel file using the button above.</span>
    `});
    $("#statusBanner").appendChild(banner);

    const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
    box.appendChild(Dom.el("div", { class: "state-icon", text: "💰" }));
    box.appendChild(Dom.el("h3", { text: "Receivables dashboard coming soon" }));
    box.appendChild(Dom.el("p", { text: "Upload a CSV file or configure the Google Sheet to get started." }));
    const btn = Dom.el("a", { class: "btn btn-secondary", href: "index.html", text: "Back to home" });
    box.appendChild(btn);
    $("#dashboard").appendChild(box);
  }

  function renderError(err) {
    $("#dashboard").innerHTML = "";
    const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
    box.appendChild(Dom.el("div", { class: "state-icon", text: "⚠️" }));
    box.appendChild(Dom.el("h3", { text: "Could not load receivables data" }));
    box.appendChild(Dom.el("p", { text: err.message }));
    const retry = Dom.el("button", { class: "btn btn-secondary", text: "Try again" });
    retry.addEventListener("click", loadFromSheet);
    box.appendChild(retry);
    $("#dashboard").appendChild(box);
  }

  /* ---------- Load from Google Sheets ---------- */

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

  /* ---------- Load from CSV upload ---------- */

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

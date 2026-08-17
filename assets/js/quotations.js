/* ============================================================
   Quotations dashboard.
   Data source is configured in assets/js/config.js.
   Column names are auto-detected from the sheet header row.
   ============================================================ */

(() => {
  const $ = (sel) => document.querySelector(sel);

  const SHEET = CONFIG.quotations;

  const headerMap = {
    timestamp: ["quote date", "quote date (date)", "date", "date submitted", "timestamp", "submitted", "created at"],
    customer: ["customer name", "customer", "client", "company", "account", "client name", "lead", "prospect", "company name"],
    number: ["quote no", "quotation no", "quote number", "quotation number", "reference", "reference no", "ref no", "quote id", "quotation id"],
    amount: ["quote value", "quotation amount", "quote amount", "amount", "total amount", "value", "quotation value", "price", "total", "amount (php)", "amount php", "cost"],
    status: ["order received ( y) or ( n)", "order received", "order", "status", "converted", "conversion status", "quotation status", "pipeline status"],
    orderValue: ["order value", "converted value", "won value", "accepted value"],
    saleValue: ["sale value", "sale", "actual sale", "sales value", "billed value"],
    invoiceNo: ["invoice no", "invoice number", "inv no", "invoice"],
    supplyDate: ["date of supply", "supply date", "delivery date"],
    remark: ["remark", "remarks", "notes", "comments", "additional notes"],
  };

  let allRows = [];
  let filterState = { search: "", status: "all" };

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

  function parseOrderReceived(v) {
    const s = String(v || "").trim().toUpperCase();
    if (!s) return null;
    return s.startsWith("Y") || s === "YES" || s.includes("ORDERED") || s.includes("RECEIVED");
  }

  function buildRows(records, cols) {
    const rows = [];
    for (const r of records) {
      const amount = cols.amount ? Format.parseNumber(r[cols.amount]) : NaN;
      let timestamp = cols.timestamp ? Format.parseDate(r[cols.timestamp]) : null;

      const orderReceived = cols.status ? parseOrderReceived(r[cols.status]) : null;
      const orderValue = cols.orderValue ? Format.parseNumber(r[cols.orderValue]) : NaN;
      const saleValue = cols.saleValue ? Format.parseNumber(r[cols.saleValue]) : NaN;

      rows.push({
        timestamp,
        customer: (r[cols.customer] || "").trim() || "—",
        number: (r[cols.number] || "").trim() || "—",
        amount: Number.isFinite(amount) ? amount : NaN,
        orderReceived,
        status: (r[cols.status] || "").trim(),
        orderValue: Number.isFinite(orderValue) ? orderValue : NaN,
        saleValue: Number.isFinite(saleValue) ? saleValue : NaN,
        invoiceNo: cols.invoiceNo ? (r[cols.invoiceNo] || "").trim() : "",
        supplyDate: cols.supplyDate ? Format.parseDate(r[cols.supplyDate]) : null,
        remark: cols.remark ? (r[cols.remark] || "").trim() : "",
      });
    }
    return rows;
  }

  function normalizeStatus(row) {
    if (row.orderReceived !== null) return row.orderReceived ? "accepted" : "open";
    const v = String(row.status || "").toLowerCase();
    if (!v) return "unknown";
    if (v.includes("accept") || v.includes("won") || v.includes("convert") || v.includes("approved") || v.includes("received") || v === "y" || v === "yes") return "accepted";
    if (v.includes("reject") || v.includes("lost") || v.includes("decline") || v.includes("declined")) return "rejected";
    if (v.includes("sent") || v.includes("issued") || v.includes("delivered") || v.includes("submitted") || v === "n" || v === "no") return "open";
    if (v.includes("draft") || v.includes("preparing")) return "draft";
    if (v.includes("follow")) return "follow up";
    return "unknown";
  }

  const STATUS_META = {
    accepted: { label: "Order received", badge: "badge-success" },
    open: { label: "Open", badge: "badge-primary" },
    rejected: { label: "Rejected", badge: "badge-danger" },
    "follow up": { label: "Follow up", badge: "badge-warning" },
    draft: { label: "Draft", badge: "badge-neutral" },
    unknown: { label: "Unknown", badge: "badge-neutral" },
  };

  function badgeFor(row) {
    const n = normalizeStatus(row);
    const meta = STATUS_META[n];
    return `<span class="badge ${meta.badge} badge-dot">${meta.label}</span>`;
  }

  function statCard(label, value, opts = {}) {
    return Dom.el("div", { class: "stat-card", html: `
      <div class="stat-label">${label}</div>
      <div class="stat-value ${opts.mono ? "mono" : ""}">${value}</div>
      ${opts.delta ? `<div class="stat-delta">${opts.delta}</div>` : ""}
    `});
  }

  function renderFilters() {
    const bar = Dom.el("div", { class: "filter-bar" });

    const search = Dom.el("div", { class: "search-box" });
    search.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></svg>`;
    const input = Dom.el("input", {
      class: "input",
      type: "search",
      placeholder: "Search customer or quote no…",
    });
    input.addEventListener("input", Dom.debounce((e) => {
      filterState.search = e.target.value.trim().toLowerCase();
      renderKpis();
      renderCharts();
      renderTable();
    }, 200));
    search.appendChild(input);
    bar.appendChild(search);

    const statusSel = Dom.el("select", { class: "select" });
    const options = [
      { v: "all", t: "All statuses" },
      { v: "accepted", t: "Order received" },
      { v: "open", t: "Open" },
      { v: "rejected", t: "Rejected" },
      { v: "draft", t: "Draft" },
      { v: "follow up", t: "Follow up" },
      { v: "unknown", t: "Unknown" },
    ];
    options.forEach((o) => {
      statusSel.appendChild(Dom.el("option", { value: o.v, text: o.t }));
    });
    statusSel.addEventListener("change", (e) => {
      filterState.status = e.target.value;
      renderKpis();
      renderCharts();
      renderTable();
    });
    bar.appendChild(statusSel);
    bar.appendChild(Dom.el("div", { style: "flex:1;" }));

    $("#filters").innerHTML = "";
    $("#filters").appendChild(bar);
  }

  function filteredRows() {
    return allRows.filter((r) => {
      if (filterState.status !== "all" && normalizeStatus(r) !== filterState.status) return false;
      if (filterState.search) {
        const hay = `${r.customer} ${r.number} ${r.invoiceNo}`.toLowerCase();
        if (!hay.includes(filterState.search)) return false;
      }
      return true;
    });
  }

  function renderKpis() {
    const rows = filteredRows();
    const withAmount = rows.filter((r) => Number.isFinite(r.amount));
    const totalCount = rows.length;
    const totalValue = withAmount.reduce((s, r) => s + r.amount, 0);

    const accepted = rows.filter((r) => normalizeStatus(r) === "accepted");
    const acceptedValue = accepted
      .map((r) => (Number.isFinite(r.orderValue) ? r.orderValue : Number.isFinite(r.amount) ? r.amount : 0))
      .reduce((s, v) => s + v, 0);
    const acceptedCount = accepted.length;

    const conversionByValue = totalValue ? (acceptedValue / totalValue) * 100 : 0;
    const conversionByCount = totalCount ? (acceptedCount / totalCount) * 100 : 0;
    const avg = withAmount.length ? totalValue / withAmount.length : 0;

    const grid = $("#kpis");
    Dom.clear(grid);
    grid.appendChild(statCard("Total quotations", Format.number(totalCount)));
    grid.appendChild(statCard("Total quote value", Format.moneyWhole(totalValue), { mono: true }));
    grid.appendChild(statCard("Order value received", Format.moneyWhole(acceptedValue), {
      mono: true,
      delta: `<span class="up">${conversionByValue.toFixed(1)}%</span> of quote value`,
    }));
    grid.appendChild(statCard("Conversion rate", `${conversionByCount.toFixed(1)}%`, {
      delta: `<span>${Format.number(acceptedCount)} of ${Format.number(totalCount)} converted</span>`,
    }));
  }

  function renderCharts() {
    const rows = filteredRows();

    const trendMap = {};
    rows.forEach((r) => {
      const d = r.timestamp;
      if (!d) return;
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      trendMap[key] = trendMap[key] || { y: 0, count: 0 };
      trendMap[key].y += Number.isFinite(r.amount) ? r.amount : 0;
      trendMap[key].count += 1;
    });
    const trendKeys = Object.keys(trendMap).sort();
    const trendPoints = trendKeys.map((k) => {
      const [y, m] = k.split("-").map(Number);
      const d = new Date(y, m - 1, 1);
      return { y: trendMap[k].y, label: d.toLocaleDateString("en-US", { month: "short", year: "2-digit" }) };
    });

    const statusMap = {};
    rows.forEach((r) => {
      const n = normalizeStatus(r);
      const v = Number.isFinite(r.orderValue) ? r.orderValue : Number.isFinite(r.amount) ? r.amount : 0;
      statusMap[n] = (statusMap[n] || 0) + v;
    });
    const statusColors = {
      accepted: "#27a644",
      open: "#5e6ad2",
      rejected: "#d33b3b",
      "follow up": "#b45309",
      draft: "#a8abb3",
      unknown: "#c9cbd2",
    };
    const segs = Object.entries(statusMap)
      .map(([k, v]) => ({ label: STATUS_META[k].label, value: v, color: statusColors[k] || "#c9cbd2" }))
      .filter((s) => s.value > 0)
      .sort((a, b) => b.value - a.value);

    const custMap = {};
    rows.filter((r) => Number.isFinite(r.amount)).forEach((r) => {
      custMap[r.customer] = (custMap[r.customer] || 0) + r.amount;
    });
    const topCustomers = Object.entries(custMap)
      .map(([k, v]) => ({ label: k, value: v }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8);

    const grid = $("#charts");
    Dom.clear(grid);

    const trendCard = Dom.el("div", { class: "card" });
    trendCard.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Quotation value</div>
        <div class="card-subtitle">Total quote value by month</div>
      </div>
    `}));
    const tWrap = Dom.el("div", { class: "chart-wrap" });
    if (trendPoints.length >= 2) {
      tWrap.appendChild(Charts.lineChart({
        points: trendPoints,
        labels: trendPoints.map((p) => p.label),
        height: 270,
        format: (v) => Format.moneyWhole(v),
      }));
    } else {
      tWrap.appendChild(Dom.el("div", { class: "state", html: "<p>Not enough data to plot a trend yet.</p>" }));
    }
    trendCard.appendChild(tWrap);
    grid.appendChild(trendCard);

    const statusCard = Dom.el("div", { class: "card" });
    statusCard.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Conversion by status</div>
        <div class="card-subtitle">Value received vs open pipeline</div>
      </div>
    `}));
    const sBody = Dom.el("div", { class: "grid-2", style: "margin-bottom:0;align-items:center;" });
    const sWrap = Dom.el("div", { class: "chart-wrap", style: "display:flex;justify-content:center;" });
    if (segs.length) {
      sWrap.appendChild(Charts.donutChart({ segments: segs, size: 210, thickness: 26, format: (v) => Format.moneyWhole(v) }));
    } else {
      sWrap.appendChild(Dom.el("div", { class: "state", html: "<p>No status data yet.</p>" }));
    }
    sBody.appendChild(sWrap);
    const legend = Dom.el("div", { class: "legend", style: "flex-direction:column;align-items:flex-start;gap:8px;" });
    segs.forEach((s) => {
      legend.appendChild(Dom.el("div", { class: "legend-item", html: `
        <span class="legend-swatch" style="background:${s.color}"></span>
        ${s.label} · ${Format.moneyWhole(s.value)}
      `}));
    });
    sBody.appendChild(legend);
    statusCard.appendChild(sBody);
    grid.appendChild(statusCard);

    const topCard = Dom.el("div", { class: "card" });
    topCard.appendChild(Dom.el("div", { class: "card-header", html: `
      <div>
        <div class="card-title">Top customers</div>
        <div class="card-subtitle">Highest quoted value</div>
      </div>
    `}));
    const cWrap = Dom.el("div", { class: "chart-wrap" });
    if (topCustomers.length) {
      cWrap.appendChild(Charts.barChart({
        items: topCustomers,
        height: 280,
        horizontal: true,
        format: (v) => Format.moneyWhole(v),
      }));
    } else {
      cWrap.appendChild(Dom.el("div", { class: "state", html: "<p>No customer data yet.</p>" }));
    }
    topCard.appendChild(cWrap);
    grid.appendChild(topCard);
  }

  function renderTable() {
    const rows = filteredRows();
    const sorted = [...rows].sort((a, b) => {
      const ta = a.timestamp ? a.timestamp.getTime() : 0;
      const tb = b.timestamp ? b.timestamp.getTime() : 0;
      return tb - ta;
    });

    const wrap = $("#quoteTable");
    Dom.clear(wrap);

    const card = Dom.el("div", { class: "card", style: "padding:0;overflow:hidden;" });
    const head = Dom.el("div", { class: "card-header", style: "padding:16px 20px;margin:0;border-bottom:1px solid var(--hairline);" });
    head.appendChild(Dom.el("div", { html: `
      <div class="card-title">Quotation detail</div>
      <div class="card-subtitle">${sorted.length} of ${allRows.length} quotations</div>
    `}));
    card.appendChild(head);

    const tWrap = Dom.el("div", { class: "table-wrap", style: "border:none;border-radius:0;" });
    const table = Dom.el("table", { class: "data-table" });
    const thead = Dom.el("thead");
    const hr = Dom.el("tr");
    ["Quote no", "Customer", "Date", "Quote value", "Order received", "Order value", "Invoice no", "Status"].forEach((h) => {
      hr.appendChild(Dom.el("th", { text: h, class: ["Quote value", "Order value"].includes(h) ? "num" : "" }));
    });
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = Dom.el("tbody");
    if (!sorted.length) {
      const tr = Dom.el("tr");
      const td = Dom.el("td", { colspan: 8, style: "text-align:center;padding:32px;color:var(--ink-subtle);" });
      td.textContent = "No quotations match the current filters.";
      tr.appendChild(td);
      tbody.appendChild(tr);
    }
    sorted.forEach((r) => {
      const tr = Dom.el("tr");
      tr.appendChild(Dom.el("td", { text: r.number, class: "cell-title" }));
      tr.appendChild(Dom.el("td", { text: r.customer }));
      tr.appendChild(Dom.el("td", { text: Format.dateShort(r.timestamp) }));
      tr.appendChild(Dom.el("td", { text: Number.isFinite(r.amount) ? Format.money(r.amount) : "—", class: "num strong" }));
      tr.appendChild(Dom.el("td", { html: r.orderReceived === null ? "—" : r.orderReceived ? '<span class="badge badge-success badge-dot">Yes</span>' : '<span class="badge badge-primary badge-dot">No</span>' }));
      tr.appendChild(Dom.el("td", { text: Number.isFinite(r.orderValue) ? Format.money(r.orderValue) : "—", class: "num" }));
      tr.appendChild(Dom.el("td", { text: r.invoiceNo || "—" }));
      tr.appendChild(Dom.el("td", { html: badgeFor(r) }));
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    tWrap.appendChild(table);
    card.appendChild(tWrap);
    wrap.appendChild(card);
  }

  function renderDashboard() {
    const dash = $("#dashboard");
    Dom.clear(dash);

    const kpis = Dom.el("div", { class: "stat-grid", id: "kpis" });
    const charts = Dom.el("div", { class: "grid-2", id: "charts" });
    const table = Dom.el("div", { id: "quoteTable" });

    dash.appendChild(kpis);
    dash.appendChild(charts);
    dash.appendChild(table);

    renderKpis();
    renderCharts();
    renderTable();
  }

  async function load() {
    const refreshBtn = $("#refreshBtn");
    refreshBtn.classList.add("loading");

    try {
      const records = await Sheets.getRows(SHEET.sheetUrl, SHEET.gid);
      if (!records.length) throw new Error("The sheet contains no data rows.");

      const headers = Object.keys(records[0]);
      const cols = matchHeader(headers);

      const banner = $("#statusBanner");
      const unmapped = [];
      if (!cols.amount) unmapped.push("quote value");
      if (!cols.customer) unmapped.push("customer name");
      if (unmapped.length) {
        banner.innerHTML = "";
        banner.appendChild(Dom.el("div", { class: "status-banner warning", html: `
          <strong>Column mapping incomplete</strong>
          <span>&nbsp;—&nbsp;Could not find ${unmapped.join(" and ")} column(s) in the sheet header:
          <code>${headers.join(", ")}</code>.</span>
        `}));
      } else {
        banner.innerHTML = "";
      }

      allRows = buildRows(records, cols);
      renderFilters();
      renderDashboard();

      const now = new Date();
      $("#lastUpdated").textContent = `Updated ${now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
    } catch (err) {
      $("#statusBanner").innerHTML = "";
      $("#filters").innerHTML = "";
      const box = Dom.el("div", { class: "card", style: "text-align:center;padding:64px 24px;" });
      box.appendChild(Dom.el("div", { class: "state-icon", text: "⚠️" }));
      box.appendChild(Dom.el("h3", { text: "Could not load quotations data" }));
      box.appendChild(Dom.el("p", { text: err.message }));
      box.appendChild(Dom.el("p", { style: "font-size:13px;color:var(--ink-subtle);max-width:520px;margin:0 auto 16px;", text: "Make sure the sheet is shared as 'Anyone with the link can view' and that the URL in assets/js/config.js is correct." }));
      const retry = Dom.el("button", { class: "btn btn-secondary", text: "Try again" });
      retry.addEventListener("click", load);
      box.appendChild(retry);
      $("#dashboard").appendChild(box);
    } finally {
      refreshBtn.classList.remove("loading");
    }
  }

  $("#refreshBtn").addEventListener("click", load);
  load();
})();

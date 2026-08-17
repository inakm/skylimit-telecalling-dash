/* ============================================================
   Shared utilities: Google Sheets CSV fetching, parsing,
   formatting, DOM helpers, and small SVG chart renderers.
   ============================================================ */

const Sheets = (() => {
  const SHEET_URL = (sheetId, gid = 0) =>
    `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&tq=&gid=${gid}`;

  const GID_MAP = { 0: "Sheet1", 1: "Sheet2", 2: "Sheet3" };

  async function fetchJson(sheetId, gid = 0) {
    const res = await fetch(SHEET_URL(sheetId, gid));
    if (!res.ok) throw new Error(`Request failed (${res.status})`);
    const text = await res.text();
    const jsonText = text
      .replace(/^\/\*O_o\*\/\s*google\.visualization\.Query\.setResponse\(/, "")
      .replace(/\);\s*$/, "");
    const data = JSON.parse(jsonText);
    if (data.status !== "ok") throw new Error(data.errors && data.errors[0] && data.errors[0].reason || "Sheet could not be read");
    return data.table;
  }

  function rowsToObjects(table) {
    const headers = (table.cols || []).map((c) => (c.label || "").trim());
    return (table.rows || [])
      .filter((r) => r.c)
      .map((r) => {
        const obj = {};
        r.c.forEach((cell, i) => {
          const key = headers[i] || `col_${i}`;
          obj[key] = cell && cell.v !== null && cell.v !== undefined ? cell.v : "";
        });
        return obj;
      })
      .filter((r) => Object.values(r).some((v) => String(v).trim() !== ""));
  }

  function findSheetId(url) {
    const m = /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/.exec(url || "");
    return m ? m[1] : null;
  }

  async function getRows(url, gid = 0) {
    const sheetId = findSheetId(url);
    if (!sheetId) throw new Error("Invalid Google Sheets URL");
    const table = await fetchJson(sheetId, gid);
    return rowsToObjects(table);
  }

  function guessGid(url, rows) {
    return rows && rows.length ? 0 : 0;
  }

  return {
    fetchJson,
    rowsToObjects,
    findSheetId,
    getRows,
    GID_MAP,
    guessGid,
  };
})();

const Format = (() => {
  const currency = new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  const currencyWhole = new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  });

  function money(v) {
    const n = parseNumber(v);
    return Number.isFinite(n) ? currency.format(n) : "—";
  }

  function moneyWhole(v) {
    const n = parseNumber(v);
    return Number.isFinite(n) ? currencyWhole.format(n) : "—";
  }

  function number(v) {
    const n = parseNumber(v);
    return Number.isFinite(n) ? new Intl.NumberFormat("en-US").format(n) : "—";
  }

  function parseNumber(v) {
    if (typeof v === "number") return v;
    if (v === null || v === undefined || v === "") return NaN;
    const cleaned = String(v).replace(/[₹P,\s]/g, "").replace(/[()]/g, "-");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function parseDate(v) {
    if (!v) return null;
    if (v instanceof Date) return v;
    if (typeof v === "number" && v > 40000) {
      const d = new Date((v - 25569) * 86400 * 1000);
      return d;
    }
    if (typeof v === "number") return new Date(v);
    const s = String(v).trim();
    const gviz = /^Date\((\d{4}),(\d{1,2}),(\d{1,2})(?:,(\d{1,2}),(\d{1,2}),(\d{1,2}))?\)$/.exec(s);
    if (gviz) {
      return new Date(
        parseInt(gviz[1], 10),
        parseInt(gviz[2], 10),
        parseInt(gviz[3], 10),
        gviz[4] ? parseInt(gviz[4], 10) : 0,
        gviz[5] ? parseInt(gviz[5], 10) : 0,
        gviz[6] ? parseInt(gviz[6], 10) : 0
      );
    }
    const parts = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
    if (parts) {
      const year = parts[3].length === 2 ? 2000 + parseInt(parts[3], 10) : parseInt(parts[3], 10);
      return new Date(year, parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  function dateShort(d) {
    if (!d) return "—";
    if (typeof d === "string") d = parseDate(d);
    if (!d) return "—";
    return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
  }

  function relativeDays(date) {
    if (!date) return null;
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    return Math.round((d - now) / 86400000);
  }

  return {
    money,
    moneyWhole,
    number,
    parseNumber,
    parseDate,
    dateShort,
    relativeDays,
  };
})();

const Dom = (() => {
  function el(tag, attrs = {}, children = []) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "html") node.innerHTML = v;
      else if (k === "text") node.textContent = v;
      else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
      else if (v !== null && v !== undefined) node.setAttribute(k, v);
    }
    for (const c of Array.isArray(children) ? children : [children]) {
      if (c) node.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    }
    return node;
  }

  function clear(node) {
    while (node.firstChild) node.removeChild(node.firstChild);
  }

  function debounce(fn, ms = 250) {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }

  return { el, clear, debounce };
})();

/* ---------- SVG chart renderers (no dependencies) ---------- */

const Charts = (() => {
  const NS = "http://www.w3.org/2000/svg";

  function svgEl(tag, attrs = {}) {
    const node = document.createElementNS(NS, tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (v !== null && v !== undefined) node.setAttribute(k, v);
    }
    return node;
  }

  const PALETTE = ["#5e6ad2", "#828fff", "#7a7fad", "#27a644", "#b45309", "#d33b3b", "#a8abb3"];

  function wrap(width, height, title) {
    const svg = svgEl("svg", { viewBox: `0 0 ${width} ${height}`, width: "100%", height });
    svg.setAttribute("role", "img");
    if (title) {
      const t = svgEl("title");
      t.textContent = title;
      svg.appendChild(t);
    }
    return svg;
  }

  function tooltipAttach(svg, tooltipEl) {
    svg.addEventListener("mousemove", (e) => {
      const rect = svg.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      tooltipEl.style.left = Math.min(x + 12, rect.width - 140) + "px";
      tooltipEl.style.top = y - 10 + "px";
    });
    svg.addEventListener("mouseleave", () => {
      tooltipEl.style.opacity = 0;
    });
  }

  function lineChart({ points, labels, width = 640, height = 260, color = PALETTE[0], format = (v) => v }) {
    const tooltip = document.querySelector(".chart-tooltip");
    const padding = { top: 16, right: 16, bottom: 28, left: 56 };
    const innerW = width - padding.left - padding.right;
    const innerH = height - padding.top - padding.bottom;

    const values = points.map((p) => p.y);
    const max = Math.max(...values, 0) * 1.08 || 1;
    const min = 0;
    const xStep = points.length > 1 ? innerW / (points.length - 1) : 0;
    const yFor = (v) => padding.top + innerH - ((v - min) / (max - min)) * innerH;

    const svg = wrap(width, height, "Trend chart");

    const xAxis = svgEl("line", { x1: padding.left, y1: padding.top + innerH, x2: padding.left + innerW, y2: padding.top + innerH, stroke: "#e6e8eb", "stroke-width": 1 });
    svg.appendChild(xAxis);

    [0, 0.25, 0.5, 0.75, 1].forEach((f) => {
      const y = padding.top + innerH - f * innerH;
      const v = min + (max - min) * f;
      const gridLine = svgEl("line", { x1: padding.left, y1: y, x2: padding.left + innerW, y2: y, stroke: "#f1f2f4", "stroke-width": 1 });
      svg.appendChild(gridLine);
      const label = svgEl("text", { x: padding.left - 8, y: y + 4, "text-anchor": "end", "font-size": 11, fill: "#8a8f98" });
      label.textContent = format(v);
      svg.appendChild(label);
    });

    const pathD = points.map((p, i) => {
      const x = padding.left + i * xStep;
      const y = yFor(p.y);
      return (i === 0 ? "M" : "L") + x + "," + y;
    }).join(" ");

    const areaPath = pathD + ` L${padding.left + (points.length - 1) * xStep},${padding.top + innerH} L${padding.left},${padding.top + innerH} Z`;
    const area = svgEl("path", { d: areaPath, fill: color, opacity: 0.08 });
    svg.appendChild(area);

    const line = svgEl("path", { d: pathD, fill: "none", stroke: color, "stroke-width": 2.5, "stroke-linecap": "round", "stroke-linejoin": "round" });
    svg.appendChild(line);

    points.forEach((p, i) => {
      const x = padding.left + i * xStep;
      const y = yFor(p.y);
      const dot = svgEl("circle", { cx: x, cy: y, r: 4, fill: "#ffffff", stroke: color, "stroke-width": 2 });
      dot.style.cursor = "pointer";
      dot.addEventListener("mousemove", (e) => {
        tooltip.innerHTML = `<strong>${labels[i] || ""}</strong><br>${format(p.y)}`;
        tooltip.style.opacity = 1;
        const rect = svg.getBoundingClientRect();
        tooltip.style.left = Math.min(e.clientX - rect.left + 12, rect.width - 160) + "px";
        tooltip.style.top = y - 24 + "px";
      });
      dot.addEventListener("mouseleave", () => { tooltip.style.opacity = 0; });
      svg.appendChild(dot);

      if (points.length <= 16) {
        const lx = svgEl("text", { x, y: padding.top + innerH + 16, "text-anchor": "middle", "font-size": 10, fill: "#a8abb3" });
        lx.textContent = labels[i];
        svg.appendChild(lx);
      }
    });

    return svg;
  }

  function barChart({ items, width = 640, height = 260, format = (v) => v, horizontal = false }) {
    const tooltip = document.querySelector(".chart-tooltip");
    const padding = { top: 16, right: 16, bottom: 28, left: horizontal ? 96 : 56 };
    const innerW = width - padding.left - padding.right;
    const innerH = height - padding.top - padding.bottom;

    const max = Math.max(...items.map((i) => i.value), 0) * 1.08 || 1;
    const svg = wrap(width, height, "Bar chart");

    if (!horizontal) {
      [0, 0.25, 0.5, 0.75, 1].forEach((f) => {
        const y = padding.top + innerH - f * innerH;
        const v = max * f;
        const gridLine = svgEl("line", { x1: padding.left, y1: y, x2: padding.left + innerW, y2: y, stroke: "#f1f2f4", "stroke-width": 1 });
        svg.appendChild(gridLine);
        const label = svgEl("text", { x: padding.left - 8, y: y + 4, "text-anchor": "end", "font-size": 11, fill: "#8a8f98" });
        label.textContent = format(v);
        svg.appendChild(label);
      });

      const slot = innerW / items.length;
      const barW = Math.min(40, slot * 0.55);
      items.forEach((item, i) => {
        const x = padding.left + i * slot + (slot - barW) / 2;
        const h = (item.value / max) * innerH;
        const y = padding.top + innerH - h;
        const bar = svgEl("rect", { x, y, width: barW, height: Math.max(h, item.value ? 2 : 0), rx: 4, fill: item.color || PALETTE[0] });
        bar.style.cursor = "pointer";
        bar.addEventListener("mousemove", (e) => {
          tooltip.innerHTML = `<strong>${item.label}</strong><br>${format(item.value)}`;
          tooltip.style.opacity = 1;
          const rect = svg.getBoundingClientRect();
          tooltip.style.left = Math.min(e.clientX - rect.left + 12, rect.width - 160) + "px";
          tooltip.style.top = y + "px";
        });
        bar.addEventListener("mouseleave", () => { tooltip.style.opacity = 0; });
        svg.appendChild(bar);
        const lx = svgEl("text", { x: x + barW / 2, y: padding.top + innerH + 16, "text-anchor": "middle", "font-size": 10, fill: "#a8abb3" });
        lx.textContent = item.label;
        svg.appendChild(lx);
      });
    } else {
      const slot = innerH / items.length;
      const barH = Math.min(24, slot * 0.55);
      const xAxis = svgEl("line", { x1: padding.left, y1: padding.top, x2: padding.left, y2: padding.top + innerH, stroke: "#e6e8eb", "stroke-width": 1 });
      svg.appendChild(xAxis);

      items.forEach((item, i) => {
        const y = padding.top + i * slot + (slot - barH) / 2;
        const w = (item.value / max) * innerW;
        const label = svgEl("text", { x: padding.left - 8, y: y + barH / 2 + 4, "text-anchor": "end", "font-size": 11, fill: "#4c4f56" });
        label.textContent = item.label;
        svg.appendChild(label);
        const bar = svgEl("rect", { x: padding.left, y, width: Math.max(w, item.value ? 2 : 0), height: barH, rx: 4, fill: item.color || PALETTE[0] });
        bar.style.cursor = "pointer";
        bar.addEventListener("mousemove", (e) => {
          tooltip.innerHTML = `<strong>${item.label}</strong><br>${format(item.value)}`;
          tooltip.style.opacity = 1;
          const rect = svg.getBoundingClientRect();
          tooltip.style.left = Math.min(e.clientX - rect.left + 12, rect.width - 160) + "px";
          tooltip.style.top = y + "px";
        });
        bar.addEventListener("mouseleave", () => { tooltip.style.opacity = 0; });
        svg.appendChild(bar);
      });
    }

    return svg;
  }

  function donutChart({ segments, size = 220, thickness = 28, format = (v) => v }) {
    const tooltip = document.querySelector(".chart-tooltip");
    const total = segments.reduce((s, x) => s + x.value, 0);
    const r = (size - thickness) / 2;
    const cx = size / 2;
    const cy = size / 2;
    const svg = wrap(size, size, "Donut chart");

    let angle = -Math.PI / 2;
    segments.forEach((seg, i) => {
      if (!seg.value) return;
      const frac = seg.value / total;
      const startAngle = angle;
      const endAngle = angle + frac * Math.PI * 2;
      const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;

      const x1 = cx + r * Math.cos(startAngle);
      const y1 = cy + r * Math.sin(startAngle);
      const x2 = cx + r * Math.cos(endAngle);
      const y2 = cy + r * Math.sin(endAngle);

      const d = `M${x1},${y1} A${r},${r} 0 ${largeArc} 1 ${x2},${y2}`;
      const arc = svgEl("path", {
        d,
        fill: "none",
        stroke: seg.color || PALETTE[i % PALETTE.length],
        "stroke-width": thickness,
        "stroke-linecap": "butt",
      });
      arc.style.cursor = "pointer";
      arc.addEventListener("mousemove", (e) => {
        tooltip.innerHTML = `<strong>${seg.label}</strong><br>${format(seg.value)}`;
        tooltip.style.opacity = 1;
        const rect = svg.getBoundingClientRect();
        tooltip.style.left = Math.min(e.clientX - rect.left + 12, rect.width - 140) + "px";
        tooltip.style.top = e.clientY - rect.top - 10 + "px";
      });
      arc.addEventListener("mouseleave", () => { tooltip.style.opacity = 0; });
      svg.appendChild(arc);
      angle = endAngle;
    });

    if (total > 0) {
      const c = svgEl("circle", { cx, cy, r: r - thickness / 2, fill: "#ffffff" });
      svg.appendChild(c);
      const val = svgEl("text", { x: cx, y: cy + 4, "text-anchor": "middle", "font-size": 18, "font-weight": 600, fill: "#191a1b" });
      val.textContent = format(total);
      svg.appendChild(val);
    }

    return svg;
  }

  function groupedBars({ groups, series, width = 640, height = 260, format = (v) => v }) {
    const tooltip = document.querySelector(".chart-tooltip");
    const padding = { top: 16, right: 16, bottom: 28, left: 56 };
    const innerW = width - padding.left - padding.right;
    const innerH = height - padding.top - padding.bottom;

    const max = Math.max(...groups.flatMap((g) => g.values)) * 1.08 || 1;
    const svg = wrap(width, height, "Grouped bar chart");

    [0, 0.25, 0.5, 0.75, 1].forEach((f) => {
      const y = padding.top + innerH - f * innerH;
      const v = max * f;
      const gridLine = svgEl("line", { x1: padding.left, y1: y, x2: padding.left + innerW, y2: y, stroke: "#f1f2f4", "stroke-width": 1 });
      svg.appendChild(gridLine);
      const label = svgEl("text", { x: padding.left - 8, y: y + 4, "text-anchor": "end", "font-size": 11, fill: "#8a8f98" });
      label.textContent = format(v);
      svg.appendChild(label);
    });

    const slot = innerW / groups.length;
    const groupW = slot * 0.62;
    const barW = groupW / series.length;

    groups.forEach((group, i) => {
      const gx = padding.left + i * slot + (slot - groupW) / 2;
      group.values.forEach((v, s) => {
        const x = gx + s * barW + (series.length > 1 ? 2 : 0);
        const h = (v / max) * innerH;
        const y = padding.top + innerH - h;
        const bar = svgEl("rect", { x, y, width: Math.max(barW - (series.length > 1 ? 4 : 0), 2), height: Math.max(h, v ? 2 : 0), rx: 3, fill: PALETTE[s % PALETTE.length] });
        bar.style.cursor = "pointer";
        bar.addEventListener("mousemove", (e) => {
          tooltip.innerHTML = `<strong>${group.label} — ${series[s].label}</strong><br>${format(v)}`;
          tooltip.style.opacity = 1;
          const rect = svg.getBoundingClientRect();
          tooltip.style.left = Math.min(e.clientX - rect.left + 12, rect.width - 160) + "px";
          tooltip.style.top = y + "px";
        });
        bar.addEventListener("mouseleave", () => { tooltip.style.opacity = 0; });
        svg.appendChild(bar);
      });
      const lx = svgEl("text", { x: gx + groupW / 2, y: padding.top + innerH + 16, "text-anchor": "middle", "font-size": 10, fill: "#a8abb3" });
      lx.textContent = group.label;
      svg.appendChild(lx);
    });

    return svg;
  }

  return {
    lineChart,
    barChart,
    donutChart,
    groupedBars,
    PALETTE,
  };
})();

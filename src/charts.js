// Charts drawn on a canvas in the client house style. One function per chart in the report.
// Each returns {png: Uint8Array, w, h, note: null}; w/h are pixels (used to keep the picture's aspect ratio).
import { NAVY, GREEN, RED, CAT5, INK, MUTED, GRID, CHART_DPI } from "./config.js";
import { MONTH_NAMES, commas } from "./util.js";

const DPI = CHART_DPI, PT = DPI / 72;
const FONT = "'DejaVu Sans', Verdana, Arial, Helvetica, sans-serif";
let provider = null;
export const setCanvasProvider = (p) => { provider = p; };

function newCanvas(w, h) {
  w = Math.round(w); h = Math.round(h);
  if (provider) return provider(w, h);
  const c = document.createElement("canvas"); c.width = w; c.height = h; return c;
}
async function toPng(c) {
  if (c.toBuffer) return new Uint8Array(c.toBuffer("image/png"));
  const blob = await new Promise((res) => c.toBlob(res, "image/png"));
  return new Uint8Array(await blob.arrayBuffer());
}
const A = (x, dp = 0) => (x < 0 ? "-" : "") + "A$" + commas(x, dp);
const short = (x) => {
  const a = Math.abs(x);
  const t = a >= 1e6 ? `${(a / 1e6).toFixed(2)}M` : a >= 1e4 ? `${(a / 1e3).toFixed(1)}k` : commas(a, 0);
  return (x < 0 ? "-" : "") + "A$" + t;
};
const fnt = (pt, { bold = false, italic = false } = {}) => `${italic ? "italic " : ""}${bold ? "bold " : ""}${pt * PT}px ${FONT}`;

class Fig {
  constructor(wIn, hIn, title) {
    this.W = Math.round(wIn * DPI); this.H = Math.round(hIn * DPI);
    this.c = newCanvas(this.W, this.H);
    this.g = this.c.getContext("2d");
    this.g.fillStyle = "#fff"; this.g.fillRect(0, 0, this.W, this.H);
    this.titleBottom = 0;
    if (title) this.title(title);
  }
  title(t) {
    const g = this.g;
    g.font = fnt(11.5, { bold: true }); g.fillStyle = NAVY; g.textAlign = "left"; g.textBaseline = "top";
    g.fillText(t, 0.012 * this.W, 0.015 * this.H);
    this.titleBottom = 0.015 * this.H + 11.5 * PT * 1.25;
  }
  text(s, x, y, { pt = 9, color = INK, align = "left", base = "alphabetic", bold = false, italic = false } = {}) {
    const g = this.g;
    g.font = fnt(pt, { bold, italic }); g.fillStyle = color; g.textAlign = align; g.textBaseline = base;
    g.fillText(s, x, y);
  }
  width(s, pt = 9, bold = false) { this.g.font = fnt(pt, { bold }); return this.g.measureText(s).width; }
  line(x1, y1, x2, y2, color, lw = 0.8) {
    const g = this.g; g.strokeStyle = color; g.lineWidth = lw * PT; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke();
  }
  rect(x, y, w, h, color) { this.g.fillStyle = color; this.g.fillRect(x, y, w, h); }
  async png(crop) {
    let c = this.c;
    if (crop) { // crop to {x,y,w,h}
      const c2 = newCanvas(crop.w, crop.h); const g2 = c2.getContext("2d");
      g2.fillStyle = "#fff"; g2.fillRect(0, 0, crop.w, crop.h); g2.drawImage(this.c, -crop.x, -crop.y); c = c2;
    }
    return { png: await toPng(c), w: c.width, h: c.height, note: null };
  }
}

function niceTicks(lo, hi, n = 6) {
  const span = hi - lo || 1;
  const raw = span / n, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) || 10 * mag;
  const out = [];
  for (let v = Math.ceil(lo / step - 1e-9) * step; v <= hi + step * 1e-9; v += step) out.push(Math.abs(v) < step * 1e-9 ? 0 : v);
  return out;
}

// Cartesian plot area with a y axis (value) and category x positions.
// opts: {yLabel, xLabel, xTicks:[labels], rot, legend:[{label,color}], ymin,ymax, hGrid:true}
function plot(fig, opts) {
  const { W, H } = fig;
  const g = fig.g;
  const tickPt = opts.tickPt || 9;
  const yTicks = niceTicks(opts.ymin, opts.ymax, opts.nTicks || 6);
  const lo = Math.min(opts.ymin, yTicks[0]), hi = Math.max(opts.ymax, yTicks[yTicks.length - 1]);
  const fmtY = opts.fmtY || ((v) => commas(v, 0).replace(/^/, v < 0 ? "-" : ""));
  const yW = Math.max(...yTicks.map((v) => fig.width(fmtY(v), tickPt)));
  const left = (opts.yLabel ? 10 * PT * 1.6 : 0) + yW + 12 * PT * 0.9 + 4 * PT;
  const right = 0.03 * W;
  const top = Math.max(fig.titleBottom + 0.035 * H, 0.1 * H);
  let bottom = 6 * PT;
  const xlabH = opts.xLabel ? 9 * PT * 1.7 : 0;
  const legH = opts.legend ? 9 * PT * 2.2 : 0;
  const tickH = opts.rot ? Math.max(...opts.xTicks.map((s) => fig.width(s, opts.xTickPt || 8))) * Math.sin(opts.rot * Math.PI / 180) + 10 * PT : (opts.xTickPt || tickPt) * PT * 0.0 + (opts.xTickPt || tickPt) * PT * 1.6;
  bottom += xlabH + legH + tickH;
  const ax = { x0: left, x1: W - right, y0: top, y1: H - bottom };
  const sy = (v) => ax.y1 - (v - lo) / (hi - lo) * (ax.y1 - ax.y0);
  // grid + ticks
  for (const v of yTicks) {
    if (opts.hGrid !== false) fig.line(ax.x0, sy(v), ax.x1, sy(v), GRID, 0.7);
    fig.line(ax.x0 - 3.5 * PT, sy(v), ax.x0, sy(v), MUTED, 0.8);
    fig.text(fmtY(v), ax.x0 - 6 * PT, sy(v), { pt: tickPt, color: MUTED, align: "right", base: "middle" });
  }
  fig.line(ax.x0, ax.y0, ax.x0, ax.y1, GRID, 0.8);
  fig.line(ax.x0, ax.y1, ax.x1, ax.y1, GRID, 0.8);
  if (opts.yLabel) {
    g.save(); g.translate(10 * PT * 0.9, (ax.y0 + ax.y1) / 2); g.rotate(-Math.PI / 2);
    fig.text(opts.yLabel, 0, 0, { pt: 9, color: MUTED, align: "center", base: "top" }); g.restore();
  }
  const n = opts.xTicks.length;
  const sx = (i) => ax.x0 + (i + 0.5) / n * (ax.x1 - ax.x0);
  const slot = (ax.x1 - ax.x0) / n;
  opts.xTicks.forEach((lab, i) => {
    const x = sx(i);
    fig.line(x, ax.y1, x, ax.y1 + 3.5 * PT, MUTED, 0.8);
    if (opts.rot) {
      g.save(); g.translate(x + 3 * PT, ax.y1 + 7 * PT); g.rotate(-opts.rot * Math.PI / 180);
      fig.text(lab, 0, 0, { pt: opts.xTickPt || 8, color: MUTED, align: "right", base: "middle" }); g.restore();
    } else {
      fig.text(lab, x, ax.y1 + 6 * PT, { pt: opts.xTickPt || tickPt, color: MUTED, align: "center", base: "top" });
    }
  });
  let yy = ax.y1 + tickH;
  if (opts.xLabel) { fig.text(opts.xLabel, (ax.x0 + ax.x1) / 2, yy + 3 * PT, { pt: 9, color: MUTED, align: "center", base: "top" }); yy += xlabH; }
  if (opts.legend) {
    const items = opts.legend, sw = 12 * PT, gap = 14 * PT;
    const widths = items.map((it) => sw + 5 * PT + fig.width(it.label, 8.5));
    const tot = widths.reduce((a, b) => a + b, 0) + gap * (items.length - 1);
    let x = (ax.x0 + ax.x1) / 2 - tot / 2;
    const ly = yy + 9 * PT;
    items.forEach((it, i) => {
      fig.rect(x, ly - 4.5 * PT, sw, 9 * PT, it.color);
      fig.text(it.label, x + sw + 5 * PT, ly, { pt: 8.5, base: "middle" });
      x += widths[i] + gap;
    });
  }
  return { ...ax, sy, sx, slot, lo, hi };
}

const yRange = (vals, pad = 0.15, includeZero = true) => {
  let mn = Math.min(...vals), mx = Math.max(...vals);
  if (includeZero) { mn = Math.min(mn, 0); mx = Math.max(mx, 0); }
  const span = mx - mn || 1;
  return [mn < 0 ? mn - pad * span : mn, mx > 0 ? mx + pad * span : mx];
};

export async function ytd(w, h, label, vals) {
  const names = Object.keys(vals), net = vals["Net Income"];
  const f = new Fig(w, h, `${label}: ${net < 0 ? "net loss" : "net profit"} of ${A(Math.abs(net))}`);
  const [lo, hi] = yRange(names.map((n) => vals[n]));
  const p = plot(f, { ymin: lo, ymax: hi, yLabel: "A$", xTicks: names });
  const cols = [CAT5[0], CAT5[4], CAT5[1], net >= 0 ? GREEN : RED];
  names.forEach((n, i) => {
    const v = vals[n], bw = p.slot * 0.55, x = p.sx(i) - bw / 2;
    f.rect(x, Math.min(p.sy(v), p.sy(0)), bw, Math.abs(p.sy(v) - p.sy(0)), cols[i]);
    f.text(short(v), p.sx(i), v >= 0 ? p.sy(v) - 3 * PT : p.sy(v) + 3 * PT, { pt: 9.5, bold: true, align: "center", base: v >= 0 ? "bottom" : "top" });
  });
  f.line(p.x0, p.sy(0), p.x1, p.sy(0), MUTED, 0.8);
  return f.png();
}

export async function fees(w, h, monthName, feeByGw) {
  const names = Object.keys(feeByGw);
  const tot = names.reduce((s, n) => s + feeByGw[n], 0);
  const top = names.reduce((a, b) => (feeByGw[b] > feeByGw[a] ? b : a));
  const f = new Fig(w, h, `${top} is ${Math.round(feeByGw[top] / tot * 100)}% of ${monthName}'s ${A(tot)} merchant fees`);
  const p = plot(f, { ymin: 0, ymax: Math.max(...names.map((n) => feeByGw[n])) * 1.2, yLabel: "A$", xTicks: names });
  names.forEach((n, i) => {
    const v = feeByGw[n], bw = p.slot * 0.55;
    f.rect(p.sx(i) - bw / 2, p.sy(v), bw, p.sy(0) - p.sy(v), CAT5[0]);
    f.text(`${Math.round(v / tot * 100)}%`, p.sx(i), p.sy(v) - 3 * PT, { align: "center", base: "bottom" });
    f.text(A(v), p.sx(i), p.sy(v) - 3 * PT - 9 * PT * 1.25, { align: "center", base: "bottom" });
  });
  return f.png();
}

export async function expenditure(w, h, items, adsTotal) {
  items = [...items].sort((a, b) => b[1] - a[1]);
  const tot = items.reduce((s, x) => s + x[1], 0), top3 = items.slice(0, 3).reduce((s, x) => s + x[1], 0);
  const f = new Fig(w, h, `Three costs make up ${Math.round(top3 / tot * 100)}% of running costs (excl. social ads)`);
  const [lo, hi] = yRange(items.map((x) => x[1]));
  const p = plot(f, { ymin: lo, ymax: hi, yLabel: "A$", xTicks: items.map((x) => x[0]), rot: 40, xTickPt: 8 });
  items.forEach(([, v], i) => {
    const bw = p.slot * 0.6;
    f.rect(p.sx(i) - bw / 2, Math.min(p.sy(v), p.sy(0)), bw, Math.abs(p.sy(v) - p.sy(0)), CAT5[0]);
    f.text(commas(v, 0).replace(/^/, v < 0 ? "-" : ""), p.sx(i), v >= 0 ? p.sy(v) - 2 * PT : p.sy(v) + 2 * PT, { pt: 8, align: "center", base: v >= 0 ? "bottom" : "top" });
  });
  f.line(p.x0, p.sy(0), p.x1, p.sy(0), MUTED, 0.8);
  f.text(`Social media ads (${A(adsTotal)}) are shown in the P&L, not here`, p.x1 - 4 * PT, p.y0 + 4 * PT, { pt: 8, color: MUTED, italic: true, align: "right", base: "top" });
  return f.png();
}

function grouped(f, p, series, nDays, bwFrac) {
  const k = series.length, bw = p.slot * bwFrac;
  series.forEach((s, j) => {
    s.vals.forEach((v, i) => {
      const x = p.sx(i) - (k * bw) / 2 + j * bw;
      f.rect(x, Math.min(p.sy(v), p.sy(0)), bw, Math.abs(p.sy(v) - p.sy(0)), s.color);
    });
  });
}

export async function daywiseSales(w, h, dates, orders, returns, ship) {
  const first = orders.slice(0, 7).reduce((a, b) => a + b, 0) / 7, last = orders.slice(-7).reduce((a, b) => a + b, 0) / 7;
  const f = new Fig(w, h, `Orders fell from about ${Math.round(first)} a day to ${Math.round(last)} a day over the month`);
  const all = [...orders, ...returns, ...ship];
  const [lo, hi] = yRange(all, 0.05);
  const series = [{ vals: orders, color: CAT5[0] }, { vals: returns, color: CAT5[1] }, { vals: ship, color: CAT5[4] }];
  const p = plot(f, { ymin: lo, ymax: hi, xTicks: dates.map((d) => String(d.d)), xTickPt: 7.5, xLabel: `Day of ${MONTH_NAMES[dates[0].m]}`,
    legend: [{ label: "Number of orders", color: CAT5[0] }, { label: "Returns (A$)", color: CAT5[1] }, { label: "Shipping charges collected (A$)", color: CAT5[4] }] });
  grouped(f, p, series, dates.length, 0.28);
  f.line(p.x0, p.sy(0), p.x1, p.sy(0), MUTED, 0.8);
  return f.png();
}

export async function daywiseCogs(w, h, dates, cogs, ship) {
  const sum = (a) => a.reduce((s, x) => s + x, 0);
  const ratio = sum(cogs) ? sum(ship) / sum(cogs) : 0;
  const f = new Fig(w, h, `Shipping costs about ${ratio.toFixed(1)}x as much as the product itself, every day`);
  const [lo, hi] = yRange([...cogs, ...ship], 0.05);
  const p = plot(f, { ymin: lo, ymax: hi, yLabel: "A$", xTicks: dates.map((d) => String(d.d)), xTickPt: 7.5, xLabel: `Day of ${MONTH_NAMES[dates[0].m]}`,
    legend: [{ label: "Cost of goods sold", color: CAT5[0] }, { label: "Shipping cost", color: CAT5[1] }] });
  grouped(f, p, [{ vals: cogs, color: CAT5[0] }, { vals: ship, color: CAT5[1] }], dates.length, 0.4);
  return f.png();
}

async function pie(w, h, title, parts, colors, legendLabels) {
  const f = new Fig(w, h, title);
  const g = f.g, tot = parts.reduce((s, x) => s + x[1], 0);
  const legW = Math.max(...legendLabels.map((l) => f.width(l, 9))) + 18 * PT;
  const D = Math.min(f.H * 0.7, f.W * 0.42, f.W * 0.94 - legW - f.W * 0.1);
  const cx = f.W * 0.05 + D / 2 + f.W * 0.04, cy = f.titleBottom + (f.H - f.titleBottom) / 2;
  let a = -Math.PI / 2;
  parts.forEach(([, v], i) => {
    const da = v / tot * Math.PI * 2;
    g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, D / 2, a, a + da); g.closePath();
    g.fillStyle = colors[i]; g.fill(); g.strokeStyle = "#fff"; g.lineWidth = 2 * PT; g.stroke();
    a += da;
  });
  const lx = cx + D / 2 + f.W * 0.05, lh = 9 * PT * 1.9, y0 = cy - (parts.length * lh) / 2 + lh / 2;
  parts.forEach((_, i) => {
    f.rect(lx, y0 + i * lh - 4.5 * PT, 12 * PT, 9 * PT, colors[i]);
    f.text(legendLabels[i], lx + 18 * PT, y0 + i * lh, { pt: 9, base: "middle" });
  });
  return f.png();
}

export async function costPie(w, h, parts) {
  parts = [...parts].sort((a, b) => b[1] - a[1]);
  const tot = parts.reduce((s, x) => s + x[1], 0), big = parts[0];
  return pie(w, h, `${big[0]} is ${Math.round(big[1] / tot * 100)}% of cost of sales (${A(tot)})`, parts, CAT5.slice(0, 4).slice(0, parts.length),
    parts.map(([n, v]) => `${n}   ${A(v)}  (${Math.round(v / tot * 100)}%)`));
}

export async function countryPie(w, h, countries, monthName, minShare = 0.015) {
  const pos = countries.map(([n, v]) => [n.replace("Product Revenue ", ""), v]);
  const tot = pos.reduce((s, x) => s + x[1], 0);
  const big = pos.filter((x) => x[1] / tot >= minShare).sort((a, b) => b[1] - a[1]);
  const rest = tot - big.reduce((s, x) => s + x[1], 0), nrest = pos.length - big.length;
  const parts = [...big, ...(nrest ? [[`All other (${nrest} countries)`, rest]] : [])];
  const t2 = big.slice(0, 2);
  let cols = [...CAT5, ...CAT5, ...CAT5].slice(0, parts.length);
  if (nrest) cols[cols.length - 1] = CAT5[4];
  if (nrest && parts.length - 1 >= 4) {
    cols = [...CAT5.slice(0, 4), ...Array(Math.max(0, parts.length - 5)).fill("#B9C1CA"), "#8E9AA6"].slice(0, parts.length);
  }
  return pie(w, h, `${t2[0][0]} and ${t2[1][0]} bring in ${Math.round((t2[0][1] + t2[1][1]) / tot * 100)}% of ${monthName} revenue`,
    parts, cols, parts.map(([n, v]) => `${n}   ${(v / tot * 100).toFixed(1)}%`));
}

export async function placeholder(w, h, title, msg) {
  const f = new Fig(w, h, null);
  const g = f.g;
  g.fillStyle = "#F2F5F8"; g.strokeStyle = GRID; g.lineWidth = 1.2 * PT;
  g.fillRect(0.01 * f.W, 0.06 * f.H, 0.98 * f.W, 0.9 * f.H); g.strokeRect(0.01 * f.W, 0.06 * f.H, 0.98 * f.W, 0.9 * f.H);
  const two = h >= 0.8;
  f.text(title + (two ? "" : "  -  " + msg), f.W / 2, f.H * (two ? 0.38 : 0.5), { pt: two ? 12 : 9.5, bold: true, color: NAVY, align: "center", base: "middle" });
  if (two) f.text(msg, f.W / 2, f.H * 0.66, { pt: 9.5, color: MUTED, align: "center", base: "middle" });
  return f.png();
}

export async function productTable(w, products, total) {
  const rows = [...products].sort((a, b) => b[1][0] - a[1][0]);
  const rh = 0.27 * 1.3 * 0.82 * DPI * 0.82, hh = rh * 1.7;     // row / header height
  const n = rows.length + 1;
  const H = Math.ceil(hh + rh * n + 0.1 * DPI);
  const f = new Fig(w, H / DPI, null);
  const g = f.g, fmt = (v) => commas(v, 2).replace(/^/, v < 0 ? "-" : "");
  const fr = [0.34, 0.13, 0.12, 0.10, 0.12, 0.095, 0.095];
  const xs = [0]; fr.forEach((x) => xs.push(xs[xs.length - 1] + x * f.W));
  const head = [["Product title"], ["Gross sales", "(A$)"], ["Discounts"], ["Refunds"], ["Net sales", "(A$)"], ["Share of", "gross %"], ["Share of", "net %"]];
  g.fillStyle = NAVY; g.fillRect(0, 0, f.W, hh);
  head.forEach((lines, c) => {
    lines.forEach((ln, i) => {
      const y = hh / 2 + (i - (lines.length - 1) / 2) * 7.5 * PT * 1.25;
      f.text(ln, c === 0 ? xs[c] + 4 * PT : xs[c + 1] - 4 * PT, y, { pt: 7.5, bold: true, color: "#fff", align: c === 0 ? "left" : "right", base: "middle" });
    });
  });
  const body = rows.map(([nm, v]) => [nm.length <= 32 ? nm : nm.slice(0, 31) + "…", fmt(v[0]), fmt(v[1]), fmt(v[2]), fmt(v[3]), v[4].toFixed(2), v[5].toFixed(2)]);
  body.push(["Total Sales", fmt(total[0]), fmt(total[1]), fmt(total[2]), fmt(total[3]), "100.00", "100.00"]);
  body.forEach((r, i) => {
    const y = hh + i * rh, last = i === body.length - 1;
    if (last) f.rect(0, y, f.W, rh, "#DCE6F1"); else if ((i + 2) % 2 === 0) f.rect(0, y, f.W, rh, "#F2F5F8");
    r.forEach((s, c) => f.text(s, c === 0 ? xs[c] + 4 * PT : xs[c + 1] - 4 * PT, y + rh / 2, { pt: 7.5, bold: last, color: last ? NAVY : INK, align: c === 0 ? "left" : "right", base: "middle" }));
    f.line(0, y + rh, f.W, y + rh, GRID, 0.6);
  });
  xs.forEach((x) => f.line(x, 0, x, hh + rh * body.length, GRID, 0.6));
  f.line(f.W - 1, 0, f.W - 1, hh + rh * body.length, GRID, 0.6);
  return f.png({ x: 0, y: 0, w: f.W, h: Math.ceil(hh + rh * body.length + 2) });
}

export async function productChart(w, h, products) {
  const rows = [...products].sort((a, b) => b[1][0] - a[1][0]).slice(0, 12).reverse();
  const f = new Fig(w, h, "Gross sales vs net sales by product (top 12 by gross sales)");
  const g = f.g;
  const maxv = Math.max(...rows.map(([, v]) => Math.max(v[0], v[3])));
  const ticks = niceTicks(0, maxv * 1.05, 6);
  const labW = Math.max(...rows.map(([n]) => f.width(n, 8))) + 10 * PT;
  const ax = { x0: labW + 6 * PT, x1: f.W * 0.97, y0: Math.max(f.titleBottom + 0.035 * f.H, 0.1 * f.H), y1: f.H - 9 * PT * 2.2 };
  const hi = ticks[ticks.length - 1];
  const sx = (v) => ax.x0 + v / hi * (ax.x1 - ax.x0);
  ticks.forEach((t) => {
    f.line(sx(t), ax.y0, sx(t), ax.y1, GRID, 0.7);
    f.line(sx(t), ax.y1, sx(t), ax.y1 + 3.5 * PT, MUTED, 0.8);
    f.text(commas(t, 0), sx(t), ax.y1 + 6 * PT, { color: MUTED, align: "center", base: "top" });
  });
  f.line(ax.x0, ax.y0, ax.x0, ax.y1, GRID, 0.8); f.line(ax.x0, ax.y1, ax.x1, ax.y1, GRID, 0.8);
  const slot = (ax.y1 - ax.y0) / rows.length, bh = slot * 0.38;
  rows.forEach(([n, v], i) => {
    const cy = ax.y1 - (i + 0.5) * slot;
    f.rect(ax.x0, cy - bh, sx(v[0]) - ax.x0, bh, CAT5[0]);
    f.rect(ax.x0, cy, sx(v[3]) - ax.x0, bh, CAT5[2]);
    f.text(n, ax.x0 - 6 * PT, cy, { pt: 8, color: MUTED, align: "right", base: "middle" });
  });
  const lg = [{ label: "Gross sales (A$)", color: CAT5[0] }, { label: "Net sales (A$)", color: CAT5[2] }];
  let y = ax.y1 - 6 * PT - 2 * 14 * PT;
  const lw = Math.max(...lg.map((l) => f.width(l.label, 8.5))) + 24 * PT;
  lg.forEach((l, i) => {
    f.rect(ax.x1 - lw, y + i * 14 * PT, 12 * PT, 9 * PT, l.color);
    f.text(l.label, ax.x1 - lw + 17 * PT, y + i * 14 * PT + 4.5 * PT, { pt: 8.5, base: "middle" });
  });
  return f.png();
}

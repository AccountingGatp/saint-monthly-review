// Readers for the Xero exports: Profit and Loss and Balance Sheet (Excel).
import * as config from "./config.js";
import { InputError, monthFromWord, monthKeyFromLabel, mk, keyY, keyM, num, round2, daysIn, MONTH_NAMES } from "./util.js";
import { openBook } from "./xl.js";

export { InputError };
const isNum = (v) => typeof v === "number";
const str = (v) => (v === null || v === undefined ? "" : String(v));

export class PL {
  constructor(periodEnd, keys, lines, ytdLabel, hasYtd, xeroTotals) {
    Object.assign(this, { periodEnd, keys, lines, ytdLabel, hasYtd, xeroTotals });
  }
  acct(name, key) {
    const l = this.lines.find((x) => x.name === name);
    return l ? (l.vals[key] ?? 0) : 0;
  }
  sect(section, key) {
    return round2(this.lines.filter((l) => l.section === section).reduce((s, l) => s + (l.vals[key] ?? 0), 0));
  }
  totals(key) {
    const inc = this.sect("income", key), cogs = this.sect("cogs", key), opex = this.sect("opex", key);
    return { income: inc, cogs, gp: round2(inc - cogs), opex, net: round2(inc - cogs - opex) };
  }
}

export async function readPL(buf, label = "P&L") {
  const ws = (await openBook(buf)).first();
  if (!str(ws.v(1, 1)).toLowerCase().includes("profit")) throw new InputError(`${label}: first sheet is not a Profit and Loss report`);
  const title = [2, 3].map((r) => str(ws.v(r, 1))).join(" ");
  const m = /ended\s+(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/.exec(title);
  if (!m) throw new InputError("P&L header must read 'For the month ended <date>' - export the exact month");
  const py = +m[3], pm = monthFromWord(m[2]), pd = +m[1];
  if (pd !== daysIn(py, pm)) throw new InputError(`P&L ends ${pd}/${pm}/${py}, which is not a month end`);
  const periodEnd = { y: py, m: pm, d: pd };

  let hdr = null;
  for (let r = 1; r < 15; r++) if (str(ws.v(r, 1)).trim().toLowerCase() === "account") { hdr = r; break; }
  if (hdr === null) throw new InputError("P&L: could not find the 'Account' header row");
  const cols = {}, keys = [];
  let ytdCol = null, ytdLabel = "";
  for (let c = 2; c <= ws.maxCol; c++) {
    const lab = ws.v(hdr, c);
    if (lab === null) continue;
    const k = monthKeyFromLabel(lab);
    if (k) { cols[c] = k; keys.push(k); }
    else if (str(lab).toLowerCase().includes("year to date")) { ytdCol = c; ytdLabel = str(lab); }
  }
  if (keys[0] !== mk(py, pm)) throw new InputError("P&L: first column is not the report month");

  const lines = [], xeroTotals = {};
  let section = null;
  const colList = Object.keys(cols).map(Number).concat(ytdCol ? [ytdCol] : []);
  for (let r = hdr + 1; r <= ws.maxRow; r++) {
    let name = ws.v(r, 1);
    if (name === null) continue;
    name = str(name).trim();
    const low = name.toLowerCase();
    const raw = {};
    for (const c of colList) raw[c] = ws.v(r, c, { formulas: true });
    const hasVals = colList.some((c) => raw[c] !== null);
    if (!hasVals) {
      if (!(low in config.PL_SECTIONS)) throw new InputError(`P&L: unknown section heading '${name}' (row ${r}). Decide where it belongs, then add it to PL_SECTIONS in js/config.js.`);
      section = config.PL_SECTIONS[low];
      continue;
    }
    if (low.startsWith("total ") || config.PL_DERIVED_ROWS.includes(low)) {
      for (const c of colList) if (isNum(raw[c])) (xeroTotals[low] ??= {})[cols[c] ?? "YTD"] = raw[c];
      continue;
    }
    if (section === null) throw new InputError(`P&L: account '${name}' appears before any section heading`);
    const vals = {};
    for (const c of Object.keys(cols)) vals[cols[c]] = num(raw[c]);
    if (ytdCol) vals.YTD = num(raw[ytdCol]);
    lines.push({ name, section, vals });
  }
  const names = lines.map((l) => l.name);
  const dup = [...new Set(names.filter((n, i) => names.indexOf(n) !== i))];
  if (dup.length) throw new InputError(`P&L: duplicate account names ${JSON.stringify(dup)} - names must be unique`);
  return new PL(periodEnd, keys, lines, ytdLabel, !!ytdCol, xeroTotals);
}

export async function readBS(buf) {
  const ws = (await openBook(buf)).first();
  if (!str(ws.v(1, 1)).toLowerCase().includes("balance")) throw new InputError("Balance Sheet: first sheet is not a Balance Sheet report");
  let hdr = null;
  for (let r = 1; r < 15; r++) if (str(ws.v(r, 2)).trim().toLowerCase() === "account") { hdr = r; break; }
  if (hdr === null) throw new InputError("Balance Sheet: could not find the 'Account' header row");
  const cols = {};
  for (let c = 3; c <= ws.maxCol; c++) {
    const lab = ws.v(hdr, c);
    if (lab === null) continue;
    if (lab instanceof Date) { cols[c] = mk(lab.getUTCFullYear(), lab.getUTCMonth() + 1); continue; }
    const m = /^\s*(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})/.exec(str(lab));
    if (m) cols[c] = mk(+m[3], monthFromWord(m[2]));
  }
  const keys = Object.values(cols);
  if (keys.length < 2) throw new InputError("Balance Sheet: need at least two date columns (SOP: two comparison columns)");

  const lines = [];
  let group = null;
  for (let r = hdr + 1; r <= ws.maxRow; r++) {
    const a = ws.v(r, 1), b = ws.v(r, 2);
    const name = str(b !== null ? b : a).trim();
    if (!name) continue;
    if (name.startsWith("[FX]")) break;
    const low = name.toLowerCase();
    const raw = {};
    for (const c of Object.keys(cols)) raw[cols[c]] = ws.v(r, +c, { formulas: true });
    const hasVals = Object.values(raw).some((v) => v !== null);
    if (low.startsWith("total ") || low === "net assets") continue;
    if (!hasVals) {
      if (!(low in config.BS_SECTIONS)) throw new InputError(`Balance Sheet: unknown heading '${name}' (row ${r}) - add to BS_SECTIONS in js/config.js`);
      const g = config.BS_SECTIONS[low];
      if (g) group = g;
      continue;
    }
    if (group === null) throw new InputError(`Balance Sheet: account '${name}' before any section`);
    if (Object.values(raw).some((v) => typeof v === "string" && v.startsWith("="))) throw new InputError(`Balance Sheet: '${name}' holds a formula, expected a Xero value`);
    const vals = {};
    for (const k of Object.keys(raw)) vals[k] = num(raw[k]);
    lines.push({ name, group, vals });
  }
  return { keys, lines };
}

// House presentation of the balance sheet. Returns {sec: {group: {name: {key: value}}}, tot: {key: {...}}}.
export function bsPresent(bs) {
  const keys = bs.keys;
  const by = new Map();
  for (const l of bs.lines) { const id = l.group + "\u0000" + l.name; if (!by.has(id)) by.set(id, l); }
  const v = (g, n, k) => { const l = by.get(g + "\u0000" + n); return l ? (l.vals[k] ?? 0) : 0; };
  const groups = ["Bank", "Current Assets", "Fixed Assets", "Current Liabilities", "Non-Current Liabilities", "Equity"];
  const sec = Object.fromEntries(groups.map((g) => [g, {}]));
  const zero = () => Object.fromEntries(keys.map((k) => [k, 0]));
  for (const l of bs.lines) {
    if (l.group === "Bank" && config.NET_ASSET_AGAINST_LIABILITY.includes(l.name)) continue;
    if (l.group === "Current Liabilities" && config.LIABILITY_SHOWN_AS_BANK.includes(l.name)) {
      sec.Bank[l.name] = Object.fromEntries(Object.entries(l.vals).map(([k, x]) => [k, -x]));
      continue;
    }
    const row = (sec[l.group][l.name] ??= zero());
    for (const [k, x] of Object.entries(l.vals)) row[k] = (row[k] ?? 0) + x;
  }
  for (const name of config.NET_ASSET_AGAINST_LIABILITY) {
    const row = (sec["Current Liabilities"][name] ??= zero());
    for (const k of keys) row[k] = v("Current Liabilities", name, k) - v("Bank", name, k);
  }
  for (const g of groups) {
    for (const n of Object.keys(sec[g])) {
      if (!keys.some((k) => Math.abs(sec[g][n][k] ?? 0) >= 0.005)) delete sec[g][n];
    }
  }
  const tot = {};
  for (const k of keys) {
    const t = {};
    for (const g of groups) t[g] = round2(Object.values(sec[g]).reduce((s, r) => s + (r[k] ?? 0), 0));
    t["Total Assets"] = round2(t.Bank + t["Current Assets"] + t["Fixed Assets"]);
    t["Net Assets"] = round2(t["Total Assets"] - t["Current Liabilities"] - t["Non-Current Liabilities"]);
    t["Total Equity"] = t.Equity;
    tot[k] = t;
  }
  return { sec, tot };
}

// Optional calendar-YTD P&L (one amount column).
export async function readPLPeriod(buf) {
  const ws = (await openBook(buf)).first();
  if (!str(ws.v(1, 1)).toLowerCase().includes("profit")) throw new InputError("YTD file: not a Profit and Loss report");
  const title = [2, 3].map((r) => str(ws.v(r, 1))).join(" ");
  let hdr = null;
  for (let r = 1; r < 15; r++) if (str(ws.v(r, 1)).trim().toLowerCase() === "account") { hdr = r; break; }
  if (hdr === null) throw new InputError("YTD P&L: no 'Account' header row");
  const sums = { income: 0, cogs: 0, opex: 0 };
  let section = null;
  for (let r = hdr + 1; r <= ws.maxRow; r++) {
    let name = ws.v(r, 1);
    if (name === null) continue;
    name = str(name).trim();
    const low = name.toLowerCase();
    const v = ws.v(r, 2, { formulas: true });
    if (v === null) {
      if (!(low in config.PL_SECTIONS)) throw new InputError(`YTD P&L: unknown section heading '${name}'`);
      section = config.PL_SECTIONS[low]; continue;
    }
    if (low.startsWith("total ") || config.PL_DERIVED_ROWS.includes(low)) continue;
    sums[section] += num(v);
  }
  const t = { income: round2(sums.income), cogs: round2(sums.cogs), opex: round2(sums.opex) };
  t.gp = round2(t.income - t.cogs);
  t.net = round2(t.gp - t.opex);
  return [t, title.trim()];
}

// SOP step 2: the tie-out gate, plus the comparisons a reviewer needs.
import * as config from "./config.js";
import { round2, sum, mk, keyY, keyM, prevKey, monthLabel, MONTH_NAMES } from "./util.js";

const FY_START_MONTH = 7;     // Australian financial year; Xero 'Year to date' starts 1 July
export { prevKey };

export function chk(group, name, expected, actual, { tol = config.TOL_EXACT, gate = false, note = "" } = {}) {
  const diff = expected === null || actual === null || expected === undefined || actual === undefined ? null : round2(actual - expected);
  let status;
  if (diff === null) status = "INFO";
  else if (Math.abs(diff) <= tol) status = "PASS";
  else status = gate ? "FAIL" : "WARN";
  return { group, name, expected: expected ?? null, actual: actual ?? null, diff, status, gate, note };
}

export function fyMonths(k) {
  const y = keyY(k), m = keyM(k);
  let cy = m >= FY_START_MONTH ? y : y - 1, cm = FY_START_MONTH;
  const out = [];
  while (mk(cy, cm) <= k) { out.push(mk(cy, cm)); cm += 1; if (cm > 12) { cm = 1; cy += 1; } }
  return out;
}

export const plLabelTotals = (t) => ({
  "Total Income": t.income, "Total Cost of Goods Sold": t.cogs, "Gross Profit": t.gp,
  "Total Operating Expenses": t.opex, "Net Income": t.net,
});

export function runChecks(pl, bs, drs, prior, k0, sec, tot) {
  const C = [];
  const k1 = prevKey(k0, 1), k2 = prevKey(k0, 2);
  const need = [k0, k1, k2];
  for (const k of need) {
    C.push(chk("Inputs", `P&L has column ${monthLabel(k)}`, 1, pl.keys.includes(k) ? 1 : 0, { gate: true }));
    C.push(chk("Inputs", `Balance Sheet has column ${monthLabel(k)}`, 1, bs.keys.includes(k) ? 1 : 0, { gate: true }));
  }
  for (const [lab, key] of [["total trading income", "income"], ["total income", "income"],
    ["total cost of sales", "cogs"], ["total cost of goods sold", "cogs"],
    ["total operating expenses", "opex"], ["gross profit", "gp"], ["net profit", "net"]]) {
    for (const [k, v] of Object.entries(pl.xeroTotals[lab] || {})) {
      C.push(chk("Xero totals", `${lab} ${k}`, v, pl.totals(k)[key], { gate: true }));
    }
  }
  if (pl.hasYtd) {
    const fm = fyMonths(k0);
    if (fm.every((m) => pl.keys.includes(m))) {
      for (const [sname, s] of [["income", "income"], ["cost of sales", "cogs"], ["operating expenses", "opex"]]) {
        C.push(chk("YTD", `Xero 'Year to date' ${sname} = sum of ${fm.length} months`,
          round2(sum(fm.map((m) => pl.sect(s, m)))), pl.sect(s, "YTD"), { gate: true }));
      }
      const bad = [];
      for (const l of pl.lines) {
        const d = round2(sum(fm.map((m) => l.vals[m] ?? 0)) - (l.vals.YTD ?? 0));
        if (Math.abs(d) > config.TOL_EXACT) bad.push([l.name, d]);
      }
      C.push(chk("YTD", "Every account: months add to 'Year to date'", 0, bad.length, {
        gate: true, note: bad.slice(0, 6).map(([n, d]) => `${n} ${d >= 0 ? "+" : "-"}${Math.abs(d).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`).join("; ") }));
    } else {
      C.push(chk("YTD", "Year to date cannot be re-added (months missing)", null, null));
    }
  }
  for (const k of need) {
    if (tot[k]) C.push(chk("Balance sheet", `Net Assets = Total Equity, ${monthLabel(k)}`, tot[k]["Total Equity"], tot[k]["Net Assets"], { gate: true }));
  }
  const cye = bs.lines.find((l) => l.name.toLowerCase() === "current year earnings");
  if (cye) {
    for (const k of need) {
      const fm = fyMonths(k);
      if (k in cye.vals && fm.every((m) => pl.keys.includes(m))) {
        C.push(chk("BS vs P&L", `Current Year Earnings ${monthLabel(k)} = P&L net since 1 July`,
          round2(sum(fm.map((m) => pl.totals(m).net))), cye.vals[k], { gate: true }));
      }
    }
  }

  // ---- DRS vs Xero (information)
  const d = drs.daily;
  const tI = pl.totals(k0);
  const drsChk = (name, drsVal, xeroVal, note = "") => {
    const tol = Math.max(config.TOL_DRS_ABS, Math.abs(xeroVal) * config.TOL_DRS_PCT);
    C.push(chk("DRS vs P&L", name, xeroVal, drsVal, { tol, note }));
  };
  drsChk("Countrywise total vs P&L total income", drs.country_total, tI.income);
  drsChk("Shipping income AU", sum(d.ship_au), pl.acct("Shipping Income - AU", k0));
  drsChk("Shipping income other", sum(d.ship_other), pl.acct("Shipping Income - Other", k0));
  drsChk("COGS (product) daily sum vs P&L Cost of Goods Sold", sum(d.cogs_product), pl.acct("Cost of Goods Sold", k0), "Daily COGS is on the supplier cost sheet");
  drsChk("Shipping cost daily sum vs P&L", sum(d.cogs_ship), pl.acct("Shipping Cost", k0));
  for (const [lab, acct, div] of config.DRS_AD_MAP) {
    const key = Object.keys(d).find((x) => x.startsWith(lab));
    if (key) drsChk(`Ad spend ${acct.split(" - ").pop()} (DRS${div !== 1 ? " ÷1.1" : ""})`, sum(d[key]) / div, pl.acct(acct, k0));
  }
  for (const [gw, acct] of Object.entries(config.FEE_ACCOUNTS)) {
    if (gw in drs.fees) drsChk(`Merchant fees ${gw} (DRS ex GST)`, drs.fees[gw].fee_ex_gst, -pl.acct(acct, k0));
  }
  const gross = sum(d.gross);
  const cov = gross ? drs.product_total[0] / gross : 0;
  C.push(chk("DRS vs P&L", "Productwise Sales gross covers DRS gross sales", gross, drs.product_total[0],
    { tol: gross * (1 - config.PRODUCT_COVERAGE_MIN), note: `coverage ${(cov * 100).toFixed(1)}%` }));
  const mm = `${String(keyM(k0)).padStart(2, "0")}/01/${keyY(k0)}`;
  for (const [nm, txt] of [["Countrywise Sales", drs.country_period], ["Productwise Sales", drs.product_period]]) {
    C.push(chk("DRS vs P&L", `${nm} tab is for ${MONTH_NAMES[keyM(k0)]} ${keyY(k0)}`, 1, txt.startsWith(mm) ? 1 : 0, { note: txt }));
  }
  // rows on the Productwise tab carry the export date; last year's rows mean the tab was never refreshed
  const pm = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const stale = (drs.product_dates || []).filter((t) => { const m = /^\d{1,2}-([A-Za-z]{3})-(\d{4})$/.exec(t); return m && !(pm[m[1].toLowerCase()] === keyM(k0) && +m[2] === keyY(k0)); });
  if ((drs.product_dates || []).length) {
    C.push(chk("DRS vs P&L", `Productwise Sales rows are dated ${MONTH_NAMES[keyM(k0)]} ${keyY(k0)}`, 1, stale.length ? 0 : 1, { note: stale.length ? `rows dated ${[...new Set(stale)].join(", ")}` : "" }));
  }
  const info = { product_coverage: cov, product_ok: cov >= config.PRODUCT_COVERAGE_MIN, product_stale: [...new Set(stale)] };

  // ---- comparison with what was published last month
  const restate = { totals: [], lines: [], bs: [] };
  const pp = prior.pl;
  const zeroT = { income: 0, cogs: 0, gp: 0, opex: 0, net: 0 };
  for (const k of prior.pl_keys) {
    if (!pl.keys.includes(k)) continue;
    const t = pl.totals(k);
    for (const [lab, v] of Object.entries(plLabelTotals(t))) {
      if (lab in pp && k in pp[lab]) restate.totals.push({ month: k, item: lab, published: pp[lab][k], now: v, diff: round2(v - pp[lab][k]) });
    }
    for (const l of pl.lines) {
      const pub = (pp[l.name] || {})[k] ?? 0, now = l.vals[k] ?? 0;
      if (Math.abs(now - pub) >= config.RESTATE_LINE_MIN) restate.lines.push({ month: k, item: l.name, published: pub, now, diff: round2(now - pub) });
    }
    const known = new Set(pl.lines.map((l) => l.name));
    const skip = new Set(["Income", "Cost of Goods Sold", "Operating Expenses", "Operating Income", "Name", "Total Cost of Goods Sold"]);
    for (const [lab, v] of Object.entries(pp)) {
      if (!known.has(lab) && !(lab in plLabelTotals(zeroT)) && !skip.has(lab) && Math.abs(v[k] ?? 0) >= 1) {
        restate.lines.push({ month: k, item: lab + " (not in new P&L)", published: v[k] ?? 0, now: 0, diff: round2(-(v[k] ?? 0)) });
      }
    }
  }
  const pb = prior.bs;
  const keyed = { "Total Bank": "Bank", "Total Current Assets": "Current Assets", "Total Assets": "Total Assets",
    "Total Current Liabilities": "Current Liabilities", "Total Non-Current Liabilities": "Non-Current Liabilities",
    "Net Assets": "Net Assets", "Total Equity": "Total Equity" };
  for (const k of prior.bs_keys) {
    if (!tot[k]) continue;
    for (const [lab, key] of Object.entries(keyed)) {
      if (lab in pb && k in pb[lab]) {
        const now = tot[k][key];
        restate.bs.push({ month: k, item: lab, published: pb[lab][k], now, diff: Math.round(now - pb[lab][k]) });
      }
    }
    for (const g of Object.keys(sec)) {
      for (const [n, r] of Object.entries(sec[g])) {
        if (n in pb && k in pb[n] && k in r && Math.abs(Math.round(r[k]) - pb[n][k]) >= 2) {
          restate.bs.push({ month: k, item: n, published: pb[n][k], now: Math.round(r[k]), diff: Math.round(r[k]) - pb[n][k] });
        }
      }
    }
  }

  // ---- new accounts
  const k1k = k1;
  const newPl = [], newBs = [];
  for (const l of pl.lines) {
    if (Math.abs(l.vals[k0] ?? 0) >= 0.005 && Math.abs(l.vals[k1k] ?? 0) < 0.005) {
      const hist = pl.keys.length > 1 ? Math.max(...pl.keys.slice(1).map((m) => Math.abs(l.vals[m] ?? 0))) : 0;
      newPl.push({ account: l.name, section: l.section, this: l.vals[k0] ?? 0, first_time: hist < 0.005 });
    }
  }
  for (const g of Object.keys(sec)) {
    for (const [n, r] of Object.entries(sec[g])) {
      if (Math.abs(r[k0] ?? 0) >= 0.005 && Math.abs(r[k1k] ?? 0) < 0.005) newBs.push({ account: n, group: g, this: r[k0] ?? 0 });
    }
  }
  const structural = new Set(["Assets", "Bank", "Current Assets", "Fixed Assets", "Current Liabilities", "Non-Current Liabilities", "Equity"]);
  const priorBsNames = new Set(Object.keys(pb).filter((n) => !(n in keyed) && !structural.has(n)));
  const nowBsNames = new Set(Object.keys(sec).flatMap((g) => Object.keys(sec[g])));
  const appeared = [...nowBsNames].filter((n) => !priorBsNames.has(n)).sort();
  const plNames = new Set(pl.lines.map((l) => l.name.toLowerCase()));
  const orphan = [];
  for (const n of nowBsNames) {
    if (n.startsWith("Accrued Social Media Ads - ")) {
      const tail = n.split(" - ").slice(1).join(" - ");
      if (!plNames.has(`social media ads - ${tail}`.toLowerCase())) orphan.push(n);
    }
  }
  return {
    checks: C, restate, new_pl: newPl, new_bs: newBs, bs_not_on_prior_report: appeared,
    orphan_accruals: orphan.sort(), info, gate_ok: !C.some((c) => c.status === "FAIL"),
  };
}

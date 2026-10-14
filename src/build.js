// Orchestrates one month-end run: read inputs -> tie-out gate -> edit last month's report -> verify -> outputs.
import * as config from "./config.js";
import { readPL, readBS, bsPresent, readPLPeriod, InputError } from "./xero.js";
import { readDRS } from "./drs.js";
import { readPrior, findTables, loadDocXml } from "./prior.js";
import { runChecks, fyMonths, prevKey } from "./checks.js";
import * as charts from "./charts.js";
import * as DB from "./docxbuild.js";
import * as reports from "./reports.js";
import { openBook } from "./xl.js";
import { MONTH_NAMES, MONTH_ABBR, mk, keyM, keyY, longLabel, monthLabel, daysIn, sum, round2 } from "./util.js";

export { InputError };
const KINDS = ["ytd", "fees", "expenditure", "daysales", "daycogs", "cogspie", "countrypie", "product_table", "product_chart"];

// Classify an uploaded Excel/Word file by its content: "pl" | "bs" | "drs" | "report" | null
export async function sniff(name, bytes) {
  const low = name.toLowerCase();
  if (low.endsWith(".docx")) {
    try { findTables(await loadDocXml(bytes)); return "report"; } catch (e) { return null; }   // only a real Month-End report counts
  }
  if (!low.endsWith(".xlsx") || low.startsWith("~$")) return null;
  try {
    const wb = await openBook(bytes);
    if (wb.names.includes("Daily Reporting")) return "drs";
    const a1 = String(wb.first().v(1, 1) || "").toLowerCase();
    if (a1.startsWith("profit")) return "pl";
    if (a1.startsWith("balance")) return "bs";
  } catch (e) { /* not a readable workbook */ }
  return null;
}

const FMT_D = (d) => `${d.d} ${MONTH_ABBR[d.m]} ${d.y}`;
// One-line description of an uploaded file, shown on the Inputs step.
export async function describeInput(role, bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (role === "pl") {
    const p = await readPL(u8);
    return `Month ended ${FMT_D(p.periodEnd)}. Columns: ${p.keys.map(monthLabel).join(", ")}${p.hasYtd ? ", year to date" : ""}.`;
  }
  if (role === "bs") {
    const b = await readBS(u8);
    return `As at ${b.keys.map(monthLabel).join(", ")}.`;
  }
  if (role === "drs") {
    const wb = await openBook(u8);
    const ws = wb.sheet("Daily Reporting");
    const need = ["Daily Reporting", "Countrywise Sales", "Productwise Sales", "Fees Calculation"].filter((n) => !wb.names.includes(n));
    if (need.length) return `Missing tabs: ${need.join(", ")}.`;
    let dates = [];
    for (let r = 1; r < 40 && !dates.length; r++) {
      if (String(ws.v(r, 1) || "").toLowerCase().includes("date")) for (let c = 2; c <= ws.maxCol; c++) { const v = ws.v(r, c); if (v instanceof Date) dates.push(v); }
    }
    return dates.length ? `${dates.length} days: ${FMT_D({ d: dates[0].getUTCDate(), m: dates[0].getUTCMonth() + 1, y: dates[0].getUTCFullYear() })} to ${FMT_D({ d: dates[dates.length - 1].getUTCDate(), m: dates[dates.length - 1].getUTCMonth() + 1, y: dates[dates.length - 1].getUTCFullYear() })}.` : "Tabs found.";
  }
  if (role === "report") {
    const pr = await readPrior(u8);
    return `Covers ${pr.pl_keys.map(monthLabel).join(", ")}. Used as the template.`;
  }
  return "";
}

// The report month of a P&L file, as "YYYY-MM" (used to pick last month's report).
export async function plMonth(bytes) {
  const p = await readPL(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  return mk(p.periodEnd.y, p.periodEnd.m);
}

const brief = (cs) => cs.map((c) => ({ name: c.name, expected: c.expected, actual: c.actual, diff: c.diff, note: c.note }));
const MIME = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  txt: "text/plain",
};

// inputs: {pl, bs, drs, report} as Uint8Array|ArrayBuffer ; opts: {links:{drs,deep,quick}, ytdPl, allowPartialProducts, skipMonthGuard}
export async function buildReport(inputs, opts = {}) {
  const log = [];
  const say = (s) => log.push(s);
  const u8 = (x) => (x instanceof Uint8Array ? x : new Uint8Array(x));
  const plR = await readPL(u8(inputs.pl));
  const k0 = mk(plR.periodEnd.y, plR.periodEnd.m), k1 = prevKey(k0, 1), k2 = prevKey(k0, 2);
  const monthName = longLabel(k0), prevName = longLabel(k1);
  const bsR = await readBS(u8(inputs.bs));
  const { sec, tot } = bsPresent(bsR);
  const d = await readDRS(u8(inputs.drs), plR.periodEnd);
  const prv = await readPrior(u8(inputs.report));
  const priorLast = prv.pl_keys[prv.pl_keys.length - 1];
  if (priorLast !== k1 && !opts.skipMonthGuard) {
    throw new InputError(`Last month's report ends ${longLabel(priorLast)} but this P&L is for ${monthName}; expected the ${prevName} report.`);
  }

  const res = runChecks(plR, bsR, d, prv, k0, sec, tot);
  const fails = res.checks.filter((c) => c.status === "FAIL"), warns = res.checks.filter((c) => c.status === "WARN");
  say(`Tie-out gate (inputs): ${res.checks.length - fails.length - warns.length} pass, ${warns.length} warn, ${fails.length} FAIL`);
  const cov = res.info.product_coverage;
  const gate = {
    month: monthName, prev_month: prevName, checks: res.checks, fails: brief(fails), warns: brief(warns), info: res.info,
    new_pl: res.new_pl, new_bs: res.new_bs, bs_not_on_prior_report: res.bs_not_on_prior_report, orphan_accruals: res.orphan_accruals,
    product: { ok: res.info.product_ok, coverage: cov, stale: res.info.product_stale, product_gross: d.product_total[0], drs_gross: sum(d.daily.gross) },
    periods: { pl: plR.keys, bs: bsR.keys, prior_last: priorLast },
    totals: { [k0]: plR.totals(k0), [k1]: plR.totals(k1) },
  };
  if (fails.length) {
    return { status: "gate_failed", summary: { status: "gate_failed", month: monthName, fails: brief(fails), warns: brief(warns) }, gate, files: [], log };
  }
  if (opts.gateOnly) return { status: "gate_ok", summary: { status: "gate_ok", month: monthName, fails: [], warns: brief(warns) }, gate, files: [], log };

  const t0 = plR.totals(k0), t1 = plR.totals(k1);
  const fm = fyMonths(k0);
  let tytd, ytdHeader, ytdLabel, ytdHeading, ytdNote;
  const ab = MONTH_ABBR;
  if (inputs.ytd) {
    const [tt, ytxt] = await readPLPeriod(u8(inputs.ytd));
    tytd = tt;
    ytdHeader = `YTD Jan-${ab[keyM(k0)]}`;
    ytdLabel = `Year-to-Date ${keyY(k0)} (Jan - ${ab[keyM(k0)]})`;
    ytdHeading = `Year-to-Date Performance (Jan - ${ab[keyM(k0)]} ${keyY(k0)}):`;
    ytdNote = `Calendar YTD taken from the supplied Jan-${ab[keyM(k0)]} Xero P&L (${ytxt}).`;
  } else {
    tytd = plR.hasYtd ? plR.totals("YTD") : Object.fromEntries(Object.keys(t0).map((k) => [k, round2(sum(fm.map((m) => plR.totals(m)[k])))]));
    const s = ab[keyM(fm[0])], e = ab[keyM(fm[fm.length - 1])];
    ytdHeader = `FYTD ${s}-${e}`;
    ytdLabel = `Financial year to date (${s} - ${e} ${keyY(k0)})`;
    ytdHeading = `Financial Year-to-Date Performance (${s} - ${e} ${keyY(k0)}):`;
    ytdNote = `Previous packs showed calendar 'YTD Jan-...'. The files supplied only support Xero's financial-year-to-date (${s}-${e}). Shown as 'FYTD' and labelled so. ` +
      `For calendar YTD add a Xero P&L for 1 Jan - ${daysIn(keyY(k0), keyM(k0))} ${MONTH_NAMES[keyM(k0)]} ${keyY(k0)} under "Optional".`;
  }

  const links = opts.links || {};
  const ctx = { k0, k1, k_list: [k2, k1, k0], t0, t1, tytd, ytd_header: ytdHeader, ytd_heading: ytdHeading,
    countries: d.countries, country_total: d.country_total, pl: plR, sec, tot, picture_kinds: KINDS,
    links: { drs: links.drs || null, deep: links.deep || null, quick: links.quick || null } };
  const D = d.daily;
  let costParts = plR.lines.filter((l) => l.section === "cogs" && (l.vals[k0] ?? 0) > 0.005)
    .map((l) => [l.name.split("COGS_").join("").replace("Cost of Goods Sold", "Cost of goods sold"), l.vals[k0]]);
  costParts = costParts.map(([n, v]) => [n.charAt(0).toUpperCase() + n.slice(1), v]);
  const runItems = plR.lines.filter((l) => l.section === "opex" && Math.abs(l.vals[k0] ?? 0) >= 0.005 &&
    !config.EXPENDITURE_EXCLUDE_PREFIXES.some((p) => l.name.toLowerCase().startsWith(p))).map((l) => [l.name, l.vals[k0]]);
  const adsTotal = sum(plR.lines.filter((l) => l.name.toLowerCase().startsWith("social media ads")).map((l) => l.vals[k0] ?? 0));
  const feeByGw = {};
  for (const [gw, acct] of Object.entries(config.FEE_ACCOUNTS)) if (Math.abs(plR.acct(acct, k0)) > 0) feeByGw[gw] = -plR.acct(acct, k0);
  const showProducts = res.info.product_ok || !!opts.allowPartialProducts;
  const prov = res.info.product_ok ? "" : " (PROVISIONAL)";
  const covTxt = `${Math.round(res.info.product_coverage * 100)}%`;
  const stale = res.info.product_stale || [];
  const openNote = stale.length
    ? `Open item: product-level sales for ${monthName} are not shown. The product sales data supplied is out of date (dated ${stale.join(", ")}) and does not cover the month. The full table will follow once the current month's export is received.`
    : `Open item: product-level sales for ${monthName} are not shown. The product export supplied covers ` +
      `A$${Math.round(d.product_total[0]).toLocaleString("en-US")} of A$${Math.round(sum(D.gross)).toLocaleString("en-US")} gross sales for the month ` +
      `(${covTxt}). The full table will follow once the complete export is received.`;
  const phText = stale.length ? "Current-month product export outstanding" : `Full-month product export outstanding (${covTxt} of gross sales supplied)`;
  const images = {
    ytd: (w, h) => charts.ytd(w, h, ytdLabel, { Revenue: tytd.income, COGS: tytd.cogs, "Operating Expenses": tytd.opex, "Net Income": tytd.net }),
    fees: (w, h) => charts.fees(w, h, monthName, feeByGw),
    expenditure: (w, h) => charts.expenditure(w, h, runItems, adsTotal),
    daysales: (w, h) => charts.daywiseSales(w, h, d.dates, D.orders, D.returns, D.ship_collected),
    daycogs: (w, h) => charts.daywiseCogs(w, h, d.dates, D.cogs_product, D.cogs_ship),
    cogspie: (w, h) => charts.costPie(w, h, costParts),
    countrypie: (w, h) => charts.countryPie(w, h, d.countries, monthName),
  };
  if (showProducts) {
    images.product_table = async (w) => {
      const r = await charts.productTable(w, d.products, d.product_total);
      if (prov) r.note = openNote.replace("Open item: product-level sales", "Provisional: product-level sales").replace("are not shown", "are shown from a partial export");
      return r;
    };
    images.product_chart = (w, h) => charts.productChart(w, h, d.products);
  } else {
    images.product_table = async (w) => {
      const r = await charts.placeholder(w, 0.9, "Product-wise sales: open item", phText);
      r.note = openNote; return r;
    };
    images.product_chart = (w) => charts.placeholder(w, 0.55, "Gross, discounts, refunds and net sales by product", "To follow with the product table");
  }

  const { bytes: docx, linkLog } = await DB.build(u8(inputs.report), ctx, images);
  const notes = ["No PDF is made in the browser. Word will ask to update fields when you open the file - click Yes so the contents page numbers refresh - then read every page (SOP step 4)."];

  const rb = await reports.readback(docx, plR, sec, tot, [k2, k1, k0], ctx);
  const rbf = rb.filter((c) => c.status === "FAIL");
  say(`Read-back of the finished report vs Xero: ${rb.length - rbf.length} agree, ${rbf.length} FAIL`);

  const extra = { product_gross: d.product_total[0], drs_gross: sum(D.gross), ytd_note: ytdNote, links: linkLog, drs_processing: sum(D.processing), product_stale: stale };
  const mname = `${MONTH_NAMES[keyM(k0)]}-${keyY(k0)}`;
  const tieout = await reports.writeWorkbook(monthName, plR, tot, [k2, k1, k0], res, rb, ctx);
  const cp = reports.callpoints(monthName, prevName, res, ctx, plR, sec, tot, extra);
  const names = { docx: `Month-End_Financial_Report_${mname}.docx`, tieout: `Tie-out_${mname}.xlsx`, callpoints: `Call_points_${mname}.txt` };
  const files = [
    { name: names.docx, mime: MIME.docx, bytes: docx },
    { name: names.tieout, mime: MIME.xlsx, bytes: tieout },
    { name: names.callpoints, mime: MIME.txt, bytes: new TextEncoder().encode(cp) },
  ];
  const recon = reports.reconTable(rb, ctx);
  const narrative = await reports.narrativeLines(docx);
  const summary = { status: rbf.length ? "readback_failed" : "ok", recon, narrative, month: monthName, n_checks: res.checks.length + rb.length,
    fails: [], warns: brief(warns), readback_ok: rb.length - rbf.length, readback_fail: brief(rbf), notes,
    links_removed: linkLog.map(([a, b]) => `${a} (${b})`), pdf: false, product_ok: res.info.product_ok, files: names };
  // SOP: nothing is released unless the finished report agrees with the Xero P&L
  return { status: summary.status, summary, gate, files: rbf.length ? [] : files, log, ctx };
}

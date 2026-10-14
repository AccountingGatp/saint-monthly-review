// Read-back verification of the finished Word file, the tie-out workbook, and the call points.
import * as config from "./config.js";
import { chk, plLabelTotals, prevKey } from "./checks.js";
import { findTables, tableByMonth, parseAmount, loadDocXml } from "./prior.js";
import { monthLabel, sum, commas } from "./util.js";

function need(obj, lab, k) {
  const v = obj[lab] && obj[lab][k];
  if (v === undefined) throw new Error(`The finished report has no '${lab}' figure for ${k}`);
  return v;
}

export async function readback(docxBytes, pl, sec, tot, kList, ctx) {
  const t = findTables(await loadDocXml(docxBytes));
  const [pld] = tableByMonth(t.pl);
  const [bsd] = tableByMonth(t.bs);
  const C = [];
  for (const k of kList) {
    const tt = pl.totals(k);
    for (const [lab, v] of Object.entries(plLabelTotals(tt))) C.push(chk("Report vs Xero P&L", `${lab} ${monthLabel(k)}`, v, need(pld, lab, k), { gate: true }));
    for (const l of pl.lines) {
      const v = l.vals[k] ?? 0;
      if (Math.abs(v) >= 0.005 || l.name in pld) {
        const got = (pld[l.name] || {})[k] ?? 0;
        if (Math.abs(got - v) > config.TOL_EXACT) C.push(chk("Report vs Xero P&L", `${l.name} ${monthLabel(k)}`, v, got, { gate: true }));
      }
    }
  }
  for (const lab of ["Total Income", "Net Income"]) {
    const row = t.pl.find((r) => r.length && r[0] === lab);
    const want = sum(kList.map((k) => pl.totals(k)[lab === "Total Income" ? "income" : "net"]));
    C.push(chk("Report vs Xero P&L", `${lab} - 3 month total column`, want, parseAmount(row[row.length - 1]), { gate: true }));
  }
  const keyed = { "Total Bank": "Bank", "Total Assets": "Total Assets", "Net Assets": "Net Assets", "Total Equity": "Total Equity",
    "Total Current Liabilities": "Current Liabilities", "Total Non-Current Liabilities": "Non-Current Liabilities" };
  for (const k of kList) {
    for (const [lab, key] of Object.entries(keyed)) C.push(chk("Report vs Xero BS", `${lab} ${monthLabel(k)}`, Math.round(tot[k][key]), need(bsd, lab, k), { tol: 1.01, gate: true }));
    C.push(chk("Report vs Xero BS", `Net Assets = Total Equity in the report, ${monthLabel(k)}`, need(bsd, "Total Equity", k), need(bsd, "Net Assets", k), { tol: 1.01, gate: true }));
  }
  const ex = {};
  for (const r of t.exec) if (r.length) ex[r[0]] = r;
  const { t0, t1, tytd: ty } = ctx;
  for (const [lab, key] of [["Revenue", "income"], ["Cost of Goods Sold", "cogs"], ["Operating Expense", "opex"], ["Net income", "net"]]) {
    const r = ex[lab];
    C.push(chk("Report vs Xero P&L", `Executive summary ${lab} (month)`, t0[key], parseAmount(r[1]), { gate: true }));
    C.push(chk("Report vs Xero P&L", `Executive summary ${lab} (prior month)`, t1[key], parseAmount(r[2]), { gate: true }));
    C.push(chk("Report vs Xero P&L", `Executive summary ${lab} (YTD)`, ty[key], parseAmount(r[4]), { gate: true }));
  }
  const ctab = t.country.find((r) => r.length && r[0] === "Total");
  C.push(chk("Report vs DRS", "Country table total = DRS Countrywise total", ctx.country_total, parseAmount(ctab[1]), { gate: true }));
  return C;
}

const argb = (hex) => "FF" + hex;
const FILL = { PASS: "D9EAD3", WARN: "FFF2CC", FAIL: "F4CCCC", INFO: "EDEDED" };
const solid = (hex) => ({ type: "pattern", pattern: "solid", fgColor: { argb: argb(hex) } });

export async function writeWorkbook(monthName, pl, tot, kList, res, rb, ctx) {
  const wb = new globalThis.ExcelJS.Workbook();
  const NAVY = "1F3A5F";
  const head = (ws, row, vals) => vals.forEach((v, i) => {
    const c = ws.getCell(row, i + 1); c.value = v; c.fill = solid(NAVY); c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.alignment = { wrapText: true, vertical: "middle" };
  });
  const widths = (ws, w) => w.forEach((x, i) => { ws.getColumn(i + 1).width = x; });

  let ws = wb.addWorksheet("Tie-out");
  ws.getCell("A1").value = `Tie-out to Xero Profit and Loss - ${monthName}`; ws.getCell("A1").font = { bold: true, size: 13, color: { argb: argb(NAVY) } };
  ws.getCell("A2").value = "Xero = line items of the Xero export, re-added. Report = figures read back out of the finished Word file.";
  let r = 4;
  const fm = '#,##0.00;[Red](#,##0.00)';
  for (const [key, lab] of [[ctx.k0, "Report month"], [ctx.k1, "Comparison month"], ["YTD", "Year to date"]]) {
    const t = key === ctx.k0 ? ctx.t0 : key === ctx.k1 ? ctx.t1 : ctx.tytd;
    ws.getCell(r, 1).value = `${lab}: ${key !== "YTD" ? monthLabel(key) : ctx.ytd_header}`; ws.getCell(r, 1).font = { bold: true };
    r++;
    head(ws, r, ["Line", "Xero P&L", "Report", "Difference"]); r++;
    for (const [nm, k2] of [["Income", "income"], ["Cost of sales", "cogs"], ["Gross profit", "gp"], ["Operating expenses", "opex"], ["Net result", "net"]]) {
      ws.getCell(r, 1).value = nm; ws.getCell(r, 2).value = t[k2];
      const lookup = { income: "Total Income", cogs: "Total Cost of Goods Sold", gp: "Gross Profit", opex: "Total Operating Expenses", net: "Net Income" }[k2];
      let m;
      if (key !== "YTD") m = rb.find((c) => c.name === `${lookup} ${monthLabel(key)}`);
      else {
        const lab2 = { income: "Revenue", cogs: "Cost of Goods Sold", opex: "Operating Expense", net: "Net income" }[k2];
        m = lab2 ? rb.find((c) => c.name === `Executive summary ${lab2} (YTD)`) : null;
      }
      if (m) { ws.getCell(r, 3).value = m.actual; ws.getCell(r, 4).value = m.diff; ws.getCell(r, 4).fill = solid(FILL[m.status]); }
      else ws.getCell(r, 3).value = "n/a (not shown in report)";
      [2, 3, 4].forEach((c) => { ws.getCell(r, c).numFmt = fm; });
      r++;
    }
    r++;
  }
  widths(ws, [34, 18, 18, 14]);

  ws = wb.addWorksheet("Checks");
  head(ws, 1, ["Group", "Check", "Expected", "Actual", "Difference", "Status", "Note"]);
  [...res.checks, ...rb].forEach((c, i) => {
    const row = i + 2;
    [c.group, c.name, c.expected, c.actual, c.diff, c.status, c.note].forEach((v, j) => { ws.getCell(row, j + 1).value = v; });
    ws.getCell(row, 6).fill = solid(FILL[c.status]);
    [3, 4, 5].forEach((j) => { ws.getCell(row, j).numFmt = '#,##0.00;[Red]-#,##0.00'; });
  });
  widths(ws, [20, 62, 16, 16, 14, 9, 50]); ws.views = [{ state: "frozen", ySplit: 1 }];

  ws = wb.addWorksheet("Restated vs published");
  ws.getCell("A1").value = "Figures in this month's Xero export compared with what the previous report published";
  ws.getCell("A1").font = { bold: true, color: { argb: argb(NAVY) } };
  head(ws, 3, ["Statement", "Month", "Item", "Published", "Now in Xero", "Change"]);
  let rr = 4;
  for (const [blk, lab] of [["totals", "P&L totals"], ["lines", "P&L lines (change >= A$1)"], ["bs", "Balance sheet"]]) {
    for (const x of res.restate[blk]) {
      [lab, monthLabel(x.month), x.item, x.published, x.now, x.diff].forEach((v, j) => { ws.getCell(rr, j + 1).value = v; });
      [4, 5, 6].forEach((j) => { ws.getCell(rr, j).numFmt = '#,##0.00;[Red]-#,##0.00'; });
      rr++;
    }
  }
  widths(ws, [26, 12, 46, 16, 16, 16]); ws.views = [{ state: "frozen", ySplit: 3 }];

  ws = wb.addWorksheet("New accounts");
  head(ws, 1, ["Ledger", "Account", "Group", `${monthName} amount`, "Note"]);
  const add = (...v) => { ws.addRow(v); };
  for (const x of res.new_pl) add("P&L", x.account, x.section, x.this, x.first_time ? "first time in the 4-month window" : "was nil last month");
  for (const x of res.new_bs) add("Balance sheet", x.account, x.group, x.this, "was nil last month");
  for (const n of res.bs_not_on_prior_report) add("Balance sheet", n, "", null, "not on the previous published report");
  for (const n of res.orphan_accruals) add("Balance sheet", n, "", null, "accrual with no matching 'Social Media Ads - ...' P&L account");
  widths(ws, [16, 48, 26, 18, 60]);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

export function callpoints(monthName, prevName, res, ctx, pl, sec, tot, extra) {
  const { k0, k1, t0, t1 } = ctx;
  const ad = (k) => pl.lines.filter((l) => l.name.startsWith("Social Media Ads")).reduce((s, l) => s + (l.vals[k] ?? 0), 0);
  const a0 = ad(k0), a1 = ad(k1);
  const pct = (a, b) => (b ? (a - b) / Math.abs(b) * 100 : 0);
  const A = (x) => (x >= 0 ? `A$${commas(x, 0)}` : `-A$${commas(x, 0)}`);
  const sg = (x) => (x < 0 ? "-" : "") + commas(x, 0);
  const p0 = (x) => `${Math.round(x * 100)}%`;
  const cb = (g, n, k) => ((sec[g][n] || {})[k]) ?? 0;
  const lines = [`CALL POINTS - Saint Eyewear month-end ${monthName}`, "(internal - do not send to the client)", ""];
  lines.push("CLIENT QUESTIONS (Taf)");
  let q = 1;
  lines.push(`${q}. Result. ${monthName} is a net ${t0.net < 0 ? "loss" : "profit"} of ${A(Math.abs(t0.net))} against a net ` +
    `${t1.net < 0 ? "loss" : "profit"} of ${A(Math.abs(t1.net))} in ${prevName}. Revenue fell ${Math.abs(pct(t0.income, t1.income)).toFixed(0)}% ` +
    `(${A(t1.income)} to ${A(t0.income)}) but social media ads fell only ${Math.abs(pct(a0, a1)).toFixed(0)}% ` +
    `(${A(a1)} to ${A(a0)}). Ads were ${p0(a0 / t0.income)} of revenue, up from ${p0(a1 / t1.income)}. ` +
    `What is the plan for ad budget while sales are falling?`); q++;
  const cu0 = cb("Bank", "CommBank USD", k0), cu1 = cb("Bank", "CommBank USD", k1);
  if (Math.abs(cu0 - cu1) > 10000) { lines.push(`${q}. CommBank USD went from ${A(cu1)} to ${A(cu0)} (${A(cu0 - cu1)}). What was moved or paid out of it?`); q++; }
  const d0 = cb("Equity", "Owner A Drawings", k0), d1 = cb("Equity", "Owner A Drawings", k1);
  if (Math.abs(d0 - d1) > 0.5) { lines.push(`${q}. Owner A Drawings ${d0 < d1 ? "increased" : "decreased"} by ${A(Math.abs(d0 - d1))} in the month (balance now ${A(d0)} drawn). Please confirm.`); q++; }
  const newx = res.new_pl.filter((x) => x.section === "opex" && x.first_time && Math.abs(x.this) > 100);
  if (newx.length) { lines.push(`${q}. New expense accounts: ` + newx.map((x) => `${x.account} ${A(x.this)}`).join("; ") + ". What are they, and are they all business costs?"); q++; }
  const tb0 = tot[k0].Bank, tb1 = tot[k1].Bank;
  lines.push(`${q}. Total bank balances fell ${A(tb1 - tb0)} (${A(tb1)} to ${A(tb0)}) after the net loss and drawings. Anything unusual we should know about?`); q++;
  if ("Trust Savings commonwealth" in sec.Bank) {
    lines.push(`${q}. Trust Savings (CommBank) holds ${A(cb("Bank", "Trust Savings commonwealth", k0))}. Please confirm this account, and the Slash and Trust accounts, belong in Xero (open item from SOP v1.0).`); q++;
  }
  lines.push("", "INTERNAL ITEMS");
  let n = 1;
  const rt = new Map(res.restate.totals.map((x) => [x.month + "|" + x.item, x]));
  if (rt.size) {
    const bits = [];
    for (const k of [k1, prevKey(k0, 2)]) {
      const x = rt.get(k + "|Net Income");
      if (x) bits.push(`${monthLabel(k)} net result ${A(x.published)} published, ${A(x.now)} now (${A(x.diff)})`);
    }
    const big = res.restate.lines.filter((x) => x.month === k1).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 4);
    const bigBs = res.restate.bs.filter((y) => y.month === k1 && !/^(Total|Net)/.test(y.item)).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 4);
    lines.push(`${n}. RESTATED COMPARATIVES. The Xero export no longer agrees with the pack that went out last month: ` + bits.join("; ") +
      ". Biggest lines for " + monthLabel(k1) + ": " + big.map((x) => `${x.item} ${A(x.published)} -> ${A(x.now)}`).join("; ") +
      ". Balance sheet comparatives moved too (biggest: " + bigBs.map((x) => `${x.item} ${sg(x.published)} -> ${sg(x.now)}`).join("; ") + "). " +
      "Find out what was posted back into closed months, and whether the client has been told. Detail in the tie-out workbook, sheet 'Restated vs published'."); n++;
  }
  if (!res.info.product_ok) {
    lines.push(`${n}. PRODUCT SALES NOT SHOWN. The DRS 'Productwise Sales' tab totals ${A(extra.product_gross)} gross, ` +
      `${p0(res.info.product_coverage)} of the ${A(extra.drs_gross)} gross sales in the Daily Reporting tab. ` +
      (extra.product_stale && extra.product_stale.length ? `The tab's rows are dated ${extra.product_stale.join(", ")} - it still holds old data and was not refreshed for the month. ` : "") + `Section left as an open item in the report. Need a full-month Shopify product export pasted into that tab.`); n++;
  }
  lines.push(`${n}. YEAR-TO-DATE BASIS. ${extra.ytd_note}`); n++;
  if (extra.links.length) { lines.push(`${n}. LINKS. ` + extra.links.map(([a, b]) => `'${a}' ${b}`).join("; ") + ". Add the new links once the DRS / Deep / Quick View are posted."); n++; }
  if (res.new_pl.length) {
    const fresh = res.new_pl.filter((x) => x.first_time);
    lines.push(`${n}. NEW P&L ACCOUNTS (first time in the window): ` + fresh.map((x) => `${x.account} ${A(x.this)}`).join("; ") + ". Confirm each is P&L, not balance sheet."); n++;
  }
  if (res.orphan_accruals.length) {
    lines.push(`${n}. ACCRUAL WITHOUT P&L ACCOUNT: ` + res.orphan_accruals.join("; ") + " has a balance but there is no matching 'Social Media Ads - ...' P&L account. Where was the cost booked?"); n++;
  }
  const cp = pl.acct("COGS_Processing fee", k0);
  if (Math.abs(cp) > 0.5 && res.new_pl.some((x) => x.account === "COGS_Processing fee")) {
    lines.push(`${n}. 'COGS_Processing fee' is a new Cost of Sales account (${A(cp)}). DRS processing fee for the month is ${A(extra.drs_processing)}. Confirm what it is and why it is in Cost of Sales.`); n++;
  }
  lines.push(`${n}. COMMBANK USD is a manual account: reconcile to the bank statement (none supplied). Balance ${A(cu0)} (was ${A(cu1)}).`); n++;
  const s0 = cb("Current Liabilities", "Suspense", k0), s1 = cb("Current Liabilities", "Suspense", k1);
  lines.push(`${n}. SUSPENSE moved ${A(Math.abs(s0 - s1))} ${s0 < s1 ? "more negative" : "less negative"}, balance ${A(s0)}. Check what went in.`); n++;
  const ap = cb("Current Liabilities", "Accounts Payable", k0);
  if (ap < 0) { lines.push(`${n}. Accounts Payable is a debit balance of ${A(ap)} - check for overpaid or unapplied supplier payments.`); n++; }
  const warns = res.checks.filter((c) => c.status === "WARN" && !c.name.includes("Productwise"));
  if (warns.length) { lines.push(`${n}. DRS vs XERO differences above tolerance: ` + warns.map((c) => `${c.name} ${c.diff >= 0 ? "+" : "-"}${commas(c.diff, 2)}`).join("; ") + "."); n++; }
  lines.push(`${n}. No General Ledger Detail was supplied. It is not needed for the month-end report, only for the expense pack.`); n++;
  return lines.join("\n") + "\n";
}

// SOP step 2: income, cost of sales, operating expenses and net result with a difference column, for both periods.
export function reconTable(rb, ctx) {
  const lookup = { income: "Total Income", cogs: "Total Cost of Goods Sold", gp: "Gross Profit", opex: "Total Operating Expenses", net: "Net Income" };
  const names = { income: "Income", cogs: "Cost of sales", gp: "Gross profit", opex: "Operating expenses", net: "Net result" };
  const out = [];
  for (const [k, t] of [[ctx.k0, ctx.t0], [ctx.k1, ctx.t1]]) {
    const rows = Object.keys(lookup).map((key) => {
      const m = rb.find((c) => c.name === `${lookup[key]} ${monthLabel(k)}`);
      return { line: names[key], xero: t[key], report: m ? m.actual : null, diff: m ? m.diff : null };
    });
    out.push({ title: monthLabel(k), rows });
  }
  const ex = { income: "Revenue", cogs: "Cost of Goods Sold", opex: "Operating Expense", net: "Net income" };
  out.push({ title: ctx.ytd_header, rows: Object.keys(ex).map((key) => {
    const m = rb.find((c) => c.name === `Executive summary ${ex[key]} (YTD)`);
    return { line: names[key], xero: ctx.tytd[key], report: m ? m.actual : null, diff: m ? m.diff : null };
  }) });
  return out;
}

// Free text in the report (outside the tables) that carries a figure or a date: the reader checks the words still match.
export async function narrativeLines(docxBytes) {
  const doc = await loadDocXml(docxBytes);
  const body = doc.documentElement.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "body")[0];
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const out = [];
  for (const p of Array.from(body.childNodes)) {
    if (p.nodeType !== 1 || p.localName !== "p") continue;
    if (p.getElementsByTagNameNS(W, "fldChar").length || p.getElementsByTagNameNS(W, "hyperlink").length) continue;     // contents page
    const t = Array.from(p.getElementsByTagNameNS(W, "t")).map((x) => x.textContent).join("").trim();
    if (t && /\d/.test(t)) out.push(t);
  }
  return out;
}

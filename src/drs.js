// Reader for the Daily Reporting Sheet (DRS) workbook exported from Google Sheets.
import { InputError, num, daysIn, MONTH_NAMES } from "./util.js";
import { openBook } from "./xl.js";

const str = (v) => (v === null || v === undefined ? "" : String(v));
const ymd = (d) => ({ y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() });

function findRow(ws, needle, col = 1, maxr = 80) {
  needle = needle.toLowerCase();
  for (let r = 1; r < maxr; r++) {
    const v = ws.v(r, col);
    if (v && str(v).toLowerCase().includes(needle)) return r;
  }
  throw new InputError(`DRS 'Daily Reporting': no row labelled like '${needle}'`);
}

export async function readDRS(buf, periodEnd) {
  const wb = await openBook(buf);
  for (const n of ["Daily Reporting", "Countrywise Sales", "Productwise Sales", "Fees Calculation"]) {
    if (!wb.names.includes(n)) throw new InputError(`DRS: sheet '${n}' is missing`);
  }
  const ws = wb.sheet("Daily Reporting");
  const drow = findRow(ws, "date");
  const dates = [], cols = [];
  for (let c = 2; c <= ws.maxCol; c++) {
    const v = ws.v(drow, c);
    if (v instanceof Date) { dates.push(ymd(v)); cols.push(c); }
  }
  if (!dates.length) throw new InputError("DRS: no dates on the Daily Reporting sheet");
  const nDays = daysIn(periodEnd.y, periodEnd.m);
  const ok = dates.filter((d) => d.y === periodEnd.y && d.m === periodEnd.m);
  if (ok.length !== nDays || ok.length !== dates.length) {
    throw new InputError(`DRS: Daily Reporting holds ${dates.length} dates, ${ok.length} in ${MONTH_NAMES[periodEnd.m]} ${periodEnd.y}; expected a full month of ${nDays}`);
  }
  const labels = new Map();
  for (let r = 1; r < 60; r++) {
    const lab = ws.v(r, 1);
    if (lab) labels.set(str(lab).trim().toLowerCase(), r);
  }
  const series = (needle) => {
    for (const [lab, r] of labels) if (lab.includes(needle)) return cols.map((c) => num(ws.v(r, c)));
    throw new InputError(`DRS: no row like '${needle}'`);
  };
  const daily = {
    orders: series("number of orders"), gross: series("gross sales"),
    discounts: series("discounts"), returns: series("returns"),
    net_sales: series("net sales"), ship_collected: series("shipping charges collected"),
    prod_rev: series("product revenue (ex gst"), ship_au: series("shipping revenue (ex gst, a$) - au"),
    ship_other: series("shipping revenue - others"), refund_ex_gst: series("refund ex gst"),
    total_rev: series("total revenue (ex gst"),
    cogs_product: series("cogs product cost"), cogs_ship: series("cogs shipping cost"),
    processing: series("processing fee"), merchant_fees: series("merchant fees"),
    ad_total: series("advertisement spend"),
  };
  for (const [lab, r] of labels) if (lab.startsWith("advertisment spend -")) daily[lab] = cols.map((c) => num(ws.v(r, c)));

  const wc = wb.sheet("Countrywise Sales");
  const countryPeriod = str(wc.v(1, 2));
  const countries = [];
  let total = null;
  for (let r = 2; r < 200; r++) {
    const n = wc.v(r, 2), v = wc.v(r, 3);
    if (n && str(n).startsWith("Product Revenue") && typeof v === "number") countries.push([str(n), v]);
    else if (n && str(n).trim().toLowerCase() === "total") { total = num(v); break; }
  }
  if (total === null) throw new InputError("DRS Countrywise Sales: no Total row");

  const wp = wb.sheet("Productwise Sales");
  const productPeriod = str(wp.v(1, 2));
  const products = [], productDates = [];
  let ptotal = null;
  for (let r = 3; r < 200; r++) {
    const dv = wp.v(r, 1);
    if (dv instanceof Date) productDates.push(`${dv.getUTCDate()}-${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][dv.getUTCMonth()]}-${dv.getUTCFullYear()}`);
    else if (typeof dv === "string" && /^\d{1,2}-[A-Za-z]{3}-\d{4}$/.test(dv.trim())) productDates.push(dv.trim());
    const n = wp.v(r, 2);
    if (!n) continue;
    const row = []; for (let c = 3; c <= 8; c++) row.push(num(wp.v(r, c)));
    if (str(n).trim().toLowerCase().startsWith("total")) { ptotal = row; break; }
    products.push([str(n), row]);
  }
  if (ptotal === null) throw new InputError("DRS Productwise Sales: no 'Total Sales' row");

  const wf = wb.sheet("Fees Calculation");
  const fees = {};
  for (let r = 2; r <= wf.maxRow; r++) {
    const d = wf.v(r, 1), gw = wf.v(r, 2);
    if (d instanceof Date && gw) {
      const t = ymd(d);
      if (t.y === periodEnd.y && t.m === periodEnd.m) {
        const f = (fees[str(gw)] ??= { orders: 0, gross: 0, fee_ex_gst: 0, gst: 0 });
        f.orders += num(wf.v(r, 3)); f.gross += num(wf.v(r, 4)); f.fee_ex_gst += num(wf.v(r, 5)); f.gst += num(wf.v(r, 6));
      }
    }
  }
  return { dates, daily, countries, country_total: total, country_period: countryPeriod,
           products, product_total: ptotal, product_period: productPeriod, product_dates: productDates, fees };
}

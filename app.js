(() => {
  // src/config.js
  var NAVY = "#1F3A5F";
  var GREEN = "#2E9E5B";
  var RED = "#C0392B";
  var CAT5 = ["#35609A", "#DE8319", "#0E9C8C", "#9B51B0", "#8E9AA6"];
  var INK = "#1a1a1a";
  var MUTED = "#5a5a5a";
  var GRID = "#DDE1E6";
  var CHART_DPI = 190;
  var PL_SECTIONS = {
    "trading income": "income",
    "income": "income",
    "revenue": "income",
    "cost of sales": "cogs",
    "cost of goods sold": "cogs",
    "direct costs": "cogs",
    "operating expenses": "opex",
    "expenses": "opex"
  };
  var PL_DERIVED_ROWS = ["gross profit", "net profit", "net income", "operating profit"];
  var BS_SECTIONS = {
    "assets": null,
    "liabilities": null,
    "bank": "Bank",
    "current assets": "Current Assets",
    "fixed assets": "Fixed Assets",
    "non-current assets": "Fixed Assets",
    "current liabilities": "Current Liabilities",
    "non-current liabilities": "Non-Current Liabilities",
    "equity": "Equity"
  };
  var LIABILITY_SHOWN_AS_BANK = ["Business Trans Acct", "Revolut USD Main"];
  var NET_ASSET_AGAINST_LIABILITY = ["Credit account (SAINT EYEWEAR LLC)"];
  var TOL_EXACT = 0.01;
  var TOL_DRS_ABS = 250;
  var TOL_DRS_PCT = 5e-3;
  var PRODUCT_COVERAGE_MIN = 0.99;
  var RESTATE_LINE_MIN = 1;
  var EXPENDITURE_EXCLUDE_PREFIXES = ["social media ads", "rounding"];
  var FEE_ACCOUNTS = { Shopify: "Fees Shopify", Paypal: "Fees Paypal", Afterpay: "Fees AfterPay" };
  var DRS_AD_MAP = [
    ["advertisment spend - meta", "Social Media Ads - Meta", 1],
    ["advertisment spend - google", "Social Media Ads - Google", 1.1],
    ["advertisment spend - pinterest", "Social Media Ads - Pinterest", 1],
    ["advertisment spend - snap", "Social Media Ads - Snapchat", 1.1],
    ["advertisment spend - tiktok", "Social Media Ads - Tiktok", 1]
  ];

  // src/util.js
  var InputError = class extends Error {
  };
  var MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var MONTH_ABBR = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var round2 = (x) => Math.round((x + Math.sign(x) * 1e-9) * 100) / 100;
  var sum = (a) => a.reduce((s, x) => s + x, 0);
  var num = (v) => typeof v === "number" && isFinite(v) ? v : 0;
  function commas(x, dp = 2) {
    const s = Math.abs(x).toFixed(dp);
    const [i, f] = s.split(".");
    return i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (f ? "." + f : "");
  }
  var mk = (y, m) => `${y}-${String(m).padStart(2, "0")}`;
  var keyY = (k) => +k.slice(0, 4);
  var keyM = (k) => +k.slice(5);
  function prevKey(k, n = 1) {
    let y = keyY(k), m = keyM(k) - n;
    while (m < 1) {
      m += 12;
      y -= 1;
    }
    return mk(y, m);
  }
  var monthLabel = (k) => `${MONTH_ABBR[keyM(k)]} ${keyY(k)}`;
  var shortLabel = (k) => `${MONTH_ABBR[keyM(k)]} '${String(keyY(k)).slice(2)}`;
  var longLabel = (k) => `${MONTH_NAMES[keyM(k)]} ${keyY(k)}`;
  var daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  function monthKeyFromLabel(label) {
    const m = /^\s*([A-Za-z]+)\.?\s+(\d{4})\s*$/.exec(String(label));
    if (!m) return null;
    const mo = monthFromWord(m[1], false);
    return mo ? mk(+m[2], mo) : null;
  }
  var MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function monthFromWord(word, strict = true) {
    const m = MON[String(word).replace(/^\.+|\.+$/g, "").toLowerCase().slice(0, 3)];
    if (!m && strict) throw new InputError(`Cannot read month name '${word}'`);
    return m || 0;
  }

  // src/xl.js
  var EJ = () => globalThis.ExcelJS;
  function norm(x, formulas) {
    if (x === null || x === void 0) return null;
    if (x instanceof Date) return x;
    const t = typeof x;
    if (t === "number" || t === "string" || t === "boolean") return x === "" ? null : x;
    if (x.formula !== void 0 || x.sharedFormula !== void 0) {
      return formulas ? "=" + (x.formula || x.sharedFormula) : norm(x.result, false);
    }
    if (x.richText) return x.richText.map((r) => r.text).join("");
    if (x.error) return x.error;
    if (x.text !== void 0) return norm(x.text, formulas);
    return null;
  }
  var Sheet = class {
    constructor(ws) {
      this.ws = ws;
      this.name = ws.name;
    }
    get maxRow() {
      return this.ws.rowCount;
    }
    get maxCol() {
      return this.ws.columnCount;
    }
    v(r, c, opts) {
      const row = this.ws.getRow(r);
      return norm(row.getCell(c).value, !!(opts && opts.formulas));
    }
  };
  var Book = class {
    constructor(wb) {
      this.wb = wb;
      this.sheets = wb.worksheets.map((w) => new Sheet(w));
      this.names = this.sheets.map((s) => s.name);
    }
    first() {
      return this.sheets[0];
    }
    sheet(name) {
      return this.sheets.find((s) => s.name === name);
    }
  };
  async function openBook(buf) {
    const wb = new (EJ()).Workbook();
    await wb.xlsx.load(buf);
    return new Book(wb);
  }

  // src/xero.js
  var isNum = (v) => typeof v === "number";
  var str = (v) => v === null || v === void 0 ? "" : String(v);
  var PL = class {
    constructor(periodEnd, keys, lines, ytdLabel, hasYtd, xeroTotals) {
      Object.assign(this, { periodEnd, keys, lines, ytdLabel, hasYtd, xeroTotals });
    }
    acct(name, key) {
      const l = this.lines.find((x) => x.name === name);
      return l ? l.vals[key] ?? 0 : 0;
    }
    sect(section, key) {
      return round2(this.lines.filter((l) => l.section === section).reduce((s, l) => s + (l.vals[key] ?? 0), 0));
    }
    totals(key) {
      const inc = this.sect("income", key), cogs = this.sect("cogs", key), opex = this.sect("opex", key);
      return { income: inc, cogs, gp: round2(inc - cogs), opex, net: round2(inc - cogs - opex) };
    }
  };
  async function readPL(buf, label = "P&L") {
    const ws = (await openBook(buf)).first();
    if (!str(ws.v(1, 1)).toLowerCase().includes("profit")) throw new InputError(`${label}: first sheet is not a Profit and Loss report`);
    const title = [2, 3].map((r) => str(ws.v(r, 1))).join(" ");
    const m = /ended\s+(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/.exec(title);
    if (!m) throw new InputError("P&L header must read 'For the month ended <date>' - export the exact month");
    const py = +m[3], pm = monthFromWord(m[2]), pd = +m[1];
    if (pd !== daysIn(py, pm)) throw new InputError(`P&L ends ${pd}/${pm}/${py}, which is not a month end`);
    const periodEnd = { y: py, m: pm, d: pd };
    let hdr = null;
    for (let r = 1; r < 15; r++) if (str(ws.v(r, 1)).trim().toLowerCase() === "account") {
      hdr = r;
      break;
    }
    if (hdr === null) throw new InputError("P&L: could not find the 'Account' header row");
    const cols = {}, keys = [];
    let ytdCol = null, ytdLabel = "";
    for (let c = 2; c <= ws.maxCol; c++) {
      const lab = ws.v(hdr, c);
      if (lab === null) continue;
      const k = monthKeyFromLabel(lab);
      if (k) {
        cols[c] = k;
        keys.push(k);
      } else if (str(lab).toLowerCase().includes("year to date")) {
        ytdCol = c;
        ytdLabel = str(lab);
      }
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
        if (!(low in PL_SECTIONS)) throw new InputError(`P&L: unknown section heading '${name}' (row ${r}). Decide where it belongs, then add it to PL_SECTIONS in js/config.js.`);
        section = PL_SECTIONS[low];
        continue;
      }
      if (low.startsWith("total ") || PL_DERIVED_ROWS.includes(low)) {
        for (const c of colList) if (isNum(raw[c])) (xeroTotals[low] ?? (xeroTotals[low] = {}))[cols[c] ?? "YTD"] = raw[c];
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
  async function readBS(buf) {
    const ws = (await openBook(buf)).first();
    if (!str(ws.v(1, 1)).toLowerCase().includes("balance")) throw new InputError("Balance Sheet: first sheet is not a Balance Sheet report");
    let hdr = null;
    for (let r = 1; r < 15; r++) if (str(ws.v(r, 2)).trim().toLowerCase() === "account") {
      hdr = r;
      break;
    }
    if (hdr === null) throw new InputError("Balance Sheet: could not find the 'Account' header row");
    const cols = {};
    for (let c = 3; c <= ws.maxCol; c++) {
      const lab = ws.v(hdr, c);
      if (lab === null) continue;
      if (lab instanceof Date) {
        cols[c] = mk(lab.getUTCFullYear(), lab.getUTCMonth() + 1);
        continue;
      }
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
        if (!(low in BS_SECTIONS)) throw new InputError(`Balance Sheet: unknown heading '${name}' (row ${r}) - add to BS_SECTIONS in js/config.js`);
        const g = BS_SECTIONS[low];
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
  function bsPresent(bs) {
    var _a, _b, _c;
    const keys = bs.keys;
    const by = /* @__PURE__ */ new Map();
    for (const l of bs.lines) {
      const id = l.group + "\0" + l.name;
      if (!by.has(id)) by.set(id, l);
    }
    const v = (g, n, k) => {
      const l = by.get(g + "\0" + n);
      return l ? l.vals[k] ?? 0 : 0;
    };
    const groups = ["Bank", "Current Assets", "Fixed Assets", "Current Liabilities", "Non-Current Liabilities", "Equity"];
    const sec = Object.fromEntries(groups.map((g) => [g, {}]));
    const zero = () => Object.fromEntries(keys.map((k) => [k, 0]));
    for (const l of bs.lines) {
      if (l.group === "Bank" && NET_ASSET_AGAINST_LIABILITY.includes(l.name)) continue;
      if (l.group === "Current Liabilities" && LIABILITY_SHOWN_AS_BANK.includes(l.name)) {
        sec.Bank[l.name] = Object.fromEntries(Object.entries(l.vals).map(([k, x]) => [k, -x]));
        continue;
      }
      const row = (_a = sec[l.group])[_b = l.name] ?? (_a[_b] = zero());
      for (const [k, x] of Object.entries(l.vals)) row[k] = (row[k] ?? 0) + x;
    }
    for (const name of NET_ASSET_AGAINST_LIABILITY) {
      const row = (_c = sec["Current Liabilities"])[name] ?? (_c[name] = zero());
      for (const k of keys) row[k] = v("Current Liabilities", name, k) - v("Bank", name, k);
    }
    for (const g of groups) {
      for (const n of Object.keys(sec[g])) {
        if (!keys.some((k) => Math.abs(sec[g][n][k] ?? 0) >= 5e-3)) delete sec[g][n];
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
  async function readPLPeriod(buf) {
    const ws = (await openBook(buf)).first();
    if (!str(ws.v(1, 1)).toLowerCase().includes("profit")) throw new InputError("YTD file: not a Profit and Loss report");
    const title = [2, 3].map((r) => str(ws.v(r, 1))).join(" ");
    let hdr = null;
    for (let r = 1; r < 15; r++) if (str(ws.v(r, 1)).trim().toLowerCase() === "account") {
      hdr = r;
      break;
    }
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
        if (!(low in PL_SECTIONS)) throw new InputError(`YTD P&L: unknown section heading '${name}'`);
        section = PL_SECTIONS[low];
        continue;
      }
      if (low.startsWith("total ") || PL_DERIVED_ROWS.includes(low)) continue;
      sums[section] += num(v);
    }
    const t = { income: round2(sums.income), cogs: round2(sums.cogs), opex: round2(sums.opex) };
    t.gp = round2(t.income - t.cogs);
    t.net = round2(t.gp - t.opex);
    return [t, title.trim()];
  }

  // src/drs.js
  var str2 = (v) => v === null || v === void 0 ? "" : String(v);
  var ymd = (d) => ({ y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() });
  function findRow(ws, needle, col = 1, maxr = 80) {
    needle = needle.toLowerCase();
    for (let r = 1; r < maxr; r++) {
      const v = ws.v(r, col);
      if (v && str2(v).toLowerCase().includes(needle)) return r;
    }
    throw new InputError(`DRS 'Daily Reporting': no row labelled like '${needle}'`);
  }
  async function readDRS(buf, periodEnd) {
    var _a;
    const wb = await openBook(buf);
    for (const n of ["Daily Reporting", "Countrywise Sales", "Productwise Sales", "Fees Calculation"]) {
      if (!wb.names.includes(n)) throw new InputError(`DRS: sheet '${n}' is missing`);
    }
    const ws = wb.sheet("Daily Reporting");
    const drow = findRow(ws, "date");
    const dates = [], cols = [];
    for (let c = 2; c <= ws.maxCol; c++) {
      const v = ws.v(drow, c);
      if (v instanceof Date) {
        dates.push(ymd(v));
        cols.push(c);
      }
    }
    if (!dates.length) throw new InputError("DRS: no dates on the Daily Reporting sheet");
    const nDays = daysIn(periodEnd.y, periodEnd.m);
    const ok = dates.filter((d) => d.y === periodEnd.y && d.m === periodEnd.m);
    if (ok.length !== nDays || ok.length !== dates.length) {
      throw new InputError(`DRS: Daily Reporting holds ${dates.length} dates, ${ok.length} in ${MONTH_NAMES[periodEnd.m]} ${periodEnd.y}; expected a full month of ${nDays}`);
    }
    const labels = /* @__PURE__ */ new Map();
    for (let r = 1; r < 60; r++) {
      const lab = ws.v(r, 1);
      if (lab) labels.set(str2(lab).trim().toLowerCase(), r);
    }
    const series = (needle) => {
      for (const [lab, r] of labels) if (lab.includes(needle)) return cols.map((c) => num(ws.v(r, c)));
      throw new InputError(`DRS: no row like '${needle}'`);
    };
    const daily = {
      orders: series("number of orders"),
      gross: series("gross sales"),
      discounts: series("discounts"),
      returns: series("returns"),
      net_sales: series("net sales"),
      ship_collected: series("shipping charges collected"),
      prod_rev: series("product revenue (ex gst"),
      ship_au: series("shipping revenue (ex gst, a$) - au"),
      ship_other: series("shipping revenue - others"),
      refund_ex_gst: series("refund ex gst"),
      total_rev: series("total revenue (ex gst"),
      cogs_product: series("cogs product cost"),
      cogs_ship: series("cogs shipping cost"),
      processing: series("processing fee"),
      merchant_fees: series("merchant fees"),
      ad_total: series("advertisement spend")
    };
    for (const [lab, r] of labels) if (lab.startsWith("advertisment spend -")) daily[lab] = cols.map((c) => num(ws.v(r, c)));
    const wc = wb.sheet("Countrywise Sales");
    const countryPeriod = str2(wc.v(1, 2));
    const countries = [];
    let total = null;
    for (let r = 2; r < 200; r++) {
      const n = wc.v(r, 2), v = wc.v(r, 3);
      if (n && str2(n).startsWith("Product Revenue") && typeof v === "number") countries.push([str2(n), v]);
      else if (n && str2(n).trim().toLowerCase() === "total") {
        total = num(v);
        break;
      }
    }
    if (total === null) throw new InputError("DRS Countrywise Sales: no Total row");
    const wp = wb.sheet("Productwise Sales");
    const productPeriod = str2(wp.v(1, 2));
    const products = [], productDates = [];
    let ptotal = null;
    for (let r = 3; r < 200; r++) {
      const dv = wp.v(r, 1);
      if (dv instanceof Date) productDates.push(`${dv.getUTCDate()}-${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][dv.getUTCMonth()]}-${dv.getUTCFullYear()}`);
      else if (typeof dv === "string" && /^\d{1,2}-[A-Za-z]{3}-\d{4}$/.test(dv.trim())) productDates.push(dv.trim());
      const n = wp.v(r, 2);
      if (!n) continue;
      const row = [];
      for (let c = 3; c <= 8; c++) row.push(num(wp.v(r, c)));
      if (str2(n).trim().toLowerCase().startsWith("total")) {
        ptotal = row;
        break;
      }
      products.push([str2(n), row]);
    }
    if (ptotal === null) throw new InputError("DRS Productwise Sales: no 'Total Sales' row");
    const wf = wb.sheet("Fees Calculation");
    const fees2 = {};
    for (let r = 2; r <= wf.maxRow; r++) {
      const d = wf.v(r, 1), gw = wf.v(r, 2);
      if (d instanceof Date && gw) {
        const t = ymd(d);
        if (t.y === periodEnd.y && t.m === periodEnd.m) {
          const f = fees2[_a = str2(gw)] ?? (fees2[_a] = { orders: 0, gross: 0, fee_ex_gst: 0, gst: 0 });
          f.orders += num(wf.v(r, 3));
          f.gross += num(wf.v(r, 4));
          f.fee_ex_gst += num(wf.v(r, 5));
          f.gst += num(wf.v(r, 6));
        }
      }
    }
    return {
      dates,
      daily,
      countries,
      country_total: total,
      country_period: countryPeriod,
      products,
      product_total: ptotal,
      product_period: productPeriod,
      product_dates: productDates,
      fees: fees2
    };
  }

  // src/dom.js
  var NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    xml: "http://www.w3.org/XML/1998/namespace"
  };
  var parseXml = (s) => {
    const d = new DOMParser().parseFromString(s, "application/xml");
    const err = d.getElementsByTagName("parsererror");
    if (err.length) throw new Error("Could not read the Word file (XML error)");
    return d;
  };
  var serializeXml = (doc) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc.documentElement);
  var isEl = (n) => n && n.nodeType === 1;
  var elems = (el) => Array.from(el.childNodes).filter(isEl);
  var kids = (el, name, ns = NS.w) => elems(el).filter((c) => c.localName === name && c.namespaceURI === ns);
  var kid = (el, name, ns = NS.w) => kids(el, name, ns)[0] || null;
  var desc = (el, name, ns = NS.w) => Array.from(el.getElementsByTagNameNS(ns, name));
  var firstDesc = (el, name, ns = NS.w) => el.getElementsByTagNameNS(ns, name)[0] || null;
  var wAttr = (el, name) => el.getAttributeNS(NS.w, name);
  var setW = (el, name, val) => el.setAttributeNS(NS.w, "w:" + name, String(val));
  function make(doc, name, attrs = {}, ns = NS.w, prefix = "w") {
    const e = doc.createElementNS(ns, prefix + ":" + name);
    for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(ns, prefix + ":" + k, String(v));
    return e;
  }
  function insertAt(parent, idx, node) {
    const ch = elems(parent);
    if (idx < ch.length) parent.insertBefore(node, ch[idx]);
    else parent.appendChild(node);
  }
  var after = (ref, node) => ref.parentNode.insertBefore(node, ref.nextSibling);
  var before = (ref, node) => ref.parentNode.insertBefore(node, ref);
  var remove = (n) => n.parentNode && n.parentNode.removeChild(n);
  var nextEl = (n) => {
    let x = n.nextSibling;
    while (x && !isEl(x)) x = x.nextSibling;
    return x;
  };
  function ptext(el) {
    return desc(el, "t").map((t) => t.textContent || "").join("");
  }

  // src/prior.js
  async function loadDocXml(bytes2) {
    const z = await globalThis.JSZip.loadAsync(bytes2);
    return parseXml(await z.file("word/document.xml").async("string"));
  }
  var cellText = (tc) => ptext(tc).replace(/ /g, " ").trim();
  function parseAmount(s) {
    s = String(s || "").replace(/\$/g, "").replace(/,/g, "").replace(/ /g, "").trim();
    if (!s) return 0;
    const neg = s.startsWith("(") && s.endsWith(")") || s.startsWith("-");
    s = s.replace(/^[()\-]+|[()\-]+$/g, "");
    const v = parseFloat(s);
    if (!isFinite(v) || !/^[0-9.]+$/.test(s)) return 0;
    return neg ? -v : v;
  }
  function loadTables(doc) {
    return desc(doc, "tbl").map((tb) => kids(tb, "tr").map((tr) => kids(tr, "tc").map(cellText)));
  }
  function findTables(doc) {
    const found = {};
    for (const rows of loadTables(doc)) {
      const flat = rows.filter((r) => r.length).map((r) => r[0]);
      const head = rows[0] || [];
      if (rows.slice(0, 3).some((r) => r.length && (r[0] === "Metrix" || r[0] === "Metric"))) found.exec = rows;
      else if (head.length && head[0] === "Country") found.country = rows;
      else if (head.length && head[0] === "Name" && flat.includes("Total Income")) found.pl = rows;
      else if (head.length && head[0] === "Name" && flat.includes("Net Assets")) found.bs = rows;
    }
    const missing = ["exec", "country", "pl", "bs"].filter((k) => !found[k]);
    if (missing.length) throw new InputError(`Report: could not find tables ${JSON.stringify(missing)}`);
    return found;
  }
  function monthKeyFromShort(label) {
    const m = /^([A-Za-z]+)\s*'(\d\d)/.exec(label.trim());
    return m ? monthKeyFromLabel(`${m[1]} 20${m[2]}`) : null;
  }
  function tableByMonth(rows) {
    const head = rows[0];
    const cols = {};
    head.forEach((h, i) => {
      if (i && monthKeyFromShort(h)) cols[i] = monthKeyFromShort(h);
    });
    const out = {};
    for (const r of rows.slice(1)) {
      if (!r.length || !r[0]) continue;
      const o = {};
      for (const [i, k] of Object.entries(cols)) if (+i < r.length) o[k] = parseAmount(r[i]);
      out[r[0]] = o;
    }
    return [out, Object.values(cols)];
  }
  async function readPrior(bytes2) {
    const t = findTables(await loadDocXml(bytes2));
    const [pl, plKeys] = tableByMonth(t.pl);
    const [bs, bsKeys] = tableByMonth(t.bs);
    return { pl, pl_keys: plKeys, bs, bs_keys: bsKeys, tables: t };
  }

  // src/checks.js
  var FY_START_MONTH = 7;
  function chk(group, name, expected, actual, { tol = TOL_EXACT, gate = false, note = "" } = {}) {
    const diff = expected === null || actual === null || expected === void 0 || actual === void 0 ? null : round2(actual - expected);
    let status;
    if (diff === null) status = "INFO";
    else if (Math.abs(diff) <= tol) status = "PASS";
    else status = gate ? "FAIL" : "WARN";
    return { group, name, expected: expected ?? null, actual: actual ?? null, diff, status, gate, note };
  }
  function fyMonths(k) {
    const y = keyY(k), m = keyM(k);
    let cy = m >= FY_START_MONTH ? y : y - 1, cm = FY_START_MONTH;
    const out = [];
    while (mk(cy, cm) <= k) {
      out.push(mk(cy, cm));
      cm += 1;
      if (cm > 12) {
        cm = 1;
        cy += 1;
      }
    }
    return out;
  }
  var plLabelTotals = (t) => ({
    "Total Income": t.income,
    "Total Cost of Goods Sold": t.cogs,
    "Gross Profit": t.gp,
    "Total Operating Expenses": t.opex,
    "Net Income": t.net
  });
  function runChecks(pl, bs, drs, prior, k0, sec, tot) {
    const C = [];
    const k1 = prevKey(k0, 1), k2 = prevKey(k0, 2);
    const need2 = [k0, k1, k2];
    for (const k of need2) {
      C.push(chk("Inputs", `P&L has column ${monthLabel(k)}`, 1, pl.keys.includes(k) ? 1 : 0, { gate: true }));
      C.push(chk("Inputs", `Balance Sheet has column ${monthLabel(k)}`, 1, bs.keys.includes(k) ? 1 : 0, { gate: true }));
    }
    for (const [lab, key] of [
      ["total trading income", "income"],
      ["total income", "income"],
      ["total cost of sales", "cogs"],
      ["total cost of goods sold", "cogs"],
      ["total operating expenses", "opex"],
      ["gross profit", "gp"],
      ["net profit", "net"]
    ]) {
      for (const [k, v] of Object.entries(pl.xeroTotals[lab] || {})) {
        C.push(chk("Xero totals", `${lab} ${k}`, v, pl.totals(k)[key], { gate: true }));
      }
    }
    if (pl.hasYtd) {
      const fm = fyMonths(k0);
      if (fm.every((m) => pl.keys.includes(m))) {
        for (const [sname, s] of [["income", "income"], ["cost of sales", "cogs"], ["operating expenses", "opex"]]) {
          C.push(chk(
            "YTD",
            `Xero 'Year to date' ${sname} = sum of ${fm.length} months`,
            round2(sum(fm.map((m) => pl.sect(s, m)))),
            pl.sect(s, "YTD"),
            { gate: true }
          ));
        }
        const bad = [];
        for (const l of pl.lines) {
          const d2 = round2(sum(fm.map((m) => l.vals[m] ?? 0)) - (l.vals.YTD ?? 0));
          if (Math.abs(d2) > TOL_EXACT) bad.push([l.name, d2]);
        }
        C.push(chk("YTD", "Every account: months add to 'Year to date'", 0, bad.length, {
          gate: true,
          note: bad.slice(0, 6).map(([n, d2]) => `${n} ${d2 >= 0 ? "+" : "-"}${Math.abs(d2).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`).join("; ")
        }));
      } else {
        C.push(chk("YTD", "Year to date cannot be re-added (months missing)", null, null));
      }
    }
    for (const k of need2) {
      if (tot[k]) C.push(chk("Balance sheet", `Net Assets = Total Equity, ${monthLabel(k)}`, tot[k]["Total Equity"], tot[k]["Net Assets"], { gate: true }));
    }
    const cye = bs.lines.find((l) => l.name.toLowerCase() === "current year earnings");
    if (cye) {
      for (const k of need2) {
        const fm = fyMonths(k);
        if (k in cye.vals && fm.every((m) => pl.keys.includes(m))) {
          C.push(chk(
            "BS vs P&L",
            `Current Year Earnings ${monthLabel(k)} = P&L net since 1 July`,
            round2(sum(fm.map((m) => pl.totals(m).net))),
            cye.vals[k],
            { gate: true }
          ));
        }
      }
    }
    const d = drs.daily;
    const tI = pl.totals(k0);
    const drsChk = (name, drsVal, xeroVal, note = "") => {
      const tol = Math.max(TOL_DRS_ABS, Math.abs(xeroVal) * TOL_DRS_PCT);
      C.push(chk("DRS vs P&L", name, xeroVal, drsVal, { tol, note }));
    };
    drsChk("Countrywise total vs P&L total income", drs.country_total, tI.income);
    drsChk("Shipping income AU", sum(d.ship_au), pl.acct("Shipping Income - AU", k0));
    drsChk("Shipping income other", sum(d.ship_other), pl.acct("Shipping Income - Other", k0));
    drsChk("COGS (product) daily sum vs P&L Cost of Goods Sold", sum(d.cogs_product), pl.acct("Cost of Goods Sold", k0), "Daily COGS is on the supplier cost sheet");
    drsChk("Shipping cost daily sum vs P&L", sum(d.cogs_ship), pl.acct("Shipping Cost", k0));
    for (const [lab, acct, div] of DRS_AD_MAP) {
      const key = Object.keys(d).find((x) => x.startsWith(lab));
      if (key) drsChk(`Ad spend ${acct.split(" - ").pop()} (DRS${div !== 1 ? " \xF71.1" : ""})`, sum(d[key]) / div, pl.acct(acct, k0));
    }
    for (const [gw, acct] of Object.entries(FEE_ACCOUNTS)) {
      if (gw in drs.fees) drsChk(`Merchant fees ${gw} (DRS ex GST)`, drs.fees[gw].fee_ex_gst, -pl.acct(acct, k0));
    }
    const gross = sum(d.gross);
    const cov = gross ? drs.product_total[0] / gross : 0;
    C.push(chk(
      "DRS vs P&L",
      "Productwise Sales gross covers DRS gross sales",
      gross,
      drs.product_total[0],
      { tol: gross * (1 - PRODUCT_COVERAGE_MIN), note: `coverage ${(cov * 100).toFixed(1)}%` }
    ));
    const mm = `${String(keyM(k0)).padStart(2, "0")}/01/${keyY(k0)}`;
    for (const [nm, txt] of [["Countrywise Sales", drs.country_period], ["Productwise Sales", drs.product_period]]) {
      C.push(chk("DRS vs P&L", `${nm} tab is for ${MONTH_NAMES[keyM(k0)]} ${keyY(k0)}`, 1, txt.startsWith(mm) ? 1 : 0, { note: txt }));
    }
    const pm = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
    const stale = (drs.product_dates || []).filter((t) => {
      const m = /^\d{1,2}-([A-Za-z]{3})-(\d{4})$/.exec(t);
      return m && !(pm[m[1].toLowerCase()] === keyM(k0) && +m[2] === keyY(k0));
    });
    if ((drs.product_dates || []).length) {
      C.push(chk("DRS vs P&L", `Productwise Sales rows are dated ${MONTH_NAMES[keyM(k0)]} ${keyY(k0)}`, 1, stale.length ? 0 : 1, { note: stale.length ? `rows dated ${[...new Set(stale)].join(", ")}` : "" }));
    }
    const info = { product_coverage: cov, product_ok: cov >= PRODUCT_COVERAGE_MIN, product_stale: [...new Set(stale)] };
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
        if (Math.abs(now - pub) >= RESTATE_LINE_MIN) restate.lines.push({ month: k, item: l.name, published: pub, now, diff: round2(now - pub) });
      }
      const known = new Set(pl.lines.map((l) => l.name));
      const skip = /* @__PURE__ */ new Set(["Income", "Cost of Goods Sold", "Operating Expenses", "Operating Income", "Name", "Total Cost of Goods Sold"]);
      for (const [lab, v] of Object.entries(pp)) {
        if (!known.has(lab) && !(lab in plLabelTotals(zeroT)) && !skip.has(lab) && Math.abs(v[k] ?? 0) >= 1) {
          restate.lines.push({ month: k, item: lab + " (not in new P&L)", published: v[k] ?? 0, now: 0, diff: round2(-(v[k] ?? 0)) });
        }
      }
    }
    const pb = prior.bs;
    const keyed = {
      "Total Bank": "Bank",
      "Total Current Assets": "Current Assets",
      "Total Assets": "Total Assets",
      "Total Current Liabilities": "Current Liabilities",
      "Total Non-Current Liabilities": "Non-Current Liabilities",
      "Net Assets": "Net Assets",
      "Total Equity": "Total Equity"
    };
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
    const k1k = k1;
    const newPl = [], newBs = [];
    for (const l of pl.lines) {
      if (Math.abs(l.vals[k0] ?? 0) >= 5e-3 && Math.abs(l.vals[k1k] ?? 0) < 5e-3) {
        const hist = pl.keys.length > 1 ? Math.max(...pl.keys.slice(1).map((m) => Math.abs(l.vals[m] ?? 0))) : 0;
        newPl.push({ account: l.name, section: l.section, this: l.vals[k0] ?? 0, first_time: hist < 5e-3 });
      }
    }
    for (const g of Object.keys(sec)) {
      for (const [n, r] of Object.entries(sec[g])) {
        if (Math.abs(r[k0] ?? 0) >= 5e-3 && Math.abs(r[k1k] ?? 0) < 5e-3) newBs.push({ account: n, group: g, this: r[k0] ?? 0 });
      }
    }
    const structural = /* @__PURE__ */ new Set(["Assets", "Bank", "Current Assets", "Fixed Assets", "Current Liabilities", "Non-Current Liabilities", "Equity"]);
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
      checks: C,
      restate,
      new_pl: newPl,
      new_bs: newBs,
      bs_not_on_prior_report: appeared,
      orphan_accruals: orphan.sort(),
      info,
      gate_ok: !C.some((c) => c.status === "FAIL")
    };
  }

  // src/charts.js
  var DPI = CHART_DPI;
  var PT = DPI / 72;
  var FONT = "'DejaVu Sans', Verdana, Arial, Helvetica, sans-serif";
  var provider = null;
  function newCanvas(w, h) {
    w = Math.round(w);
    h = Math.round(h);
    if (provider) return provider(w, h);
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    return c;
  }
  async function toPng(c) {
    if (c.toBuffer) return new Uint8Array(c.toBuffer("image/png"));
    const blob = await new Promise((res) => c.toBlob(res, "image/png"));
    return new Uint8Array(await blob.arrayBuffer());
  }
  var A = (x, dp = 0) => (x < 0 ? "-" : "") + "A$" + commas(x, dp);
  var short = (x) => {
    const a = Math.abs(x);
    const t = a >= 1e6 ? `${(a / 1e6).toFixed(2)}M` : a >= 1e4 ? `${(a / 1e3).toFixed(1)}k` : commas(a, 0);
    return (x < 0 ? "-" : "") + "A$" + t;
  };
  var fnt = (pt, { bold = false, italic = false } = {}) => `${italic ? "italic " : ""}${bold ? "bold " : ""}${pt * PT}px ${FONT}`;
  var Fig = class {
    constructor(wIn, hIn, title) {
      this.W = Math.round(wIn * DPI);
      this.H = Math.round(hIn * DPI);
      this.c = newCanvas(this.W, this.H);
      this.g = this.c.getContext("2d");
      this.g.fillStyle = "#fff";
      this.g.fillRect(0, 0, this.W, this.H);
      this.titleBottom = 0;
      if (title) this.title(title);
    }
    title(t) {
      const g = this.g;
      g.font = fnt(11.5, { bold: true });
      g.fillStyle = NAVY;
      g.textAlign = "left";
      g.textBaseline = "top";
      g.fillText(t, 0.012 * this.W, 0.015 * this.H);
      this.titleBottom = 0.015 * this.H + 11.5 * PT * 1.25;
    }
    text(s, x, y, { pt = 9, color = INK, align = "left", base = "alphabetic", bold = false, italic = false } = {}) {
      const g = this.g;
      g.font = fnt(pt, { bold, italic });
      g.fillStyle = color;
      g.textAlign = align;
      g.textBaseline = base;
      g.fillText(s, x, y);
    }
    width(s, pt = 9, bold = false) {
      this.g.font = fnt(pt, { bold });
      return this.g.measureText(s).width;
    }
    line(x1, y1, x2, y2, color, lw = 0.8) {
      const g = this.g;
      g.strokeStyle = color;
      g.lineWidth = lw * PT;
      g.beginPath();
      g.moveTo(x1, y1);
      g.lineTo(x2, y2);
      g.stroke();
    }
    rect(x, y, w, h, color) {
      this.g.fillStyle = color;
      this.g.fillRect(x, y, w, h);
    }
    async png(crop) {
      let c = this.c;
      if (crop) {
        const c2 = newCanvas(crop.w, crop.h);
        const g2 = c2.getContext("2d");
        g2.fillStyle = "#fff";
        g2.fillRect(0, 0, crop.w, crop.h);
        g2.drawImage(this.c, -crop.x, -crop.y);
        c = c2;
      }
      return { png: await toPng(c), w: c.width, h: c.height, note: null };
    }
  };
  function niceTicks(lo, hi, n = 6) {
    const span = hi - lo || 1;
    const raw = span / n, mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) || 10 * mag;
    const out = [];
    for (let v = Math.ceil(lo / step - 1e-9) * step; v <= hi + step * 1e-9; v += step) out.push(Math.abs(v) < step * 1e-9 ? 0 : v);
    return out;
  }
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
    const tickH = opts.rot ? Math.max(...opts.xTicks.map((s) => fig.width(s, opts.xTickPt || 8))) * Math.sin(opts.rot * Math.PI / 180) + 10 * PT : (opts.xTickPt || tickPt) * PT * 0 + (opts.xTickPt || tickPt) * PT * 1.6;
    bottom += xlabH + legH + tickH;
    const ax = { x0: left, x1: W - right, y0: top, y1: H - bottom };
    const sy = (v) => ax.y1 - (v - lo) / (hi - lo) * (ax.y1 - ax.y0);
    for (const v of yTicks) {
      if (opts.hGrid !== false) fig.line(ax.x0, sy(v), ax.x1, sy(v), GRID, 0.7);
      fig.line(ax.x0 - 3.5 * PT, sy(v), ax.x0, sy(v), MUTED, 0.8);
      fig.text(fmtY(v), ax.x0 - 6 * PT, sy(v), { pt: tickPt, color: MUTED, align: "right", base: "middle" });
    }
    fig.line(ax.x0, ax.y0, ax.x0, ax.y1, GRID, 0.8);
    fig.line(ax.x0, ax.y1, ax.x1, ax.y1, GRID, 0.8);
    if (opts.yLabel) {
      g.save();
      g.translate(10 * PT * 0.9, (ax.y0 + ax.y1) / 2);
      g.rotate(-Math.PI / 2);
      fig.text(opts.yLabel, 0, 0, { pt: 9, color: MUTED, align: "center", base: "top" });
      g.restore();
    }
    const n = opts.xTicks.length;
    const sx = (i) => ax.x0 + (i + 0.5) / n * (ax.x1 - ax.x0);
    const slot = (ax.x1 - ax.x0) / n;
    opts.xTicks.forEach((lab, i) => {
      const x = sx(i);
      fig.line(x, ax.y1, x, ax.y1 + 3.5 * PT, MUTED, 0.8);
      if (opts.rot) {
        g.save();
        g.translate(x + 3 * PT, ax.y1 + 7 * PT);
        g.rotate(-opts.rot * Math.PI / 180);
        fig.text(lab, 0, 0, { pt: opts.xTickPt || 8, color: MUTED, align: "right", base: "middle" });
        g.restore();
      } else {
        fig.text(lab, x, ax.y1 + 6 * PT, { pt: opts.xTickPt || tickPt, color: MUTED, align: "center", base: "top" });
      }
    });
    let yy = ax.y1 + tickH;
    if (opts.xLabel) {
      fig.text(opts.xLabel, (ax.x0 + ax.x1) / 2, yy + 3 * PT, { pt: 9, color: MUTED, align: "center", base: "top" });
      yy += xlabH;
    }
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
  var yRange = (vals, pad = 0.15, includeZero = true) => {
    let mn = Math.min(...vals), mx = Math.max(...vals);
    if (includeZero) {
      mn = Math.min(mn, 0);
      mx = Math.max(mx, 0);
    }
    const span = mx - mn || 1;
    return [mn < 0 ? mn - pad * span : mn, mx > 0 ? mx + pad * span : mx];
  };
  async function ytd(w, h, label, vals) {
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
  async function fees(w, h, monthName, feeByGw) {
    const names = Object.keys(feeByGw);
    const tot = names.reduce((s, n) => s + feeByGw[n], 0);
    const top = names.reduce((a, b) => feeByGw[b] > feeByGw[a] ? b : a);
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
  async function expenditure(w, h, items, adsTotal) {
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
        const x = p.sx(i) - k * bw / 2 + j * bw;
        f.rect(x, Math.min(p.sy(v), p.sy(0)), bw, Math.abs(p.sy(v) - p.sy(0)), s.color);
      });
    });
  }
  async function daywiseSales(w, h, dates, orders, returns, ship) {
    const first = orders.slice(0, 7).reduce((a, b) => a + b, 0) / 7, last = orders.slice(-7).reduce((a, b) => a + b, 0) / 7;
    const f = new Fig(w, h, `Orders fell from about ${Math.round(first)} a day to ${Math.round(last)} a day over the month`);
    const all = [...orders, ...returns, ...ship];
    const [lo, hi] = yRange(all, 0.05);
    const series = [{ vals: orders, color: CAT5[0] }, { vals: returns, color: CAT5[1] }, { vals: ship, color: CAT5[4] }];
    const p = plot(f, {
      ymin: lo,
      ymax: hi,
      xTicks: dates.map((d) => String(d.d)),
      xTickPt: 7.5,
      xLabel: `Day of ${MONTH_NAMES[dates[0].m]}`,
      legend: [{ label: "Number of orders", color: CAT5[0] }, { label: "Returns (A$)", color: CAT5[1] }, { label: "Shipping charges collected (A$)", color: CAT5[4] }]
    });
    grouped(f, p, series, dates.length, 0.28);
    f.line(p.x0, p.sy(0), p.x1, p.sy(0), MUTED, 0.8);
    return f.png();
  }
  async function daywiseCogs(w, h, dates, cogs, ship) {
    const sum2 = (a) => a.reduce((s, x) => s + x, 0);
    const ratio = sum2(cogs) ? sum2(ship) / sum2(cogs) : 0;
    const f = new Fig(w, h, `Shipping costs about ${ratio.toFixed(1)}x as much as the product itself, every day`);
    const [lo, hi] = yRange([...cogs, ...ship], 0.05);
    const p = plot(f, {
      ymin: lo,
      ymax: hi,
      yLabel: "A$",
      xTicks: dates.map((d) => String(d.d)),
      xTickPt: 7.5,
      xLabel: `Day of ${MONTH_NAMES[dates[0].m]}`,
      legend: [{ label: "Cost of goods sold", color: CAT5[0] }, { label: "Shipping cost", color: CAT5[1] }]
    });
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
      g.beginPath();
      g.moveTo(cx, cy);
      g.arc(cx, cy, D / 2, a, a + da);
      g.closePath();
      g.fillStyle = colors[i];
      g.fill();
      g.strokeStyle = "#fff";
      g.lineWidth = 2 * PT;
      g.stroke();
      a += da;
    });
    const lx = cx + D / 2 + f.W * 0.05, lh = 9 * PT * 1.9, y0 = cy - parts.length * lh / 2 + lh / 2;
    parts.forEach((_, i) => {
      f.rect(lx, y0 + i * lh - 4.5 * PT, 12 * PT, 9 * PT, colors[i]);
      f.text(legendLabels[i], lx + 18 * PT, y0 + i * lh, { pt: 9, base: "middle" });
    });
    return f.png();
  }
  async function costPie(w, h, parts) {
    parts = [...parts].sort((a, b) => b[1] - a[1]);
    const tot = parts.reduce((s, x) => s + x[1], 0), big = parts[0];
    return pie(
      w,
      h,
      `${big[0]} is ${Math.round(big[1] / tot * 100)}% of cost of sales (${A(tot)})`,
      parts,
      CAT5.slice(0, 4).slice(0, parts.length),
      parts.map(([n, v]) => `${n}   ${A(v)}  (${Math.round(v / tot * 100)}%)`)
    );
  }
  async function countryPie(w, h, countries, monthName, minShare = 0.015) {
    const pos = countries.map(([n, v]) => [n.replace("Product Revenue ", ""), v]);
    const tot = pos.reduce((s, x) => s + x[1], 0);
    const big = pos.filter((x) => x[1] / tot >= minShare).sort((a, b) => b[1] - a[1]);
    const rest = tot - big.reduce((s, x) => s + x[1], 0), nrest = pos.length - big.length;
    const parts = [...big, ...nrest ? [[`All other (${nrest} countries)`, rest]] : []];
    const t2 = big.slice(0, 2);
    let cols = [...CAT5, ...CAT5, ...CAT5].slice(0, parts.length);
    if (nrest) cols[cols.length - 1] = CAT5[4];
    if (nrest && parts.length - 1 >= 4) {
      cols = [...CAT5.slice(0, 4), ...Array(Math.max(0, parts.length - 5)).fill("#B9C1CA"), "#8E9AA6"].slice(0, parts.length);
    }
    return pie(
      w,
      h,
      `${t2[0][0]} and ${t2[1][0]} bring in ${Math.round((t2[0][1] + t2[1][1]) / tot * 100)}% of ${monthName} revenue`,
      parts,
      cols,
      parts.map(([n, v]) => `${n}   ${(v / tot * 100).toFixed(1)}%`)
    );
  }
  async function placeholder(w, h, title, msg) {
    const f = new Fig(w, h, null);
    const g = f.g;
    g.fillStyle = "#F2F5F8";
    g.strokeStyle = GRID;
    g.lineWidth = 1.2 * PT;
    g.fillRect(0.01 * f.W, 0.06 * f.H, 0.98 * f.W, 0.9 * f.H);
    g.strokeRect(0.01 * f.W, 0.06 * f.H, 0.98 * f.W, 0.9 * f.H);
    const two = h >= 0.8;
    f.text(title + (two ? "" : "  -  " + msg), f.W / 2, f.H * (two ? 0.38 : 0.5), { pt: two ? 12 : 9.5, bold: true, color: NAVY, align: "center", base: "middle" });
    if (two) f.text(msg, f.W / 2, f.H * 0.66, { pt: 9.5, color: MUTED, align: "center", base: "middle" });
    return f.png();
  }
  async function productTable(w, products, total) {
    const rows = [...products].sort((a, b) => b[1][0] - a[1][0]);
    const rh = 0.27 * 1.3 * 0.82 * DPI * 0.82, hh = rh * 1.7;
    const n = rows.length + 1;
    const H = Math.ceil(hh + rh * n + 0.1 * DPI);
    const f = new Fig(w, H / DPI, null);
    const g = f.g, fmt2 = (v) => commas(v, 2).replace(/^/, v < 0 ? "-" : "");
    const fr = [0.34, 0.13, 0.12, 0.1, 0.12, 0.095, 0.095];
    const xs = [0];
    fr.forEach((x) => xs.push(xs[xs.length - 1] + x * f.W));
    const head = [["Product title"], ["Gross sales", "(A$)"], ["Discounts"], ["Refunds"], ["Net sales", "(A$)"], ["Share of", "gross %"], ["Share of", "net %"]];
    g.fillStyle = NAVY;
    g.fillRect(0, 0, f.W, hh);
    head.forEach((lines, c) => {
      lines.forEach((ln, i) => {
        const y = hh / 2 + (i - (lines.length - 1) / 2) * 7.5 * PT * 1.25;
        f.text(ln, c === 0 ? xs[c] + 4 * PT : xs[c + 1] - 4 * PT, y, { pt: 7.5, bold: true, color: "#fff", align: c === 0 ? "left" : "right", base: "middle" });
      });
    });
    const body = rows.map(([nm, v]) => [nm.length <= 32 ? nm : nm.slice(0, 31) + "\u2026", fmt2(v[0]), fmt2(v[1]), fmt2(v[2]), fmt2(v[3]), v[4].toFixed(2), v[5].toFixed(2)]);
    body.push(["Total Sales", fmt2(total[0]), fmt2(total[1]), fmt2(total[2]), fmt2(total[3]), "100.00", "100.00"]);
    body.forEach((r, i) => {
      const y = hh + i * rh, last = i === body.length - 1;
      if (last) f.rect(0, y, f.W, rh, "#DCE6F1");
      else if ((i + 2) % 2 === 0) f.rect(0, y, f.W, rh, "#F2F5F8");
      r.forEach((s, c) => f.text(s, c === 0 ? xs[c] + 4 * PT : xs[c + 1] - 4 * PT, y + rh / 2, { pt: 7.5, bold: last, color: last ? NAVY : INK, align: c === 0 ? "left" : "right", base: "middle" }));
      f.line(0, y + rh, f.W, y + rh, GRID, 0.6);
    });
    xs.forEach((x) => f.line(x, 0, x, hh + rh * body.length, GRID, 0.6));
    f.line(f.W - 1, 0, f.W - 1, hh + rh * body.length, GRID, 0.6);
    return f.png({ x: 0, y: 0, w: f.W, h: Math.ceil(hh + rh * body.length + 2) });
  }
  async function productChart(w, h, products) {
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
    f.line(ax.x0, ax.y0, ax.x0, ax.y1, GRID, 0.8);
    f.line(ax.x0, ax.y1, ax.x1, ax.y1, GRID, 0.8);
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

  // src/docxbuild.js
  var SHADE = "dce6f1";
  var RED_TXT = "9c0006";
  var MONTH_RE = MONTH_NAMES.slice(1).join("|");
  function fmtAcc(x, blankZero = true) {
    if (x === null || x === void 0 || blankZero && Math.abs(x) < 5e-3) return "";
    const s = commas(x, 2);
    return x < -4e-3 ? `(${s})` : s;
  }
  function fmtUsd0(x, blankZero = true) {
    if (x === null || x === void 0 || blankZero && Math.abs(x) < 0.5) return "";
    return x < -0.5 ? `-$${commas(x, 0)}` : `$${commas(x, 0)}`;
  }
  var fmtUsd2 = (x) => x < 0 ? `-$${commas(x, 2)}` : `$${commas(x, 2)}`;
  function rprSet(rpr, bold, color) {
    const doc = rpr.ownerDocument;
    if (bold !== null && bold !== void 0) {
      for (const tag of ["b", "bCs"]) kids(rpr, tag).forEach(remove);
      if (bold) {
        const pos = kid(rpr, "rFonts") ? 1 : 0;
        ["b", "bCs"].forEach((tag, i) => insertAt(rpr, pos + i, make(doc, tag, { val: "1" })));
      }
    }
    if (color !== null && color !== void 0) {
      kids(rpr, "color").forEach(remove);
      const e = make(doc, "color", { val: color });
      const sz = kid(rpr, "sz");
      if (sz) before(sz, e);
      else rpr.appendChild(e);
    }
  }
  function baseRpr(tc) {
    for (const r of desc(tc, "r")) {
      if (kid(r, "t") && kid(r, "rPr")) return kid(r, "rPr").cloneNode(true);
    }
    return null;
  }
  function setCell(tc, text, base, { bold = null, color = null, align = null, fill = "keep" } = {}) {
    const doc = tc.ownerDocument;
    const ps = kids(tc, "p");
    ps.slice(1).forEach(remove);
    const p = ps[0];
    for (const ch of Array.from(p.childNodes)) if (!(ch.nodeType === 1 && ch.localName === "pPr" && ch.namespaceURI === NS.w)) p.removeChild(ch);
    let ppr = kid(p, "pPr");
    if (!ppr) {
      ppr = make(doc, "pPr");
      p.insertBefore(ppr, p.firstChild);
    }
    if (align) {
      kids(ppr, "jc").forEach(remove);
      const jc = make(doc, "jc", { val: align });
      const rp = kid(ppr, "rPr");
      if (rp) before(rp, jc);
      else ppr.appendChild(jc);
    }
    if (text !== "") {
      const rpr = base ? base.cloneNode(true) : make(doc, "rPr");
      rprSet(rpr, bold, color);
      const r = make(doc, "r");
      r.appendChild(rpr);
      const t = make(doc, "t");
      t.textContent = text;
      t.setAttributeNS(NS.xml, "xml:space", "preserve");
      r.appendChild(t);
      p.appendChild(r);
    }
    if (fill !== "keep") {
      const tcpr = kid(tc, "tcPr");
      let shd = kid(tcpr, "shd");
      if (fill === null) {
        if (shd) remove(shd);
      } else {
        if (!shd) {
          shd = make(doc, "shd");
          const anchor = kid(tcpr, "tcMar");
          if (anchor) before(anchor, shd);
          else tcpr.appendChild(shd);
        }
        setW(shd, "fill", fill);
        setW(shd, "val", "clear");
      }
    }
  }
  function setRowProps(tr, header = false) {
    const doc = tr.ownerDocument;
    let trpr = kid(tr, "trPr");
    if (!trpr) {
      trpr = make(doc, "trPr");
      insertAt(tr, 0, trpr);
    }
    elems(trpr).forEach(remove);
    trpr.appendChild(make(doc, "cantSplit", { val: "1" }));
    if (header) trpr.appendChild(make(doc, "tblHeader", { val: "1" }));
  }
  function clearRows(tbl, keep = 0) {
    kids(tbl, "tr").slice(keep).forEach(remove);
  }
  var RowFactory = class {
    constructor(templateTr) {
      this.t = templateTr;
      this.base = kids(templateTr, "tc").map(baseRpr);
    }
    make(texts, { boldFirst = false, boldNums = null, colors = null, fill = null, numAlign = "right", firstAlign = null } = {}) {
      const tr = this.t.cloneNode(true);
      kids(tr, "tc").forEach((tc, i) => {
        const base = this.base[i] || this.base.find((b) => b) || null;
        const txt = i < texts.length ? texts[i] : "";
        if (i === 0) setCell(tc, txt, base, { bold: boldFirst, color: null, align: firstAlign, fill });
        else setCell(tc, txt, base, { bold: boldNums, color: colors ? colors[i] : null, align: numAlign, fill });
      });
      setRowProps(tr);
      return tr;
    }
  };
  function findTables2(root) {
    const out = {};
    for (const tb of desc(root, "tbl")) {
      const rows = kids(tb, "tr");
      const first = rows.length ? kids(rows[0], "tc").map((tc) => ptext(tc).trim()) : [];
      const labels = rows.filter((r) => kids(r, "tc").length).map((r) => ptext(kids(r, "tc")[0]).trim());
      if (labels.slice(0, 3).includes("Metrix") || labels.slice(0, 3).includes("Metric")) out.exec = tb;
      else if (first[0] === "Country") out.country = tb;
      else if (first[0] === "Name" && labels.includes("Total Income")) out.pl = tb;
      else if (first[0] === "Name" && labels.includes("Net Assets")) out.bs = tb;
    }
    const miss = ["exec", "country", "pl", "bs"].filter((k) => !out[k]);
    if (miss.length) throw new InputError(`Template: cannot find tables ${JSON.stringify(miss)}`);
    return out;
  }
  var rowLabels = (rows) => rows.map((r) => ptext(kids(r, "tc")[0]).trim());
  function buildExec(tbl, ctx) {
    const rows = kids(tbl, "tr");
    const hdr = rows[1];
    const bases = kids(hdr, "tc").map(baseRpr);
    const k0 = ctx.k0, k1 = ctx.k1;
    let full0 = `${MONTH_NAMES[keyM(k0)]}-${keyY(k0)}`;
    if (full0.length > 11) full0 = `${MONTH_ABBR[keyM(k0)]}-${keyY(k0)}`;
    const heads = ["Metrix", full0, MONTH_NAMES[keyM(k1)], `${MONTH_ABBR[keyM(k0)]} vs ${MONTH_ABBR[keyM(k1)]}`, ctx.ytd_header];
    kids(hdr, "tc").forEach((tc, i) => setCell(tc, heads[i], bases[i], { align: "center" }));
    const { t0, t1, tytd } = ctx;
    const spec = [
      ["Revenue", "income", true],
      ["Cost of Goods Sold", "cogs", false],
      ["Operating Expense", "opex", false],
      ["Net income", "net", true]
    ];
    const BLUE = "2f75b5", REDC = "c00000", GREENC = "008000";
    spec.forEach(([label, key, goodUp], j) => {
      const tr = rows[2 + j];
      const tcs = kids(tr, "tc");
      const bs = tcs.map(baseRpr);
      const cur = t0[key], prev = t1[key], yt = tytd[key];
      const chg = Math.abs(prev) > 5e-3 ? (cur - prev) / Math.abs(prev) : 0;
      const arrow = chg > 0 ? "\u25B2" : "\u25BC";
      const improved = chg > 0 === goodUp;
      const pct = Math.abs(chg) < 1 ? `${(Math.abs(chg) * 100).toFixed(1)}%` : `${(Math.abs(chg) * 100).toFixed(0)}%`;
      const money = (x) => x < 0 ? `(${commas(x, 2)})` : commas(x, 2);
      const vals = [label, money(cur), money(prev), `${arrow}\xA0${pct}`, money(yt)];
      tcs.forEach((tc, i) => {
        let color = BLUE;
        if ([1, 2, 4].includes(i) && vals[i].startsWith("(")) color = REDC;
        if (i === 3) color = improved ? GREENC : REDC;
        setCell(tc, vals[i], bs[i], { bold: i === 0 ? true : null, color, align: "center" });
      });
      setRowProps(tr);
    });
    setRowProps(rows[0]);
    setRowProps(rows[1], false);
  }
  function buildCountry(tbl, countries, total) {
    const rows = kids(tbl, "tr");
    const head = rows[0], dataT = rows[1], blankT = rows[rows.length - 2], totalT = rows[rows.length - 1];
    const fd = new RowFactory(dataT), fb = new RowFactory(blankT), ft = new RowFactory(totalT);
    clearRows(tbl, 1);
    setRowProps(head, true);
    for (const [n, v] of countries) tbl.appendChild(fd.make([n, fmtUsd2(v), (v / total * 100).toFixed(2)], { numAlign: "right" }));
    tbl.appendChild(fb.make(["", "", ""]));
    tbl.appendChild(ft.make(["Total", fmtUsd2(total), "100.00"], { boldFirst: true }));
  }
  var negColor = (x) => x < -4e-3 ? RED_TXT : "000000";
  function buildPL(tbl, pl, kList) {
    const rows = kids(tbl, "tr");
    const head = rows[0];
    const labels = rowLabels(rows);
    const iInc = labels.indexOf("Income"), iTot = labels.indexOf("Total Cost of Goods Sold");
    const sectT = rows[iInc >= 0 ? iInc : 1], dataT = rows[2], totalT = rows[iTot >= 0 ? iTot : rows.length - 1];
    const fsec = new RowFactory(sectT), fdat = new RowFactory(dataT), ftot = new RowFactory(totalT);
    const hb = kids(head, "tc").map(baseRpr);
    const heads = ["Name", ...kList.map(shortLabel), "Total"];
    kids(head, "tc").forEach((tc, i) => setCell(tc, heads[i], hb[i]));
    setRowProps(head, true);
    clearRows(tbl, 1);
    let idx = 0;
    const fill = () => idx++ % 2 === 0 ? SHADE : null;
    const addSection = (name) => tbl.appendChild(fsec.make([name, "", "", "", ""], { boldFirst: true, fill: fill() }));
    const addData = (name, vals) => {
      const tot = vals.reduce((a, b) => a + b, 0);
      tbl.appendChild(fdat.make(
        [name, ...vals.map((v) => fmtAcc(v)), fmtAcc(tot)],
        { colors: [null, ...vals.map(negColor), negColor(tot)], fill: fill() }
      ));
    };
    const addTotal = (name, vals) => {
      const tot = vals.reduce((a, b) => a + b, 0);
      tbl.appendChild(ftot.make(
        [name, ...vals.map((v) => fmtAcc(v, false)), fmtAcc(tot, false)],
        { boldFirst: true, boldNums: true, colors: [null, ...vals.map(negColor), negColor(tot)], fill: fill() }
      ));
    };
    const valsOf = (l) => kList.map((k) => l.vals[k] ?? 0);
    const sums = {};
    for (const [section, title, totalName] of [
      ["income", "Income", "Total Income"],
      ["cogs", "Cost of Goods Sold", "Total Cost of Goods Sold"],
      ["opex", "Operating Expenses", "Total Operating Expenses"]
    ]) {
      addSection(title);
      for (const l of pl.lines) if (l.section === section && valsOf(l).some((x) => Math.abs(x) >= 5e-3)) addData(l.name, valsOf(l));
      sums[section] = kList.map((k) => pl.sect(section, k));
      addTotal(totalName, sums[section]);
      if (section === "cogs") addTotal("Gross Profit", sums.income.map((a, i) => a - sums.cogs[i]));
    }
    const net = sums.income.map((a, i) => a - sums.cogs[i] - sums.opex[i]);
    addTotal("Operating Income", net);
    addTotal("Net Income", net);
  }
  var lcSort = (a, b) => {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  };
  function buildBS(tbl, sec, tot, kList) {
    const rows = kids(tbl, "tr");
    const head = rows[0];
    const labels = rowLabels(rows);
    const i = labels.indexOf("Business Saving Account");
    const fdat = new RowFactory(rows[i >= 0 ? i : 2]);
    const hb = kids(head, "tc").map(baseRpr);
    const heads = ["Name", ...kList.map(shortLabel)];
    kids(head, "tc").forEach((tc, j) => setCell(tc, heads[j], hb[j]));
    setRowProps(head, true);
    clearRows(tbl, 1);
    let idx = 0;
    const fill = () => idx++ % 2 === 0 ? SHADE : null;
    const headRow = (name) => tbl.appendChild(fdat.make([name, "", "", ""], { boldFirst: true, fill: fill() }));
    const acctRow = (name, vals, force = false) => tbl.appendChild(fdat.make([name, ...vals.map((v) => fmtUsd0(v, !force))], { boldFirst: true, boldNums: true, fill: fill() }));
    const group = (title, g, totalName) => {
      headRow(title);
      for (const n of Object.keys(sec[g]).sort(lcSort)) acctRow(n, kList.map((k) => sec[g][n][k] ?? 0));
      acctRow(totalName, kList.map((k) => tot[k][g]), true);
    };
    headRow("Assets");
    group("Bank", "Bank", "Total Bank");
    group("Current Assets", "Current Assets", "Total Current Assets");
    group("Fixed Assets", "Fixed Assets", "Total Fixed Assets");
    acctRow("Total Assets", kList.map((k) => tot[k]["Total Assets"]), true);
    group("Current Liabilities", "Current Liabilities", "Total Current Liabilities");
    group("Non-Current Liabilities", "Non-Current Liabilities", "Total Non-Current Liabilities");
    acctRow("Net Assets", kList.map((k) => tot[k]["Net Assets"]), true);
    headRow("Equity");
    for (const n of Object.keys(sec.Equity).sort(lcSort)) acctRow(n, kList.map((k) => sec.Equity[n][k] ?? 0));
    acctRow("Total Equity", kList.map((k) => tot[k]["Total Equity"]), true);
  }
  function pictureDrawings(root, rels) {
    const out = [];
    for (const dr of desc(root, "drawing")) {
      const inl = kid(dr, "inline", NS.wp);
      if (!inl) continue;
      const blip = firstDesc(inl, "blip", NS.a);
      if (!blip) continue;
      const target = rels[blip.getAttributeNS(NS.r, "embed")] || "";
      const ext = kid(inl, "extent", NS.wp);
      const cx = +ext.getAttribute("cx"), cy = +ext.getAttribute("cy");
      if (/\.jpe?g$/i.test(target) || cx < 1e6) continue;
      out.push({ inline: inl, target, cx, cy });
    }
    return out;
  }
  function setPictureHeight(inl, cy) {
    kid(inl, "extent", NS.wp).setAttribute("cy", String(cy));
    for (const e of desc(inl, "ext", NS.a)) if (e.hasAttribute("cx")) e.setAttribute("cy", String(cy));
  }
  function replaceTexts(root, ctx) {
    const monthName = longLabel(ctx.k0);
    const ks = ctx.k_list;
    const rng = `${longLabel(ks[0])} - ${longLabel(ks[ks.length - 1])}`;
    const re1 = new RegExp(`^(?:${MONTH_RE}) 20\\d\\d$`);
    const re2 = new RegExp(`^(?:${MONTH_RE}) 20\\d\\d - (?:${MONTH_RE}) 20\\d\\d$`);
    for (const t of desc(root, "t")) {
      const s = (t.textContent || "").trim();
      if (re1.test(s)) t.textContent = monthName;
      else if (re2.test(s)) t.textContent = rng;
      else if (s.startsWith("Year-to-Date Performance") || s.startsWith("Financial Year-to-Date Performance")) t.textContent = ctx.ytd_heading;
    }
  }
  function handleLinks(root, relsRoot, ctx) {
    const relById = {};
    for (const r of Array.from(relsRoot.childNodes).filter((n) => n.nodeType === 1)) relById[r.getAttribute("Id")] = r;
    const changed = [];
    const monthName = longLabel(ctx.k0);
    for (const hl of desc(root, "hyperlink")) {
      const rid = hl.getAttributeNS(NS.r, "id");
      if (!rid) continue;
      const txt = ptext(hl).trim();
      let nw;
      if (txt === monthName) nw = ctx.links.drs;
      else if (txt.toLowerCase() === "deep") nw = ctx.links.deep;
      else if (txt.toLowerCase().startsWith("quick")) nw = ctx.links.quick;
      else continue;
      if (nw) {
        relById[rid].setAttribute("Target", nw);
        changed.push([txt, "updated"]);
      } else {
        const parent = hl.parentNode;
        for (const r of elems(hl)) {
          const rpr = kid(r, "rPr");
          if (rpr) {
            kids(rpr, "u").forEach(remove);
            if (txt === monthName) rprSet(rpr, null, "2b2b2b");
          }
          parent.insertBefore(r, hl);
        }
        parent.removeChild(hl);
        if (relById[rid]) remove(relById[rid]);
        changed.push([txt, "link removed - pointed at last month's file"]);
      }
    }
    return changed;
  }
  function pageBreakBeforeHeadings(root, labels = ["Profit and Loss (last 3 months)"]) {
    const body = kid(root, "body");
    for (const p of desc(root, "p")) {
      if (kid(p, "hyperlink") || p.parentNode !== body) continue;
      const txt = ptext(p).trim();
      if (labels.some((l) => txt.startsWith(l)) && !/\d$/.test(txt.split(")")[0])) {
        let ppr = kid(p, "pPr");
        if (!ppr) {
          ppr = make(root.ownerDocument, "pPr");
          p.insertBefore(ppr, p.firstChild);
        }
        if (!kid(ppr, "pageBreakBefore")) {
          let pos = 0;
          elems(ppr).forEach((ch, i) => {
            if (["pStyle", "keepNext", "keepLines"].includes(ch.localName)) pos = i + 1;
          });
          insertAt(ppr, pos, make(root.ownerDocument, "pageBreakBefore", { val: "1" }));
        }
      }
    }
  }
  function fixTocTab(root) {
    for (const p of desc(root, "p")) {
      if (!kid(p, "hyperlink")) continue;
      const ppr = kid(p, "pPr"), tabs = ppr && kid(ppr, "tabs");
      if (tabs) {
        for (const tab of kids(tabs, "tab")) if (wAttr(tab, "pos") === "10330") setW(tab, "pos", "10200");
      }
    }
  }
  function tocToFields(root) {
    const doc = root.ownerDocument;
    let n = 0;
    const body = kid(root, "body");
    const marks = new Set(desc(root, "bookmarkStart").map((b) => wAttr(b, "name")));
    let nextId = Math.max(100, ...desc(root, "bookmarkStart").map((b) => +wAttr(b, "id") || 0)) + 1;
    const shared = /* @__PURE__ */ new Map();
    for (const h of desc(root, "hyperlink")) {
      const a = h.getAttributeNS(NS.w, "anchor");
      if (a) shared.set(a, (shared.get(a) || 0) + 1);
    }
    const heading = (label) => elems(body).find((p) => p.localName === "p" && !kid(p, "hyperlink") && !firstDesc(p, "fldChar") && ptext(p).trim().startsWith(label));
    for (const p of desc(root, "p")) {
      if (!kid(p, "hyperlink")) continue;
      const label = kids(p, "r").map((r) => kid(r, "t") ? kid(r, "t").textContent : "").join("").trim();
      for (const hl of kids(p, "hyperlink")) {
        let anchor = hl.getAttributeNS(NS.w, "anchor");
        if (anchor && (!marks.has(anchor) || shared.get(anchor) > 1) && label) {
          const h = heading(label);
          if (h) {
            anchor = "_toc" + nextId;
            const bs = make(doc, "bookmarkStart", { id: nextId, name: anchor }), be = make(doc, "bookmarkEnd", { id: nextId });
            nextId++;
            const ppr = kid(h, "pPr");
            if (ppr) after(ppr, bs);
            else h.insertBefore(bs, h.firstChild);
            h.appendChild(be);
            marks.add(anchor);
          }
        }
        const t = anchor && desc(hl, "t")[0];
        if (!t || !/^\d+$/.test((t.textContent || "").trim())) continue;
        const run = kid(hl, "r");
        const rprOf = () => kid(run, "rPr") ? kid(run, "rPr").cloneNode(true) : null;
        const mkRun = (child) => {
          const r = make(doc, "r");
          const rp = rprOf();
          if (rp) r.appendChild(rp);
          r.appendChild(child);
          return r;
        };
        const parts = [
          mkRun(make(doc, "fldChar", { fldCharType: "begin" })),
          (() => {
            const it = make(doc, "instrText");
            it.setAttributeNS(NS.xml, "xml:space", "preserve");
            it.textContent = ` PAGEREF ${anchor} \\h `;
            return mkRun(it);
          })(),
          mkRun(make(doc, "fldChar", { fldCharType: "separate" })),
          run,
          mkRun(make(doc, "fldChar", { fldCharType: "end" }))
        ];
        for (const x of parts) hl.parentNode.insertBefore(x, hl);
        hl.parentNode.removeChild(hl);
        n++;
      }
    }
    return n;
  }
  function setToc(root, pages) {
    let n = 0;
    for (const p of desc(root, "p")) {
      if (!kid(p, "hyperlink")) continue;
      const first = kids(p, "r").map((r) => kid(r, "t") ? kid(r, "t").textContent : "").join("").trim();
      for (const [label, pg] of Object.entries(pages)) {
        if (first === label) {
          for (const hl of kids(p, "hyperlink")) {
            if (hl.getAttributeNS(NS.w, "anchor")) {
              for (const t of desc(hl, "t")) if (/^\d+$/.test((t.textContent || "").trim())) {
                t.textContent = String(pg);
                n++;
              }
            }
          }
        }
      }
    }
    return n;
  }
  var paraOf = (el) => {
    while (el && !(el.localName === "p" && el.namespaceURI === NS.w)) el = el.parentNode;
    return el;
  };
  function tightenBetween(a, b) {
    const pa = paraOf(a), pb = paraOf(b);
    let el = nextEl(pa);
    while (el && el !== pb) {
      const nxt = nextEl(el);
      if (el.localName === "p" && !ptext(el).trim() && !firstDesc(el, "drawing") && !firstDesc(el, "sectPr") && !firstDesc(el, "br") && !firstDesc(el, "bookmarkStart")) remove(el);
      el = nxt;
    }
  }
  function addNoteAfter(root, inline, text) {
    const doc = root.ownerDocument;
    const dp = paraOf(inline);
    const note = make(doc, "p");
    const r = make(doc, "r");
    note.appendChild(r);
    const rpr = make(doc, "rPr");
    r.appendChild(rpr);
    const f = make(doc, "rFonts");
    for (const a of ["ascii", "cs", "eastAsia", "hAnsi"]) setW(f, a, "Arial");
    rpr.appendChild(f);
    rpr.appendChild(make(doc, "i"));
    rpr.appendChild(make(doc, "color", { val: "5a5a5a" }));
    rpr.appendChild(make(doc, "sz", { val: "18" }));
    const t = make(doc, "t");
    t.textContent = text;
    t.setAttributeNS(NS.xml, "xml:space", "preserve");
    r.appendChild(t);
    if (firstDesc(dp, "sectPr")) before(dp, note);
    else after(dp, note);
  }
  async function build(priorBytes, ctx, images, tocPages = null) {
    const zin = await globalThis.JSZip.loadAsync(priorBytes);
    const order = Object.keys(zin.files);
    const files = {};
    for (const n of order) if (!zin.files[n].dir) files[n] = await zin.files[n].async("uint8array");
    const dec = (u8) => new TextDecoder().decode(u8);
    const docXml = parseXml(dec(files["word/document.xml"]));
    const relsXml = parseXml(dec(files["word/_rels/document.xml.rels"]));
    const root = docXml.documentElement;
    const rels = {};
    for (const r of elems(relsXml.documentElement)) rels[r.getAttribute("Id")] = r.getAttribute("Target");
    const T = findTables2(root);
    buildExec(T.exec, ctx);
    buildCountry(T.country, ctx.countries, ctx.country_total);
    buildPL(T.pl, ctx.pl, ctx.k_list);
    buildBS(T.bs, ctx.sec, ctx.tot, ctx.k_list);
    replaceTexts(root, ctx);
    const linkLog = handleLinks(root, relsXml.documentElement, ctx);
    pageBreakBeforeHeadings(root);
    fixTocTab(root);
    if (tocPages) setToc(root, tocPages);
    tocToFields(root);
    for (const p of desc(root, "p")) {
      const t = ptext(p).trim();
      if ((t.startsWith("Open item: product-level") || t.startsWith("Provisional: product-level")) && !firstDesc(p, "drawing")) remove(p);
    }
    const pics = pictureDrawings(root, rels);
    const kinds = ctx.picture_kinds;
    if (pics.length !== kinds.length) throw new InputError(`Template has ${pics.length} chart pictures, expected ${kinds.length}: ${kinds.join(", ")}`);
    const notes = [];
    for (let i = 0; i < pics.length; i++) {
      const pic = pics[i], kind = kinds[i];
      const wIn = pic.cx / 914400, hIn = pic.cy / 914400;
      const img = await images[kind](wIn, hIn);
      setPictureHeight(pic.inline, Math.trunc(pic.cx * img.h / img.w));
      const path = ("word/" + pic.target).replace(/\\/g, "/");
      if (!path.endsWith(".png")) throw new InputError(`Picture ${path} is not a PNG`);
      files[path] = img.png;
      if (img.note) notes.push([pic.inline, img.note]);
    }
    for (const [inl, text] of notes) addNoteAfter(root, inl, text);
    tightenBetween(pics[pics.length - 2].inline, pics[pics.length - 1].inline);
    if (files["word/settings.xml"]) {
      const st = parseXml(dec(files["word/settings.xml"]));
      const sr = st.documentElement;
      if (!kid(sr, "updateFields")) {
        const anchor = elems(sr).find((e) => ["hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "themeFontLang", "clrSchemeMapping", "decimalSymbol", "listSeparator"].includes(e.localName));
        const uf = make(st, "updateFields", { val: "true" });
        if (anchor) sr.insertBefore(uf, anchor);
        else sr.appendChild(uf);
      }
      files["word/settings.xml"] = new TextEncoder().encode(serializeXml(st));
    }
    files["word/document.xml"] = new TextEncoder().encode(serializeXml(docXml));
    files["word/_rels/document.xml.rels"] = new TextEncoder().encode(serializeXml(relsXml));
    const zo = new globalThis.JSZip();
    for (const n of order) if (n in files) zo.file(n, files[n]);
    const bytes2 = await zo.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { bytes: bytes2, linkLog };
  }

  // src/reports.js
  function need(obj, lab, k) {
    const v = obj[lab] && obj[lab][k];
    if (v === void 0) throw new Error(`The finished report has no '${lab}' figure for ${k}`);
    return v;
  }
  async function readback(docxBytes, pl, sec, tot, kList, ctx) {
    const t = findTables(await loadDocXml(docxBytes));
    const [pld] = tableByMonth(t.pl);
    const [bsd] = tableByMonth(t.bs);
    const C = [];
    for (const k of kList) {
      const tt = pl.totals(k);
      for (const [lab, v] of Object.entries(plLabelTotals(tt))) C.push(chk("Report vs Xero P&L", `${lab} ${monthLabel(k)}`, v, need(pld, lab, k), { gate: true }));
      for (const l of pl.lines) {
        const v = l.vals[k] ?? 0;
        if (Math.abs(v) >= 5e-3 || l.name in pld) {
          const got = (pld[l.name] || {})[k] ?? 0;
          if (Math.abs(got - v) > TOL_EXACT) C.push(chk("Report vs Xero P&L", `${l.name} ${monthLabel(k)}`, v, got, { gate: true }));
        }
      }
    }
    for (const lab of ["Total Income", "Net Income"]) {
      const row = t.pl.find((r) => r.length && r[0] === lab);
      const want = sum(kList.map((k) => pl.totals(k)[lab === "Total Income" ? "income" : "net"]));
      C.push(chk("Report vs Xero P&L", `${lab} - 3 month total column`, want, parseAmount(row[row.length - 1]), { gate: true }));
    }
    const keyed = {
      "Total Bank": "Bank",
      "Total Assets": "Total Assets",
      "Net Assets": "Net Assets",
      "Total Equity": "Total Equity",
      "Total Current Liabilities": "Current Liabilities",
      "Total Non-Current Liabilities": "Non-Current Liabilities"
    };
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
  var argb = (hex) => "FF" + hex;
  var FILL = { PASS: "D9EAD3", WARN: "FFF2CC", FAIL: "F4CCCC", INFO: "EDEDED" };
  var solid = (hex) => ({ type: "pattern", pattern: "solid", fgColor: { argb: argb(hex) } });
  async function writeWorkbook(monthName, pl, tot, kList, res, rb, ctx) {
    const wb = new globalThis.ExcelJS.Workbook();
    const NAVY2 = "1F3A5F";
    const head = (ws2, row, vals) => vals.forEach((v, i) => {
      const c = ws2.getCell(row, i + 1);
      c.value = v;
      c.fill = solid(NAVY2);
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.alignment = { wrapText: true, vertical: "middle" };
    });
    const widths = (ws2, w) => w.forEach((x, i) => {
      ws2.getColumn(i + 1).width = x;
    });
    let ws = wb.addWorksheet("Tie-out");
    ws.getCell("A1").value = `Tie-out to Xero Profit and Loss - ${monthName}`;
    ws.getCell("A1").font = { bold: true, size: 13, color: { argb: argb(NAVY2) } };
    ws.getCell("A2").value = "Xero = line items of the Xero export, re-added. Report = figures read back out of the finished Word file.";
    let r = 4;
    const fm = "#,##0.00;[Red](#,##0.00)";
    for (const [key, lab] of [[ctx.k0, "Report month"], [ctx.k1, "Comparison month"], ["YTD", "Year to date"]]) {
      const t = key === ctx.k0 ? ctx.t0 : key === ctx.k1 ? ctx.t1 : ctx.tytd;
      ws.getCell(r, 1).value = `${lab}: ${key !== "YTD" ? monthLabel(key) : ctx.ytd_header}`;
      ws.getCell(r, 1).font = { bold: true };
      r++;
      head(ws, r, ["Line", "Xero P&L", "Report", "Difference"]);
      r++;
      for (const [nm, k2] of [["Income", "income"], ["Cost of sales", "cogs"], ["Gross profit", "gp"], ["Operating expenses", "opex"], ["Net result", "net"]]) {
        ws.getCell(r, 1).value = nm;
        ws.getCell(r, 2).value = t[k2];
        const lookup = { income: "Total Income", cogs: "Total Cost of Goods Sold", gp: "Gross Profit", opex: "Total Operating Expenses", net: "Net Income" }[k2];
        let m;
        if (key !== "YTD") m = rb.find((c) => c.name === `${lookup} ${monthLabel(key)}`);
        else {
          const lab2 = { income: "Revenue", cogs: "Cost of Goods Sold", opex: "Operating Expense", net: "Net income" }[k2];
          m = lab2 ? rb.find((c) => c.name === `Executive summary ${lab2} (YTD)`) : null;
        }
        if (m) {
          ws.getCell(r, 3).value = m.actual;
          ws.getCell(r, 4).value = m.diff;
          ws.getCell(r, 4).fill = solid(FILL[m.status]);
        } else ws.getCell(r, 3).value = "n/a (not shown in report)";
        [2, 3, 4].forEach((c) => {
          ws.getCell(r, c).numFmt = fm;
        });
        r++;
      }
      r++;
    }
    widths(ws, [34, 18, 18, 14]);
    ws = wb.addWorksheet("Checks");
    head(ws, 1, ["Group", "Check", "Expected", "Actual", "Difference", "Status", "Note"]);
    [...res.checks, ...rb].forEach((c, i) => {
      const row = i + 2;
      [c.group, c.name, c.expected, c.actual, c.diff, c.status, c.note].forEach((v, j) => {
        ws.getCell(row, j + 1).value = v;
      });
      ws.getCell(row, 6).fill = solid(FILL[c.status]);
      [3, 4, 5].forEach((j) => {
        ws.getCell(row, j).numFmt = "#,##0.00;[Red]-#,##0.00";
      });
    });
    widths(ws, [20, 62, 16, 16, 14, 9, 50]);
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws = wb.addWorksheet("Restated vs published");
    ws.getCell("A1").value = "Figures in this month's Xero export compared with what the previous report published";
    ws.getCell("A1").font = { bold: true, color: { argb: argb(NAVY2) } };
    head(ws, 3, ["Statement", "Month", "Item", "Published", "Now in Xero", "Change"]);
    let rr = 4;
    for (const [blk, lab] of [["totals", "P&L totals"], ["lines", "P&L lines (change >= A$1)"], ["bs", "Balance sheet"]]) {
      for (const x of res.restate[blk]) {
        [lab, monthLabel(x.month), x.item, x.published, x.now, x.diff].forEach((v, j) => {
          ws.getCell(rr, j + 1).value = v;
        });
        [4, 5, 6].forEach((j) => {
          ws.getCell(rr, j).numFmt = "#,##0.00;[Red]-#,##0.00";
        });
        rr++;
      }
    }
    widths(ws, [26, 12, 46, 16, 16, 16]);
    ws.views = [{ state: "frozen", ySplit: 3 }];
    ws = wb.addWorksheet("New accounts");
    head(ws, 1, ["Ledger", "Account", "Group", `${monthName} amount`, "Note"]);
    const add = (...v) => {
      ws.addRow(v);
    };
    for (const x of res.new_pl) add("P&L", x.account, x.section, x.this, x.first_time ? "first time in the 4-month window" : "was nil last month");
    for (const x of res.new_bs) add("Balance sheet", x.account, x.group, x.this, "was nil last month");
    for (const n of res.bs_not_on_prior_report) add("Balance sheet", n, "", null, "not on the previous published report");
    for (const n of res.orphan_accruals) add("Balance sheet", n, "", null, "accrual with no matching 'Social Media Ads - ...' P&L account");
    widths(ws, [16, 48, 26, 18, 60]);
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }
  function callpoints(monthName, prevName, res, ctx, pl, sec, tot, extra) {
    const { k0, k1, t0, t1 } = ctx;
    const ad = (k) => pl.lines.filter((l) => l.name.startsWith("Social Media Ads")).reduce((s, l) => s + (l.vals[k] ?? 0), 0);
    const a0 = ad(k0), a1 = ad(k1);
    const pct = (a, b) => b ? (a - b) / Math.abs(b) * 100 : 0;
    const A2 = (x) => x >= 0 ? `A$${commas(x, 0)}` : `-A$${commas(x, 0)}`;
    const sg = (x) => (x < 0 ? "-" : "") + commas(x, 0);
    const p0 = (x) => `${Math.round(x * 100)}%`;
    const cb = (g, n2, k) => (sec[g][n2] || {})[k] ?? 0;
    const lines = [`CALL POINTS - Saint Eyewear month-end ${monthName}`, "(internal - do not send to the client)", ""];
    lines.push("CLIENT QUESTIONS (Taf)");
    let q = 1;
    lines.push(`${q}. Result. ${monthName} is a net ${t0.net < 0 ? "loss" : "profit"} of ${A2(Math.abs(t0.net))} against a net ${t1.net < 0 ? "loss" : "profit"} of ${A2(Math.abs(t1.net))} in ${prevName}. Revenue fell ${Math.abs(pct(t0.income, t1.income)).toFixed(0)}% (${A2(t1.income)} to ${A2(t0.income)}) but social media ads fell only ${Math.abs(pct(a0, a1)).toFixed(0)}% (${A2(a1)} to ${A2(a0)}). Ads were ${p0(a0 / t0.income)} of revenue, up from ${p0(a1 / t1.income)}. What is the plan for ad budget while sales are falling?`);
    q++;
    const cu0 = cb("Bank", "CommBank USD", k0), cu1 = cb("Bank", "CommBank USD", k1);
    if (Math.abs(cu0 - cu1) > 1e4) {
      lines.push(`${q}. CommBank USD went from ${A2(cu1)} to ${A2(cu0)} (${A2(cu0 - cu1)}). What was moved or paid out of it?`);
      q++;
    }
    const d0 = cb("Equity", "Owner A Drawings", k0), d1 = cb("Equity", "Owner A Drawings", k1);
    if (Math.abs(d0 - d1) > 0.5) {
      lines.push(`${q}. Owner A Drawings ${d0 < d1 ? "increased" : "decreased"} by ${A2(Math.abs(d0 - d1))} in the month (balance now ${A2(d0)} drawn). Please confirm.`);
      q++;
    }
    const newx = res.new_pl.filter((x) => x.section === "opex" && x.first_time && Math.abs(x.this) > 100);
    if (newx.length) {
      lines.push(`${q}. New expense accounts: ` + newx.map((x) => `${x.account} ${A2(x.this)}`).join("; ") + ". What are they, and are they all business costs?");
      q++;
    }
    const tb0 = tot[k0].Bank, tb1 = tot[k1].Bank;
    lines.push(`${q}. Total bank balances fell ${A2(tb1 - tb0)} (${A2(tb1)} to ${A2(tb0)}) after the net loss and drawings. Anything unusual we should know about?`);
    q++;
    if ("Trust Savings commonwealth" in sec.Bank) {
      lines.push(`${q}. Trust Savings (CommBank) holds ${A2(cb("Bank", "Trust Savings commonwealth", k0))}. Please confirm this account, and the Slash and Trust accounts, belong in Xero (open item from SOP v1.0).`);
      q++;
    }
    lines.push("", "INTERNAL ITEMS");
    let n = 1;
    const rt = new Map(res.restate.totals.map((x) => [x.month + "|" + x.item, x]));
    if (rt.size) {
      const bits = [];
      for (const k of [k1, prevKey(k0, 2)]) {
        const x = rt.get(k + "|Net Income");
        if (x) bits.push(`${monthLabel(k)} net result ${A2(x.published)} published, ${A2(x.now)} now (${A2(x.diff)})`);
      }
      const big = res.restate.lines.filter((x) => x.month === k1).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 4);
      const bigBs = res.restate.bs.filter((y) => y.month === k1 && !/^(Total|Net)/.test(y.item)).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 4);
      lines.push(`${n}. RESTATED COMPARATIVES. The Xero export no longer agrees with the pack that went out last month: ` + bits.join("; ") + ". Biggest lines for " + monthLabel(k1) + ": " + big.map((x) => `${x.item} ${A2(x.published)} -> ${A2(x.now)}`).join("; ") + ". Balance sheet comparatives moved too (biggest: " + bigBs.map((x) => `${x.item} ${sg(x.published)} -> ${sg(x.now)}`).join("; ") + "). Find out what was posted back into closed months, and whether the client has been told. Detail in the tie-out workbook, sheet 'Restated vs published'.");
      n++;
    }
    if (!res.info.product_ok) {
      lines.push(`${n}. PRODUCT SALES NOT SHOWN. The DRS 'Productwise Sales' tab totals ${A2(extra.product_gross)} gross, ${p0(res.info.product_coverage)} of the ${A2(extra.drs_gross)} gross sales in the Daily Reporting tab. ` + (extra.product_stale && extra.product_stale.length ? `The tab's rows are dated ${extra.product_stale.join(", ")} - it still holds old data and was not refreshed for the month. ` : "") + `Section left as an open item in the report. Need a full-month Shopify product export pasted into that tab.`);
      n++;
    }
    lines.push(`${n}. YEAR-TO-DATE BASIS. ${extra.ytd_note}`);
    n++;
    if (extra.links.length) {
      lines.push(`${n}. LINKS. ` + extra.links.map(([a, b]) => `'${a}' ${b}`).join("; ") + ". Add the new links once the DRS / Deep / Quick View are posted.");
      n++;
    }
    if (res.new_pl.length) {
      const fresh = res.new_pl.filter((x) => x.first_time);
      lines.push(`${n}. NEW P&L ACCOUNTS (first time in the window): ` + fresh.map((x) => `${x.account} ${A2(x.this)}`).join("; ") + ". Confirm each is P&L, not balance sheet.");
      n++;
    }
    if (res.orphan_accruals.length) {
      lines.push(`${n}. ACCRUAL WITHOUT P&L ACCOUNT: ` + res.orphan_accruals.join("; ") + " has a balance but there is no matching 'Social Media Ads - ...' P&L account. Where was the cost booked?");
      n++;
    }
    const cp = pl.acct("COGS_Processing fee", k0);
    if (Math.abs(cp) > 0.5 && res.new_pl.some((x) => x.account === "COGS_Processing fee")) {
      lines.push(`${n}. 'COGS_Processing fee' is a new Cost of Sales account (${A2(cp)}). DRS processing fee for the month is ${A2(extra.drs_processing)}. Confirm what it is and why it is in Cost of Sales.`);
      n++;
    }
    lines.push(`${n}. COMMBANK USD is a manual account: reconcile to the bank statement (none supplied). Balance ${A2(cu0)} (was ${A2(cu1)}).`);
    n++;
    const s0 = cb("Current Liabilities", "Suspense", k0), s1 = cb("Current Liabilities", "Suspense", k1);
    lines.push(`${n}. SUSPENSE moved ${A2(Math.abs(s0 - s1))} ${s0 < s1 ? "more negative" : "less negative"}, balance ${A2(s0)}. Check what went in.`);
    n++;
    const ap = cb("Current Liabilities", "Accounts Payable", k0);
    if (ap < 0) {
      lines.push(`${n}. Accounts Payable is a debit balance of ${A2(ap)} - check for overpaid or unapplied supplier payments.`);
      n++;
    }
    const warns = res.checks.filter((c) => c.status === "WARN" && !c.name.includes("Productwise"));
    if (warns.length) {
      lines.push(`${n}. DRS vs XERO differences above tolerance: ` + warns.map((c) => `${c.name} ${c.diff >= 0 ? "+" : "-"}${commas(c.diff, 2)}`).join("; ") + ".");
      n++;
    }
    lines.push(`${n}. No General Ledger Detail was supplied. It is not needed for the month-end report, only for the expense pack.`);
    n++;
    return lines.join("\n") + "\n";
  }
  function reconTable(rb, ctx) {
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
  async function narrativeLines(docxBytes) {
    const doc = await loadDocXml(docxBytes);
    const body = doc.documentElement.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "body")[0];
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const out = [];
    for (const p of Array.from(body.childNodes)) {
      if (p.nodeType !== 1 || p.localName !== "p") continue;
      if (p.getElementsByTagNameNS(W, "fldChar").length || p.getElementsByTagNameNS(W, "hyperlink").length) continue;
      const t = Array.from(p.getElementsByTagNameNS(W, "t")).map((x) => x.textContent).join("").trim();
      if (t && /\d/.test(t)) out.push(t);
    }
    return out;
  }

  // src/build.js
  var KINDS = ["ytd", "fees", "expenditure", "daysales", "daycogs", "cogspie", "countrypie", "product_table", "product_chart"];
  async function sniff(name, bytes2) {
    const low = name.toLowerCase();
    if (low.endsWith(".docx")) {
      try {
        findTables(await loadDocXml(bytes2));
        return "report";
      } catch (e) {
        return null;
      }
    }
    if (!low.endsWith(".xlsx") || low.startsWith("~$")) return null;
    try {
      const wb = await openBook(bytes2);
      if (wb.names.includes("Daily Reporting")) return "drs";
      const a1 = String(wb.first().v(1, 1) || "").toLowerCase();
      if (a1.startsWith("profit")) return "pl";
      if (a1.startsWith("balance")) return "bs";
    } catch (e) {
    }
    return null;
  }
  var FMT_D = (d) => `${d.d} ${MONTH_ABBR[d.m]} ${d.y}`;
  async function describeInput(role, bytes2) {
    const u8 = bytes2 instanceof Uint8Array ? bytes2 : new Uint8Array(bytes2);
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
      const need2 = ["Daily Reporting", "Countrywise Sales", "Productwise Sales", "Fees Calculation"].filter((n) => !wb.names.includes(n));
      if (need2.length) return `Missing tabs: ${need2.join(", ")}.`;
      let dates = [];
      for (let r = 1; r < 40 && !dates.length; r++) {
        if (String(ws.v(r, 1) || "").toLowerCase().includes("date")) for (let c = 2; c <= ws.maxCol; c++) {
          const v = ws.v(r, c);
          if (v instanceof Date) dates.push(v);
        }
      }
      return dates.length ? `${dates.length} days: ${FMT_D({ d: dates[0].getUTCDate(), m: dates[0].getUTCMonth() + 1, y: dates[0].getUTCFullYear() })} to ${FMT_D({ d: dates[dates.length - 1].getUTCDate(), m: dates[dates.length - 1].getUTCMonth() + 1, y: dates[dates.length - 1].getUTCFullYear() })}.` : "Tabs found.";
    }
    if (role === "report") {
      const pr = await readPrior(u8);
      return `Covers ${pr.pl_keys.map(monthLabel).join(", ")}. Used as the template.`;
    }
    return "";
  }
  async function plMonth(bytes2) {
    const p = await readPL(bytes2 instanceof Uint8Array ? bytes2 : new Uint8Array(bytes2));
    return mk(p.periodEnd.y, p.periodEnd.m);
  }
  var brief = (cs) => cs.map((c) => ({ name: c.name, expected: c.expected, actual: c.actual, diff: c.diff, note: c.note }));
  var MIME = {
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    txt: "text/plain"
  };
  async function buildReport(inputs2, opts = {}) {
    const log = [];
    const say = (s) => log.push(s);
    const u8 = (x) => x instanceof Uint8Array ? x : new Uint8Array(x);
    const plR = await readPL(u8(inputs2.pl));
    const k0 = mk(plR.periodEnd.y, plR.periodEnd.m), k1 = prevKey(k0, 1), k2 = prevKey(k0, 2);
    const monthName = longLabel(k0), prevName = longLabel(k1);
    const bsR = await readBS(u8(inputs2.bs));
    const { sec, tot } = bsPresent(bsR);
    const d = await readDRS(u8(inputs2.drs), plR.periodEnd);
    const prv = await readPrior(u8(inputs2.report));
    const priorLast = prv.pl_keys[prv.pl_keys.length - 1];
    if (priorLast !== k1 && !opts.skipMonthGuard) {
      throw new InputError(`Last month's report ends ${longLabel(priorLast)} but this P&L is for ${monthName}; expected the ${prevName} report.`);
    }
    const res = runChecks(plR, bsR, d, prv, k0, sec, tot);
    const fails = res.checks.filter((c) => c.status === "FAIL"), warns = res.checks.filter((c) => c.status === "WARN");
    say(`Tie-out gate (inputs): ${res.checks.length - fails.length - warns.length} pass, ${warns.length} warn, ${fails.length} FAIL`);
    const cov = res.info.product_coverage;
    const gate = {
      month: monthName,
      prev_month: prevName,
      checks: res.checks,
      fails: brief(fails),
      warns: brief(warns),
      info: res.info,
      new_pl: res.new_pl,
      new_bs: res.new_bs,
      bs_not_on_prior_report: res.bs_not_on_prior_report,
      orphan_accruals: res.orphan_accruals,
      product: { ok: res.info.product_ok, coverage: cov, stale: res.info.product_stale, product_gross: d.product_total[0], drs_gross: sum(d.daily.gross) },
      periods: { pl: plR.keys, bs: bsR.keys, prior_last: priorLast },
      totals: { [k0]: plR.totals(k0), [k1]: plR.totals(k1) }
    };
    if (fails.length) {
      return { status: "gate_failed", summary: { status: "gate_failed", month: monthName, fails: brief(fails), warns: brief(warns) }, gate, files: [], log };
    }
    if (opts.gateOnly) return { status: "gate_ok", summary: { status: "gate_ok", month: monthName, fails: [], warns: brief(warns) }, gate, files: [], log };
    const t0 = plR.totals(k0), t1 = plR.totals(k1);
    const fm = fyMonths(k0);
    let tytd, ytdHeader, ytdLabel, ytdHeading, ytdNote;
    const ab = MONTH_ABBR;
    if (inputs2.ytd) {
      const [tt, ytxt] = await readPLPeriod(u8(inputs2.ytd));
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
      ytdNote = `Previous packs showed calendar 'YTD Jan-...'. The files supplied only support Xero's financial-year-to-date (${s}-${e}). Shown as 'FYTD' and labelled so. For calendar YTD add a Xero P&L for 1 Jan - ${daysIn(keyY(k0), keyM(k0))} ${MONTH_NAMES[keyM(k0)]} ${keyY(k0)} under "Optional".`;
    }
    const links = opts.links || {};
    const ctx = {
      k0,
      k1,
      k_list: [k2, k1, k0],
      t0,
      t1,
      tytd,
      ytd_header: ytdHeader,
      ytd_heading: ytdHeading,
      countries: d.countries,
      country_total: d.country_total,
      pl: plR,
      sec,
      tot,
      picture_kinds: KINDS,
      links: { drs: links.drs || null, deep: links.deep || null, quick: links.quick || null }
    };
    const D = d.daily;
    let costParts = plR.lines.filter((l) => l.section === "cogs" && (l.vals[k0] ?? 0) > 5e-3).map((l) => [l.name.split("COGS_").join("").replace("Cost of Goods Sold", "Cost of goods sold"), l.vals[k0]]);
    costParts = costParts.map(([n, v]) => [n.charAt(0).toUpperCase() + n.slice(1), v]);
    const runItems = plR.lines.filter((l) => l.section === "opex" && Math.abs(l.vals[k0] ?? 0) >= 5e-3 && !EXPENDITURE_EXCLUDE_PREFIXES.some((p) => l.name.toLowerCase().startsWith(p))).map((l) => [l.name, l.vals[k0]]);
    const adsTotal = sum(plR.lines.filter((l) => l.name.toLowerCase().startsWith("social media ads")).map((l) => l.vals[k0] ?? 0));
    const feeByGw = {};
    for (const [gw, acct] of Object.entries(FEE_ACCOUNTS)) if (Math.abs(plR.acct(acct, k0)) > 0) feeByGw[gw] = -plR.acct(acct, k0);
    const showProducts = res.info.product_ok || !!opts.allowPartialProducts;
    const prov = res.info.product_ok ? "" : " (PROVISIONAL)";
    const covTxt = `${Math.round(res.info.product_coverage * 100)}%`;
    const stale = res.info.product_stale || [];
    const openNote = stale.length ? `Open item: product-level sales for ${monthName} are not shown. The product sales data supplied is out of date (dated ${stale.join(", ")}) and does not cover the month. The full table will follow once the current month's export is received.` : `Open item: product-level sales for ${monthName} are not shown. The product export supplied covers A$${Math.round(d.product_total[0]).toLocaleString("en-US")} of A$${Math.round(sum(D.gross)).toLocaleString("en-US")} gross sales for the month (${covTxt}). The full table will follow once the complete export is received.`;
    const phText = stale.length ? "Current-month product export outstanding" : `Full-month product export outstanding (${covTxt} of gross sales supplied)`;
    const images = {
      ytd: (w, h) => ytd(w, h, ytdLabel, { Revenue: tytd.income, COGS: tytd.cogs, "Operating Expenses": tytd.opex, "Net Income": tytd.net }),
      fees: (w, h) => fees(w, h, monthName, feeByGw),
      expenditure: (w, h) => expenditure(w, h, runItems, adsTotal),
      daysales: (w, h) => daywiseSales(w, h, d.dates, D.orders, D.returns, D.ship_collected),
      daycogs: (w, h) => daywiseCogs(w, h, d.dates, D.cogs_product, D.cogs_ship),
      cogspie: (w, h) => costPie(w, h, costParts),
      countrypie: (w, h) => countryPie(w, h, d.countries, monthName)
    };
    if (showProducts) {
      images.product_table = async (w) => {
        const r = await productTable(w, d.products, d.product_total);
        if (prov) r.note = openNote.replace("Open item: product-level sales", "Provisional: product-level sales").replace("are not shown", "are shown from a partial export");
        return r;
      };
      images.product_chart = (w, h) => productChart(w, h, d.products);
    } else {
      images.product_table = async (w) => {
        const r = await placeholder(w, 0.9, "Product-wise sales: open item", phText);
        r.note = openNote;
        return r;
      };
      images.product_chart = (w) => placeholder(w, 0.55, "Gross, discounts, refunds and net sales by product", "To follow with the product table");
    }
    const { bytes: docx, linkLog } = await build(u8(inputs2.report), ctx, images);
    const notes = ["No PDF is made in the browser. Word will ask to update fields when you open the file - click Yes so the contents page numbers refresh - then read every page (SOP step 4)."];
    const rb = await readback(docx, plR, sec, tot, [k2, k1, k0], ctx);
    const rbf = rb.filter((c) => c.status === "FAIL");
    say(`Read-back of the finished report vs Xero: ${rb.length - rbf.length} agree, ${rbf.length} FAIL`);
    const extra = { product_gross: d.product_total[0], drs_gross: sum(D.gross), ytd_note: ytdNote, links: linkLog, drs_processing: sum(D.processing), product_stale: stale };
    const mname = `${MONTH_NAMES[keyM(k0)]}-${keyY(k0)}`;
    const tieout = await writeWorkbook(monthName, plR, tot, [k2, k1, k0], res, rb, ctx);
    const cp = callpoints(monthName, prevName, res, ctx, plR, sec, tot, extra);
    const names = { docx: `Month-End_Financial_Report_${mname}.docx`, tieout: `Tie-out_${mname}.xlsx`, callpoints: `Call_points_${mname}.txt` };
    const files = [
      { name: names.docx, mime: MIME.docx, bytes: docx },
      { name: names.tieout, mime: MIME.xlsx, bytes: tieout },
      { name: names.callpoints, mime: MIME.txt, bytes: new TextEncoder().encode(cp) }
    ];
    const recon = reconTable(rb, ctx);
    const narrative = await narrativeLines(docx);
    const summary = {
      status: rbf.length ? "readback_failed" : "ok",
      recon,
      narrative,
      month: monthName,
      n_checks: res.checks.length + rb.length,
      fails: [],
      warns: brief(warns),
      readback_ok: rb.length - rbf.length,
      readback_fail: brief(rbf),
      notes,
      links_removed: linkLog.map(([a, b]) => `${a} (${b})`),
      pdf: false,
      product_ok: res.info.product_ok,
      files: names
    };
    return { status: summary.status, summary, gate, files: rbf.length ? [] : files, log, ctx };
  }

  // src/app.js
  var $ = (s) => document.querySelector(s);
  var esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
  var fmt = (v) => v == null ? "" : Number(v).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  var ROLES = [
    ["pl", "Profit and Loss", "Xero .xlsx"],
    ["bs", "Balance Sheet", "Xero .xlsx"],
    ["drs", "Daily Reporting Sheet", ".xlsx"],
    ["report", "Last month's report", ".docx (template)"]
  ];
  var picked = {};
  var ytdFile = null;
  var urls = [];
  var gateOk = false;
  var stored = null;
  var bytes = async (f) => new Uint8Array(await f.arrayBuffer());
  var prevMonth = (k) => {
    let y = +k.slice(0, 4), m = +k.slice(5) - 1;
    if (m < 1) {
      m = 12;
      y--;
    }
    return `${y}-${String(m).padStart(2, "0")}`;
  };
  var idb = () => new Promise((res, rej) => {
    const r = indexedDB.open("saint-reports", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("reports");
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  async function idbGet(key) {
    try {
      const db = await idb();
      return await new Promise((res) => {
        const q = db.transaction("reports").objectStore("reports").get(key);
        q.onsuccess = () => res(q.result || null);
        q.onerror = () => res(null);
      });
    } catch {
      return null;
    }
  }
  async function idbPut(key, val) {
    try {
      const db = await idb();
      await new Promise((res) => {
        const t = db.transaction("reports", "readwrite");
        t.objectStore("reports").put(val, key);
        t.oncomplete = res;
        t.onerror = res;
      });
    } catch {
    }
  }
  async function refreshStored() {
    stored = null;
    if (picked.pl && !picked.report) {
      try {
        const k = prevMonth(await plMonth(picked.pl.bytes));
        const b = await idbGet(k);
        if (b) stored = { key: k, bytes: b };
      } catch {
      }
    }
    resetGate();
    renderInputs();
  }
  function renderInputs() {
    $("#inputs").innerHTML = ROLES.map(([r, name, hint]) => {
      const p = picked[r];
      let state;
      if (p) state = `<span class="ok">${esc(p.name)}</span>`;
      else if (r === "report" && stored) state = `<span class="ok">Remembered (${esc(stored.key)})</span>`;
      else state = `<span class="no">not supplied</span>`;
      const ok = !!p || r === "report" && !!stored;
      return `<tr><td class="st ${ok ? "ok" : "no"}">${ok ? "&#10003;" : "&#9675;"}</td><td><b>${name}</b> <span class="muted">${hint}</span></td><td>${state}</td></tr>`;
    }).join("");
    const missing = ROLES.filter(([r]) => !(picked[r] || r === "report" && stored)).map(([, n]) => n);
    $("#gate").disabled = missing.length > 0;
    $("#missing").innerHTML = missing.length ? `<div class="banner warn"><b>Still needed: ${missing.map(esc).join("; ")}</b>${missing.some((n) => /^Last/.test(n)) ? "Drop last month's finished report once; after a successful run it is remembered." : ""}</div>` : "";
  }
  function resetGate() {
    gateOk = false;
    $("#go").disabled = true;
    $("#gateOut").innerHTML = "";
    $("#buildOut").innerHTML = "";
    $("#after").hidden = true;
  }
  async function addFiles(list) {
    const msgs = [];
    for (const f of list) {
      const b = await bytes(f);
      const role = await sniff(f.name, b);
      if (!role) {
        msgs.push(`${f.name}: not recognised. Use the Xero Profit and Loss, Xero Balance Sheet, Daily Reporting Sheet (.xlsx) or a finished Month-End Financial Report (.docx).`);
        continue;
      }
      let info = "";
      try {
        info = await describeInput(role, b);
      } catch (e) {
        info = "Could not read: " + (e.message || e);
      }
      picked[role] = { name: f.name, bytes: b, info };
    }
    $("#problems").innerHTML = msgs.map((m) => `<div class="banner bad">${esc(m)}</div>`).join("");
    await refreshStored();
  }
  var reportBytes = () => picked.report ? picked.report.bytes : stored.bytes;
  var inputs = () => ({ pl: picked.pl.bytes, bs: picked.bs.bytes, drs: picked.drs.bytes, report: reportBytes(), ytd: ytdFile });
  var options = () => ({ links: { drs: $("#drsLink").value.trim(), deep: $("#deepLink").value.trim(), quick: $("#quickLink").value.trim() } });
  var linksOk = () => ["drsLink", "deepLink", "quickLink"].every((k) => {
    const v = $("#" + k).value.trim();
    return !v || /^https:\/\/\S+$/.test(v);
  });
  var busy = (btn, text) => {
    btn.disabled = true;
    btn.dataset.t = btn.textContent;
    btn.textContent = text;
    return new Promise((r) => setTimeout(r, 30));
  };
  var unbusy = (btn) => {
    btn.textContent = btn.dataset.t;
  };
  function statusCell(c) {
    return `<td class="${c.status === "PASS" ? "pass" : c.status === "FAIL" ? "fail" : "warn"}">${c.status === "PASS" ? "Agrees" : c.status === "FAIL" ? "Differs" : c.status === "WARN" ? "Look" : "Info"}</td>`;
  }
  function gateHtml(g, status) {
    const gateChecks = g.checks.filter((c) => c.group !== "DRS vs P&L");
    const drsChecks = g.checks.filter((c) => c.group === "DRS vs P&L");
    const bad = gateChecks.filter((c) => c.status === "FAIL");
    let h = "";
    if (status === "gate_ok") h += `<div class="banner"><b>Gate passed</b>All ${gateChecks.length} checks agree to the cent.</div>`;
    else h += `<div class="banner bad"><b>Stopped: ${bad.length} difference${bad.length === 1 ? "" : "s"}</b>Nothing is built. Chase each one, re-export from Xero on the same day and run again.</div><table class="t"><tr><th>Check</th><th class="n">Expected</th><th class="n">Actual</th><th class="n">Difference</th></tr>${bad.map((c) => `<tr><td>${esc(c.group)}: ${esc(c.name)}</td><td class="n">${fmt(c.expected)}</td><td class="n">${fmt(c.actual)}</td><td class="n fail">${fmt(c.diff)}</td></tr>`).join("")}</table>`;
    const pr = g.product;
    if (!pr.ok) h += `<div class="banner warn"><b>Product sales will be an open item</b>${pr.stale && pr.stale.length ? `The Productwise tab still holds old data (${esc(pr.stale.join(", "))}).` : `The Productwise tab covers ${Math.round(pr.coverage * 100)}% of gross sales.`} Paste this month's Shopify product sales into it and run again to include them.</div>`;
    const rows = [
      ...g.new_pl.map((x) => ["P&L", x.account, x.this]),
      ...g.new_bs.map((x) => ["Balance sheet", x.account, x.this]),
      ...g.bs_not_on_prior_report.map((n) => ["Balance sheet", n, null]),
      ...g.orphan_accruals.map((n) => ["Balance sheet", n, null])
    ];
    if (rows.length) h += `<details><summary>${rows.length} new accounts - confirm they are in the right place</summary><table class="t"><tr><th>Ledger</th><th>Account</th><th class="n">This month</th></tr>${rows.map((r) => `<tr><td>${r[0]}</td><td>${esc(r[1])}</td><td class="n">${fmt(r[2])}</td></tr>`).join("")}</table></details>`;
    const ks = Object.keys(g.totals);
    h += `<details><summary>Details</summary><table class="t"><tr><th></th><th class="n">${esc(ks[0])}</th><th class="n">${esc(ks[1])}</th></tr>` + [["Income", "income"], ["Cost of sales", "cogs"], ["Gross profit", "gp"], ["Operating expenses", "opex"], ["Net result", "net"]].map(([n, k]) => `<tr><td>${n}</td><td class="n">${fmt(g.totals[ks[0]][k])}</td><td class="n">${fmt(g.totals[ks[1]][k])}</td></tr>`).join("") + `</table><table class="t"><tr><th>Check</th><th class="n">Expected</th><th class="n">Actual</th><th class="n">Difference</th><th></th></tr>` + gateChecks.concat(drsChecks).map((c) => `<tr><td>${esc(c.name)}</td><td class="n">${fmt(c.expected)}</td><td class="n">${fmt(c.actual)}</td><td class="n">${fmt(c.diff)}</td>${statusCell(c)}</tr>`).join("") + `</table></details>`;
    return h;
  }
  async function runGate() {
    const btn = $("#gate");
    await busy(btn, "Checking...");
    resetGate();
    try {
      if (!linksOk()) throw new Error("Links must start with https://");
      const r = await buildReport(inputs(), { ...options(), gateOnly: true });
      $("#gateOut").innerHTML = gateHtml(r.gate, r.status);
      gateOk = r.status === "gate_ok";
      $("#go").disabled = !gateOk;
    } catch (e) {
      $("#gateOut").innerHTML = `<div class="banner bad"><b>Could not run the tie-out</b>${esc(e.message || e)}</div>`;
      console.error(e);
    }
    unbusy(btn);
    btn.disabled = !(picked.pl && picked.bs && picked.drs && (picked.report || stored));
  }
  function reconHtml(recon) {
    return `<div class="three">${recon.map((t) => `<div><h3>${esc(t.title)}</h3><table class="t"><tr><th></th><th class="n">Xero</th><th class="n">Report</th><th class="n">Diff</th></tr>${t.rows.map((r) => `<tr><td>${r.line}</td><td class="n">${fmt(r.xero)}</td><td class="n">${r.report == null ? "n/a" : fmt(r.report)}</td><td class="n ${r.diff === 0 ? "pass" : "fail"}">${r.diff == null ? "" : fmt(r.diff)}</td></tr>`).join("")}</table></div>`).join("")}</div>`;
  }
  async function runBuild() {
    const btn = $("#go");
    await busy(btn, "Building...");
    $("#buildOut").innerHTML = "";
    $("#after").hidden = true;
    urls.forEach((u) => URL.revokeObjectURL(u));
    urls = [];
    try {
      const r = await buildReport(inputs(), options());
      const s = r.summary;
      if (s.status === "readback_failed") {
        $("#buildOut").innerHTML = `<div class="banner bad"><b>Do not send - nothing released</b>The report does not agree with the Xero P&amp;L. Differences are below.</div>`;
        $("#reviewOut").innerHTML = reconHtml(s.recon) + `<table class="t"><tr><th>Check</th><th class="n">Xero</th><th class="n">Report</th><th class="n">Diff</th></tr>${s.readback_fail.map((c) => `<tr><td>${esc(c.name)}</td><td class="n">${fmt(c.expected)}</td><td class="n">${fmt(c.actual)}</td><td class="n">${fmt(c.diff)}</td></tr>`).join("")}</table>`;
        $("#deliverOut").innerHTML = "";
        $("#after").hidden = false;
      } else {
        $("#buildOut").innerHTML = `<div class="banner"><b>Built - ${esc(s.month)}</b>${s.readback_ok} figures checked against Xero - all agree.</div>`;
        const link = (f) => {
          const u = URL.createObjectURL(new Blob([f.bytes], { type: f.mime }));
          urls.push(u);
          return u;
        };
        let rv = "";
        if (!r.gate.product.ok) rv += `<div class="banner warn"><b>Open item</b>Product sales are shown as an open item in the report.</div>`;
        if (s.links_removed.length) rv += `<p class="muted">Old links removed. Add new ones under Optional in step 1 if needed.</p>`;
        rv += `<details><summary>Report vs Xero</summary>${reconHtml(s.recon)}</details>`;
        rv += `<details><summary>Check the wording still matches the numbers</summary><ul class="chk">${s.narrative.map((t) => `<li>${esc(t)}</li>`).join("")}</ul></details>`;
        rv += `<p class="muted">Open the .docx in Word, click Yes to update fields, and read every page before sending.</p>`;
        $("#reviewOut").innerHTML = rv;
        const files = r.files.filter((f) => !f.name.startsWith("Call_points"));
        const cp = r.files.find((f) => f.name.startsWith("Call_points"));
        let dv = `<div class="dl">${files.map((f) => `<a href="${link(f)}" download="${esc(f.name)}">${esc(f.name)}</a>`).join("")}</div>`;
        dv += `<details><summary>Call points (internal - not for the client)</summary><textarea readonly>${esc(new TextDecoder().decode(cp.bytes))}</textarea><div class="dl"><button class="ghost" id="copy">Copy</button> <a href="${link(cp)}" download="${esc(cp.name)}">Download .txt</a></div></details>`;
        $("#deliverOut").innerHTML = dv;
        $("#copy").onclick = async (e) => {
          await navigator.clipboard.writeText(new TextDecoder().decode(cp.bytes));
          e.target.textContent = "Copied";
        };
        $("#after").hidden = false;
        await idbPut(r.ctx.k0, r.files.find((f) => f.name.endsWith(".docx")).bytes);
      }
    } catch (e) {
      $("#buildOut").innerHTML = `<div class="banner bad"><b>Could not build</b>${esc(e.message || e)}</div>`;
      console.error(e);
    }
    unbusy(btn);
    btn.disabled = !gateOk;
  }
  $("#files").addEventListener("change", async (e) => {
    const l = [...e.target.files];
    e.target.value = "";
    await addFiles(l);
  });
  $("#ytd").addEventListener("change", async (e) => {
    const f = e.target.files[0];
    ytdFile = f ? await bytes(f) : null;
    $("#ytdName").textContent = f ? f.name : "none";
    resetGate();
  });
  ["drsLink", "deepLink", "quickLink"].forEach((k) => $("#" + k).addEventListener("input", resetGate));
  $("#gate").addEventListener("click", runGate);
  $("#go").addEventListener("click", runBuild);
  var dz = $("#drop");
  ["dragover", "dragenter"].forEach((ev) => dz.addEventListener(ev, (e) => {
    e.preventDefault();
    dz.classList.add("over");
  }));
  ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => {
    e.preventDefault();
    dz.classList.remove("over");
  }));
  dz.addEventListener("drop", (e) => addFiles([...e.dataTransfer.files]));
  renderInputs();
  ["dragover", "drop"].forEach((ev) => window.addEventListener(ev, (e) => {
    if (!e.target.closest || !e.target.closest("#drop")) e.preventDefault();
  }));
  window.__saintStarted = true;
})();

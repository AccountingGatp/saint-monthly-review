// Edit last month's Month-End Financial Report (.docx) into this month's.
// The prior report is the template: only the cells, text and pictures that carry numbers are rewritten.
import { NS, parseXml, serializeXml, kids, kid, desc, firstDesc, wAttr, setW, make, insertAt, after, before, remove,
         nextEl, indexOfEl, elems, ptext } from "./dom.js";
import { InputError, MONTH_NAMES, MONTH_ABBR, commas, keyM, keyY, longLabel, shortLabel } from "./util.js";

const SHADE = "dce6f1";
const RED_TXT = "9c0006";
const MONTH_RE = MONTH_NAMES.slice(1).join("|");

// ------------------------------------------------------------ formatting
export function fmtAcc(x, blankZero = true) {
  if (x === null || x === undefined || (blankZero && Math.abs(x) < 0.005)) return "";
  const s = commas(x, 2);
  return x < -0.004 ? `(${s})` : s;
}
export function fmtUsd0(x, blankZero = true) {
  if (x === null || x === undefined || (blankZero && Math.abs(x) < 0.5)) return "";
  return x < -0.5 ? `-$${commas(x, 0)}` : `$${commas(x, 0)}`;
}
export const fmtUsd2 = (x) => (x < 0 ? `-$${commas(x, 2)}` : `$${commas(x, 2)}`);
export { shortLabel, longLabel };

// ------------------------------------------------------------ cells & rows
function rprSet(rpr, bold, color) {
  const doc = rpr.ownerDocument;
  if (bold !== null && bold !== undefined) {
    for (const tag of ["b", "bCs"]) kids(rpr, tag).forEach(remove);
    if (bold) {
      const pos = kid(rpr, "rFonts") ? 1 : 0;
      ["b", "bCs"].forEach((tag, i) => insertAt(rpr, pos + i, make(doc, tag, { val: "1" })));
    }
  }
  if (color !== null && color !== undefined) {
    kids(rpr, "color").forEach(remove);
    const e = make(doc, "color", { val: color });
    const sz = kid(rpr, "sz");
    if (sz) before(sz, e); else rpr.appendChild(e);
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
  if (!ppr) { ppr = make(doc, "pPr"); p.insertBefore(ppr, p.firstChild); }
  if (align) {
    kids(ppr, "jc").forEach(remove);
    const jc = make(doc, "jc", { val: align });
    const rp = kid(ppr, "rPr");
    if (rp) before(rp, jc); else ppr.appendChild(jc);
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
    if (fill === null) { if (shd) remove(shd); }
    else {
      if (!shd) {
        shd = make(doc, "shd");
        const anchor = kid(tcpr, "tcMar");
        if (anchor) before(anchor, shd); else tcpr.appendChild(shd);
      }
      setW(shd, "fill", fill); setW(shd, "val", "clear");
    }
  }
}

function setRowProps(tr, header = false) {
  const doc = tr.ownerDocument;
  let trpr = kid(tr, "trPr");
  if (!trpr) { trpr = make(doc, "trPr"); insertAt(tr, 0, trpr); }
  elems(trpr).forEach(remove);
  trpr.appendChild(make(doc, "cantSplit", { val: "1" }));
  if (header) trpr.appendChild(make(doc, "tblHeader", { val: "1" }));
}

function clearRows(tbl, keep = 0) { kids(tbl, "tr").slice(keep).forEach(remove); }

class RowFactory {
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
}

// ------------------------------------------------------------ tables
export function findTables(root) {
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

const rowLabels = (rows) => rows.map((r) => ptext(kids(r, "tc")[0]).trim());

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
  const spec = [["Revenue", "income", true], ["Cost of Goods Sold", "cogs", false],
    ["Operating Expense", "opex", false], ["Net income", "net", true]];
  const BLUE = "2f75b5", REDC = "c00000", GREENC = "008000";
  spec.forEach(([label, key, goodUp], j) => {
    const tr = rows[2 + j];
    const tcs = kids(tr, "tc");
    const bs = tcs.map(baseRpr);
    const cur = t0[key], prev = t1[key], yt = tytd[key];
    const chg = Math.abs(prev) > 0.005 ? (cur - prev) / Math.abs(prev) : 0;
    const arrow = chg > 0 ? "▲" : "▼";
    const improved = (chg > 0) === goodUp;
    const pct = Math.abs(chg) < 1 ? `${(Math.abs(chg) * 100).toFixed(1)}%` : `${(Math.abs(chg) * 100).toFixed(0)}%`;
    const money = (x) => (x < 0 ? `(${commas(x, 2)})` : commas(x, 2));
    const vals = [label, money(cur), money(prev), `${arrow} ${pct}`, money(yt)];
    tcs.forEach((tc, i) => {
      let color = BLUE;
      if ([1, 2, 4].includes(i) && vals[i].startsWith("(")) color = REDC;
      if (i === 3) color = improved ? GREENC : REDC;
      setCell(tc, vals[i], bs[i], { bold: i === 0 ? true : null, color, align: "center" });
    });
    setRowProps(tr);
  });
  setRowProps(rows[0]); setRowProps(rows[1], false);
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

const negColor = (x) => (x < -0.004 ? RED_TXT : "000000");

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
  const fill = () => (idx++ % 2 === 0 ? SHADE : null);
  const addSection = (name) => tbl.appendChild(fsec.make([name, "", "", "", ""], { boldFirst: true, fill: fill() }));
  const addData = (name, vals) => {
    const tot = vals.reduce((a, b) => a + b, 0);
    tbl.appendChild(fdat.make([name, ...vals.map((v) => fmtAcc(v)), fmtAcc(tot)],
      { colors: [null, ...vals.map(negColor), negColor(tot)], fill: fill() }));
  };
  const addTotal = (name, vals) => {
    const tot = vals.reduce((a, b) => a + b, 0);
    tbl.appendChild(ftot.make([name, ...vals.map((v) => fmtAcc(v, false)), fmtAcc(tot, false)],
      { boldFirst: true, boldNums: true, colors: [null, ...vals.map(negColor), negColor(tot)], fill: fill() }));
  };
  const valsOf = (l) => kList.map((k) => l.vals[k] ?? 0);
  const sums = {};
  for (const [section, title, totalName] of [["income", "Income", "Total Income"],
    ["cogs", "Cost of Goods Sold", "Total Cost of Goods Sold"], ["opex", "Operating Expenses", "Total Operating Expenses"]]) {
    addSection(title);
    for (const l of pl.lines) if (l.section === section && valsOf(l).some((x) => Math.abs(x) >= 0.005)) addData(l.name, valsOf(l));
    sums[section] = kList.map((k) => pl.sect(section, k));
    addTotal(totalName, sums[section]);
    if (section === "cogs") addTotal("Gross Profit", sums.income.map((a, i) => a - sums.cogs[i]));
  }
  const net = sums.income.map((a, i) => a - sums.cogs[i] - sums.opex[i]);
  addTotal("Operating Income", net);
  addTotal("Net Income", net);
}

const lcSort = (a, b) => { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; };

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
  const fill = () => (idx++ % 2 === 0 ? SHADE : null);
  const headRow = (name) => tbl.appendChild(fdat.make([name, "", "", ""], { boldFirst: true, fill: fill() }));
  const acctRow = (name, vals, force = false) =>
    tbl.appendChild(fdat.make([name, ...vals.map((v) => fmtUsd0(v, !force))], { boldFirst: true, boldNums: true, fill: fill() }));
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

// ------------------------------------------------------------ pictures
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
    if (/\.jpe?g$/i.test(target) || cx < 1_000_000) continue;       // cover logo
    out.push({ inline: inl, target, cx, cy });
  }
  return out;
}

function setPictureHeight(inl, cy) {
  kid(inl, "extent", NS.wp).setAttribute("cy", String(cy));
  for (const e of desc(inl, "ext", NS.a)) if (e.hasAttribute("cx")) e.setAttribute("cy", String(cy));
}

// ------------------------------------------------------------ text bits
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
    if (!rid) continue;                                   // internal TOC anchor
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
      if (relById[rid]) remove(relById[rid]);            // drop the dead relationship (old Google link)
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
      if (!ppr) { ppr = make(root.ownerDocument, "pPr"); p.insertBefore(ppr, p.firstChild); }
      if (!kid(ppr, "pageBreakBefore")) {
        let pos = 0;
        elems(ppr).forEach((ch, i) => { if (["pStyle", "keepNext", "keepLines"].includes(ch.localName)) pos = i + 1; });
        insertAt(ppr, pos, make(root.ownerDocument, "pageBreakBefore", { val: "1" }));
      }
    }
  }
}

function fixTocTab(root) {
  for (const p of desc(root, "p")) {
    if (!kid(p, "hyperlink")) continue;
    const ppr = kid(p, "pPr"), tabs = ppr && kid(ppr, "tabs");
    if (tabs) for (const tab of kids(tabs, "tab")) if (wAttr(tab, "pos") === "10330") setW(tab, "pos", "10200");
  }
}

// The contents page numbers cannot be known without laying the document out, so each number becomes a
// PAGEREF field that Word refreshes when the file is opened (settings.xml asks Word to update fields).
function tocToFields(root) {
  const doc = root.ownerDocument;
  let n = 0;
  const body = kid(root, "body");
  const marks = new Set(desc(root, "bookmarkStart").map((b) => wAttr(b, "name")));
  let nextId = Math.max(100, ...desc(root, "bookmarkStart").map((b) => +wAttr(b, "id") || 0)) + 1;
  const shared = new Map();
  for (const h of desc(root, "hyperlink")) { const a = h.getAttributeNS(NS.w, "anchor"); if (a) shared.set(a, (shared.get(a) || 0) + 1); }
  const heading = (label) => elems(body).find((p) => p.localName === "p" && !kid(p, "hyperlink") && !firstDesc(p, "fldChar") && ptext(p).trim().startsWith(label));
  for (const p of desc(root, "p")) {
    if (!kid(p, "hyperlink")) continue;
    const label = kids(p, "r").map((r) => (kid(r, "t") ? kid(r, "t").textContent : "")).join("").trim();
    for (const hl of kids(p, "hyperlink")) {
      let anchor = hl.getAttributeNS(NS.w, "anchor");
      if (anchor && (!marks.has(anchor) || shared.get(anchor) > 1) && label) {          // Google Docs left some contents links without a target: give the heading a bookmark
        const h = heading(label);
        if (h) {
          anchor = "_toc" + nextId;
          const bs = make(doc, "bookmarkStart", { id: nextId, name: anchor }), be = make(doc, "bookmarkEnd", { id: nextId });
          nextId++;
          const ppr = kid(h, "pPr");
          if (ppr) after(ppr, bs); else h.insertBefore(bs, h.firstChild);
          h.appendChild(be);
          marks.add(anchor);
        }
      }
      const t = anchor && desc(hl, "t")[0];
      if (!t || !/^\d+$/.test((t.textContent || "").trim())) continue;
      const run = kid(hl, "r");
      const rprOf = () => (kid(run, "rPr") ? kid(run, "rPr").cloneNode(true) : null);
      const mkRun = (child) => { const r = make(doc, "r"); const rp = rprOf(); if (rp) r.appendChild(rp); r.appendChild(child); return r; };
      const parts = [
        mkRun(make(doc, "fldChar", { fldCharType: "begin" })),
        (() => { const it = make(doc, "instrText"); it.setAttributeNS(NS.xml, "xml:space", "preserve"); it.textContent = ` PAGEREF ${anchor} \\h `; return mkRun(it); })(),
        mkRun(make(doc, "fldChar", { fldCharType: "separate" })),
        run,
        mkRun(make(doc, "fldChar", { fldCharType: "end" })),
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
    const first = kids(p, "r").map((r) => (kid(r, "t") ? kid(r, "t").textContent : "")).join("").trim();
    for (const [label, pg] of Object.entries(pages)) {
      if (first === label) {
        for (const hl of kids(p, "hyperlink")) {
          if (hl.getAttributeNS(NS.w, "anchor")) {
            for (const t of desc(hl, "t")) if (/^\d+$/.test((t.textContent || "").trim())) { t.textContent = String(pg); n++; }
          }
        }
      }
    }
  }
  return n;
}

const paraOf = (el) => { while (el && !(el.localName === "p" && el.namespaceURI === NS.w)) el = el.parentNode; return el; };

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
  if (firstDesc(dp, "sectPr")) before(dp, note);     // a section break follows: keep the note on its page
  else after(dp, note);
}

// ------------------------------------------------------------ main
// images: kind -> async (w_in, h_in) => {png: Uint8Array, w, h, note}
export async function build(priorBytes, ctx, images, tocPages = null) {
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

  const T = findTables(root);
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

  // notes added by a previous run (open item / provisional) are dropped; this run re-adds them if still needed
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
      if (anchor) sr.insertBefore(uf, anchor); else sr.appendChild(uf);
    }
    files["word/settings.xml"] = new TextEncoder().encode(serializeXml(st));
  }
  files["word/document.xml"] = new TextEncoder().encode(serializeXml(docXml));
  files["word/_rels/document.xml.rels"] = new TextEncoder().encode(serializeXml(relsXml));
  const zo = new globalThis.JSZip();
  for (const n of order) if (n in files) zo.file(n, files[n]);
  const bytes = await zo.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { bytes, linkLog };
}

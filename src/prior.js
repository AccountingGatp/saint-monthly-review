// Read the figures published in a Month-End Financial Report (.docx).
import { NS, parseXml, kids, desc, ptext } from "./dom.js";
import { InputError, monthKeyFromLabel, mk } from "./util.js";

export async function loadDocXml(bytes) {
  const z = await globalThis.JSZip.loadAsync(bytes);
  return parseXml(await z.file("word/document.xml").async("string"));
}
export const cellText = (tc) => ptext(tc).replace(/ /g, " ").trim();

export function parseAmount(s) {
  s = String(s || "").replace(/\$/g, "").replace(/,/g, "").replace(/ /g, "").trim();
  if (!s) return 0;
  const neg = (s.startsWith("(") && s.endsWith(")")) || s.startsWith("-");
  s = s.replace(/^[()\-]+|[()\-]+$/g, "");
  const v = parseFloat(s);
  if (!isFinite(v) || !/^[0-9.]+$/.test(s)) return 0;
  return neg ? -v : v;
}

export function loadTables(doc) {
  return desc(doc, "tbl").map((tb) => kids(tb, "tr").map((tr) => kids(tr, "tc").map(cellText)));
}

// raw row-lists by content: exec, country, pl, bs
export function findTables(doc) {
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

export function monthKeyFromShort(label) {                 // "Aug '26" -> "2026-08"
  const m = /^([A-Za-z]+)\s*'(\d\d)/.exec(label.trim());
  return m ? monthKeyFromLabel(`${m[1]} 20${m[2]}`) : null;
}

export function tableByMonth(rows) {
  const head = rows[0];
  const cols = {};
  head.forEach((h, i) => { if (i && monthKeyFromShort(h)) cols[i] = monthKeyFromShort(h); });
  const out = {};
  for (const r of rows.slice(1)) {
    if (!r.length || !r[0]) continue;
    const o = {};
    for (const [i, k] of Object.entries(cols)) if (+i < r.length) o[k] = parseAmount(r[i]);
    out[r[0]] = o;
  }
  return [out, Object.values(cols)];
}

export async function readPrior(bytes) {
  const t = findTables(await loadDocXml(bytes));
  const [pl, plKeys] = tableByMonth(t.pl);
  const [bs, bsKeys] = tableByMonth(t.bs);
  return { pl, pl_keys: plKeys, bs, bs_keys: bsKeys, tables: t };
}

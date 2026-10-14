// Small shared helpers (number formatting, month keys, errors).
export class InputError extends Error {}

export const MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const MONTH_ABBR = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export const round2 = (x) => Math.round((x + Math.sign(x) * 1e-9) * 100) / 100;
export const round0 = (x) => Math.round(x + Math.sign(x) * 1e-9);
export const sum = (a) => a.reduce((s, x) => s + x, 0);
export const num = (v) => (typeof v === "number" && isFinite(v) ? v : 0);

export function commas(x, dp = 2) {                    // 1234.5 -> "1,234.50" (x must be >= 0)
  const s = Math.abs(x).toFixed(dp);
  const [i, f] = s.split(".");
  return i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (f ? "." + f : "");
}

// month keys are "YYYY-MM"
export const mk = (y, m) => `${y}-${String(m).padStart(2, "0")}`;
export const keyY = (k) => +k.slice(0, 4);
export const keyM = (k) => +k.slice(5);
export function prevKey(k, n = 1) {
  let y = keyY(k), m = keyM(k) - n;
  while (m < 1) { m += 12; y -= 1; }
  return mk(y, m);
}
export const monthLabel = (k) => `${MONTH_ABBR[keyM(k)]} ${keyY(k)}`;       // "Sep 2026"
export const shortLabel = (k) => `${MONTH_ABBR[keyM(k)]} '${String(keyY(k)).slice(2)}`;   // "Sep '26"
export const longLabel = (k) => `${MONTH_NAMES[keyM(k)]} ${keyY(k)}`;       // "September 2026"
export const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();

export function monthKeyFromLabel(label) {                                   // "Sept 2026" -> "2026-09"
  const m = /^\s*([A-Za-z]+)\.?\s+(\d{4})\s*$/.exec(String(label));
  if (!m) return null;
  const mo = monthFromWord(m[1], false);
  return mo ? mk(+m[2], mo) : null;
}
const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
export function monthFromWord(word, strict = true) {
  const m = MON[String(word).replace(/^\.+|\.+$/g, "").toLowerCase().slice(0, 3)];
  if (!m && strict) throw new InputError(`Cannot read month name '${word}'`);
  return m || 0;
}

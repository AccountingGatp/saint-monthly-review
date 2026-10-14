// Thin reader over ExcelJS that mimics what the report tool needs:
// v(r,c) -> number | string | Date | null ; with {formulas:true} formula cells come back as "=...".
const EJ = () => globalThis.ExcelJS;

function norm(x, formulas) {
  if (x === null || x === undefined) return null;
  if (x instanceof Date) return x;
  const t = typeof x;
  if (t === "number" || t === "string" || t === "boolean") return x === "" ? null : x;
  if (x.formula !== undefined || x.sharedFormula !== undefined) {
    return formulas ? "=" + (x.formula || x.sharedFormula) : norm(x.result, false);
  }
  if (x.richText) return x.richText.map((r) => r.text).join("");
  if (x.error) return x.error;
  if (x.text !== undefined) return norm(x.text, formulas);
  return null;
}

export class Sheet {
  constructor(ws) { this.ws = ws; this.name = ws.name; }
  get maxRow() { return this.ws.rowCount; }
  get maxCol() { return this.ws.columnCount; }
  v(r, c, opts) {
    const row = this.ws.getRow(r);
    return norm(row.getCell(c).value, !!(opts && opts.formulas));
  }
}

export class Book {
  constructor(wb) {
    this.wb = wb;
    this.sheets = wb.worksheets.map((w) => new Sheet(w));
    this.names = this.sheets.map((s) => s.name);
  }
  first() { return this.sheets[0]; }
  sheet(name) { return this.sheets.find((s) => s.name === name); }
}

export async function openBook(buf) {
  const wb = new (EJ().Workbook)();
  await wb.xlsx.load(buf);
  return new Book(wb);
}

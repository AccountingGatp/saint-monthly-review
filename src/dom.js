// XML helpers over the browser DOM (also works under jsdom in Node tests).
export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  xml: "http://www.w3.org/XML/1998/namespace",
};

export const parseXml = (s) => {
  const d = new DOMParser().parseFromString(s, "application/xml");
  const err = d.getElementsByTagName("parsererror");
  if (err.length) throw new Error("Could not read the Word file (XML error)");
  return d;
};
export const serializeXml = (doc) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc.documentElement);

export const isEl = (n) => n && n.nodeType === 1;
export const elems = (el) => Array.from(el.childNodes).filter(isEl);
export const local = (el) => el.localName;
// direct children by namespace + local name
export const kids = (el, name, ns = NS.w) => elems(el).filter((c) => c.localName === name && c.namespaceURI === ns);
export const kid = (el, name, ns = NS.w) => kids(el, name, ns)[0] || null;
// all descendants (document order) by local name
export const desc = (el, name, ns = NS.w) => Array.from(el.getElementsByTagNameNS(ns, name));
export const firstDesc = (el, name, ns = NS.w) => el.getElementsByTagNameNS(ns, name)[0] || null;

export const wAttr = (el, name) => el.getAttributeNS(NS.w, name);
export const setW = (el, name, val) => el.setAttributeNS(NS.w, "w:" + name, String(val));

export function make(doc, name, attrs = {}, ns = NS.w, prefix = "w") {
  const e = doc.createElementNS(ns, prefix + ":" + name);
  for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(ns, prefix + ":" + k, String(v));
  return e;
}
export function insertAt(parent, idx, node) {          // idx counts element children (as lxml does)
  const ch = elems(parent);
  if (idx < ch.length) parent.insertBefore(node, ch[idx]);
  else parent.appendChild(node);
}
export const after = (ref, node) => ref.parentNode.insertBefore(node, ref.nextSibling);
export const before = (ref, node) => ref.parentNode.insertBefore(node, ref);
export const remove = (n) => n.parentNode && n.parentNode.removeChild(n);
export const nextEl = (n) => { let x = n.nextSibling; while (x && !isEl(x)) x = x.nextSibling; return x; };
export const indexOfEl = (el) => elems(el.parentNode).indexOf(el);
export function ptext(el) { return desc(el, "t").map((t) => t.textContent || "").join(""); }

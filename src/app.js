import { buildReport, sniff, plMonth, describeInput } from "./build.js";

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fmt = (v) => (v == null ? "" : Number(v).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const A0 = (v) => "A$" + Math.round(v).toLocaleString("en-AU");
const ROLES = [
  ["pl", "Profit and Loss", "Xero .xlsx"],
  ["bs", "Balance Sheet", "Xero .xlsx"],
  ["drs", "Daily Reporting Sheet", ".xlsx"],
  ["report", "Last month's report", ".docx (template)"],
];
const picked = {};                       // role -> {name, bytes, info}
let ytdFile = null, urls = [], gateOk = false, stored = null;

const bytes = async (f) => new Uint8Array(await f.arrayBuffer());

// ---- remembered report (each finished report becomes next month's template) ----
const prevMonth = (k) => { let y = +k.slice(0, 4), m = +k.slice(5) - 1; if (m < 1) { m = 12; y--; } return `${y}-${String(m).padStart(2, "0")}`; };
const idb = () => new Promise((res, rej) => {
  const r = indexedDB.open("saint-reports", 1);
  r.onupgradeneeded = () => r.result.createObjectStore("reports");
  r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
});
async function idbGet(key) { try { const db = await idb(); return await new Promise((res) => { const q = db.transaction("reports").objectStore("reports").get(key); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); }); } catch { return null; } }
async function idbPut(key, val) { try { const db = await idb(); await new Promise((res) => { const t = db.transaction("reports", "readwrite"); t.objectStore("reports").put(val, key); t.oncomplete = res; t.onerror = res; }); } catch { /* storage blocked */ } }

async function refreshStored() {
  stored = null;
  if (picked.pl && !picked.report) {
    try { const k = prevMonth(await plMonth(picked.pl.bytes)); const b = await idbGet(k); if (b) stored = { key: k, bytes: b }; } catch { /* reported at tie-out */ }
  }
  resetGate(); renderInputs();
}

function renderInputs() {
  $("#inputs").innerHTML = ROLES.map(([r, name, hint]) => {
    const p = picked[r];
    let state;
    if (p) state = `<span class="ok">${esc(p.name)}</span>`;
    else if (r === "report" && stored) state = `<span class="ok">Remembered (${esc(stored.key)})</span>`;
    else state = `<span class="no">not supplied</span>`;
    const ok = !!p || (r === "report" && !!stored);
    return `<tr><td class="st ${ok ? "ok" : "no"}">${ok ? "&#10003;" : "&#9675;"}</td><td><b>${name}</b> <span class="muted">${hint}</span></td><td>${state}</td></tr>`;
  }).join("");
  const missing = ROLES.filter(([r]) => !(picked[r] || (r === "report" && stored))).map(([, n]) => n);
  $("#gate").disabled = missing.length > 0;
  $("#missing").innerHTML = missing.length
    ? `<div class="banner warn"><b>Still needed: ${missing.map(esc).join("; ")}</b>${missing.some((n) => /^Last/.test(n)) ? "Drop last month's finished report once; after a successful run it is remembered." : ""}</div>` : "";
}

function resetGate() { gateOk = false; $("#go").disabled = true; $("#gateOut").innerHTML = ""; $("#buildOut").innerHTML = ""; $("#after").hidden = true; }

async function addFiles(list) {
  const msgs = [];
  for (const f of list) {
    const b = await bytes(f);
    const role = await sniff(f.name, b);
    if (!role) { msgs.push(`${f.name}: not recognised. Use the Xero Profit and Loss, Xero Balance Sheet, Daily Reporting Sheet (.xlsx) or a finished Month-End Financial Report (.docx).`); continue; }
    let info = "";
    try { info = await describeInput(role, b); } catch (e) { info = "Could not read: " + (e.message || e); }
    picked[role] = { name: f.name, bytes: b, info };
  }
  $("#problems").innerHTML = msgs.map((m) => `<div class="banner bad">${esc(m)}</div>`).join("");
  await refreshStored();
}

const reportBytes = () => (picked.report ? picked.report.bytes : stored.bytes);
const inputs = () => ({ pl: picked.pl.bytes, bs: picked.bs.bytes, drs: picked.drs.bytes, report: reportBytes(), ytd: ytdFile });
const options = () => ({ links: { drs: $("#drsLink").value.trim(), deep: $("#deepLink").value.trim(), quick: $("#quickLink").value.trim() } });
const linksOk = () => ["drsLink", "deepLink", "quickLink"].every((k) => { const v = $("#" + k).value.trim(); return !v || /^https:\/\/\S+$/.test(v); });
const busy = (btn, text) => { btn.disabled = true; btn.dataset.t = btn.textContent; btn.textContent = text; return new Promise((r) => setTimeout(r, 30)); };
const unbusy = (btn) => { btn.textContent = btn.dataset.t; };

// ---------------------------------------------------------------- step 2
function statusCell(c) { return `<td class="${c.status === "PASS" ? "pass" : c.status === "FAIL" ? "fail" : "warn"}">${c.status === "PASS" ? "Agrees" : c.status === "FAIL" ? "Differs" : c.status === "WARN" ? "Look" : "Info"}</td>`; }

function gateHtml(g, status) {
  const gateChecks = g.checks.filter((c) => c.group !== "DRS vs P&L");
  const drsChecks = g.checks.filter((c) => c.group === "DRS vs P&L");
  const bad = gateChecks.filter((c) => c.status === "FAIL");
  let h = "";
  if (status === "gate_ok") h += `<div class="banner"><b>Gate passed</b>All ${gateChecks.length} checks agree to the cent.</div>`;
  else h += `<div class="banner bad"><b>Stopped: ${bad.length} difference${bad.length === 1 ? "" : "s"}</b>Nothing is built. Chase each one, re-export from Xero on the same day and run again.</div>` +
    `<table class="t"><tr><th>Check</th><th class="n">Expected</th><th class="n">Actual</th><th class="n">Difference</th></tr>${bad.map((c) => `<tr><td>${esc(c.group)}: ${esc(c.name)}</td><td class="n">${fmt(c.expected)}</td><td class="n">${fmt(c.actual)}</td><td class="n fail">${fmt(c.diff)}</td></tr>`).join("")}</table>`;

  const pr = g.product;
  if (!pr.ok) h += `<div class="banner warn"><b>Product sales will be an open item</b>${pr.stale && pr.stale.length ? `The Productwise tab still holds old data (${esc(pr.stale.join(", "))}).` : `The Productwise tab covers ${Math.round(pr.coverage * 100)}% of gross sales.`} Paste this month's Shopify product sales into it and run again to include them.</div>`;

  const rows = [
    ...g.new_pl.map((x) => ["P&L", x.account, x.this]),
    ...g.new_bs.map((x) => ["Balance sheet", x.account, x.this]),
    ...g.bs_not_on_prior_report.map((n) => ["Balance sheet", n, null]),
    ...g.orphan_accruals.map((n) => ["Balance sheet", n, null]),
  ];
  if (rows.length) h += `<details><summary>${rows.length} new accounts - confirm they are in the right place</summary><table class="t"><tr><th>Ledger</th><th>Account</th><th class="n">This month</th></tr>${rows.map((r) => `<tr><td>${r[0]}</td><td>${esc(r[1])}</td><td class="n">${fmt(r[2])}</td></tr>`).join("")}</table></details>`;

  const ks = Object.keys(g.totals);
  h += `<details><summary>Details</summary><table class="t"><tr><th></th><th class="n">${esc(ks[0])}</th><th class="n">${esc(ks[1])}</th></tr>` +
    [["Income", "income"], ["Cost of sales", "cogs"], ["Gross profit", "gp"], ["Operating expenses", "opex"], ["Net result", "net"]]
      .map(([n, k]) => `<tr><td>${n}</td><td class="n">${fmt(g.totals[ks[0]][k])}</td><td class="n">${fmt(g.totals[ks[1]][k])}</td></tr>`).join("") + `</table>` +
    `<table class="t"><tr><th>Check</th><th class="n">Expected</th><th class="n">Actual</th><th class="n">Difference</th><th></th></tr>` +
    gateChecks.concat(drsChecks).map((c) => `<tr><td>${esc(c.name)}</td><td class="n">${fmt(c.expected)}</td><td class="n">${fmt(c.actual)}</td><td class="n">${fmt(c.diff)}</td>${statusCell(c)}</tr>`).join("") + `</table></details>`;
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
  unbusy(btn); btn.disabled = !(picked.pl && picked.bs && picked.drs && (picked.report || stored));
}

// ---------------------------------------------------------------- steps 3-5
function reconHtml(recon) {
  return `<div class="three">${recon.map((t) => `<div><h3>${esc(t.title)}</h3><table class="t"><tr><th></th><th class="n">Xero</th><th class="n">Report</th><th class="n">Diff</th></tr>${t.rows.map((r) =>
    `<tr><td>${r.line}</td><td class="n">${fmt(r.xero)}</td><td class="n">${r.report == null ? "n/a" : fmt(r.report)}</td><td class="n ${r.diff === 0 ? "pass" : "fail"}">${r.diff == null ? "" : fmt(r.diff)}</td></tr>`).join("")}</table></div>`).join("")}</div>`;
}

async function runBuild() {
  const btn = $("#go");
  await busy(btn, "Building...");
  $("#buildOut").innerHTML = ""; $("#after").hidden = true;
  urls.forEach((u) => URL.revokeObjectURL(u)); urls = [];
  try {
    const r = await buildReport(inputs(), options());
    const s = r.summary;
    if (s.status === "readback_failed") {
      $("#buildOut").innerHTML = `<div class="banner bad"><b>Do not send - nothing released</b>The report does not agree with the Xero P&amp;L. Differences are below.</div>`;
      $("#reviewOut").innerHTML = reconHtml(s.recon) + `<table class="t"><tr><th>Check</th><th class="n">Xero</th><th class="n">Report</th><th class="n">Diff</th></tr>${s.readback_fail.map((c) => `<tr><td>${esc(c.name)}</td><td class="n">${fmt(c.expected)}</td><td class="n">${fmt(c.actual)}</td><td class="n">${fmt(c.diff)}</td></tr>`).join("")}</table>`;
      $("#deliverOut").innerHTML = ""; $("#after").hidden = false;
    } else {
      $("#buildOut").innerHTML = `<div class="banner"><b>Built - ${esc(s.month)}</b>${s.readback_ok} figures checked against Xero - all agree.</div>`;
      const link = (f) => { const u = URL.createObjectURL(new Blob([f.bytes], { type: f.mime })); urls.push(u); return u; };
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
      $("#copy").onclick = async (e) => { await navigator.clipboard.writeText(new TextDecoder().decode(cp.bytes)); e.target.textContent = "Copied"; };
      $("#after").hidden = false;
      await idbPut(r.ctx.k0, r.files.find((f) => f.name.endsWith(".docx")).bytes);
    }
  } catch (e) {
    $("#buildOut").innerHTML = `<div class="banner bad"><b>Could not build</b>${esc(e.message || e)}</div>`;
    console.error(e);
  }
  unbusy(btn); btn.disabled = !gateOk;
}

$("#files").addEventListener("change", async (e) => { const l = [...e.target.files]; e.target.value = ""; await addFiles(l); });
$("#ytd").addEventListener("change", async (e) => { const f = e.target.files[0]; ytdFile = f ? await bytes(f) : null; $("#ytdName").textContent = f ? f.name : "none"; resetGate(); });
["drsLink", "deepLink", "quickLink"].forEach((k) => $("#" + k).addEventListener("input", resetGate));
$("#gate").addEventListener("click", runGate);
$("#go").addEventListener("click", runBuild);
const dz = $("#drop");
["dragover", "dragenter"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("over"); }));
dz.addEventListener("drop", (e) => addFiles([...e.dataTransfer.files]));
renderInputs();

// A file dropped outside the box must not make the browser open it instead.
["dragover", "drop"].forEach((ev) => window.addEventListener(ev, (e) => { if (!e.target.closest || !e.target.closest("#drop")) e.preventDefault(); }));
window.__saintStarted = true;

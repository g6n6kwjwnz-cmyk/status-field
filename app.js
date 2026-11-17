// Status Report Field - PWA to prepare the Dashboard Timing Status report from the shop floor.
// Everything is stored on the phone (IndexedDB) until it is sent as a .zip that the Excel macro
// "Mobile Import" (modMovil) loads into the Status report.
"use strict";

const APP_VERSION = "0.13.0";
const MAX_PHOTOS = 12;                 // same limit as the Status report (MAX_FOTOS_TOOL)
const PHOTO_MAX_SIDE = 1600;           // px, enough for a slide and light to send
const PHOTO_QUALITY = 0.85;

const STATUS = [
  { code: 1, name: "ON Going", color: "var(--green)" },
  { code: 2, name: "Pay Attention", color: "var(--amber)" },
  { code: 3, name: "ON RISK", color: "var(--red)" },
  { code: 4, name: "On HOLD", color: "var(--grey)" },
  { code: 5, name: "Cancelled", color: "var(--grey)" },
  { code: 6, name: "HLTO", color: "var(--grey)" },
  { code: 7, name: "CLOSED", color: "var(--grey)" },
];
const STATUS_PICK = STATUS.slice(0, 3);   // on the phone only ON Going / Pay Attention / ON RISK
// Same list as clsSelFotosPPT.ListaComponentes
const COMPONENTS = ["Blank Holder", "Cams", "CF (Checking Fixture)", "Common Plates", "Cutting Steels",
  "Electronic Standards", "Forming Steels", "Lifters", "Lower Die", "Lower Pad", "Piercing Components",
  "Punch", "Raw Material", "Standards (Columns, Guide plates...)", "Upper Die", "Upper Pad"];
const MANUAL = "Manual entry...";

const DEMO_REFS = {
  format: "DashboardTiming-refs", version: 1, created: new Date().toISOString(), owner: "Demo",
  refs: [
    { id: "DEMO-0001", owner: "Demo", project: "DEMO_PROJECT", supplier: "Supplier A", pn: "PN-1001", no: "OP10 Draw", status: 1 },
    { id: "DEMO-0002", owner: "Demo", project: "DEMO_PROJECT", supplier: "Supplier A", pn: "PN-1002", no: "OP20 Trim", status: 2 },
    { id: "DEMO-0003", owner: "Demo", project: "OTHER_PROJECT", supplier: "Supplier B", pn: "PN-2001", no: "Blanking", status: 3 },
  ],
};

// ---------------- Storage (IndexedDB) ----------------
const db = (() => {
  let p;
  function open() {
    if (p) return p;
    p = new Promise((res, rej) => {
      const r = indexedDB.open("status-report-field", 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore("kv");
        d.createObjectStore("drafts", { keyPath: "refId" });
        const ph = d.createObjectStore("photos", { keyPath: "id" });
        ph.createIndex("refId", "refId");
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return p;
  }
  async function tx(store, mode, fn) {
    const d = await open();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode);
      const s = t.objectStore(store);
      let out;
      Promise.resolve(fn(s)).then(v => { out = v; });
      t.oncomplete = () => res(out);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error);
    });
  }
  const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  return {
    get: (store, key) => tx(store, "readonly", s => req(s.get(key))),
    getAll: (store) => tx(store, "readonly", s => req(s.getAll())),
    put: (store, val, key) => tx(store, "readwrite", s => req(key === undefined ? s.put(val) : s.put(val, key))),
    del: (store, key) => tx(store, "readwrite", s => req(s.delete(key))),
    clear: (store) => tx(store, "readwrite", s => req(s.clear())),
    byRef: (refId) => tx("photos", "readonly", s => req(s.index("refId").getAll(refId))),
    // kind: "tool" (Tool advance, default), "cf" (C. Fixture box) or "mat" (Material box)
    toolPhotos: async (refId) => (await tx("photos", "readonly", s => req(s.index("refId").getAll(refId)))).filter(p => !p.kind || p.kind === "tool"),
    boxPhoto: async (refId, kind) => (await tx("photos", "readonly", s => req(s.index("refId").getAll(refId)))).find(p => p.kind === kind) || null,
  };
})();

// ---------------- State ----------------
const state = { refs: null, drafts: new Map(), photoCounts: new Map(), current: null, author: "", project: null };
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

// REFS: "PN" = part name, "No." = part number
function refName(r) { return r.pn || r.no || r.id; }
const pct = v => (v === null || v === undefined || v === "") ? "" : `${Math.round(Number(v) * (Number(v) <= 1.5 ? 100 : 1))}%`;
function qualityText(r) {
  const q = r.quality;
  if (!q || !q.parts || !q.parts.length) return "";
  return q.parts.map(([n, v]) => `${n} ${pct(v)}`).join(" · ");
}
// Card: only Part Number, Supplier, Production Plant and Quality %
function cardLines(r) {
  const q = qualityText(r);
  const t = timingSummary(r);
  return `<div class="sub">${r.no ? "PN " + esc(r.no) : ""}</div>
    <div class="sub">${esc([r.supplier, r.plant].filter(Boolean).join(" · "))}</div>
    <div class="sub">Quality: ${q ? esc(q) : "–"}</div>${t ? `<div class="sub">${t}</div>` : ""}`;
}

// ---------------- Timing, C/F, shipping... (refs file version 4, Excel 15f) ----------------
function isoKey(d) {                                   // ISO week key yyyyww (same as Excel ClaveISO)
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));   // Thursday of this week decides the year
  const y = t.getUTCFullYear();
  return y * 100 + Math.ceil(((t - Date.UTC(y, 0, 1)) / 86400000 + 1) / 7);
}
const curWeek = () => isoKey(new Date());
const wkText = k => { k = Number(k); if (!k) return ""; const y = Math.floor(k / 100); return `W${k % 100}${y !== Math.floor(curWeek() / 100) ? "/" + String(y).slice(2) : ""}`; };
const codeInfo = c => (state.refs && state.refs.codes && state.refs.codes[c]) || null;
function chip(c) {
  if (!c) return "";
  const i = codeInfo(c);
  return `<span class="code" style="${i ? `background:${i[1]};color:${i[2]}` : ""}">${esc(c)}</span>`;
}
// Latest plan line: Plan3 > Plan2 > Plan
function planLine(t) {
  let best = null, n = -1;
  for (const [name, cells] of t.lines || []) {
    const m = /^Plan(\d*)$/i.exec(String(name).trim());
    if (m) { const k = m[1] ? Number(m[1]) : 1; if (k > n) { n = k; best = [name, cells]; } }
  }
  return best;
}
function lineCells(t, name) { const l = (t.lines || []).find(([n]) => String(n).toLowerCase() === name.toLowerCase()); return l ? l[1] : null; }
// "W47: Plan R · Actual R · next PBO W49, SP W52"
function timingSummary(r) {
  const t = r.timing;
  if (!t || !Array.isArray(t.w) || !t.w.length) return "";
  const cw = curWeek();
  const i = t.w.indexOf(cw);
  const pl = planLine(t), act = lineCells(t, "Actual");
  const parts = [];
  if (i >= 0) {
    if (pl && pl[1][i]) parts.push(`${esc(pl[0])} ${chip(pl[1][i])}`);
    if (act && act[i]) parts.push(`Actual ${chip(act[i])}`);
  }
  const next = [];
  if (pl) {
    const from = t.w.findIndex(k => k >= cw);
    if (from >= 0) for (const m of ["PBO", "BO", "FBO", "SP"]) {
      const j = pl[1].findIndex((c, k) => k >= from && String(c).toUpperCase() === m);
      if (j >= 0) next.push(`${m} ${wkText(t.w[j])}`);
    }
  }
  if (!parts.length && !next.length) return "";
  return `Timing ${wkText(cw)}: ${parts.join(" · ") || "–"}${next.length ? ` · next ${esc(next.join(", "))}` : ""}`;
}
function timingGrid(t) {
  const cw = curWeek();
  const head = t.w.map(k => `<th class="${k === cw ? "now" : ""}">${k % 100}</th>`).join("");
  const rows = (t.lines || []).filter(([, cells]) => cells.some(c => c)).map(([name, cells]) =>
    `<tr><th class="ln">${esc(name)}</th>${cells.map((c, j) => {
      const i = c ? codeInfo(c) : null;
      return `<td data-k="${t.w[j]}" data-l="${esc(name)}" data-c="${esc(c)}" class="${t.w[j] === cw ? "now" : ""}" style="${i ? `background:${i[1]};color:${i[2]}` : ""}">${esc(c)}</td>`;
    }).join("")}</tr>`).join("");
  return `<div class="tgrid"><table><thead><tr><th class="ln"></th>${head}</tr></thead><tbody>${rows}</tbody></table></div>`;
}
const pctCell = v => (v === null || v === undefined || v === "") ? "" : pct(v);
function refDataHtml(r) {
  const sec = (title, n, body, open) => body ? `<details class="sec"${open ? " open" : ""}><summary>${title}${n ? ` <span class="muted">(${n})</span>` : ""}</summary>${body}</details>` : "";
  let h = "";
  if (r.timing && Array.isArray(r.timing.w) && r.timing.w.length) h += sec("Timing", 0, timingGrid(r.timing) + `<p class="muted small">Tap a cell to see the code.</p>`, true);
  if (Array.isArray(r.cf) && r.cf.length) {
    const rows = [...r.cf].sort((a, b) => String(b[0]).localeCompare(String(a[0])) || (b[1] || 0) - (a[1] || 0));
    const extra = [7, 8, 9, 10].filter(k => rows.some(x => x[k] !== null && x[k] !== undefined && x[k] !== ""));
    const names = { 7: "P3", 8: "P4", 9: "P5", 10: "P6" };
    h += sec("C/F results", rows.length, `<div class="tscroll"><table class="dt"><thead><tr><th>Date</th><th>Wk</th><th>Stage</th><th>PIST</th><th>LH</th><th>RH</th>${extra.map(k => `<th>${names[k]}</th>`).join("")}<th>Notes</th></tr></thead><tbody>${rows.map(x =>
      `<tr><td>${esc(x[0])}</td><td>${esc(wkText(x[1]))}</td><td>${esc(x[2])}</td><td>${pctCell(x[3])}</td><td>${pctCell(x[4])}</td><td>${pctCell(x[5])}</td>${extra.map(k => `<td>${pctCell(x[k])}</td>`).join("")}<td>${esc(x[6])}</td></tr>`).join("")}</tbody></table></div>`, true);
  }
  if (Array.isArray(r.ship) && r.ship.length) {
    const rows = [...r.ship].sort((a, b) => (b[0] || 0) - (a[0] || 0));
    h += sec("Shipping (SP)", rows.length, rows.map(x => {
      const wk = [wkText(x[0]), x[1] && x[1] !== x[0] ? wkText(x[1]) : ""].filter(Boolean).join(" – ");
      const link = /^https?:\/\//i.test(String(x[4] || "")) ? ` <a href="${esc(x[4])}" target="_blank" rel="noopener">tracking</a>` : (x[4] ? " " + esc(x[4]) : "");
      return `<div class="item"><b>${esc(wk)}</b> ${esc([x[2], x[3]].filter(Boolean).join(" · "))}${link}${x[5] || x[6] ? `<div class="muted small">${x[5] ? "ETD " + esc(x[5]) : ""} ${x[6] ? "ETA " + esc(x[6]) : ""}</div>` : ""}${x[7] ? `<div class="pre small">${esc(x[7])}</div>` : ""}</div>`;
    }).join(""));
  }
  if (Array.isArray(r.hlto) && r.hlto.length) {
    const rows = [...r.hlto].sort((a, b) => String(b[0]).localeCompare(String(a[0])));
    const open = rows.filter(x => !/clos/i.test(String(x[5] || ""))).length;
    h += sec(`HLTO${open ? ` <span class="code red">${open} open</span>` : ""}`, rows.length, rows.map(x =>
      `<div class="item"><b>${esc(x[0])}</b> · ${esc(x[1])} · <b>${esc(x[5] || "")}</b>${x[6] ? " " + esc(x[6]) : ""}<div class="pre">${esc(x[2])}</div>${x[3] ? `<div class="pre small">Action: ${esc(x[3])}</div>` : ""}${x[4] ? `<div class="muted small">Responsible: ${esc(x[4])}</div>` : ""}</div>`).join(""), open > 0);
  }
  if (Array.isArray(r.comments) && r.comments.length) {
    const rows = [...r.comments].sort((a, b) => String(b[0]).localeCompare(String(a[0])));
    h += sec("Comments", rows.length, rows.map(x => `<div class="item"><b>${esc(x[0])}</b> · ${esc(x[1])}<div class="pre">${esc(x[2])}</div></div>`).join(""));
  }
  if (Array.isArray(r.notes) && r.notes.length) {
    const rows = [...r.notes].sort((a, b) => (b[0] || 0) - (a[0] || 0));
    h += sec("Timing notes", rows.length, rows.map(x => `<div class="item"><b>${esc(wkText(x[0]))}</b> · ${esc(x[1])}<div class="pre">${esc(x[2])}</div></div>`).join(""));
  }
  return h;
}
// Current week in the middle of the timing grid (only once the view is visible: hidden it has no size)
function scrollTimingNow() {
  const g = document.querySelector("#refData .tgrid");
  const now = g && g.querySelector("th.now");
  if (g) g.scrollLeft = now ? Math.max(0, now.offsetLeft - g.clientWidth / 2) : g.scrollWidth;
}
function renderRefData(r) {
  const box = $("refData");
  box.innerHTML = refDataHtml(r);
  const g = box.querySelector(".tgrid");
  if (g) {
    g.onclick = e => {
      const td = e.target.closest("td[data-c]");
      if (!td || !td.dataset.c) return;
      const i = codeInfo(td.dataset.c);
      toast(`${wkText(td.dataset.k)} ${td.dataset.l}: ${td.dataset.c}${i ? " – " + i[0] : ""}`);
    };
  }
}
// "i": everything else about the reference
function showInfo(r) {
  const i = r.info || {};
  const cur = STATUS.find(s => s.code === Number(r.status));
  const q = r.quality;
  const rows = [
    ["Part name", r.pn], ["Part number", r.no], ["Project", r.project], ["Supplier", r.supplier], ["Plant", r.plant],
    ["Status in Excel", cur ? cur.name : "No status"],
    ["Quality (C/F)", q ? `${qualityText(r)}${q.stage ? " · " + q.stage : ""}${q.week ? " · W" + String(q.week).slice(4) + "/" + String(q.week).slice(0, 4) : ""}` : ""],
    ["Type", i.type], ["CAD level", i.cad], ["ND", i.nd], ["Material", i.material], ["Thickness", i.thickness],
    ["Width", i.width], ["Pitch", i.pitch], ["Press", i.press], ["Weight", i.weight], ["Product", i.product], ["Parts per die", i.partsDie], ["Last change", i.lastChange],
  ].filter(([, v]) => v !== undefined && v !== null && String(v).trim() !== "");
  ask(`${r.thumb ? `<img class="bigthumb" src="${r.thumb}" alt="">` : ""}<table class="info">${rows.map(([k, v]) =>
    `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`).join("")}</table>`, { cancel: "" });
}
function hasContent(d, nPhotos) {
  return !!d && (d.status || (d.comment || "").trim() || (d.cfixture || "").trim() || (d.material || "").trim() || nPhotos > 0);
}

// ---------------- UI helpers ----------------
let toastTimer;
function toast(msg) {
  const t = $("toast");
  t.textContent = msg;
  t.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add("hidden"), 3000);
}
function ask(html, { input = false, area = false, value = "", ok = "OK", cancel = "Cancel" } = {}) {
  return new Promise(res => {
    $("dlgText").innerHTML = html;
    $("dlgInput").classList.toggle("hidden", !input);
    $("dlgInput").value = value;
    $("dlgArea").classList.toggle("hidden", !area);
    $("dlgArea").value = "";
    $("dlgYes").textContent = ok;
    $("dlgNo").textContent = cancel;
    $("dlgNo").classList.toggle("hidden", !cancel);
    $("dlg").classList.remove("hidden"); syncHistory();
    if (input) setTimeout(() => $("dlgInput").focus(), 50);
    if (area) setTimeout(() => $("dlgArea").focus(), 50);
    const done = v => { $("dlg").classList.add("hidden"); $("dlgYes").onclick = $("dlgNo").onclick = null; syncHistory(); res(v); };
    $("dlgYes").onclick = () => done(area ? $("dlgArea").value : input ? $("dlgInput").value : true);
    $("dlgNo").onclick = () => done(input || area ? null : false);
  });
}

function show(view) {
  for (const v of ["viewList", "viewRef", "viewSettings", "viewReport"]) $(v).classList.toggle("hidden", v !== view);
  $("btnBack").classList.toggle("hidden", view === "viewList" && state.project === null);
  $("btnSettings").classList.toggle("hidden", view !== "viewList" || state.project !== null);
  $("btnFilter").classList.toggle("hidden", view !== "viewList" || !(state.refs && state.refs.refs.length));
  closeFilters();
  syncHistory();
  $("footer").classList.toggle("hidden", view !== "viewList" || !state.refs);
  if (view === "viewList") $("title").textContent = state.project === null ? "Projects" : (state.project || "(No project)");
  if (view === "viewList") $("search").placeholder = state.project === null ? "Search project or reference..." : "Search reference, supplier...";
  if (view === "viewSettings") $("title").textContent = "Settings";
  if (view === "viewReport") $("title").textContent = "Report";
  window.scrollTo(0, 0);
}

// ---------------- References ----------------
async function loadRefsFromFile(file) { return loadRefsFromText(await file.text()); }

// (0.11.0) Several refs_mobile_<name>.json at once (e.g. a whole team): one list with all their references
async function loadRefsFromFiles(files) {
  files = [...(files || [])];
  if (files.length <= 1) return files[0] ? loadRefsFromFile(files[0]) : undefined;
  const parts = [], bad = [];
  for (const f of files) {
    try { parts.push(parseRefsText(await f.text())); }
    catch (e) { bad.push(`${esc(f.name)}: ${esc(e.message)}`); }
  }
  if (!parts.length) return ask(`<b>The files could not be loaded.</b><br>${bad.join("<br>")}`, { cancel: "" });
  await saveRefs(mergeRefs(parts), bad);
}

// Joins several refs files: same reference in two files = the newest file wins
function mergeRefs(parts) {
  parts.sort((x, y) => String(x.created || "").localeCompare(String(y.created || "")));
  const byId = new Map(), codes = {}, owners = [];
  let week = "", version = 0;
  for (const d of parts) {
    for (const r of d.refs) byId.set(r.id, Object.assign({ owner: d.owner || "" }, r));
    if (d.codes) Object.assign(codes, d.codes);
    if (d.week && String(d.week) > week) week = String(d.week);
    if (d.owner && !owners.includes(String(d.owner))) owners.push(String(d.owner));
    version = Math.max(version, Number(d.version) || 0);
  }
  const m = Object.assign({}, parts[parts.length - 1]);
  m.refs = [...byId.values()];
  m.codes = codes; m.version = version; m.files = parts.length;
  if (week) m.week = isNaN(Number(week)) ? week : Number(week);
  m.owner = owners.length === 1 ? owners[0] : owners.join(" + ");
  m.owners = owners;
  return m;
}

// Copy / paste: the text of refs_mobile_<name>.json (e-mail, Notes, Teams...). Anything around the JSON is ignored.
async function pasteRefs() {
  let t = "";
  try { if (navigator.clipboard && navigator.clipboard.readText) t = await navigator.clipboard.readText(); } catch (e) { t = ""; }
  if (!t || t.indexOf('"refs"') < 0) {
    t = await ask("<b>Paste the references</b><br>Long-press in the box and choose <i>Paste</i>.", { area: true, ok: "Load" });
    if (!t) return;
  }
  await loadRefsFromText(t);
}

// Excel (VBA Str) writes 0.85 as ".85" and -0.5 as "-.5", which is not valid JSON (quality %).
// Outside the texts in quotes, a "." without a digit before it gets a "0".
function fixVbaNumbers(t) {
  let out = "", inStr = false;
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (inStr) {
      out += ch;
      if (ch === "\\") { out += t[i + 1] || ""; i++; }
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === "." && /\d/.test(t[i + 1] || "") && !/\d/.test(t[i - 1] || "")) out += "0";
    out += ch;
  }
  return out;
}
// JSON.parse; where the browser gives the position (Chrome, Edge), show the text around it
function parseRefsJson(t) {
  try { return JSON.parse(t); }
  catch (e) {
    const m = /position (\d+)/.exec(e.message);
    if (!m) throw e;
    const p = Number(m[1]);
    throw new Error(`${e.message} - near: ${t.slice(Math.max(0, p - 60), p + 40).replace(/data:image[^"]*/g, "(picture)")}`);
  }
}

function parseRefsText(text) {
  text = String(text || "").replace(/^\uFEFF/, "");
  const a = text.indexOf("{"), b = text.lastIndexOf("}");
  if (a < 0 || b < a) throw new Error("No references found in the text.");
  const data = parseRefsJson(fixVbaNumbers(text.slice(a, b + 1)));
  if (!data || !Array.isArray(data.refs)) throw new Error("This is not a Dashboard Timing references file.");
  data.refs = data.refs.filter(r => r && r.id);
  return data;
}

async function saveRefs(data, bad) {
  await db.put("kv", data, "refs");
  state.refs = data;
  state.project = null;
  const owners = data.owners || (data.owner ? [String(data.owner)] : []);
  if (owners.length === 1) {                          // one person's file: that is the author of the comments
    state.author = owners[0];
    await db.put("kv", state.author, "author");
    $("txtAuthor").value = state.author;
  }
  if (bad && bad.length) await ask(`<b>${data.refs.length} references loaded from ${data.files} files.</b><br>Not loaded:<br>${bad.join("<br>")}`, { cancel: "" });
  else toast(data.files > 1 ? `${data.refs.length} references loaded from ${data.files} files` : `${data.refs.length} references loaded`);
  await renderList();
  show("viewList");
}

async function loadRefsFromText(text) {
  try { await saveRefs(parseRefsText(text)); }
  catch (e) {
    await ask(`<b>The file could not be loaded.</b><br>${esc(e.message)}`, { cancel: "" });
  }
}

async function refreshDrafts() {
  state.drafts.clear();
  state.photoCounts.clear();
  for (const d of await db.getAll("drafts")) state.drafts.set(d.refId, d);
  for (const p of await db.getAll("photos")) state.photoCounts.set(p.refId, (state.photoCounts.get(p.refId) || 0) + 1);
}

async function renderList() {
  await refreshDrafts();
  const has = !!(state.refs && state.refs.refs.length);
  $("noRefs").classList.toggle("hidden", has);
  $("listTools").classList.toggle("hidden", !has);
  $("btnFilter").classList.toggle("hidden", !has || $("viewList").classList.contains("hidden"));
  // Old file with everybody's references (before the files per person): ask for the personal one
  const nOwners = has ? new Set(state.refs.refs.map(r => r.owner).filter(Boolean)).size : 0;
  $("oldFile").classList.toggle("hidden", !has || !!state.refs.owner || nOwners < 2);
  $("footer").classList.toggle("hidden", !has || !$("viewRef").classList.contains("hidden") || !$("viewSettings").classList.contains("hidden"));
  const ul = $("refList");
  ul.innerHTML = "";
  if (!has) return;
  renderActiveFilters();
  if (!$("fPanel").classList.contains("hidden")) renderFilterPanel();
  const q = $("search").value.trim().toLowerCase();
  const only = $("onlyDraft").checked;
  let n = 0;
  const proj = r => String(r.project || "").trim();

  // First screen: one line per project (searching shows the matching references of every project)
  if (state.project === null && !q) {
    const groups = new Map();
    for (const r of state.refs.refs) {
      if (!groups.has(proj(r))) groups.set(proj(r), []);
      groups.get(proj(r)).push(r);
    }
    const names = [...groups.keys()].sort((a, b) => (a || "~").localeCompare(b || "~"));
    for (const name of names) {
      const total = groups.get(name).length;
      const refs = groups.get(name).filter(r => passFilters(r));
      if (!refs.length) continue;
      const nRep = refs.filter(r => hasContent(state.drafts.get(r.id), state.photoCounts.get(r.id) || 0)).length;
      if (only && !nRep) continue;
      const th = refs.find(r => r.thumb);
      const sup = [...new Set(refs.map(r => r.supplier).filter(Boolean))];
      const li = document.createElement("li");
      li.innerHTML = `${th ? `<img class="thumb" src="${th.thumb}" alt="" loading="lazy">` : `<span class="thumb empty"></span>`}
        <div class="txt"><div class="name">${esc(name || "(No project)")}</div>
        <div class="sub">${refs.length < total ? `${refs.length} of ${total}` : refs.length} reference${total > 1 ? "s" : ""}${sup.length ? " · " + esc(sup.slice(0, 2).join(", ")) + (sup.length > 2 ? "..." : "") : ""}</div></div>
        ${nRep ? `<span class="badge">${nRep} ✎</span>` : ""}<span class="chev">&#8250;</span>`;
      li.onclick = () => { state.project = name; $("search").value = ""; renderList(); show("viewList"); };
      ul.appendChild(li);
      n++;
    }
    if (!n) ul.innerHTML = `<li class="muted">No projects match${nFilters() ? " the filters" : ""}.</li>`;
  }

  for (const r of (state.project === null && !q) ? [] : state.refs.refs) {
    if (state.project !== null && proj(r) !== state.project) continue;
    const d = state.drafts.get(r.id);
    const np = state.photoCounts.get(r.id) || 0;
    const withReport = hasContent(d, np);
    if (only && !withReport) continue;
    if (!passFilters(r)) continue;
    const hay = [r.no, r.pn, r.project, r.supplier].join(" ").toLowerCase();
    if (q && !hay.includes(q)) continue;
    const st = STATUS.find(s => s.code === Number((d && d.status) || r.status));
    const li = document.createElement("li");
    li.classList.add("refcard");
    li.innerHTML = `<button class="ibtn" type="button" aria-label="Information">i</button>
      ${r.thumb ? `<img class="thumb" src="${r.thumb}" alt="" loading="lazy">` : `<span class="thumb empty"></span>`}
      <span class="dot" style="background:${st ? st.color : "transparent"}"></span>
      <div class="txt"><div class="name">${esc(refName(r))}</div>
      ${state.project === null && r.project ? `<div class="sub">${esc(r.project)}</div>` : ""}${cardLines(r)}</div>
      ${withReport ? `<span class="badge">${np} 📷${d && d.status ? " · ●" : ""}${d && (d.comment || "").trim() ? " · ✎" : ""}</span>` : ""}`;
    li.onclick = () => openRef(r.id);
    li.querySelector(".ibtn").onclick = e => { e.stopPropagation(); showInfo(r); };
    ul.appendChild(li);
    n++;
  }
  if (!n && !(state.project === null && !q)) ul.innerHTML = `<li class="muted">No references match${nFilters() ? " the filters" : ""}.</li>`;
  const nRep = state.refs.refs.filter(r => hasContent(state.drafts.get(r.id), state.photoCounts.get(r.id) || 0)).length;
  $("btnExport").textContent = nRep ? `Send report (${nRep} reference${nRep > 1 ? "s" : ""})` : "Send report";
  $("btnExport").disabled = !nRep;
}

// ---------------- Filters (0.12.0, like the PC Project view: several values per field, Milestone) ----------------
const FILTER_KEY = "sf.filters";
const FIELDS = [
  ["project", "Project", r => r.project],
  ["supplier", "Supplier", r => r.supplier],
  ["plant", "Plant", r => r.plant],
  ["type", "Type", r => (r.info || {}).type],
  ["owner", "Person", r => r.owner],
  ["status", "Status", r => { const st = STATUS.find(s => s.code === Number(curStatus(r))); return st ? st.name : ""; }],
  ["material", "Material", r => (r.info || {}).material],
  ["press", "Press", r => (r.info || {}).press],
  ["product", "Product", r => (r.info || {}).product],
];
const MS_RANGES = [["last4", "Last 4 weeks", -4, -1], ["now", "This week", 0, 0], ["next2", "Next 2 weeks", 0, 1], ["next4", "Next 4 weeks", 0, 3], ["next8", "Next 8 weeks", 0, 7]];
let filt = loadFilters();
let fOpen = "";                                         // field open in the panel
function loadFilters() {
  try { const f = JSON.parse(localStorage.getItem(FILTER_KEY) || "null"); if (f && f.sel) return { sel: f.sel, ms: f.ms || [], range: f.range || "next4" }; } catch (e) {}
  return { sel: {}, ms: [], range: "next4" };
}
function saveFilters() { try { localStorage.setItem(FILTER_KEY, JSON.stringify(filt)); } catch (e) {} }
function curStatus(r) { const d = state.drafts.get(r.id); return (d && d.status) || r.status; }
const fval = (f, r) => String(f[2](r) ?? "").trim();
const weekPlus = n => isoKey(new Date(Date.now() + n * 7 * 86400000));
function msRange() { const g = MS_RANGES.find(x => x[0] === filt.range) || MS_RANGES[3]; return [weekPlus(g[2]), weekPlus(g[3]), g]; }
// Milestone like the PC: Actual for past weeks, the latest plan from this week on
function msCodes(r) {
  const t = r.timing, out = new Set();
  if (!t || !Array.isArray(t.w)) return out;
  const cw = curWeek(), [a, b] = msRange(), pl = planLine(t), act = lineCells(t, "Actual");
  t.w.forEach((k, i) => {
    if (k < a || k > b) return;
    const c = k < cw ? (act && act[i]) : (pl && pl[1][i]);
    if (c) out.add(String(c).toUpperCase());
  });
  return out;
}
function nFilters() { return Object.values(filt.sel).filter(v => v && v.length).length + (filt.ms.length ? 1 : 0); }
// Does reference r pass the filters (except field "skip")?
function passFilters(r, skip) {
  for (const f of FIELDS) {
    const v = filt.sel[f[0]];
    if (f[0] === skip || !v || !v.length) continue;
    if (!v.includes(fval(f, r).toLowerCase())) return false;
  }
  if (skip !== "ms" && filt.ms.length) { const c = msCodes(r); if (!filt.ms.some(x => c.has(x))) return false; }
  return true;
}
function scopeRefs() {
  const refs = (state.refs && state.refs.refs) || [];
  return state.project === null ? refs : refs.filter(r => String(r.project || "").trim() === state.project);
}
function toggleIn(arr, v) { const i = arr.indexOf(v); if (i >= 0) arr.splice(i, 1); else arr.push(v); return arr; }
async function filtersChanged() { saveFilters(); renderFilterPanel(); await renderList(); }

function renderFilterPanel() {
  const refs = scopeRefs();
  let h = "";
  for (const f of FIELDS) {
    if (f[0] === "project" && state.project !== null) continue;
    const sel = filt.sel[f[0]] || [];
    const all = new Map();                               // value (lowercase) -> [text, count]
    for (const r of refs) { const v = fval(f, r); if (!all.has(v.toLowerCase())) all.set(v.toLowerCase(), [v, 0]); }
    if (all.size < 2 && !sel.length) continue;           // one value only: nothing to filter
    for (const r of refs) if (passFilters(r, f[0])) all.get(fval(f, r).toLowerCase())[1]++;
    const vals = [...all.entries()]                       // (0.12.1) every value, also the ones the other filters leave at 0
      .sort((x, y) => f[0] === "status" ? (STATUS.findIndex(s => s.name === x[1][0]) - STATUS.findIndex(s => s.name === y[1][0])) : (x[1][0] || "~").localeCompare(y[1][0] || "~"));
    const chips = vals.map(([k, [t, n]]) => {
      const st = f[0] === "status" ? STATUS.find(s => s.name === t) : null;
      return `<button type="button" class="comp${sel.includes(k) ? " on" : ""}${n ? "" : " zero"}" data-f="${f[0]}" data-v="${esc(k)}">${st ? `<span class="sdot" style="background:${st.color}"></span>` : ""}${esc(t || "(Empty)")} <small>${n}</small></button>`;
    }).join("");
    const sum = sel.length ? (sel.length > 2 ? `${sel.length} selected` : sel.map(k => (all.get(k) || [k])[0] || "(Empty)").join(", ")) : `All (${all.size})`;
    h += `<div class="frow${fOpen === f[0] ? " open" : ""}" data-row="${f[0]}"><button type="button" class="fsum"><span>${esc(f[1])}</span><span class="fval${sel.length ? " on" : ""}">${esc(sum)}</span><span class="chev">&#8250;</span></button>
      <div class="comps">${chips}</div></div>`;
  }
  // Milestone
  const order = Object.keys((state.refs && state.refs.codes) || {}).map(c => c.toUpperCase());
  const cnt = new Map();
  for (const r of refs) if (passFilters(r, "ms")) for (const c of msCodes(r)) cnt.set(c, (cnt.get(c) || 0) + 1);
  const codes = [...new Set([...cnt.keys(), ...filt.ms])].sort((x, y) => (order.indexOf(x) + 1 || 999) - (order.indexOf(y) + 1 || 999) || x.localeCompare(y));
  const [a, b, g] = msRange();
  const msSum = filt.ms.length ? `${filt.ms.join(", ")} · ${g[1]}` : "All";
  h += `<div class="frow${fOpen === "ms" ? " open" : ""}" data-row="ms"><button type="button" class="fsum"><span>Milestone</span><span class="fval${filt.ms.length ? " on" : ""}">${esc(msSum)}</span><span class="chev">&#8250;</span></button>
    <div class="fms"><div class="muted small">When: ${wkText(a)}${b !== a ? " – " + wkText(b) : ""} (Actual for past weeks, plan from this week)</div>
    <div class="comps">${MS_RANGES.map(x => `<button type="button" class="comp${filt.range === x[0] ? " on" : ""}" data-range="${x[0]}">${x[1]}</button>`).join("")}</div>
    <div class="muted small" style="margin-top:8px">Code:</div>
    <div class="comps">${codes.length ? codes.map(c => `<button type="button" class="comp${filt.ms.includes(c) ? " on" : ""}" data-ms="${esc(c)}">${chip(c)} ${esc((codeInfo(c) || [""])[0])} <small>${cnt.get(c) || 0}</small></button>`).join("") : `<span class="muted small">No timing codes in these weeks.</span>`}</div></div></div>`;
  $("fBody").innerHTML = h;
  const n = refs.filter(r => passFilters(r)).length;
  $("fShow").textContent = `Show ${n} reference${n === 1 ? "" : "s"}`;
  $("fClear").disabled = !nFilters();
  for (const row of $("fBody").querySelectorAll(".frow")) row.querySelector(".fsum").onclick = () => { fOpen = fOpen === row.dataset.row ? "" : row.dataset.row; renderFilterPanel(); };
  for (const bt of $("fBody").querySelectorAll("[data-f]")) bt.onclick = () => { const f = bt.dataset.f; filt.sel[f] = toggleIn(filt.sel[f] || [], bt.dataset.v); filtersChanged(); };
  for (const bt of $("fBody").querySelectorAll("[data-ms]")) bt.onclick = () => { toggleIn(filt.ms, bt.dataset.ms); filtersChanged(); };
  for (const bt of $("fBody").querySelectorAll("[data-range]")) bt.onclick = () => { filt.range = bt.dataset.range; filtersChanged(); };
}
function openFilters(field) {
  if (field !== undefined) fOpen = field;
  renderFilterPanel();
  $("fPanel").style.top = Math.round(document.querySelector("header.bar").getBoundingClientRect().bottom) + "px";
  $("fPanel").classList.remove("hidden"); $("fBack").classList.remove("hidden"); syncHistory();
}
function closeFilters() { $("fPanel").classList.add("hidden"); $("fBack").classList.add("hidden"); syncHistory(); }

// ---------------- Back (0.13.0) ----------------
// The phone's back button / back gesture does the same as the app's arrow: it closes the open window
// (message, filters) or goes back one screen. Only on the Projects screen does it leave the app.
let popSkip = false;
function atRoot() {
  return !$("viewList").classList.contains("hidden") && state.project === null &&
    $("fPanel").classList.contains("hidden") && $("dlg").classList.contains("hidden");
}
function syncHistory() {
  if (popSkip) return;
  const guard = !!(history.state && history.state.sf);
  if (!atRoot() && !guard) history.pushState({ sf: 1 }, "");
  else if (atRoot() && guard) { popSkip = true; history.back(); }
}
async function backAction() {
  if (!$("dlg").classList.contains("hidden")) { $("dlgNo").click(); return; }
  if (!$("fPanel").classList.contains("hidden")) { closeFilters(); return; }
  if (atRoot()) return;
  if (!$("viewList").classList.contains("hidden")) {       // reference list -> projects
    state.project = null; $("search").value = "";
  } else {
    await flushTexts(); state.current = null;
  }
  await renderList(); show("viewList");
}
function goBack() { if (history.state && history.state.sf) history.back(); else backAction(); }
window.addEventListener("popstate", async () => {
  if (popSkip) { popSkip = false; syncHistory(); return; }
  await backAction();
  syncHistory();
});
// Swipe right from the left edge of the screen = back (for phones without a back gesture, e.g. iPhone app)
let swipe0 = null;
document.addEventListener("touchstart", e => {
  const t = e.touches[0];
  swipe0 = e.touches.length === 1 && t.clientX < 28 ? { x: t.clientX, y: t.clientY } : null;
}, { passive: true });
document.addEventListener("touchend", e => {
  if (!swipe0) return;
  const t = e.changedTouches[0], dx = t.clientX - swipe0.x, dy = Math.abs(t.clientY - swipe0.y);
  swipe0 = null;
  if (dx > 70 && dy < 60 && !atRoot()) goBack();
}, { passive: true });
// Active filters as small chips under the search box (x = remove)
function renderActiveFilters() {
  const n = nFilters();
  $("fCount").textContent = n; $("fCount").classList.toggle("hidden", !n);
  const box = $("fActive");
  const items = [];
  for (const f of FIELDS) {
    const v = filt.sel[f[0]];
    if (!v || !v.length) continue;
    const names = v.map(k => { const r = state.refs && state.refs.refs.find(x => fval(f, x).toLowerCase() === k); return r ? fval(f, r) : k; }).map(t => t || "(Empty)");
    items.push([f[0], `${f[1]}: ${names.length > 2 ? names.slice(0, 2).join(", ") + ` +${names.length - 2}` : names.join(", ")}`]);
  }
  if (filt.ms.length) items.push(["ms", `${filt.ms.join(", ")} · ${msRange()[2][1]}`]);
  box.innerHTML = items.map(([k, t]) => `<span class="fchip" data-k="${k}"><span class="ft">${esc(t)}</span><button type="button" aria-label="Remove">&times;</button></span>`).join("");
  box.classList.toggle("hidden", !items.length);
  for (const c of box.querySelectorAll(".fchip")) {
    c.querySelector(".ft").onclick = () => openFilters(c.dataset.k);
    c.querySelector("button").onclick = async () => { if (c.dataset.k === "ms") filt.ms = []; else delete filt.sel[c.dataset.k]; saveFilters(); await renderList(); };
  }
}

// ---------------- One reference ----------------
async function getDraft(refId) {
  return (await db.get("drafts", refId)) || { refId, status: null, comment: "", cfixture: "", material: "", updated: null };
}
async function saveDraft(d) {
  d.updated = new Date().toISOString();
  await db.put("drafts", d);
}

async function openRef(refId) {
  const r = state.refs.refs.find(x => x.id === refId);
  if (!r) return;
  state.current = refId;
  const d = await getDraft(refId);
  $("title").textContent = refName(r);
  $("btnInfoRef").onclick = () => showInfo(r);
  $("refName").textContent = refName(r);
  $("refThumb").classList.toggle("hidden", !r.thumb);
  if (r.thumb) $("refThumb").src = r.thumb; else $("refThumb").removeAttribute("src");
  const cur = STATUS.find(s => s.code === Number(r.status));
  $("refSub").innerHTML = cardLines(r) + `<div class="sub">In Excel: ${cur ? esc(cur.name) : "No status"}</div>`;
  renderRefData(r);
  renderStatus(d);
  $("txtCom").value = d.comment || "";
  $("txtCF").value = d.cfixture || "";
  $("txtMat").value = d.material || "";
  await renderPhotos();
  await renderBox("cf");
  await renderBox("mat");
  show("viewRef");
  scrollTimingNow();
}

function renderStatus(d) {
  const g = $("statusGrid");
  g.innerHTML = "";
  for (const s of STATUS_PICK) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = Number(d.status) === s.code ? "on" : "";
    b.innerHTML = `<span class="dot" style="background:${s.color}"></span>${esc(s.name)}`;
    b.onclick = async () => {
      const dd = await getDraft(state.current);
      dd.status = Number(dd.status) === s.code ? null : s.code;   // tap again = do not change in Excel
      await saveDraft(dd);
      renderStatus(dd);
    };
    g.appendChild(b);
  }
}

// Components added by the user on this phone (kept between sessions)
const CUSTOM_KEY = "sf.customComponents";
function customComponents() {
  try { const v = JSON.parse(localStorage.getItem(CUSTOM_KEY) || "[]"); return Array.isArray(v) ? v.filter(x => typeof x === "string" && x.trim()) : []; }
  catch { return []; }
}
function saveCustomComponents(list) {
  try { localStorage.setItem(CUSTOM_KEY, JSON.stringify(list)); } catch { }
}
function allComponents() {
  const extra = customComponents().filter(c => !COMPONENTS.includes(c));
  return [...COMPONENTS, ...extra].sort((a, b) => a.localeCompare(b));
}

// Like the PC: select photos (red), then tap their component. It is written on them.
const photoSel = { ref: null, ids: new Set() };

async function renderPhotos() {
  const list = (await db.toolPhotos(state.current)).sort((a, b) => a.taken.localeCompare(b.taken));
  if (photoSel.ref !== state.current) { photoSel.ref = state.current; photoSel.ids.clear(); }
  for (const id of [...photoSel.ids]) if (!list.some(p => p.id === id)) photoSel.ids.delete(id);
  $("photoCount").textContent = `(${list.length}/${MAX_PHOTOS})`;
  const box = $("photos");
  for (const img of box.querySelectorAll("img")) URL.revokeObjectURL(img.src);
  box.innerHTML = "";
  for (const p of list) {
    const div = document.createElement("div");
    const on = photoSel.ids.has(p.id);
    div.className = "ph sel" + (on ? " on" : "");
    div.innerHTML = `<img src="${URL.createObjectURL(p.blob)}" alt="">
      <div class="tag${p.component ? "" : " none"}">${esc(p.component || "No component")}</div>
      <div class="tick">${on ? "&#10003;" : ""}</div>`;
    div.onclick = () => {
      if (photoSel.ids.has(p.id)) photoSel.ids.delete(p.id); else photoSel.ids.add(p.id);
      renderPhotos();
    };
    box.appendChild(div);
  }
  renderCompPanel(list);
}

function renderCompPanel(list) {
  const pan = $("compPanel");
  if (!list.length) { pan.innerHTML = ""; return; }
  const sel = list.filter(p => photoSel.ids.has(p.id));
  const n = sel.length;
  const common = n && sel.every(p => p.component === sel[0].component) ? sel[0].component : "";
  const custom = customComponents();
  let h = `<div class="row small">
      <button type="button" class="btn grow" data-a="all">Select all</button>
      <button type="button" class="btn grow" data-a="none"${n ? "" : " disabled"}>Clear</button>
      <button type="button" class="btn grow danger" data-a="del"${n ? "" : " disabled"}>Remove</button>
    </div>`;
  if (!n) h += `<p class="muted small">Tap the photos to select them, then choose their component.</p>`;
  else {
    h += `<p class="small"><b>${n} selected.</b> Choose their component in the COMPONENTS LIST:</p><div class="comps">`;
    for (const c of allComponents()) {
      const mine = custom.includes(c) && !COMPONENTS.includes(c);
      h += `<button type="button" class="comp${c === common ? " on" : ""}" data-c="${esc(c)}">${esc(c)}${mine ? ` <span class="x" data-x="${esc(c)}" title="Remove from list">&times;</span>` : ""}</button>`;
    }
    h += `<button type="button" class="comp add" data-a="add">+ New component...</button>`;
    if (sel.some(p => p.component)) h += `<button type="button" class="comp clr" data-a="clr">No component</button>`;
    h += `</div>`;
  }
  pan.innerHTML = h;
  pan.onclick = async ev => {
    const x = ev.target.closest("[data-x]");
    if (x) {
      ev.stopPropagation();
      const c = x.dataset.x;
      if (await ask(`Remove <b>${esc(c)}</b> from the list?<br>(Photos that already have it keep it.)`, { ok: "Remove" })) {
        saveCustomComponents(customComponents().filter(v => v !== c));
        renderCompPanel(list);
      }
      return;
    }
    const b = ev.target.closest("button");
    if (!b || b.disabled) return;
    const a = b.dataset.a;
    if (a === "all") { for (const p of list) photoSel.ids.add(p.id); return renderPhotos(); }
    if (a === "none") { photoSel.ids.clear(); return renderPhotos(); }
    if (a === "del") {
      if (await ask(n === 1 ? "Remove this photo from the report?" : `Remove these ${n} photos from the report?`, { ok: "Remove" })) {
        for (const p of sel) await db.del("photos", p.id);
        photoSel.ids.clear();
        await renderPhotos();
      }
      return;
    }
    let comp;
    if (a === "add") {
      const t = await ask("Component name:", { input: true, ok: "Add" });
      comp = (t || "").trim().replace(/\s+/g, " ");
      if (!comp) return;
      const known = allComponents().find(c => c.toLowerCase() === comp.toLowerCase());
      if (known) comp = known;
      else saveCustomComponents([...customComponents(), comp]);
    } else if (a === "clr") comp = "";
    else if (b.dataset.c !== undefined) comp = b.dataset.c;
    else return;
    for (const p of sel) { p.component = comp; await db.put("photos", p); }
    photoSel.ids.clear();                      // done: the component is now written on them
    await renderPhotos();
    toast(comp ? `${n} photo(s): ${comp}` : `${n} photo(s) without component`);
  };
}

// Reduce the photo (keeps EXIF orientation) -> JPEG Blob
async function shrink(file) {
  let src;
  try { src = await createImageBitmap(file, { imageOrientation: "from-image" }); }
  catch {
    src = await new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = () => rej(new Error("Not an image"));
      i.src = URL.createObjectURL(file);
    });
  }
  const w0 = src.width, h0 = src.height;
  const f = Math.min(1, PHOTO_MAX_SIDE / Math.max(w0, h0));
  const c = document.createElement("canvas");
  c.width = Math.round(w0 * f);
  c.height = Math.round(h0 * f);
  c.getContext("2d").drawImage(src, 0, 0, c.width, c.height);
  if (src.close) src.close();
  return new Promise(res => c.toBlob(res, "image/jpeg", PHOTO_QUALITY));
}

async function addPhotos(files) {
  if (!files || !files.length) return;
  const refId = state.current;
  let n = (await db.toolPhotos(refId)).length;
  let skipped = 0;
  for (const file of files) {
    if (n >= MAX_PHOTOS) { skipped++; continue; }
    try {
      const blob = await shrink(file);
      const taken = new Date(file.lastModified || Date.now()).toISOString();
      await db.put("photos", { id: `${refId}|${Date.now()}|${Math.random().toString(36).slice(2, 8)}`, refId, kind: "tool", blob, component: "", taken });
      n++;
    } catch (e) { skipped++; }
  }
  await renderPhotos();
  if (skipped) toast(`${skipped} photo(s) not added (maximum ${MAX_PHOTOS} per reference)`);
}

// C. Fixture / Material: one photo per box (CF PICTURE / MAT PICTURE of the Status report menu)
async function renderBox(kind) {
  const box = $(kind === "cf" ? "boxCF" : "boxMat");
  const p = await db.boxPhoto(state.current, kind);
  const old = box.querySelector("img");
  if (old) URL.revokeObjectURL(old.src);
  box.innerHTML = p
    ? `<div class="ph"><img src="${URL.createObjectURL(p.blob)}" alt=""><button class="del" type="button">Remove photo</button></div>`
    : `<div class="row"><label class="btn grow">&#128247; Photo<input type="file" accept="image/*" capture="environment" hidden></label>
       <label class="btn grow">&#128444; Gallery<input type="file" accept="image/*" hidden></label></div>`;
  if (p) {
    box.querySelector(".del").onclick = async () => {
      if (await ask("Remove this photo?", { ok: "Remove" })) { await db.del("photos", p.id); await renderBox(kind); }
    };
  } else {
    for (const inp of box.querySelectorAll("input")) inp.onchange = async e => {
      const f = e.target.files[0];
      if (!f) return;
      try {
        const blob = await shrink(f);
        await db.put("photos", { id: `${state.current}|${kind}|${Date.now()}`, refId: state.current, kind, blob, component: "", taken: new Date(f.lastModified || Date.now()).toISOString() });
      } catch (err) { toast("The photo could not be added"); }
      await renderBox(kind);
    };
  }
}

// ---------------- Export ----------------
const pad = n => String(n).padStart(2, "0");
const stamp = d => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
const safe = s => String(s || "").replace(/[^A-Za-z0-9._-]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 40);
// Text for report.txt: one line, tab separated. Line breaks as |n| (same as the Dashboard registry texts)
const cell = s => String(s ?? "").replace(/\t/g, " ").replace(/\r\n|\r|\n/g, "|n|");

async function exportReport() {
  if (!state.author) {
    const name = await ask("<b>Your name</b><br>It is saved as \"By\" in the comments (as in the Dashboard).", { input: true, value: "" });
    if (!name || !name.trim()) return;
    state.author = name.trim();
    await db.put("kv", state.author, "author");
    $("txtAuthor").value = state.author;
  }
  await refreshDrafts();
  const refs = state.refs.refs.filter(r => hasContent(state.drafts.get(r.id), state.photoCounts.get(r.id) || 0));
  if (!refs.length) return;

  const noComp = [];
  for (const r of refs) {
    for (const p of await db.toolPhotos(r.id)) if (!p.component) { noComp.push(r); break; }
  }
  if (noComp.length && !(await ask(`<b>Photos without component</b><br>${noComp.map(r => esc(refName(r))).join("<br>")}<br><br>Send anyway? (You can choose it later in Excel.)`, { ok: "Send" }))) return;

  const now = new Date();
  const enc = new TextEncoder();
  const files = [];
  const json = { format: "DashboardTiming-mobile", version: 1, app: APP_VERSION, author: state.author, created: now.toISOString(), refs: [] };
  const lines = [["DTMOBILE", "1", cell(state.author), now.toISOString(), APP_VERSION].join("\t")];

  for (const r of refs) {
    const d = state.drafts.get(r.id) || {};
    const photos = (await db.toolPhotos(r.id)).sort((a, b) => a.taken.localeCompare(b.taken));
    const jr = { id: r.id, status: d.status || null, comment: d.comment || "", cfixture: d.cfixture || "", material: d.material || "", photos: [] };
    lines.push(["REF", r.id, d.status || "", cell(d.comment), cell(d.cfixture), cell(d.material)].join("\t"));
    let k = 0;
    for (const p of photos) {
      k++;
      const name = `photos/${safe(r.id)}/${safe(r.id)}_${stamp(new Date(p.taken))}_${pad(k)}${p.component ? "_" + safe(p.component) : ""}.jpg`;
      files.push({ name, data: new Uint8Array(await p.blob.arrayBuffer()) });
      jr.photos.push({ file: name, component: p.component || "", taken: p.taken });
      lines.push(["PHOTO", r.id, name, cell(p.component), p.taken].join("\t"));
    }
    for (const [kind, tag] of [["cf", "CFPHOTO"], ["mat", "MATPHOTO"]]) {
      const p = await db.boxPhoto(r.id, kind);
      if (!p) continue;
      const name = `photos/${safe(r.id)}/${safe(r.id)}_${kind === "cf" ? "CFixture" : "Material"}_${stamp(new Date(p.taken))}.jpg`;
      files.push({ name, data: new Uint8Array(await p.blob.arrayBuffer()) });
      jr[kind === "cf" ? "cfixturePhoto" : "materialPhoto"] = name;
      lines.push([tag, r.id, name, "", p.taken].join("\t"));
    }
    json.refs.push(jr);
  }
  files.unshift(
    { name: "report.txt", data: enc.encode("﻿" + lines.join("\r\n") + "\r\n") },
    { name: "report.json", data: enc.encode(JSON.stringify(json, null, 2)) },
  );
  const fileName = `StatusMobile_${safe(state.author)}_${stamp(now)}.zip`;
  const blob = MiniZip.makeZip(files);
  const file = new File([blob], fileName, { type: "application/zip" });

  let sent = false;
  try {
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: fileName, text: `Status report from the shop floor (${refs.length} references)` });
      sent = true;
    }
  } catch (e) {
    if (e && e.name === "AbortError") return;      // the user closed the share menu
  }
  if (!sent) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  const mb = (blob.size / 1048576).toFixed(1);
  const how = sent ? "" : `<br><br>Saved in <b>Downloads</b>. To send it by e-mail or Bluetooth: open <i>Files &gt; Downloads</i>, hold the file and tap <i>Share</i>.`;
  if (await ask(`<b>Report ready</b> (${fileName}, ${mb} MB).${how}<br><br>Once it has arrived, clear these ${refs.length} report(s) from the phone?`, { ok: "Clear", cancel: "Keep" })) {
    for (const r of refs) await clearRef(r.id);
    await renderList();
  }
}

async function clearRef(refId) {
  await db.del("drafts", refId);
  for (const p of await db.byRef(refId)) await db.del("photos", p.id);
}

// ---------------- View a report (.zip) ----------------
// The same file that "Mobile Import" loads in Excel: report.txt + photos/
async function readZip(blob) {
  const buf = new Uint8Array(await blob.arrayBuffer());
  const dv = new DataView(buf.buffer);
  let e = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
  if (e < 0) throw new Error("This is not a .zip file.");
  const n = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const dec = new TextDecoder();
  const files = new Map();
  for (let k = 0; k < n; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("Damaged .zip file.");
    const method = dv.getUint16(p + 10, true), size = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    const name = dec.decode(buf.subarray(p + 46, p + 46 + nl)).replace(/\\/g, "/");
    const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
    const raw = buf.subarray(start, start + size);
    files.set(name, { method, raw });
    p += 46 + nl + xl + cl;
  }
  return {
    has: name => files.has(name),
    async bytes(name) {
      const f = files.get(name);
      if (!f) return null;
      if (f.method === 0) return f.raw;
      if (f.method === 8 && "DecompressionStream" in window) {
        const ds = new Blob([f.raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        return new Uint8Array(await new Response(ds).arrayBuffer());
      }
      throw new Error("Unsupported compression in the .zip file.");
    },
  };
}

let reportUrls = [];
async function openReport(file) {
  try {
    const z = await readZip(file);
    const txt = await z.bytes("report.txt");
    if (!txt) throw new Error("This is not a Status report from the app (report.txt missing).");
    const lines = new TextDecoder().decode(txt).replace(/^\uFEFF/, "").split(/\r?\n/).map(l => l.split("\t"));
    if (lines[0][0] !== "DTMOBILE") throw new Error("Unknown report format.");
    const unc = t => String(t || "").split("|n|").join("\n");
    const refs = new Map();
    const get = id => { if (!refs.has(id)) refs.set(id, { id, status: "", comment: "", cf: "", mat: "", photos: [], cfPhoto: null, matPhoto: null }); return refs.get(id); };
    for (const c of lines.slice(1)) {
      if (c.length < 2) continue;
      const r = get(c[1]);
      if (c[0] === "REF") Object.assign(r, { status: c[2] || "", comment: unc(c[3]), cf: unc(c[4]), mat: unc(c[5]) });
      if (c[0] === "PHOTO") r.photos.push({ file: c[2], component: unc(c[3]) });
      if (c[0] === "CFPHOTO") r.cfPhoto = c[2];
      if (c[0] === "MATPHOTO") r.matPhoto = c[2];
    }
    for (const u of reportUrls) URL.revokeObjectURL(u);
    reportUrls = [];
    const img = async f => {
      const b = f && await z.bytes(f);
      if (!b) return "";
      const u = URL.createObjectURL(new Blob([b], { type: "image/jpeg" }));
      reportUrls.push(u);
      return u;
    };
    const known = id => (state.refs && state.refs.refs.find(r => r.id === id)) || null;
    let html = `<div class="card"><b>${esc(file.name)}</b><div class="muted small">By ${esc(lines[0][2] || "?")} · ${esc(new Date(lines[0][3]).toLocaleString())} · ${refs.size} reference(s)</div></div>`;
    for (const r of refs.values()) {
      const k = known(r.id);
      const st = STATUS.find(s => s.code === Number(r.status));
      let ph = "";
      for (const p of r.photos) ph += `<div class="ph"><img src="${await img(p.file)}" alt=""><div class="cap">${esc(p.component || "No component")}</div></div>`;
      const box = async (label, t, f) => (t.trim() || f) ? `<h3>${label}</h3>${t.trim() ? `<p class="pre">${esc(t)}</p>` : ""}${f ? `<div class="ph boxph"><img src="${await img(f)}" alt=""></div>` : ""}` : "";
      html += `<div class="card rep">
        <div class="rephead">${k && k.thumb ? `<img class="thumb" src="${k.thumb}" alt="">` : ""}
          <div><div class="refname">${esc(k ? refName(k) : r.id)}</div>${k ? cardLines(k) : ""}</div></div>
        <p>${st ? `<span class="pill"><span class="dot" style="background:${st.color}"></span>${esc(st.name)}</span>` : `<span class="muted small">Status not changed</span>`}</p>
        ${r.comment.trim() ? `<h3>Comments</h3><p class="pre">${esc(r.comment)}</p>` : ""}
        ${ph ? `<h3>Tool advance (${r.photos.length})</h3><div class="photos">${ph}</div>` : ""}
        ${await box("C. Fixture", r.cf, r.cfPhoto)}${await box("Material", r.mat, r.matPhoto)}
      </div>`;
    }
    $("reportBody").innerHTML = html;
    show("viewReport");
  } catch (e) {
    await ask(`<b>The report could not be opened.</b><br>${esc(e.message)}`, { cancel: "" });
  }
}

// File shared to the app from another app (Android share menu, see sw.js)
async function openShared() {
  if (!location.search.includes("shared=1") || !("caches" in window)) return;
  history.replaceState(null, "", location.pathname);
  try {
    const c = await caches.open("status-field-shared");
    const r = await c.match("./shared-file");
    if (!r) return;
    await c.delete("./shared-file");
    const name = decodeURIComponent(r.headers.get("X-File-Name") || "shared");
    const blob = await r.blob();
    const head = new Uint8Array(await blob.slice(0, 2).arrayBuffer());
    if (head[0] === 0x50 && head[1] === 0x4b) await openReport(new File([blob], name));   // "PK" = .zip report
    else await loadRefsFromText(await blob.text());                                       // references .json
  } catch (e) {
    await ask(`<b>The shared file could not be opened.</b><br>${esc(e.message)}`, { cancel: "" });
  }
}

// ---------------- Wiring ----------------
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
async function flushTexts() {
  if (!state.current) return;
  const d = await getDraft(state.current);
  d.comment = $("txtCom").value; d.cfixture = $("txtCF").value; d.material = $("txtMat").value;
  await saveDraft(d);
}
const saveTexts = debounce(flushTexts, 300);

async function init() {
  $("appVersion").textContent = `Version ${APP_VERSION}`;
  state.refs = (await db.get("kv", "refs")) || null;
  state.author = (await db.get("kv", "author")) || "";
  $("txtAuthor").value = state.author;

  $("fileRefs").onchange = e => { loadRefsFromFiles(e.target.files); e.target.value = ""; };
  $("fileRefs2").onchange = e => { loadRefsFromFiles(e.target.files); e.target.value = ""; };
  $("btnPaste").onclick = () => pasteRefs();
  $("btnPaste2").onclick = () => pasteRefs();
  $("fileReport").onchange = async e => { if (e.target.files[0]) await openReport(e.target.files[0]); e.target.value = ""; };
  $("fileRefs3").onchange = e => { loadRefsFromFiles(e.target.files); e.target.value = ""; };
  $("btnDemo").onclick = async () => { await db.put("kv", DEMO_REFS, "refs"); state.refs = DEMO_REFS; await renderList(); show("viewList"); };
  $("search").oninput = () => renderList();
  $("onlyDraft").onchange = () => renderList();
  $("btnFilter").onclick = () => $("fPanel").classList.contains("hidden") ? openFilters() : closeFilters();
  $("fBack").onclick = closeFilters;
  $("fShow").onclick = closeFilters;
  $("fClear").onclick = () => { filt.sel = {}; filt.ms = []; filtersChanged(); };
  $("btnBack").onclick = goBack;
  $("btnSettings").onclick = () => {
    $("refsInfo").textContent = state.refs ? `${state.refs.refs.length} references${state.refs.files > 1 ? ` from ${state.refs.files} files (${state.refs.owner})` : ""}${state.refs.created ? ", file of " + new Date(state.refs.created).toLocaleString() : ""}` : "No references loaded.";
    show("viewSettings");
  };
  $("txtAuthor").onchange = async () => { state.author = $("txtAuthor").value.trim(); await db.put("kv", state.author, "author"); };
  for (const id of ["txtCom", "txtCF", "txtMat"]) $(id).oninput = saveTexts;
  $("filePhotoCam").onchange = async e => { await addPhotos([...e.target.files]); e.target.value = ""; };
  $("filePhotoGal").onchange = async e => { await addPhotos([...e.target.files]); e.target.value = ""; };
  $("btnClearRef").onclick = async () => {
    if (await ask("Delete the photos, status and texts of this reference from the phone?", { ok: "Delete" })) {
      await clearRef(state.current);
      await openRef(state.current);
    }
  };
  $("btnClearAll").onclick = async () => {
    if (await ask("Delete ALL reports (photos and texts) stored on this phone?", { ok: "Delete all" })) {
      await db.clear("drafts"); await db.clear("photos");
      toast("All reports deleted");
    }
  };
  $("btnExport").onclick = () => exportReport().catch(e => ask(`<b>The report could not be created.</b><br>${esc(e.message)}`, { cancel: "" }));

  await renderList();
  show("viewList");
  await openShared();

  // Offline + install (only over http(s); opened as a file it works without it)
  if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
}

init().catch(e => { document.body.insertAdjacentHTML("beforeend", `<p style="padding:16px;color:#b42318">Error: ${esc(e.message)}</p>`); });

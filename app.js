// Status Report Field - PWA to prepare the Dashboard Timing Status report from the shop floor.
// Everything is stored on the phone (IndexedDB) until it is sent as a .zip that the Excel macro
// "Mobile Import" (modMovil) loads into the Status report.
"use strict";

const APP_VERSION = "0.8.0";
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
  return `<div class="sub">${r.no ? "PN " + esc(r.no) : ""}</div>
    <div class="sub">${esc([r.supplier, r.plant].filter(Boolean).join(" · "))}</div>
    <div class="sub">Quality: ${q ? esc(q) : "–"}</div>`;
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
    ["Width", i.width], ["Pitch", i.pitch], ["Press", i.press], ["Weight", i.weight], ["Last change", i.lastChange],
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
    $("dlg").classList.remove("hidden");
    if (input) setTimeout(() => $("dlgInput").focus(), 50);
    if (area) setTimeout(() => $("dlgArea").focus(), 50);
    const done = v => { $("dlg").classList.add("hidden"); $("dlgYes").onclick = $("dlgNo").onclick = null; res(v); };
    $("dlgYes").onclick = () => done(area ? $("dlgArea").value : input ? $("dlgInput").value : true);
    $("dlgNo").onclick = () => done(input || area ? null : false);
  });
}

function show(view) {
  for (const v of ["viewList", "viewRef", "viewSettings", "viewReport"]) $(v).classList.toggle("hidden", v !== view);
  $("btnBack").classList.toggle("hidden", view === "viewList" && state.project === null);
  $("btnSettings").classList.toggle("hidden", view !== "viewList" || state.project !== null);
  $("footer").classList.toggle("hidden", view !== "viewList" || !state.refs);
  if (view === "viewList") $("title").textContent = state.project === null ? "Projects" : (state.project || "(No project)");
  if (view === "viewList") $("search").placeholder = state.project === null ? "Search project or reference..." : "Search reference, supplier...";
  if (view === "viewSettings") $("title").textContent = "Settings";
  if (view === "viewReport") $("title").textContent = "Report";
  window.scrollTo(0, 0);
}

// ---------------- References ----------------
async function loadRefsFromFile(file) { return loadRefsFromText(await file.text()); }

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

async function loadRefsFromText(text) {
  try {
    text = String(text || "").replace(/^\uFEFF/, "");
    const a = text.indexOf("{"), b = text.lastIndexOf("}");
    if (a < 0 || b < a) throw new Error("No references found in the text.");
    const data = JSON.parse(text.slice(a, b + 1));
    if (!data || !Array.isArray(data.refs)) throw new Error("This is not a Dashboard Timing references file.");
    data.refs = data.refs.filter(r => r && r.id);
    await db.put("kv", data, "refs");
    state.refs = data;
    state.project = null;
    if (data.owner) {                                   // each person receives only their own file
      state.author = String(data.owner);
      await db.put("kv", state.author, "author");
      $("txtAuthor").value = state.author;
    }
    toast(`${data.refs.length} references loaded`);
    await renderList();
    show("viewList");
  } catch (e) {
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
  // Old file with everybody's references (before the files per person): ask for the personal one
  const nOwners = has ? new Set(state.refs.refs.map(r => r.owner).filter(Boolean)).size : 0;
  $("oldFile").classList.toggle("hidden", !has || !!state.refs.owner || nOwners < 2);
  $("footer").classList.toggle("hidden", !has || !$("viewRef").classList.contains("hidden") || !$("viewSettings").classList.contains("hidden"));
  const ul = $("refList");
  ul.innerHTML = "";
  if (!has) return;
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
      const refs = groups.get(name);
      const nRep = refs.filter(r => hasContent(state.drafts.get(r.id), state.photoCounts.get(r.id) || 0)).length;
      if (only && !nRep) continue;
      const th = refs.find(r => r.thumb);
      const sup = [...new Set(refs.map(r => r.supplier).filter(Boolean))];
      const li = document.createElement("li");
      li.innerHTML = `${th ? `<img class="thumb" src="${th.thumb}" alt="" loading="lazy">` : `<span class="thumb empty"></span>`}
        <div class="txt"><div class="name">${esc(name || "(No project)")}</div>
        <div class="sub">${refs.length} reference${refs.length > 1 ? "s" : ""}${sup.length ? " · " + esc(sup.slice(0, 2).join(", ")) + (sup.length > 2 ? "..." : "") : ""}</div></div>
        ${nRep ? `<span class="badge">${nRep} ✎</span>` : ""}<span class="chev">&#8250;</span>`;
      li.onclick = () => { state.project = name; $("search").value = ""; renderList(); show("viewList"); };
      ul.appendChild(li);
      n++;
    }
    if (!n) ul.innerHTML = `<li class="muted">No projects match.</li>`;
  }

  for (const r of (state.project === null && !q) ? [] : state.refs.refs) {
    if (state.project !== null && proj(r) !== state.project) continue;
    const d = state.drafts.get(r.id);
    const np = state.photoCounts.get(r.id) || 0;
    const withReport = hasContent(d, np);
    if (only && !withReport) continue;
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
  if (!n && !(state.project === null && !q)) ul.innerHTML = `<li class="muted">No references match.</li>`;
  const nRep = state.refs.refs.filter(r => hasContent(state.drafts.get(r.id), state.photoCounts.get(r.id) || 0)).length;
  $("btnExport").textContent = nRep ? `Send report (${nRep} reference${nRep > 1 ? "s" : ""})` : "Send report";
  $("btnExport").disabled = !nRep;
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
  renderStatus(d);
  $("txtCom").value = d.comment || "";
  $("txtCF").value = d.cfixture || "";
  $("txtMat").value = d.material || "";
  await renderPhotos();
  await renderBox("cf");
  await renderBox("mat");
  show("viewRef");
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

async function renderPhotos() {
  const list = (await db.toolPhotos(state.current)).sort((a, b) => a.taken.localeCompare(b.taken));
  $("photoCount").textContent = `(${list.length}/${MAX_PHOTOS})`;
  const box = $("photos");
  for (const img of box.querySelectorAll("img")) URL.revokeObjectURL(img.src);
  box.innerHTML = "";
  for (const p of list) {
    const div = document.createElement("div");
    div.className = "ph";
    const url = URL.createObjectURL(p.blob);
    const isManual = p.component && !COMPONENTS.includes(p.component);
    div.innerHTML = `<img src="${url}" alt="">
      <select>${["", ...COMPONENTS, MANUAL].map(c =>
        `<option value="${esc(c)}"${(c === p.component || (isManual && c === MANUAL)) ? " selected" : ""}>${c ? esc(c) : "Component..."}</option>`).join("")}</select>
      <input type="text" placeholder="Component" value="${isManual ? esc(p.component) : ""}" class="${isManual ? "" : "hidden"}">
      <button class="del" type="button">Remove</button>`;
    const sel = div.querySelector("select"), inp = div.querySelector("input");
    sel.onchange = async () => {
      if (sel.value === MANUAL) { inp.classList.remove("hidden"); inp.focus(); p.component = inp.value.trim(); }
      else { inp.classList.add("hidden"); p.component = sel.value; }
      await db.put("photos", p);
    };
    inp.onchange = async () => { p.component = inp.value.trim(); await db.put("photos", p); };
    div.querySelector(".del").onclick = async () => {
      if (await ask("Remove this photo from the report?", { ok: "Remove" })) {
        await db.del("photos", p.id);
        await renderPhotos();
      }
    };
    box.appendChild(div);
  }
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

  $("fileRefs").onchange = e => e.target.files[0] && loadRefsFromFile(e.target.files[0]);
  $("fileRefs2").onchange = e => e.target.files[0] && loadRefsFromFile(e.target.files[0]);
  $("btnPaste").onclick = () => pasteRefs();
  $("btnPaste2").onclick = () => pasteRefs();
  $("fileReport").onchange = async e => { if (e.target.files[0]) await openReport(e.target.files[0]); e.target.value = ""; };
  $("fileRefs3").onchange = e => e.target.files[0] && loadRefsFromFile(e.target.files[0]);
  $("btnDemo").onclick = async () => { await db.put("kv", DEMO_REFS, "refs"); state.refs = DEMO_REFS; await renderList(); show("viewList"); };
  $("search").oninput = () => renderList();
  $("onlyDraft").onchange = () => renderList();
  $("btnBack").onclick = async () => {
    if (!$("viewList").classList.contains("hidden")) {       // reference list -> projects
      state.project = null; $("search").value = "";
    } else {
      await flushTexts(); state.current = null;
    }
    await renderList(); show("viewList");
  };
  $("btnSettings").onclick = () => {
    $("refsInfo").textContent = state.refs ? `${state.refs.refs.length} references${state.refs.created ? ", file of " + new Date(state.refs.created).toLocaleString() : ""}` : "No references loaded.";
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

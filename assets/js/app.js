/* HH Signing Package Builder – app.js (repo-aware paths, global mapping)
 * Repo layout expected:
 *   /config/standard_pkg.json, /config/map.json
 *   /data/templates/*.pdf  (lender docs, headers_signing_package.pdf, etc.)
 *   /assets/js/app.js      (this file)
 *
 * Requires (loaded in HTML):
 *   - pdf-lib  (PDFLib global)
 *   - jszip    (JSZip global)
 */

//////////////////////////
// DOM ELEMENTS
//////////////////////////
const els = {
  status: document.getElementById("status"),
  lender: document.getElementById("lender"),
  pkg: document.getElementById("pkg"),
  upCommitment: document.getElementById("upCommitment"),
  upApplication: document.getElementById("upApplication"),
  upMPP: document.getElementById("upMPP"),
  upAPA: document.getElementById("upAPA"),
  apaWrap: document.getElementById("apaWrap"),
  btnPkg: document.getElementById("btnPkg"),
  btnSingles: document.getElementById("btnSingles"),
  btnInfo: document.getElementById("btnInfo"),
  manualJson: document.getElementById("manualJson"),
};

function setStatus(msg, cls = "muted") {
  if (!els.status) return;
  els.status.className = cls;
  els.status.textContent = msg;
}

// Small UX helper to prevent double-clicks and show progress text
async function withBusy(btn, runningLabel, fn){
  if (!btn) return fn();
  const prev = btn.textContent;
  btn.disabled = true;
  btn.textContent = runningLabel;
  try { return await fn(); }
  finally { btn.disabled = false; btn.textContent = prev; }
}

//////////////////////////
// PATHS
//////////////////////////
const _parts = location.pathname.split("/").filter(Boolean);
const REPO_BASE = _parts.length ? `/${_parts[0]}/` : "/";
const CONFIG_BASE = `${REPO_BASE}config/`;
const TEMPL_BASE  = `${REPO_BASE}data/templates/`;

//////////////////////////
// CONSTANTS
//////////////////////////
const LENDER_OPTIONS = ["TD","MCAP","CMLS","BEEM","CWB","FirstNat","Haventree","Lendwise","Scotia","Strive","Bridgewater","Private"];

const HEADER_INDEX = {
  WMB: 1,
  MPP_OR_PROSPR: 2,
  COMMITMENT: 3,
  COB: 4,
  FORM10: 5,
  APPLICATION: 6
};

const DEBUG_FILL = false;
function dbg(...a){ if (DEBUG_FILL) console.log(...a); }
function dwarn(...a){ if (DEBUG_FILL) console.warn(...a); }
function derr(...a){ if (DEBUG_FILL) console.error(...a); }

//////////////////////////
// PAYLOAD READERS
//////////////////////////
function b64urlToB64(s) {
  const n = s.replace(/-/g, "+").replace(/_/g, "/").replace(/\s/g, "+");
  return n.padEnd(Math.ceil(n.length / 4) * 4, "=");
}
function tryJSON(t) { try { const o = JSON.parse(t); return (o && typeof o === "object") ? o : null; } catch { return null; } }
function readFromQuery() {
  try {
    const p = new URLSearchParams(location.search);
    let raw = p.get("data");
    if (!raw) return null;
    raw = decodeURIComponent(raw.trim());
    const decoded = atob(b64urlToB64(raw));
    const parsed = tryJSON(decoded);
    if (!parsed) throw new Error("Payload invalid or truncated—relaunch from CRM.");
    return parsed;
  } catch { return null; }
}
function readFromWindowName() {
  try {
    if (!window.name) return null;
    const decoded = atob(b64urlToB64(window.name));
    const parsed = tryJSON(decoded);
    if (!parsed) throw new Error("Payload invalid or truncated—relaunch from CRM.");
    return parsed;
  } catch { return null; }
}

//////////////////////////
// UTILS
//////////////////////////
function sanitizeName(s) {
  return (s || "").toString().normalize("NFKD").replace(/[^\w\- ]+/g, "").trim();
}
function lastNameFromContact(contactName) {
  if (!contactName) return "Client";
  const parts = contactName.trim().split(/\s+/);
  return sanitizeName(parts[parts.length - 1] || "Client");
}
function inferLenderPrefix(lenderName) {
  const n = (lenderName || "").toLowerCase();
  const candidates = ["td","mcap","cmls","beem","cwb","firstnat","haventree","lendwise","scotia","strive","bridgewater","private"];
  for (const c of candidates) if (n.includes(c)) return c.toUpperCase();
  if (n) return sanitizeName(n.split(/\s+/)[0]).toUpperCase();
  return "TD";
}
function isTD(prefix) { return (prefix || "").toUpperCase().includes("TD"); }
function lenderFile(prefix, suffix) {
  return `${prefix}_${suffix}`.replace(/\.pdf$/i, "") + ".pdf";
}
function populateLenderOptions(selected) {
  if (!els.lender) return;
  els.lender.innerHTML = LENDER_OPTIONS
    .map(v => `<option value="${v}" ${v === selected ? "selected" : ""}>${v}</option>`)
    .join("");
}

// Replace non-WinAnsi chars with ASCII-safe versions for pdf-lib built-in fonts
function safeText(s) {
  const str = String(s ?? "");
  return str
    .replace(/\u2192/g, "->")              // arrows
    .replace(/[\u2012\u2013\u2014\u2015]/g, "-") // dashes → hyphen
    .replace(/\u2026/g, "...")             // ellipsis
    .replace(/[^\x00-\x7F]/g, "");         // strip other non-ASCII
}

// Tiny helper for blanks
function isBlank(v){ return v === undefined || v === null || String(v).trim() === ""; }

//////////////////////////
// FETCH HELPERS
//////////////////////////
async function fetchJson(url) {
  const r = await fetch(url, { cache: "no-store" });
  if (!r.ok) throw new Error(`Failed to fetch ${url} (${r.status})`);
  return r.json();
}
async function fetchPdfBytes(path) {
  const r = await fetch(path, { cache: "no-store" });
  if (!r.ok) throw new Error(`Missing PDF: ${path} (${r.status})`);
  return await r.arrayBuffer();
}
async function tryFetchPdfBytes(path) { try { return await fetchPdfBytes(path); } catch { return null; } }

// Load map.json (Zoho field → PDF field mapping) – tolerate flat or {__default:{...}}
let MAP = {};
(async () => {
  try {
    const raw = await fetch("./config/map.json", { cache: "no-store" }).then(r => r.json());
    MAP = raw && typeof raw === "object"
      ? (raw.__default ? raw : { __default: raw })
      : { __default: {} };
    console.log("MAP loaded:", Object.keys(MAP.__default || {}).length, "fields");
  } catch (err) {
    console.error("Could not load map.json:", err);
    MAP = { __default: {} };
  }
})();

//////////////////////////
// NUMBER + AMORTIZATION UTILS
//////////////////////////
function nnum(v){
  if (typeof v === "string") v = v.replace(/[, ]+/g, "");
  const f = parseFloat(v);
  return isFinite(f) ? f : 0;
}
function money(v){ return (Math.round(nnum(v)*100)/100).toFixed(2); }

// Explicit map first; then safe fallbacks
function ppy(freq){
  const t = (freq || "").trim().toLowerCase();
  const exact = {
    "monthly": 12,
    "semi-monthly": 24,
    "semi monthly": 24,
    "bi-weekly": 26,
    "biweekly": 26,
    "accelerated bi-weekly": 26,
    "accelerated biweekly": 26,
    "weekly": 52
  };
  if (t in exact) return exact[t];
  if (t.includes("accelerat")) return 26;
  if (t.includes("bi")) return 26;
  if (t.includes("semi-month")) return 24;
  if (t.includes("week")) return 52;
  if (t.includes("month")) return 12;
  return 12;
}
function effAnnualFromNominal(nomPct, comp){
  const r = nnum(nomPct)/100;
  const c = (comp || "").toLowerCase();
  if (c.includes("semi")) return Math.pow(1 + r/2, 2) - 1;
  if (c.includes("annual")) return r;
  return r;
}
function perPeriodRate(nomPct, comp, freq){
  const eff = effAnnualFromNominal(nomPct, comp);
  return Math.pow(1 + eff, 1/ppy(freq)) - 1;
}
function paymentFor(P, nomPct, comp, freq, amortMonths){
  const r = perPeriodRate(nomPct, comp, freq);
  const nper = Math.round((nnum(amortMonths) / 12) * ppy(freq));
  if (nper <= 0) return 0;
  if (r === 0) return P / nper;
  return P * r / (1 - Math.pow(1 + r, -nper));
}
function balanceAfter(P, nomPct, comp, freq, amortMonths, paidMonths, payAmount){
  const r = perPeriodRate(nomPct, comp, freq);
  const N = Math.round((nnum(amortMonths)/12) * ppy(freq));
  const k = Math.round((nnum(paidMonths)/12) * ppy(freq));
  const A = (payAmount && nnum(payAmount)>0) ? nnum(payAmount)
      : paymentFor(P, nomPct, comp, freq, amortMonths);
  if (N <= 0 || k <= 0) return Math.max(0, P);
  if (r === 0){
    const princPaid = A * k;
    return Math.max(0, P - princPaid);
  }
  const bal = P * Math.pow(1 + r, k) - A * ((Math.pow(1 + r, k) - 1) / r);
  return Math.max(0, bal);
}

//////////////////////////
// COB CALCULATIONS
//////////////////////////
function computeCOB(record) {
  // ---- Extract Inputs ----
  const P        = nnum(record.Total_Mortgage_Amount_incl_Insurance || record["Mortgage Amount"]);
  const nomPct   = nnum(record.Current_Rate || record.Mortgage_Rate);
  const comp     = (record.Compounding || "Semi-Annual").trim();
  const freq     = (record.Payment_Frequency || record.Mtg_Pmt_Freq || "Monthly").trim();

  const termM  = nnum(record.Term_Months || record.Term_Years || 0);
  const amortM = nnum(
    record.Mtg_Amortization ||
    record.Amortization_Months ||
    record.Amortization_Years ||
    record.Amortization ||
    0
  );

  const A_input = record.Payment_Amount || record.Mtg_Pmt_Amount;
  const A = (A_input && nnum(A_input) > 0)
    ? nnum(A_input)
    : paymentFor(P, nomPct, comp, freq, amortM);

  const periodsPerYear = ppy(freq);
  const r = perPeriodRate(nomPct, comp, freq);
  const N_term  = Math.round((termM / 12) * periodsPerYear);

  // ---- Balance at Maturity ----
  let Bal = 0;
  if (r === 0) {
    Bal = Math.max(0, P - A * N_term);
  } else {
    const g = Math.pow(1 + r, N_term);
    Bal = Math.max(0, P * g - A * ((g - 1) / r));
  }

  // ---- Interest to Term ----
  const totalPaid = A * N_term;
  const principalRepaid = Math.max(0, P - Bal);
  const interestToTerm  = Math.max(0, totalPaid - principalRepaid);

  // ---- Fee Definitions (28–34 deduct, 35–41 APR; Borrower’s lawyer = never) ----
  const feeDefs = [
    { field: "Title_Insurance",       ded: "Check Box28", apr: "Check Box35", dedInclude: true,  aprInclude: true  },
    { field: "Appraisal_AVM_Fees",    ded: "Check Box29", apr: "Check Box36", dedInclude: true,  aprInclude: true  },
    { field: "Borrowers_Lawyer_Fees", ded: "Check Box30", apr: "Check Box37", dedInclude: false, aprInclude: false }, // never tick
    { field: "Lenders_Lawyer_Fees",   ded: "Check Box31", apr: "Check Box38", dedInclude: true,  aprInclude: true  },
    { field: "Brokerage_Fee",         ded: "Check Box32", apr: "Check Box39", dedInclude: true,  aprInclude: true  },
    { field: "Lender_Fees",           ded: "Check Box33", apr: "Check Box40", dedInclude: true,  aprInclude: true  },
    { field: "Other_Fees",            ded: "Check Box34", apr: "Check Box41", dedInclude: true,  aprInclude: true  }
  ];

  let totalFees = 0;
  const checkboxStates = {};

  for (const f of feeDefs) {
    const val = nnum(record[f.field]);
    const hasVal = val > 0;

    // Deduct from proceeds
    checkboxStates[f.ded] = hasVal && f.dedInclude;

    // Include in APR
    checkboxStates[f.apr] = hasVal && f.aprInclude;

    // Only fees marked aprInclude affect APR/COB fee bucket
    if (hasVal && f.aprInclude) totalFees += val;
  }

  const totalCOB = interestToTerm + totalFees;

  // ---- APR via IRR ----
  // CF0 = P - totalFees (net advance); then A each period; Bal outstanding at maturity
  const CF0 = P - totalFees;
  const n = N_term;

  function solveAPR() {
    if (n === 0) return nomPct;

    let low = 0, high = 1, guess = 0.05;

    const f = (rate) => {
      let sum = CF0;
      for (let i = 1; i <= n; i++) {
        sum -= A / Math.pow(1 + rate, i);
      }
      sum -= Bal / Math.pow(1 + rate, n);
      return sum;
    };

    while (f(high) < 0) high *= 2;
    while (f(low) > 0) low  /= 2;

    for (let i = 0; i < 50; i++) {
      guess = (low + high) / 2;
      const val = f(guess);
      if (Math.abs(val) < 1e-9) break;
      if (val > 0) low = guess; else high = guess;
    }

    const effAnnual = Math.pow(1 + guess, periodsPerYear) - 1;
    return effAnnual * 100;
  }

  let APR = solveAPR();
  const contract = nomPct;

  if (APR < contract) APR = contract;
  if (totalFees > 0 && APR <= contract) APR = contract + 0.01;

  // ---- FINAL OUTPUT ----
  return {
    ...checkboxStates,
    "Mtg_Pmt_Amount":         money(A),
    "Balance_Calc":           money(Bal),
    "Total_Interest_To_Term": money(interestToTerm),
    "Total_Fees_Calc":        money(totalFees),
    "Total_COB_Calc":         money(totalCOB),
    "APR_Calc":               APR.toFixed(2)
  };
}

//////////////////////////
// PDF FILLING + MERGING
//////////////////////////
function normalize(v) {
  if (v == null) return "";
  if (typeof v === "boolean") return v ? "Yes" : "No";
  return String(v);
}

async function fillTemplateBytes(templatePath, record, map, preloadedBytes) {
  const bytes = preloadedBytes || await fetchPdfBytes(templatePath);
  const pdfDoc = await PDFLib.PDFDocument.load(bytes);
  const form = pdfDoc.getForm?.();
  const fields = form?.getFields?.() || [];

  if (!fields.length) {
    throw new Error("Template appears flattened/XFA (no AcroForm fields).");
  }

  const fieldNames = new Set(
    fields.map(f => {
      try { return f.getName(); } catch { return ""; }
    }).filter(Boolean)
  );
  dbg(`Template ${templatePath}: found ${fieldNames.size} form fields`);

  const usedMap = map && (map.__default || map) || {};
  let hits = 0, misses = 0;

  for (const [crmKeyRaw, pdfField] of Object.entries(usedMap)) {
    const keyStr = (typeof crmKeyRaw === "string") ? crmKeyRaw : String(crmKeyRaw ?? "");
    const leaf = keyStr.includes(".") ? keyStr.split(".").pop() : keyStr;
    const hasKey  = keyStr && Object.prototype.hasOwnProperty.call(record || {}, keyStr);
    const hasLeaf = leaf  && Object.prototype.hasOwnProperty.call(record || {}, leaf);
    const val = (hasKey ? record[keyStr] : undefined) ??
                (hasLeaf ? record[leaf] : undefined) ?? "";
    const v = normalize(val);

    if (!fieldNames.has(pdfField)) { misses++; continue; }

    try {
      const field = form.getField(pdfField);
      const type = field.constructor.name;

      if (type === "PDFCheckBox") {
        // Standard checkbox logic – COB-specific overrides are handled via computeCOB values/mapping
        const shouldCheck = /^(yes|true|1|x)$/i.test(String(val));
        if (shouldCheck) field.check();
        else             field.uncheck();
      } else if (type === "PDFDropdown") {
        try { field.select(v); } catch { field.setText(v); }
      } else if (type === "PDFRadioGroup") {
        try { field.select(v); } catch {}
      } else if (field.setText) {
        field.setText(v);
      }
      hits++;
    } catch (e) {
      misses++;
      dwarn(`Field set failed for ${pdfField}: ${e.message}`);
    }
  }

  try {
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    form.updateFieldAppearances(font);
  } catch {}

  dbg(`Filled ${hits} fields (misses: ${misses}) in ${templatePath}`);
  if (hits === 0) throw new Error("No fields filled (check mapping vs PDF field names, or template type)");
  return await pdfDoc.save({ updateFieldAppearances: false });
}

async function flattenAllPages(bytes) {
  const doc = await PDFLib.PDFDocument.load(bytes);
  try { const form = doc.getForm(); form && form.flatten(); } catch {}
  return await doc.save({ updateFieldAppearances: true });
}

async function mergeDocsFlattened(listOfBytes) {
  const out = await PDFLib.PDFDocument.create();
  for (const b of listOfBytes) {
    const flat = await flattenAllPages(b);
    const src = await PDFLib.PDFDocument.load(flat);
    const pages = await out.copyPages(src, src.getPageIndices());
    pages.forEach(p => out.addPage(p));
  }
  return await out.save({ updateFieldAppearances: true });
}

async function makeHeaderPageBytes(headersPdfPath, headerIndex1Based) {
  const headersBytes = await fetchPdfBytes(headersPdfPath);
  const src = await PDFLib.PDFDocument.load(headersBytes);
  const out = await PDFLib.PDFDocument.create();
  const [page] = await out.copyPages(src, [headerIndex1Based - 1]);
  out.addPage(page);
  return await out.save();
}

//////////////////////////
// GLOBAL MAP-AWARE ADDER
//////////////////////////
async function addTemplateToMerge(relPath, record, fieldMap, bucket, preloadedBytes) {
  const fullPath = `${TEMPL_BASE}${relPath}`;
  try {
    const isCOB = /COB/i.test(relPath);
    const rec2 = isCOB
      ? { ...record, ...computeCOB(record) }   // inject COB + checkbox values
      : record;

    const filled = await fillTemplateBytes(fullPath, rec2, fieldMap, preloadedBytes);
    bucket.push(filled);
  } catch (err) {
    dwarn(`Could not fill ${relPath}, using static version instead: ${err.message}`);
    const bytes = preloadedBytes || await fetchPdfBytes(fullPath);
    bucket.push(bytes);
  }
}

//////////////////////////
// ZIP + INFO COVER
//////////////////////////
async function downloadZip(namedBytes) {
  const zip = new JSZip();
  for (const it of namedBytes) zip.file(it.name, it.bytes);
  const blob = await zip.generateAsync({ type: "blob" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "Editable_Singles.zip";
  document.body.appendChild(a); a.click(); a.remove();
}

// Info Cover Page (unchanged display re: months per your choice)
async function makeInfoCoverPdf(record, mapDefault) {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  const toStr = v => (v === null || v === undefined || v === "" ? "(blank)" : String(v));
  const isBlankVal = v => v === "(blank)";
  const trimCell = (s, max) => {
    const txt = safeText(String(s || ""));
    return txt.length > max ? txt.slice(0, max - 3) + "..." : txt;
  };

  const values = { ...record };

  const rows = [];
  for (const [zohoKey, pdfField] of Object.entries(mapDefault || {})) {
    const val = toStr(values[zohoKey]);
    rows.push({ zohoKey, pdfField, val, blank: isBlankVal(val) });
  }
  for (const k of Object.keys(values || {})) {
    if (!(k in (mapDefault || {}))) {
      const val = toStr(values[k]);
      rows.push({ zohoKey: `${k} (unmapped)`, pdfField: "-", val, blank: isBlankVal(val) });
    }
  }

  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontB = await doc.embedFont(StandardFonts.HelveticaBold);

  const pageSize = [595, 842];
  const margin = 36;
  const colX = { key: margin, map: margin + 210, val: margin + 380 };
  const lineH = 14;

  let page = doc.addPage(pageSize);
  let y = pageSize[1] - margin;

  const address = [values.Street, values.City, values.Province, values.Postal_Code].filter(Boolean).join(", ");
  page.drawText(safeText("HH - Info Cover Page"), { x: margin, y, size: 16, font: fontB }); y -= 24;
  const headerLines = [
    ["Deal / File Name:", values.Deal_Name || values.Name],
    ["Contact Name:", values.Contact_Name],
    ["Lender:", values.Lender_Name],
    ["Address:", address],
    ["Closing Date:", values.Closing_Date],
    ["COF Date:", values.COF_Date],
    ["Total Mtg Amt (incl. ins.):", values.Total_Mortgage_Amount_incl_Insurance],
    ["Email (B1):", values.Email],
    ["Email (B2):", values.Contact_Email_2],
  ];
  for (const [lbl, val] of headerLines) {
    page.drawText(safeText(lbl), { x: margin, y, size: 11, font: fontB });
    page.drawText(safeText(String(val || "")), { x: margin + 160, y, size: 11, font });
    y -= 16;
  }
  y -= 6;

  const drawHeaders = () => {
    page.drawText(safeText("Zoho Key"), { x: colX.key, y, size: 11, font: fontB });
    page.drawText(safeText("-> PDF Field"), { x: colX.map, y, size: 11, font: fontB });
    page.drawText(safeText("Value"), { x: colX.val, y, size: 11, font: fontB });
    y -= 12;
  };
  drawHeaders();

  for (const r of rows) {
    if (y < margin + 24) {
      page = doc.addPage(pageSize);
      y = pageSize[1] - margin;
      page.drawText(safeText("HH - Info Cover Page (cont.)"), { x: margin, y, size: 12, font: fontB });
      y -= 18;
      drawHeaders();
    }
    const keyTxt = trimCell(r.zohoKey, 40);
    const mapTxt = "-> " + trimCell(r.pdfField, 28);
    const valTxt = trimCell(r.val, 48);

    page.drawText(safeText(keyTxt), { x: colX.key, y, size: 10, font });
    page.drawText(safeText(mapTxt), { x: colX.map, y, size: 10, font });
    page.drawText(safeText(valTxt), { x: colX.val, y, size: 10, font, color: r.blank ? rgb(0.65, 0, 0) : rgb(0, 0, 0) });
    y -= lineH;
  }

  const bytes = await doc.save({ updateFieldAppearances: true });
  return new Blob([bytes], { type: "application/pdf" });
}

//////////////////////////
// MAIN
//////////////////////////
(async function run() {
  // Startup guards
  if (!window.PDFLib) { setStatus("PDFLib not loaded. Include pdf-lib before app.js.", "err"); throw new Error("PDFLib missing"); }
  if (!window.JSZip)  { setStatus("JSZip not loaded. Include jszip before app.js.", "err");  throw new Error("JSZip missing"); }

  setStatus("Loading payload...");
  let record = readFromQuery() || readFromWindowName();
  if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());

  // Default First_Payment_Date text ONLY if payload missing
  if (record && isBlank(record.First_Payment_Date)) {
    record.First_Payment_Date = "See Lender Commitment";
  }

  // Manifest (safe defaults/guards)
  const manifest = await fetchJson(`${CONFIG_BASE}standard_pkg.json`);
  const sequence = Array.isArray(manifest.sequence) ? manifest.sequence : [];
  const prosprFallback = manifest.prospr_fallback || "PROSPR.pdf";

  // Map (normalized above); also allow local override for Singles
  let fieldMap = {};
  try {
    const rawMap = await fetchJson(`${CONFIG_BASE}map.json`);
    fieldMap = rawMap && typeof rawMap === "object"
      ? (rawMap.__default ? rawMap : { __default: rawMap })
      : { __default: {} };
  } catch {
    fieldMap = { __default: {} };
  }

  const detected = inferLenderPrefix(record?.Lender_Name || "");
  populateLenderOptions(detected || "TD");
  function refreshAPA() { if (els.apaWrap) els.apaWrap.style.display = isTD(els.lender.value) ? "block" : "none"; }
  refreshAPA();
  els.lender?.addEventListener("change", refreshAPA);

  // ---- Build Signing Package ----
  els.btnPkg?.addEventListener("click", () => withBusy(els.btnPkg, "Building…", async () => {
    try {
      if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());
      if (!record) throw new Error("No record payload. Paste JSON or launch from CRM.");
      if (!els.upCommitment?.files[0]) throw new Error("Commitment is required.");
      if (!els.upApplication?.files[0]) throw new Error("Mortgage Application is required.");
      setStatus("Building signing package...", "muted");

      const lenderPrefix = els.lender.value;
      const pkgType = (els.pkg.value || "Standard").toLowerCase();
      const headersPath = `${TEMPL_BASE}Headers_Signing_Package.pdf`;

      // Verify headers exist now (faster fail)
      const headersProbe = await tryFetchPdfBytes(headersPath);
      if (!headersProbe) throw new Error(`Missing required header file: ${headersPath}`);

      const upCommitmentBytes = await els.upCommitment.files[0].arrayBuffer();
      const upApplicationBytes = await els.upApplication.files[0].arrayBuffer();
      const upMPPBytes = els.upMPP?.files[0] ? await els.upMPP.files[0].arrayBuffer() : null;
      const upAPABytes = els.upAPA?.files[0] ? await els.upAPA.files[0].arrayBuffer() : null;
      const toMerge = [];

      async function pushHeader(key) {
        const idx = HEADER_INDEX[key];
        if (!idx) return;
        toMerge.push(await makeHeaderPageBytes(headersPath, idx));
      }
      async function pushBytes(bytes) { toMerge.push(bytes); }

      for (const s of sequence) {
        const spec = s && typeof s === "object" ? s : {};
        const cond = spec.condition || "";
        if (cond === "lender_has_TD" && !isTD(lenderPrefix)) continue;
        if (cond === "package_is_HELOC" && pkgType !== "heloc") continue;
        if (cond === "package_is_STANDARD" && pkgType !== "standard") continue;

        if (spec.no_header !== true && spec.header) await pushHeader(spec.header);

        if (spec.doc_if_upload || spec.doc_if_missing) {
          if (upMPPBytes) await pushBytes(upMPPBytes);
          else await addTemplateToMerge(prosprFallback, record, fieldMap, toMerge);
          continue;
        }
        if (spec.doc === "UPLOAD_Commitment.pdf") { await pushBytes(upCommitmentBytes); continue; }
        if (spec.doc === "UPLOAD_Application.pdf") { await pushBytes(upApplicationBytes); continue; }
        if (spec.doc === "UPLOAD_APA_TD.pdf") { if (isTD(lenderPrefix) && upAPABytes) await pushBytes(upAPABytes); continue; }

        if (spec.doc && spec.doc.startsWith("LENDER_")) {
          const suffix = spec.doc.replace(/^LENDER_/, "");
          const rel = `${lenderFile(lenderPrefix, suffix)}`;
          const path = `${TEMPL_BASE}${rel}`;
          const preload = await tryFetchPdfBytes(path);
          if (!preload) {
            if (spec.optional) continue;
            throw new Error(`Missing required lender doc for ${lenderPrefix}: ${rel}`);
          }
          await addTemplateToMerge(rel, record, fieldMap, toMerge, preload);
          continue;
        }

        if (spec.doc) {
          await addTemplateToMerge(spec.doc, record, fieldMap, toMerge);
        }
      }

      const merged = await mergeDocsFlattened(toMerge);
      const lastName = lastNameFromContact(record?.Contact_Name || "");

      const d = new Date();
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth()+1).padStart(2,"0");
      const dd = String(d.getDate()).padStart(2,"0");
      const outName = `${lastName} - Signing Package - ${yyyy}${mm}${dd}.pdf`;

      const blob = new Blob([merged], { type: "application/pdf" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = outName;
      document.body.appendChild(a); a.click(); a.remove();

      setStatus(`Signing package created: ${outName}`, "ok");
    } catch (e) {
      console.error(e);
      setStatus(`Error: ${e.message}`, "err");
      alert(e.message);
    }
  }));

  // ---- Editable Singles ----
  els.btnSingles?.addEventListener("click", () => withBusy(els.btnSingles, "Generating…", async () => {
    try {
      if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());
      if (!record) throw new Error("No record payload. Paste JSON or relaunch from CRM.");
      setStatus("Generating editable singles...", "muted");

      const localMap = fieldMap;
      const lenderPrefix = els.lender.value;
      const singles = [];

      // KYC/IDV (required)
      {
        const path = `${TEMPL_BASE}KYC_IDV_Template.pdf`;
        const probe = await tryFetchPdfBytes(path);
        if (!probe) throw new Error(`Missing required template: ${path}`);
        const filled = await fillTemplateBytes(path, record, localMap, probe);
        singles.push({ name: "Identification Verification.pdf", bytes: filled });
      }

      // FCT Auth (optional)
      {
        const rel = `${lenderFile(lenderPrefix, "FCT_Auth.pdf")}`;
        const path = `${TEMPL_BASE}${rel}`;
        const exists = await tryFetchPdfBytes(path);
        if (exists) {
          const filled = await fillTemplateBytes(path, record, localMap, exists);
          singles.push({ name: "FCT Authorization.pdf", bytes: filled });
        }
      }

      // Gift Letter (optional)
      {
        const rel = `${lenderFile(lenderPrefix, "Gift_Letter.pdf")}`;
        const path = `${TEMPL_BASE}${rel}`;
        const exists = await tryFetchPdfBytes(path);
        if (exists) {
          const filled = await fillTemplateBytes(path, record, localMap, exists);
          singles.push({ name: "Gift Letter.pdf", bytes: filled });
        }
      }

      if (singles.length === 0) throw new Error("No editable singles available for this lender.");
      await downloadZip(singles);
      setStatus("Editable singles downloaded (ZIP).", "ok");
    } catch (e) {
      console.error(e);
      setStatus(`Error: ${e.message}`, "err");
      alert(e.message);
    }
  }));

  // ---- Info Cover Page ----
  els.btnInfo?.addEventListener("click", () => withBusy(els.btnInfo, "Generating…", async () => {
    try {
      const rec = record || tryJSON(els.manualJson?.value?.trim()) || {};
      if (!Object.keys(rec).length) throw new Error("No record payload. Paste JSON or relaunch from CRM.");
      setStatus("Generating info cover page...", "muted");
      const blob = await makeInfoCoverPdf(rec, MAP.__default);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `HH_Data_Snapshot_${rec?.Lender_Name || "Lender"}_${Date.now()}.pdf`;
      document.body.appendChild(a); a.click(); a.remove();
      setStatus("Info cover page downloaded.", "ok");
    } catch (e) {
      console.error(e);
      setStatus(`Error: ${e.message}`, "err");
      alert(e.message);
    }
  }));

  populateLenderOptions(detected || "TD");
  if (els.pkg && !els.pkg.value) els.pkg.value = "Standard";

  setStatus(
    record
      ? `Payload OK. Lender approx ${detected || "TD"}. Choose options and upload files.`
      : "Paste JSON or launch from CRM.",
    record ? "ok" : "warn"
  );
})();

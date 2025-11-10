/* HH Signing Package Builder – app.js (repo-aware paths, global mapping)
 * Repo layout expected:
 *   /config/standard_pkg.json, /config/map.json
 *   /data/templates/*.pdf  (lender docs, headers-signing-package.pdf, etc.)
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
const LENDER_OPTIONS = ["TD","MCAP","CMLS","CWB","FirstNat","Haventree","Lendwise","Scotia","Strive","Bridgewater","Private"];

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
    return tryJSON(atob(b64urlToB64(raw)));
  } catch { return null; }
}
function readFromWindowName() {
  try { if (!window.name) return null; return tryJSON(atob(b64urlToB64(window.name))); } catch { return null; }
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
  const candidates = ["td","mcap","cmls","cwb","firstnat","haventree","lendwise","scotia","strive","bridgewater","private"];
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

//////////////////////////
// NUMBER + AMORTIZATION UTILS
//////////////////////////
function nnum(v){
  if (typeof v === "string") v = v.replace(/[, ]+/g, "");
  const f = parseFloat(v);
  return isFinite(f) ? f : 0;
}
function money(v){ return (Math.round(nnum(v)*100)/100).toFixed(2); }

function ppy(freq){
  const t = (freq || "").toLowerCase();
  if (t.includes("weekly")) return 52;
  if (t.includes("semi-month")) return 24;
  if (t.includes("accelerat")) return 26;
  if (t.includes("bi")) return 26;
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
function computeCOB(record){
  const P        = nnum(record.Total_Mortgage_Amount_incl_Insurance || record["Mortgage Amount"]);
  const nomPct   = nnum(record.Interest_Rate || record.Mortgage_Rate);
  const comp     = record.Compounding || "Semi-Annual";
  const freq     = record.Payment_Frequency || record.Mtg_Pmt_Freq || "Monthly";
  const termM    = nnum(record.Term_Months);
  const amortM   = record.Mtg_Amortization ? nnum(record.Mtg_Amortization)
                    : (record.Amortization ? nnum(record.Amortization)*12
                    : nnum(record.Amortization_Years));
  const A_input  = record.Payment_Amount || record.Mtg_Pmt_Amount;

  const A = (A_input && nnum(A_input)>0) ? nnum(A_input)
        : paymentFor(P, nomPct, comp, freq, amortM);

  const Bal = balanceAfter(P, nomPct, comp, freq, amortM, termM, A);
  const termPayments = Math.round((termM/12) * ppy(freq));
  const totalPaid = A * termPayments;
  const principalRepaid = Math.max(0, P - Bal);
  const interestToTerm = Math.max(0, totalPaid - principalRepaid);

  const feeLogic = [
    { field: "Title_Insurance",        cb: "Check Box33", include: true },
    { field: "Appraisal_AVM_Fees",     cb: "Check Box36", include: true },
    { field: "Borrowers_Lawyer_Fees",  cb: "Check Box35", include: false },
    { field: "Lenders_Lawyer_Fees",    cb: "Check Box32", include: true },
    { field: "Brokerage_Fee",          cb: "Check Box31", include: true },
    { field: "Lender_Fees",            cb: "Check Box29", include: true },
    { field: "Other_Fees",             cb: "Check Box34", include: true }
  ];

  let feesIncluded = 0;
  const checkboxStates = {};
  for (const { field, cb, include } of feeLogic) {
    const val = nnum(record[field]);
    const hasValue = val > 0;
    checkboxStates[cb] = hasValue && include;
    if (include && hasValue) feesIncluded += val;
  }

  const totalFees = feesIncluded;
  const totalCOB  = interestToTerm + totalFees;
  const termYears = nnum(termM)/12 || 1;
  const avgPrincipal = (P + Bal) / 2 || P;
  const APR = (avgPrincipal > 0)
    ? 100 * ((totalCOB) / (termYears * avgPrincipal))
    : 0;

  return {
    ...checkboxStates,
    "Mtg_Pmt_Amount": money(A),
    "Balance_Calc": money(Bal),
    "Total_Interest_To_Term": money(interestToTerm),
    "Total_Fees_Calc": money(totalFees),
    "Total_COB_Calc": money(totalCOB),
    "APR_Calc": (Math.round(APR*100)/100).toFixed(2)
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

async function fillTemplateBytes(templatePath, record, map) {
  const bytes = await fetchPdfBytes(templatePath);
  const pdfDoc = await PDFLib.PDFDocument.load(bytes);
  const form = pdfDoc.getForm?.();
  const fields = form?.getFields?.() || [];
  const fieldNames = new Set(fields.map(f => { try { return f.getName(); } catch { return ""; } }).filter(Boolean));

  dbg(`Template ${templatePath}: found ${fieldNames.size} form fields`);
  if (!fieldNames.size) throw new Error("No AcroForm fields detected (flattened or XFA template?)");

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
        /^(yes|true|1|x)$/i.test(String(val)) ? field.check() : field.uncheck();
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
async function addTemplateToMerge(relPath, record, fieldMap, bucket) {
  const fullPath = `${TEMPL_BASE}${relPath}`;
  try {
    const rec2 = /COB/i.test(relPath) ? { ...record, ...computeCOB(record) } : record;
    const filled = await fillTemplateBytes(fullPath, rec2, fieldMap);
    bucket.push(filled);
  } catch (err) {
    dwarn(`Could not fill ${relPath}, using static version instead: ${err.message}`);
    const bytes = await fetchPdfBytes(fullPath);
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

async function makeInfoCoverPdf(record) {
  const doc = await PDFLib.PDFDocument.create();
  const page = doc.addPage([595, 842]);
  const font = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  const fontB = await doc.embedFont(PDFLib.StandardFonts.HelveticaBold);
  let y = 800;
  function line(lbl, val) {
    page.drawText(lbl, { x: 40, y, size: 11, font: fontB });
    page.drawText(String(val || ""), { x: 200, y, size: 11, font });
    y -= 18;
  }
  page.drawText("HH - Info Cover Page", { x: 40, y, size: 16, font: fontB }); y -= 28;
  line("Deal / File Name:", record.Deal_Name || record.Name);
  line("Contact Name:", record.Contact_Name);
  line("Lender:", record.Lender_Name);
  line("Address:", [record.Street, record.City, record.Province, record.Postal_Code].filter(Boolean).join(", "));
  line("Closing Date:", record.Closing_Date);
  line("COF Date:", record.COF_Date);
  line("Total Mtg Amt (incl. ins.):", record.Total_Mortgage_Amount_incl_Insurance);
  line("Email (B1):", record.Email);
  line("Email (B2):", record.Contact_Email_2);
  const bytes = await doc.save({ updateFieldAppearances: true });
  return new Blob([bytes], { type: "application/pdf" });
}

//////////////////////////
// MAIN
//////////////////////////
(async function run() {
  setStatus("Loading payload...");
  let record = readFromQuery() || readFromWindowName();
  if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());

  const manifest = await fetchJson(`${CONFIG_BASE}standard_pkg.json`);
  let fieldMap = {};
  try { fieldMap = await fetchJson(`${CONFIG_BASE}map.json`); } catch { fieldMap = {}; }

  const detected = inferLenderPrefix(record?.Lender_Name || "");
  populateLenderOptions(detected || "TD");
  function refreshAPA() { if (els.apaWrap) els.apaWrap.style.display = isTD(els.lender.value) ? "block" : "none"; }
  refreshAPA();
  els.lender?.addEventListener("change", refreshAPA);

  // ---- Build Signing Package ----
  els.btnPkg?.addEventListener("click", async () => {
    try {
      if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());
      if (!record) throw new Error("No record payload. Paste JSON or launch from CRM.");
      if (!els.upCommitment?.files[0]) throw new Error("Commitment is required.");
      if (!els.upApplication?.files[0]) throw new Error("Mortgage Application is required.");
      setStatus("Building signing package...", "muted");

      const lenderPrefix = els.lender.value;
      const pkgType = (els.pkg.value || "Standard").toLowerCase();
      const headersPath = `${TEMPL_BASE}${manifest.headers_pdf}`;
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

      for (const spec of manifest.sequence) {
        if (spec.condition === "lender_has_TD" && !isTD(lenderPrefix)) continue;
        if (spec.condition === "package_is_HELOC" && pkgType !== "heloc") continue;
        if (spec.condition === "package_is_STANDARD" && pkgType !== "standard") continue;
        if (spec.no_header !== true && spec.header) await pushHeader(spec.header);

        if (spec.doc_if_upload || spec.doc_if_missing) {
          if (upMPPBytes) await pushBytes(upMPPBytes);
          else await addTemplateToMerge(manifest.prospr_fallback, record, fieldMap, toMerge);
          continue;
        }
        if (spec.doc === "UPLOAD_Commitment.pdf") { await pushBytes(upCommitmentBytes); continue; }
        if (spec.doc === "UPLOAD_Application.pdf") { await pushBytes(upApplicationBytes); continue; }
        if (spec.doc === "UPLOAD_APA_TD.pdf") { if (isTD(lenderPrefix) && upAPABytes) await pushBytes(upAPABytes); continue; }

        if (spec.doc && spec.doc.startsWith("LENDER_")) {
          const suffix = spec.doc.replace(/^LENDER_/, "");
          const rel = `${lenderFile(lenderPrefix, suffix)}`;
          const exists = await tryFetchPdfBytes(`${TEMPL_BASE}${rel}`);
          if (!exists) {
            if (spec.optional) continue;
            throw new Error(`Missing required lender doc: ${rel}`);
          }
          await addTemplateToMerge(rel, record, fieldMap, toMerge);
          continue;
        }

        await addTemplateToMerge(spec.doc, record, fieldMap, toMerge);
      }

      const merged = await mergeDocsFlattened(toMerge);
      const lastName = lastNameFromContact(record?.Contact_Name || "");
      const outName = `${lastName} - Signing Pkg.pdf`;

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
  });

  // ---- Editable Singles ----
  els.btnSingles?.addEventListener("click", async () => {
    try {
      if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());
      if (!record) throw new Error("No record payload. Paste JSON or relaunch from CRM.");
      setStatus("Generating editable singles...", "muted");
      let localMap = {};
      try { localMap = await fetchJson(`${CONFIG_BASE}map.json`); } catch {}

      const lenderPrefix = els.lender.value;
      const singles = [];

      // KYC/IDV
      {
        const path = `${TEMPL_BASE}KYC_IDV_Template.pdf`;
        const filled = await fillTemplateBytes(path, record, localMap);
        singles.push({ name: "Identification Verification.pdf", bytes: filled });
      }

      // FCT Auth
      {
        const path = `${TEMPL_BASE}${lenderFile(lenderPrefix, "FCT_Auth.pdf")}`;
        const exists = await tryFetchPdfBytes(path);
        if (exists) {
          const filled = await fillTemplateBytes(path, record, localMap);
          singles.push({ name: "FCT Authorization.pdf", bytes: filled });
        }
      }

      // Gift Letter
      {
        const path = `${TEMPL_BASE}${lenderFile(lenderPrefix, "Gift_Letter.pdf")}`;
        const exists = await tryFetchPdfBytes(path);
        if (exists) {
          const filled = await fillTemplateBytes(path, record, localMap);
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
  });

  // ---- Info Cover Page ----
  els.btnInfo?.addEventListener("click", async () => {
    try {
      const rec = record || tryJSON(els.manualJson?.value?.trim()) || {};
      if (!Object.keys(rec).length) throw new Error("No record payload. Paste JSON or relaunch from CRM.");
      setStatus("Generating info cover page...", "muted");
      const blob = await makeInfoCoverPdf(rec);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "Info Cover Page.pdf";
      document.body.appendChild(a); a.click(); a.remove();
      setStatus("Info cover page downloaded.", "ok");
    } catch (e) {
      console.error(e);
      setStatus(`Error: ${e.message}`, "err");
      alert(e.message);
    }
  });

  populateLenderOptions(detected || "TD");
  if (els.pkg && !els.pkg.value) els.pkg.value = "Standard";
  setStatus(
    record
      ? `Payload OK. Lender approx ${detected || "TD"}. Choose options and upload files.`
      : "Paste JSON or launch from CRM.",
    record ? "ok" : "warn"
  );
})();

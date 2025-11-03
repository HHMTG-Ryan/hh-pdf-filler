/* HH Signing Package Builder – app.js (repo-aware paths)
 * Repo layout expected:
 *   /config/standard_pkg.json, /config/map.json
 *   /data/templates/*.pdf  (lender docs, headers-signing-package.pdf, etc.)
 *   /assets/js/app.js      (this file)
 */

//////////////////////////
// DOM
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
// Paths (GitHub Pages project base)
//////////////////////////
// e.g. https://user.github.io/HH-LENDER-DOCS-FILLER/app/
//      location.pathname = "/HH-LENDER-DOCS-FILLER/app/"
const _parts = location.pathname.split("/").filter(Boolean);
// repo root like "/HH-LENDER-DOCS-FILLER/"
const REPO_BASE = _parts.length ? `/${_parts[0]}/` : "/";
const CONFIG_BASE = `${REPO_BASE}config/`;
const TEMPL_BASE  = `${REPO_BASE}data/templates/`;

//////////////////////////
// Constants
//////////////////////////
const LENDER_OPTIONS = ["TD","MCAP","CMLS","CWB","FirstNat","Haventree","Lendwise","Scotia","Strive","Bridgewater","Private"];

// Header page indices inside headers-signing-package.pdf (1-based)
const HEADER_INDEX = {
  WMB: 1,
  MPP_OR_PROSPR: 2,
  COMMITMENT: 3,
  COB: 4,
  FORM10: 5,
  APPLICATION: 6
};

//////////////////////////
// Payload readers
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
// Utils
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
// Fetch helpers
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
// PDF filling & merging
//////////////////////////
function normalize(v) { if (v == null) return ""; if (typeof v === "boolean") return v ? "Yes" : "No"; return String(v); }

async function fillTemplateBytes(templatePath, record, map) {
  const bytes = await fetchPdfBytes(templatePath);
  const pdfDoc = await PDFLib.PDFDocument.load(bytes);
  const form = pdfDoc.getForm();

  for (const [crmKey, pdfField] of Object.entries(map || {})) {
    const val = record[crmKey] ?? record[crmKey.split(".").pop()];
    const v = normalize(val);
    try {
      const field = form.getField(pdfField);
      const type = field.constructor.name;
      if (type === "PDFTextField") {
        field.setText(v);
      } else if (type === "PDFDropdown") {
        try { field.select(v); } catch { field.setText(v); }
      } else if (type === "PDFCheckBox") {
        /^(yes|true|1|x)$/i.test(v) ? field.check() : field.uncheck();
      } else if (type === "PDFRadioGroup") {
        try { field.select(v); } catch {}
      } else if (field.setText) {
        field.setText(v);
      }
    } catch {}
  }
  try {
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    pdfDoc.getForm().updateFieldAppearances(font);
  } catch {}
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
// ZIP + Info Cover
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
  page.drawText("HH – Info Cover Page", { x: 40, y, size: 16, font: fontB }); y -= 28;
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
// Main
//////////////////////////
(async function run() {
  setStatus("Loading payload...");

  // 1) payload
  let record = readFromQuery() || readFromWindowName();
  if (!record && els.manualJson?.value?.trim()) record = tryJSON(els.manualJson.value.trim());

  // 2) config / manifest
  const manifest = await fetchJson(`${CONFIG_BASE}standard_pkg.json`);
  let fieldMap = {};
  try { fieldMap = await fetchJson(`${CONFIG_BASE}map.json`); } catch { fieldMap = {}; }

  // 3) lender detection + dropdown
  const detected = inferLenderPrefix(record?.Lender_Name || "");
  populateLenderOptions(detected || "TD");
  function refreshAPA() { if (els.apaWrap) els.apaWrap.style.display = isTD(els.lender.value) ? "block" : "none"; }
  refreshAPA();
  els.lender?.addEventListener("change", refreshAPA);

  // 4) Buttons
  els.btnPkg?.addEventListener("click", async () => {
    try {
      if (!record && !els.manualJson?.value?.trim()) throw new Error("No record payload. Paste JSON or launch from CRM.");
      if (!els.upCommitment?.files[0]) throw new Error("Commitment is required.");
      if (!els.upApplication?.files[0]) throw new Error("Mortgage Application is required.");

      setStatus("Building signing package…", "muted");

      const lenderPrefix = els.lender.value;
      const pkgType = els.pkg.value; // "Standard" or "HELOC"
      const headersPath = `${TEMPL_BASE}${manifest.headers_pdf}`;

      // Upload bytes
      const upCommitmentBytes = await els.upCommitment.files[0].arrayBuffer();
      const upApplicationBytes = await els.upApplication.files[0].arrayBuffer();
      const upMPPBytes = els.upMPP?.files[0] ? await els.upMPP.files[0].arrayBuffer() : null;
      const upAPABtes = els.upAPA?.files[0] ? await els.upAPA.files[0].arrayBuffer() : null;

      const toMerge = [];

      async function pushHeader(key) {
        const idx = HEADER_INDEX[key];
        if (!idx) return;
        toMerge.push(await makeHeaderPageBytes(headersPath, idx));
      }
      async function pushStatic(relPath) { toMerge.push(await fetchPdfBytes(`${TEMPL_BASE}${relPath}`)); }
      async function pushBytes(bytes) { toMerge.push(bytes); }

      // Walk manifest sequence & conditions
      for (const spec of manifest.sequence) {
        if (spec.condition === "lender_has_TD" && !isTD(lenderPrefix)) continue;
        if (spec.condition === "package_is_HELOC" && pkgType.toLowerCase() !== "heloc") continue;
        if (spec.condition === "package_is_STANDARD" && pkgType.toLowerCase() !== "standard") continue;

        if (spec.no_header !== true && spec.header) await pushHeader(spec.header);

        // MPP vs Prospr
        if (spec.doc_if_upload || spec.doc_if_missing) {
          if (upMPPBytes) await pushBytes(upMPPBytes);
          else await pushStatic(manifest.prospr_fallback);
          continue;
        }

        // Uploads
        if (spec.doc === "UPLOAD_Commitment.pdf") { await pushBytes(upCommitmentBytes); continue; }
        if (spec.doc === "UPLOAD_Application.pdf") { await pushBytes(upApplicationBytes); continue; }
        if (spec.doc === "UPLOAD_APA_TD.pdf") { if (isTD(lenderPrefix) && upAPABtes) await pushBytes(upAPABtes); continue; }

        // LENDER_* auto-pick
        if (spec.doc.startsWith("LENDER_")) {
          const suffix = spec.doc.replace(/^LENDER_/, "");
          const path = `${lenderFile(lenderPrefix, suffix)}`;
          const b = await tryFetchPdfBytes(`${TEMPL_BASE}${path}`);
          if (!b) { if (spec.optional) continue; else throw new Error(`Missing required lender doc: ${path}`); }
          await pushBytes(b);
          continue;
        }

        // Plain relative template path
        await pushStatic(spec.doc);
      }

      // Merge + flatten
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

  els.btnSingles?.addEventListener("click", async () => {
    try {
      if (!record && !els.manualJson?.value?.trim()) throw new Error("No record payload. Paste JSON or relaunch from CRM.");
      setStatus("Generating editable singles…", "muted");

      // reload map just in case
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

      // FCT Authorization (if lender file exists)
      {
        const path = `${TEMPL_BASE}${lenderFile(lenderPrefix, "FCT_Auth.pdf")}`;
        const exists = await tryFetchPdfBytes(path);
        if (exists) {
          const filled = await fillTemplateBytes(path, record, localMap);
          singles.push({ name: "FCT Authorization.pdf", bytes: filled });
        }
      }

      // Gift Letter (if lender file exists)
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

  els.btnInfo?.addEventListener("click", async () => {
    try {
      const rec = record || tryJSON(els.manualJson?.value?.trim()) || {};
      if (!Object.keys(rec).length) throw new Error("No record payload. Paste JSON or relaunch from CRM.");
      setStatus("Generating info cover page…", "muted");
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

  // Initial UI state
  populateLenderOptions(detected || "TD");
  if (els.pkg && !els.pkg.value) els.pkg.value = "Standard";
  setStatus(record ? `Payload OK. Lender ≈ ${detected || "TD"}. Choose options and upload files.` : "Paste JSON or launch from CRM.", record ? "ok" : "warn");
})();

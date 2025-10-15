/* HH PDF Filler – client-side only
 * ---------------------------------------------------------------
 * - Reads record data from window.name (set by Zoho Function) or manual JSON
 * - Loads /docs/map.json for CRM→PDF field mapping
 * - Loads chosen template PDF from /docs/templates/<key>.pdf
 * - Fills AcroForm fields, keeps them EDITABLE, then downloads the file
 */

const TEMPLATE_KEYS = [
  // EDIT this list to match the PDFs in /docs/templates (filenames without .pdf)
  "mcap_std",
  "mcap_heloc",
  "scotia_std",
  "scotia_heloc",
  "firstnat_std",
  "bridgewater_std",
  "cmls_std",
  "cwb_optimum_std",
  "haventree_std",
  "lendwise_std",
  "private_std",
  "strive_std",
  "td_std",
  "td_heloc",
  "signing_package"
];

const els = {
  sel: document.getElementById("template"),
  btn: document.getElementById("fillBtn"),
  txt: document.getElementById("manualJson"),
  status: document.getElementById("status"),
  fieldsBox: document.getElementById("fieldsBox")
};

function setStatus(msg, cls = "muted") {
  els.status.className = `row ${cls}`;
  els.status.textContent = msg;
}

function toTitleCase(str) {
  return str
    .replace(/_/g, " ")
    .replace(/\b\w/g, ch => ch.toUpperCase());
}

function loadTemplates() {
  els.sel.innerHTML = "";
  TEMPLATE_KEYS.forEach(k => {
    const opt = document.createElement("option");
    opt.value = k;
    opt.textContent = toTitleCase(k);
    els.sel.appendChild(opt);
  });
}

function readPayloadFromWindowName() {
  try {
    if (!window.name) return null;
    const decoded = atob(window.name);
    const obj = JSON.parse(decoded);
    return obj && typeof obj === "object" ? obj : null;
  } catch (e) {
    return null;
  }
}

async function fetchJson(url) {
  const r = await fetch(url, { cache: "no-store" });
  if (!r.ok) throw new Error(`Failed to fetch ${url}`);
  return r.json();
}

function normalize(v) {
  if (v == null) return "";
  if (typeof v === "boolean") return v ? "Yes" : "No";
  return String(v);
}

function deriveFilename(templateKey, rec) {
  const name = rec.Deal_Name || rec.Name || "document";
  const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
  return `${templateKey}_${name}_${stamp}.pdf`;
}

/* ---------------------------------------------------------------
 * Pretty date formatter: 2025-10-15 → Oct 15, 2025
 * -------------------------------------------------------------*/
function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return d.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "2-digit"
  });
}

/* ---------------------------------------------------------------
 * Fill the PDF fields (keeps fields EDITABLE – no flatten)
 * -------------------------------------------------------------*/
async function fillPdf(templateKey, record, map) {
  const pdfUrl = `./templates/${templateKey}.pdf`;

  const res = await fetch(pdfUrl, { cache: "no-store" });
  if (!res.ok) {
    throw new Error(`Template fetch failed ${res.status} ${res.statusText} at ${pdfUrl}`);
  }

  const ct = (res.headers.get("content-type") || "").toLowerCase();
  if (!ct.includes("pdf")) {
    const text = await res.text();
    throw new Error(`Expected PDF but got ${ct || "unknown"} from ${pdfUrl}. First 200 chars:\n` + text.slice(0, 200));
  }

  const pdfBytes = await res.arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();

  // Best-effort: ensure fields are not read-only if template set them
  try {
    for (const f of form.getFields()) {
      if (typeof f.enableReadOnly === "function") {
        // pdf-lib uses setReadOnly/enableReadOnly in newer versions; fall through gracefully
      }
      if (typeof f.disableReadOnly === "function") f.disableReadOnly();
      if (typeof f.setReadOnly === "function") f.setReadOnly(false);
    }
  } catch {}

  // Apply mapping
  for (const [crmField, pdfField] of Object.entries(map)) {
    const v = normalize(record[crmField] ?? record[crmField.split(".").pop()]);
    try {
      const field = form.getField(pdfField);
      const type = field.constructor.name;

      if (type === "PDFTextField") {
        field.setText(v);
      } else if (type === "PDFDropdown") {
        try { field.select(v); } catch { field.setText(v); }
      } else if (type === "PDFCheckBox") {
        const on = /^(yes|true|1|x)$/i.test(v);
        on ? field.check() : field.uncheck();
      } else if (type === "PDFRadioGroup") {
        try { field.select(v); } catch {}
      } else if (field.setText) {
        field.setText(v);
      }
    } catch (e) {
      // Missing field in PDF—skip silently
    }
  }

  // Keep fields editable; do NOT flatten.
  // Improve appearances so filled values are visible in Acrobat/Preview/Chrome.
  try {
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    form.updateFieldAppearances(font);
  } catch {
    // If embedding fails, most viewers still render text fine.
  }

  const outBytes = await pdfDoc.save({ updateFieldAppearances: false }); // retain AcroForm
  return new Blob([outBytes], { type: "application/pdf" });
}

/* ---------------------------------------------------------------
 * List all fields in current template (for debugging)
 * -------------------------------------------------------------*/
async function listPdfFields(templateKey) {
  const pdfUrl = `./templates/${templateKey}.pdf`;
  const res = await fetch(pdfUrl, { cache: "no-store" });
  if (!res.ok) throw new Error(`Template fetch failed ${res.status} ${res.statusText} at ${pdfUrl}`);
  const pdfBytes = await res.arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();
  const names = form.getFields().map(f => f.getName());
  return names.sort();
}

/* ---------------------------------------------------------------
 * Main app
 * -------------------------------------------------------------*/
async function run() {
  loadTemplates();
  let record = readPayloadFromWindowName();

  if (record) {
    setStatus("Record loaded from Zoho (secure window.name).", "ok");
  } else {
    setStatus("No Zoho payload detected. You can paste JSON below for testing.", "warn");
  }

  // Hook up list-fields button
  const listBtn = document.getElementById("listBtn");
  if (listBtn) {
    listBtn.addEventListener("click", async () => {
      try {
        const templateKey = els.sel.value || "mcap_std";
        setStatus(`Reading fields from ${templateKey}.pdf…`);
        const names = await listPdfFields(templateKey);
        els.fieldsBox.textContent = names.join("\n");
        setStatus(`Found ${names.length} fields.`, "ok");
      } catch (e) {
        console.error(e);
        setStatus(`Error listing fields: ${e.message}`, "err");
      }
    });
  }

  // Generate button
  els.btn.addEventListener("click", async () => {
    try {
      const templateKey = els.sel.value || "mcap_std";
      if (!TEMPLATE_KEYS.includes(templateKey)) throw new Error("Invalid template");

      // Fallback to manual JSON if no Zoho payload
      if (!record) {
        if (!els.txt.value.trim()) throw new Error("No record data. Launch from Zoho or paste JSON.");
        record = JSON.parse(els.txt.value);
      }

      // Format key dates
      [
        "Closing_Date",
        "COF_Date",
        "Borrower1_DOB",
        "Contact_Date_of_Birth_2",
        "Maturity_Date"
      ].forEach(k => {
        if (record[k]) record[k] = fmtDate(record[k]);
      });

      const map = await fetchJson("./map.json");
      setStatus("Generating PDF…", "muted");

      const blob = await fillPdf(templateKey, record, map);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = deriveFilename(templateKey, record);
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(a.href);

      setStatus("Done. The PDF has editable fields.", "ok");
    } catch (err) {
      console.error(err);
      setStatus(`Error: ${err.message}`, "err");
      alert(err.message);
    }
  });
}

run();

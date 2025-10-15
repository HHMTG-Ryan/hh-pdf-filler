/* HH PDF Filler – client-side only
 * ---------------------------------------------------------------
 * - Reads record data from window.name (set by Zoho Function)
 * - Loads /docs/map.json for CRM→PDF field mapping
 * - Loads chosen template PDF from /docs/templates/<key>.pdf
 * - Fills AcroForm fields, then downloads the file
 */

const TEMPLATE_KEYS = [
  "default",
  "scotia_std",
  "scotia_heloc",
  "firstnat_std"
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

function loadTemplates() {
  TEMPLATE_KEYS.forEach(k => {
    const opt = document.createElement("option");
    opt.value = k;
    opt.textContent = k.replace(/_/g, " ").toUpperCase();
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
 * Fill the PDF fields
 * -------------------------------------------------------------*/
async function fillPdf(templateKey, record, map) {
  const pdfUrl = `./templates/${templateKey}.pdf`;
  const pdfBytes = await (await fetch(pdfUrl, { cache: "no-store" })).arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();

  // console.log("PDF fields:", form.getFields().map(f => f.getName()));

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
      // console.warn(`Missing field in PDF: ${pdfField}`);
    }
  }

  form.flatten();
  const outBytes = await pdfDoc.save();
  return new Blob([outBytes], { type: "application/pdf" });
}

/* ---------------------------------------------------------------
 * Optional helper: list all fields in current template
 * -------------------------------------------------------------*/
async function listPdfFields(templateKey) {
  const pdfUrl = `./templates/${templateKey}.pdf`;
  const pdfBytes = await (await fetch(pdfUrl, { cache: "no-store" })).arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();
  const names = form.getFields().map(f => f.getName());
  return names.sort();
}

/* ---------------------------------------------------------------
 * Main logic
 * -------------------------------------------------------------*/
async function run() {
  loadTemplates();
  let record = readPayloadFromWindowName();

  if (record) {
    setStatus("Record loaded from Zoho (secure window.name).", "ok");
  } else {
    setStatus("No Zoho payload detected. You can paste JSON below for testing.", "warn");
  }

  // List-fields button
  const listBtn = document.getElementById("listBtn");
  if (listBtn) {
    listBtn.addEventListener("click", async () => {
      try {
        const templateKey = els.sel.value || "default";
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

  els.btn.addEventListener("click", async () => {
    try {
      const templateKey = els.sel.value || "default";
      if (!TEMPLATE_KEYS.includes(templateKey)) throw new Error("Invalid template");

      if (!record) {
        if (!els.txt.value.trim()) throw new Error("No record data. Launch from Zoho or paste JSON.");
        record = JSON.parse(els.txt.value);
      }

      // Format the key date fields before filling
      const dateKeys = [
        "Closing_Date",
        "COF_Date",
        "Date_Of_Birth",
        "Contact_Date_of_Birth_2",
        "Maturity_Date"
      ];
      dateKeys.forEach(k => {
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

      setStatus("Done. Check your download.", "ok");
    } catch (err) {
      console.error(err);
      setStatus(`Error: ${err.message}`, "err");
      alert(err.message);
    }
  });
}

run();

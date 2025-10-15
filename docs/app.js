/* HH PDF Filler – client-side only
 * - Reads record data from window.name (set by a Zoho Function redirect) or manual JSON
 * - Loads /docs/map.json for CRM→PDF field mapping
 * - Loads chosen template PDF from /docs/templates/<key>.pdf
 * - Fills AcroForm fields, then downloads as attachment
 */

// ---- Config: list your packages (filename without .pdf) ----
const TEMPLATE_KEYS = [
  "default",
  "scotia_std",
  "scotia_heloc",
  "firstnat_std"
  // add more keys that match files in /docs/templates/<key>.pdf
];

const els = {
  sel: document.getElementById("template"),
  btn: document.getElementById("fillBtn"),
  txt: document.getElementById("manualJson"),
  status: document.getElementById("status"),
};

function setStatus(msg, cls="muted") {
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
    // Expect base64 JSON set by Zoho Function
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
  const stamp = new Date().toISOString().slice(0,16).replace(/[-:T]/g,"");
  return `${templateKey}_${name}_${stamp}.pdf`;
}

async function fillPdf(templateKey, record, map) {
  const pdfUrl = `./templates/${templateKey}.pdf`;
  const pdfBytes = await (await fetch(pdfUrl, { cache: "no-store" })).arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();

  // For debugging missing fields:
  // console.log("PDF fields:", form.getFields().map(f => f.getName()));

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
        // tick for truthy strings commonly used
        const on = /^(yes|true|1|x)$/i.test(v);
        on ? field.check() : field.uncheck();
      } else if (type === "PDFRadioGroup") {
        try { field.select(v); } catch {}
      } else {
        // fallback to text when possible
        if (field.setText) field.setText(v);
      }
    } catch (e) {
      // Field not found—skip silently
      // console.warn(`Missing field in PDF: ${pdfField}`);
    }
  }

  // Flatten by making fields read-only
  form.flatten();

  const outBytes = await pdfDoc.save();
  return new Blob([outBytes], { type: "application/pdf" });
}

async function run() {
  loadTemplates();

  // Prefer payload from Zoho (window.name). If absent, rely on manual textarea.
  let record = readPayloadFromWindowName();
  if (record) {
    setStatus("Record loaded from Zoho (secure window.name).", "ok");
  } else {
    setStatus("No Zoho payload detected. You can paste JSON below for testing.", "warn");
  }

  els.btn.addEventListener("click", async () => {
    try {
      const templateKey = els.sel.value || "default";
      if (!TEMPLATE_KEYS.includes(templateKey)) throw new Error("Invalid template");

      // If no Zoho payload, try manual JSON
      if (!record) {
        if (!els.txt.value.trim()) throw new Error("No record data. Launch from Zoho or paste JSON.");
        record = JSON.parse(els.txt.value);
      }

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

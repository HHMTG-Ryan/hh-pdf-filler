/* HH PDF Filler – client-side only
 * ---------------------------------------------------------------
 * - Reads record data from ?data=<base64> (preferred) or window.name (legacy)
 * - Loads ./map.json for CRM→PDF field mapping
 * - Loads ./templates/<key>.pdf
 * - Fills AcroForm fields, keeps them EDITABLE, then downloads the file
 */

const TEMPLATE_KEYS = [
  "mcap_std","mcap_heloc","scotia_std","scotia_heloc","firstnat_std",
  "bridgewater_std","cmls_std","cwb_optimum_std","haventree_std",
  "lendwise_std","private_std","strive_std","td_std","td_heloc","signing_package"
];

const els = {
  sel: document.getElementById("template"),
  btn: document.getElementById("fillBtn"),
  txt: document.getElementById("manualJson"),
  status: document.getElementById("status"),
  fieldsBox: document.getElementById("fieldsBox"),
  listBtn: document.getElementById("listBtn")
};

function setStatus(msg, cls = "muted") {
  els.status.className = `row ${cls}`;
  els.status.textContent = msg;
}

function toTitleCase(str) {
  return str.replace(/_/g, " ").replace(/\b\w/g, ch => ch.toUpperCase());
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

/* -------------------- Payload readers -------------------- */

function b64urlToB64(s) {
  const norm = s.replace(/-/g, "+").replace(/_/g, "/").replace(/\s/g, "+");
  return norm.padEnd(Math.ceil(norm.length / 4) * 4, "=");
}

function tryParseJson(text) {
  try { const o = JSON.parse(text); return (o && typeof o === "object") ? o : null; }
  catch { return null; }
}

function readPayloadFromQuery() {
  try {
    const params = new URLSearchParams(location.search);
    let raw = params.get("data");
    if (!raw) return null;

    raw = decodeURIComponent(raw.trim());              // URL decode
    const decoded = atob(b64urlToB64(raw));            // Base64/Base64URL → bytes
    const obj = tryParseJson(decoded);                 // JSON parse

    if (obj) {
      const url = new URL(location.href);             // strip ?data= for clean refresh
      url.searchParams.delete("data");
      history.replaceState({}, "", url);
    }
    return obj;
  } catch (e) {
    console.error("Failed to parse ?data= payload", e);
    return null;
  }
}

function readPayloadFromWindowName() {
  try {
    if (!window.name) return null;
    const decoded = atob(b64urlToB64(window.name));
    return tryParseJson(decoded);
  } catch {
    return null;
  }
}

/* -------------------- Fetch helpers -------------------- */

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

/* -------------------- Formatting helpers -------------------- */

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "2-digit" });
}

/* -------------------- PDF helpers -------------------- */

async function listPdfFields(templateKey) {
  const pdfUrl = `./templates/${templateKey}.pdf`;
  const res = await fetch(pdfUrl, { cache: "no-store" });
  if (!res.ok) throw new Error(`Template fetch failed ${res.status} ${res.statusText} at ${pdfUrl}`);
  const pdfBytes = await res.arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  return pdfDoc.getForm().getFields().map(f => f.getName()).sort();
}

async function fillPdf(templateKey, record, map) {
  const pdfUrl = `./templates/${templateKey}.pdf`;
  const res = await fetch(pdfUrl, { cache: "no-store" });
  if (!res.ok) throw new Error(`Template fetch failed ${res.status} ${res.statusText} at ${pdfUrl}`);

  const ct = (res.headers.get("content-type") || "").toLowerCase();
  if (!ct.includes("pdf")) {
    const text = await res.text();
    throw new Error(`Expected PDF but got ${ct || "unknown"} from ${pdfUrl}. First 200 chars:\n` + text.slice(0, 200));
  }

  const pdfBytes = await res.arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();

  // Ensure fields are editable (best effort)
  try {
    for (const f of form.getFields()) {
      if (typeof f.disableReadOnly === "function") f.disableReadOnly();
      if (typeof f.setReadOnly === "function") f.setReadOnly(false);
    }
  } catch {}

  // Apply mapping
  for (const [crmField, pdfField] of Object.entries(map)) {
    const val = record[crmField];
    const alt = record[crmField.split(".").pop()];
    const v = normalize(val != null ? val : alt);

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
    } catch {
      // Missing field in PDF — ignore and continue
    }
  }

  // Improve appearances; keep fields editable
  try {
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    form.updateFieldAppearances(font);
  } catch {}

  const outBytes = await pdfDoc.save({ updateFieldAppearances: false });
  return new Blob([outBytes], { type: "application/pdf" });
}

/* -------------------- Main -------------------- */

(async function run() {
  loadTemplates();

  // Prefer ?data=, fallback to window.name
  let record = readPayloadFromQuery();
  if (record) {
    setStatus("Record loaded from Zoho (?data=).", "ok");
  } else {
    record = readPayloadFromWindowName();
    if (record) setStatus("Record loaded from Zoho (window.name).", "ok");
    else setStatus("No Zoho payload detected. You can paste JSON below for testing.", "warn");
  }

  if (els.listBtn) {
    els.listBtn.addEventListener("click", async () => {
      try {
        const key = els.sel.value || TEMPLATE_KEYS[0];
        setStatus(`Reading fields from ${key}.pdf…`);
        const names = await listPdfFields(key);
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
      const key = els.sel.value || TEMPLATE_KEYS[0];
      if (!TEMPLATE_KEYS.includes(key)) throw new Error("Invalid template");

      // Fallback to manual JSON if no Zoho payload
      let rec = record;
      if (!rec) {
        const raw = els.txt.value.trim();
        if (!raw) throw new Error("No record data. Launch from Zoho or paste JSON.");
        rec = JSON.parse(raw);
      }

      // Pretty common dates
      [
        "Closing_Date","COF_Date","Borrower1_DOB",
        "Contact_Date_of_Birth_2","Maturity_Date","Date_of_Birth"
      ].forEach(k => { if (rec[k]) rec[k] = fmtDate(rec[k]); });

      const map = await fetchJson("./map.json");
      setStatus("Generating PDF…", "muted");
      const blob = await fillPdf(key, rec, map);

      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = deriveFilename(key, rec);
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
})();

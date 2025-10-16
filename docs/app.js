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
  // Normalize base64url and stray spaces (+ can become space via URL encoding)
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

    // 1) URL-decode, 2) base64url→base64, 3) atob, 4) JSON.parse
    raw = decodeURIComponent(raw.trim());
    const decoded = atob(b64urlToB64(raw));
    const obj = tryParseJson(decoded);

    if (obj) {
      // Strip ?data= so refresh doesn't re-run decode
      const url = new URL(location.href);
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
  return d.toLocaleDateString(undefined, { year: "numeric", mont

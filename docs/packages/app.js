/* HH Signing Package Builder – Client-only
 * - Reads payload from ?data= (base64/url-safe) or window.name
 * - Uses packages/standard_pkg.json for order & rules
 * - Fetches templates from ./templates
 * - Builds:
 *    1) Signing Package (merged + flattened) → "<LASTNAME> - Signing Pkg.pdf"
 *    2) Editable Singles (ZIP): KYC/IDV, FCT, Gift (if present)
 *    3) Info Cover Page (1-pager with key payload data)
 */

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

function setStatus(msg, cls="muted"){ els.status.className = cls; els.status.textContent = msg; }

// ---------- Payload readers ----------
function b64urlToB64(s){
  const n = s.replace(/-/g,"+").replace(/_/g,"/").replace(/\s/g,"+");
  return n.padEnd(Math.ceil(n.length/4)*4,"=");
}
function tryJSON(t){ try{const o=JSON.parse(t); return o&&typeof o==="object"?o:null;}catch{return null;} }
function readFromQuery(){
  try{
    const p = new URLSearchParams(location.search);
    let raw = p.get("data"); if(!raw) return null;
    raw = decodeURIComponent(raw.trim());
    const obj = tryJSON(atob(b64urlToB64(raw)));
    if(obj){
      const u=new URL(location.href); u.searchParams.delete("data"); history.replaceState({}, "", u);
    }
    return obj;
  }catch{ return null; }
}
function readFromWindowName(){
  try{ if(!window.name) return null; return tryJSON(atob(b64urlToB64(window.name))); }catch{ return null; }
}

// ---------- Utils ----------
function sanitizeName(s){
  return (s||"").toString().normalize("NFKD").replace(/[^\w\- ]+/g,"").trim();
}
function lastNameFromContact(contactName){
  if(!contactName) return "Client";
  const parts = contactName.trim().split(/\s+/);
  return sanitizeName(parts[parts.length-1] || "Client");
}
function inferLenderPrefix(lenderName){
  const n = (lenderName||"").toLowerCase();
  // add aliases here as needed:
  const candidates = [
    "td","mcap","cmls","cwb","firstnat","haventree","lendwise","scotia","strive","bridgewater","private"
  ];
  for(const c of candidates){ if(n.includes(c)) return c.toUpperCase(); }
  // fallback: titlecase first token
  if(n) return sanitizeName(n.split(/\s+/)[0]).toUpperCase();
  return "TD";
}
function isTD(prefix){ return (prefix||"").toUpperCase().includes("TD"); }

async function fetchJson(url){
  const r = await fetch(url, { cache: "no-store" });
  if(!r.ok) throw new Error(`Failed to fetch ${url}`);
  return r.json();
}
async function fetchPdfBytes(path){ // returns ArrayBuffer
  const r = await fetch(path, { cache:"no-store" });
  if(!r.ok) throw new Error(`Missing PDF: ${path}`);
  return await r.arrayBuffer();
}
async function tryFetchPdfBytes(path){ // returns ArrayBuffer|null
  try{ return await fetchPdfBytes(path); } catch{ return null; }
}

// ---------- Filling helpers ----------
function normalize(v){ if(v==null) return ""; if(typeof v==="boolean") return v?"Yes":"No"; return String(v); }
async function fillTemplateBytes(templatePath, record, map){ // returns Uint8Array
  const bytes = await fetchPdfBytes(templatePath);
  const pdfDoc = await PDFLib.PDFDocument.load(bytes);
  const form = pdfDoc.getForm();
  for (const [crmKey, pdfField] of Object.entries(map||{})){
    const val = record[crmKey] ?? record[crmKey.split(".").pop()];
    const v = normalize(val);
    try{
      const field = form.getField(pdfField);
      const type = field.constructor.name;
      if (type === "PDFTextField") field.setText(v);
      else if (type === "PDFDropdown"){ try{ field.select(v); }catch{ field.setText(v); } }
      else if (type === "PDFCheckBox"){ /^(yes|true|1|x)$/i.test(v) ? field.check() : field.uncheck(); }
      else if (type === "PDFRadioGroup"){ try{ field.select(v); }catch{} }
      else if (field.setText) field.setText(v);
    }catch{ /* field not present, skip */ }
  }
  // Improve appearances; keep editable for singles, will flatten for package later
  try{
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    pdfDoc.getForm().updateFieldAppearances(font);
  }catch{}
  return await pdfDoc.save({ updateFieldAppearances:false });
}

// ---------- Merge/flatten ----------
async function flattenAllPages(bytes){ // returns flattened Uint8Array
  const doc = await PDFLib.PDFDocument.load(bytes);
  try{ const form = doc.getForm(); form && form.flatten(); }catch{}
  return await doc.save({ updateFieldAppearances:true });
}
async function mergeDocsFlattened(listOfBytes){ // Array<Uint8Array>
  const out = await PDFLib.PDFDocument.create();
  for (const b of listOfBytes){
    // ensure flattened before merge
    const flat = await flattenAllPages(b);
    const src = await PDFLib.PDFDocument.load(flat);
    const pages = await out.copyPages(src, src.getPageIndices());
    pages.forEach(p => out.addPage(p));
  }
  return await out.save({ updateFieldAppearances:true });
}

// ---------- Headers slicing ----------
async function makeHeaderPageBytes(headersPdfPath, headerIndex1Based){
  const headersBytes = await fetchPdfBytes(headersPdfPath);
  const src = await PDFLib.PDFDocument.load(headersBytes);
  const out = await PDFLib.PDFDocument.create();
  const [page] = await out.copyPages(src, [headerIndex1Based-1]);
  out.addPage(page);
  return await out.save();
}

// ---------- Singles ZIP ----------
async function downloadZip(namedBytes){
  const zip = new JSZip();
  for(const it of namedBytes){ zip.file(it.name, it.bytes); }
  const blob = await zip.generateAsync({ type:"blob" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "Editable_Singles.zip";
  document.body.appendChild(a); a.click(); a.remove();
}

// ---------- Info Cover Page ----------
async function makeInfoCoverPdf(record){
  const doc = await PDFLib.PDFDocument.create();
  const page = doc.addPage([595, 842]); // A4-ish
  const { width } = page.getSize();
  const font = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  const fontB = await doc.embedFont(PDFLib.StandardFonts.HelveticaBold);
  let y = 800;
  function line(lbl, val){ page.drawText(lbl, { x:40, y: y, size:11, font:fontB }); page.drawText(String(val||""), { x:200, y: y, size:11, font }); y -= 18; }

  page.drawText("HH – Info Cover Page", { x:40, y: y, size:16, font:fontB }); y -= 28;
  line("Deal / File Name:", record.Deal_Name || record.Name);
  line("Contact Name:", record.Contact_Name);
  line("Lender:", record.Lender_Name);
  line("Address:", [record.Street, record.City, record.Province, record.Postal_Code].filter(Boolean).join(", "));
  line("Closing Date:", record.Closing_Date);
  line("COF Date:", record.COF_Date);
  line("Total Mtg Amt (incl. ins.):", record.Total_Mortgage_Amount_incl_Insurance);
  line("Email (B1):", record.Email);
  line("Email (B2):", record.Contact_Email_2);

  const bytes = await doc.save({ updateFieldAppearances:true });
  return new Blob([bytes], { type:"application/pdf" });
}

// ---------- Lender file resolver ----------
function lenderFile(prefix, suffix){ return `${prefix}_${suffix}.pdf`; }

// ---------- Main run ----------
(async function run(){
  setStatus("Loading payload…");

  // 1) payload
  let record = readFromQuery() || readFromWindowName();
  if(!record && els.manualJson.value.trim()){ record = tryJSON(els.manualJson.value.trim()); }
  if(!record){ setStatus("No Zoho payload detected. Paste JSON or relaunch from CRM.", "warn"); }

  // 2) package manifest
  const manifest = await fetchJson("./packages/standard_pkg.json");

  // 3) lender detection and dropdown
  const detectedPrefix = inferLenderPrefix(record?.Lender_Name || "");
  const lenderOptions = ["TD","MCAP","CMLS","CWB","FirstNat","Haventree","Lendwise","Scotia","Strive","Bridgewater","Private"];
  els.lender.innerHTML = lenderOptions.map(v => `<option value="${v}" ${v===detectedPrefix?"selected":""}>${v}</option>`).join("");

  // TD APA visibility
  function refreshAPA(){ els.apaWrap.style.display = isTD(els.lender.value) ? "block":"none"; }
  refreshAPA();
  els.lender.addEventListener("change", refreshAPA);

  // 4) map.json (for filling fields)
  let map = {};
  try{ map = await fetchJson("./map.json"); }catch{ map = {}; }

  // 5) Buttons
  els.btnPkg.addEventListener("click", async ()=>{
    try{
      // validations
      if(!record && !els.manualJson.value.trim()){ throw new Error("No record payload. Paste JSON or relaunch from CRM."); }
      if(!els.upCommitment.files[0]) throw new Error("Commitment is required.");
      if(!els.upApplication.files[0]) throw new Error("Mortgage Application is required.");

      setStatus("Building signing package…","muted");
      const lenderPrefix = els.lender.value;                 // e.g., TD
      const pkgType = els.pkg.value;                         // "standard" or "heloc"
      const headersPath = `./templates/${manifest.headers_pdf}`;

      // Resolve header index by section order
      const headerIndex = {
        WMB: 1, MPP_OR_PROSPR: 2, COMMITMENT: 3, COB: 4, FORM10: 5, APPLICATION: 6
      };

      // Preload uploads as bytes
      const upCommitmentBytes = await els.upCommitment.files[0].arrayBuffer();
      const upApplicationBytes = await els.upApplication.files[0].arrayBuffer();
      const upMPPBytes = els.upMPP.files[0] ? await els.upMPP.files[0].arrayBuffer() : null;
      const upAPABtes = els.upAPA.files[0] ? await els.upAPA.files[0].arrayBuffer() : null;

      // Build list of doc bytes (flattened) in final order, inserting headers where needed
      const toMerge = [];

      // Helper: push header page
      async function pushHeader(key){
        const idx = headerIndex[key];
        if(!idx) return;
        const hb = await makeHeaderPageBytes(headersPath, idx);
        toMerge.push(hb);
      }

      // Helper: push template or upload
      async function pushDocBySpec(spec){
        if(spec.no_header !== true && spec.header){ await pushHeader(spec.header); }

        if(spec.doc_if_upload || spec.doc_if_missing){
          // MPP vs Prospr
          if(upMPPBytes){ toMerge.push(upMPPBytes); } else {
            const path = `./templates/${manifest.prospr_fallback}`;
            const bytes = await fetchPdfBytes(path);
            toMerge.push(bytes);
          }
          return;
        }

        if(spec.doc === "UPLOAD_Commitment.pdf"){ toMerge.push(upCommitmentBytes); return; }
        if(spec.doc === "UPLOAD_Application.pdf"){ toMerge.push(upApplicationBytes); return; }
        if(spec.doc === "UPLOAD_APA_TD.pdf"){
          if(isTD(lenderPrefix) && upAPABtes){ toMerge.push(upAPABtes); }
          return;
        }

        // conditionals
        if(spec.condition === "lender_has_TD" && !isTD(lenderPrefix)) return;
        if(spec.condition === "package_is_HELOC" && pkgType !== "heloc") return;
        if(spec.condition === "package_is_STANDARD" && pkgType !== "standard") return;

        // LENDER_* resolution
        let path;
        if(spec.doc.startsWith("LENDER_")){
          const suffix = spec.doc.replace(/^LENDER_/, "");
          path = `./templates/${lenderFile(lenderPrefix, suffix)}`;
          const tryBytes = await tryFetchPdfBytes(path);
          if(!tryBytes){
            if(spec.optional) return; // skip missing optional
            else throw new Error(`Missing required lender doc: ${path}`);
          }
          toMerge.push(tryBytes);
          return;
        }

        // Static template
        path = `./templates/${spec.doc}`;
        const bytes = await fetchPdfBytes(path);
        toMerge.push(bytes);
      }

      // Walk manifest sequence
      for(const spec of manifest.sequence){ await pushDocBySpec(spec); }

      // HELOC rule: already encoded in manifest (COB both; FORM10 helox only)

      // Flatten & merge
      const flattenedMerged = await mergeDocsFlattened(toMerge);

      // Build filename
      const lastName = lastNameFromContact(record?.Contact_Name || "");
      const outName = `${lastName} - Signing Pkg.pdf`;

      // Download
      const blob = new Blob([flattenedMerged], { type:"application/pdf" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = outName;
      document.body.appendChild(a); a.click(); a.remove();

      setStatus(`Signing package created: ${outName}`,"ok");
    }catch(e){
      console.error(e);
      setStatus(`Error: ${e.message}`,"err");
      alert(e.message);
    }
  });

  els.btnSingles.addEventListener("click", async ()=>{
    try{
      if(!record && !els.manualJson.value.trim()){ throw new Error("No record payload. Paste JSON or relaunch from CRM."); }
      setStatus("Generating editable singles…","muted");

      const lenderPrefix = els.lender.value;
      const map = await (async()=>{ try{ return await fetchJson("./map.json"); }catch{return {};} })();

      const singles = [];
      // KYC/IDV (always)
      {
        const path = "./templates/KYC_IDV_Template.pdf";
        const filled = await fillTemplateBytes(path, record, map);
        singles.push({ name: "Identification Verification.pdf", bytes: filled });
      }
      // FCT Authorization (if exists)
      {
        const path = `./templates/${lenderFile(lenderPrefix, "FCT_Auth.pdf")}`;
        const b = await tryFetchPdfBytes(path);
        if(b){
          const filled = await fillTemplateBytes(path, record, map);
          singles.push({ name: "FCT Authorization.pdf", bytes: filled });
        }
      }
      // Gift Letter (if exists)
      {
        const path = `./templates/${lenderFile(lenderPrefix, "Gift_Letter.pdf")}`;
        const b = await tryFetchPdfBytes(path);
        if(b){
          const filled = await fillTemplateBytes(path, record, map);
          singles.push({ name: "Gift Letter.pdf", bytes: filled });
        }
      }

      if(singles.length === 0){ throw new Error("No editable singles available for this lender."); }
      await downloadZip(singles);
      setStatus("Editable singles downloaded (ZIP).","ok");
    }catch(e){
      console.error(e);
      setStatus(`Error: ${e.message}`,"err");
      alert(e.message);
    }
  });

  els.btnInfo.addEventListener("click", async ()=>{
    try{
      if(!record && !els.manualJson.value.trim()){ throw new Error("No record payload. Paste JSON or relaunch from CRM."); }
      setStatus("Generating info cover page…","muted");
      const rec = record || tryJSON(els.manualJson.value.trim()) || {};
      const blob = await makeInfoCoverPdf(rec);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "Info Cover Page.pdf";
      document.body.appendChild(a); a.click(); a.remove();
      setStatus("Info cover page downloaded.","ok");
    }catch(e){
      console.error(e);
      setStatus(`Error: ${e.message}`,"err");
      alert(e.message);
    }
  });

  // Initial status
  const ln = record?.Lender_Name || "";
  const auto = inferLenderPrefix(ln);
  setStatus(record ? `Payload OK. Lender ≈ ${auto}. Choose options and upload files.` : "Paste JSON or launch from CRM.", record ? "ok" : "warn");
})();

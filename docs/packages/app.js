// ======= CONFIG mirrored from Python (will externalize in Step 7) =======
const CLOSING_ORDER = [
  "commitment only","kyc","idv","pep","mtg application","insurance","consent","cost of credit","form 10 only","form 10 to lender"
];
const MULTI_INCLUDE_SLOTS = new Set(["kyc","pep","consent","insurance"]);
const DISPLAY_NAMES = {
  "commitment only":"Commitment Only","kyc":"KYC","idv":"IDV","pep":"PEP","mtg application":"Mtg Application","insurance":"Insurance","consent":"Consent","cost of credit":"Cost of Borrowing","form 10 only":"Form 10 Only","form 10 to Lender":"Form 10 to Lender",
};
const CLIENT_PKG_ORDER = ["thank you","amortization","commitment only","form 10 only","cost of credit"];
const SYNONYMS = {
  "commitment only":["commit only","commitment"],
  "mtg application":["mortgage application","mtg app","mortgage app","application"],
  "idv":["idv","verification"],
  "cost of credit":["cob","cost of borrowing","disclosure"],
  "insurance":["prospr","mpp","manulife","sun life","sunlife","indem","assurant"],
  "form 10 only":["form 10","form10","broker compensation"],
  "form 10 to lender":["form 10 to lender","form10 to lender","compensation to lender"],
  "kyc":["know your client","know-your-client","client information"],
  "pep":["pep","politically exposed"],
  "consent":["consent","authorization","authorisation"],
  "thank you":["thank you","thank-you","thanks"],
  "amortization":["amortization","amortisation","amort sched","amort schedule"],
};

// ======= State / utils =======
let pickedFiles = [];
let pickedMap = new Map();
let dirHandle = null;
let lastClosingUsedNames = new Set();
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

const $ = (sel) => document.querySelector(sel);
function log(m){
  const el = document.getElementById('log');
  if(!el) return;
  el.textContent += m + '\n';
  // Optional cap to prevent runaway logs
  if (el.textContent.length > 200000) el.textContent = el.textContent.slice(-150000);
  el.scrollTop = el.scrollHeight;
}
const norm = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const natKey = (s) => s.toLowerCase().split(/(\d+)/).map(t=>/^\d+$/.test(t)?Number(t):t);

// ======= Status UI (auto-create if missing) =======
function ensureStatus(){
  let el = document.getElementById('status');
  if (!el) {
    el = document.createElement('div');
    el.id = 'status';
    el.className = 'status';
    el.setAttribute('role','status');
    el.style.margin = '12px 0';
    el.style.fontSize = '14px';
    el.innerHTML = `<span class="spinner" style="display:none;margin-right:6px">⏳</span><span class="status-text"></span>`;
    const main = document.querySelector('main') || document.body;
    main.insertBefore(el, main.firstChild);
  }
  return el;
}
function setStatus(text, state='running'){
  const el = ensureStatus();
  el.removeAttribute('hidden');
  el.dataset.state = state; // running | ready | error
  const txt = el.querySelector('.status-text'); if(txt) txt.textContent = text || '';
  const spin = el.querySelector('.spinner'); if(spin) spin.style.display = (state === 'running') ? 'inline-block' : 'none';
}
function clearStatus(){
  const el = ensureStatus();
  el.dataset.state = 'ready';
  const txt = el.querySelector('.status-text'); if(txt) txt.textContent = '';
  el.setAttribute('hidden','');
}

// Button busy helper: disable & swap label while a task runs
async function withBusy(btnId, runningText, fn){
  const btn = document.getElementById(btnId);
  const original = btn ? btn.textContent : '';
  try{
    if(btn){ btn.disabled = true; btn.textContent = runningText; }
    setActionsEnabled(false); // belt-and-suspenders: disable all actions during run
    return await fn();
  } finally {
    if(btn){ btn.disabled = false; btn.textContent = original; }
    setActionsEnabled(pickedFiles.length > 0);
  }
}

// ======= FS Access helpers =======
async function grantFolderAccess(){
  if(!('showDirectoryPicker' in window)){
    log('This browser does not support Local Folder Access. Use Chrome/Edge for write mode.');
    return;
  }
  try{
    dirHandle = await window.showDirectoryPicker();
    $('#fsBadge').textContent = 'Folder Write: ON (' + (dirHandle.name || 'selected') + ')';
    $('#fsBadge').classList.remove('muted');
    log('Folder access granted. Writes/moves will target: ' + (dirHandle.name || 'selected folder'));
  }catch(e){ log('Folder access cancelled.'); }
}

async function ensureDir(parent, name){
  return await parent.getDirectoryHandle(name, { create: true });
}
async function writeFile(parent, name, blob){
  const fileHandle = await parent.getFileHandle(name, { create: true });
  const w = await fileHandle.createWritable();
  await w.write(blob); await w.close();
}
async function readTopLevelPDFsFromFS(){
  const files = [];
  if(!dirHandle) return files;
  for await (const [name, handle] of dirHandle.entries()){
    if(handle.kind === 'file' && name.toLowerCase().endsWith('.pdf')){
      const f = await handle.getFile();
      files.push(f);
    }
  }
  files.sort((a,b)=>{ const ak=natKey(a.name), bk=natKey(b.name);
    for(let i=0;i<Math.max(ak.length,bk.length);i++){
      if(ak[i]==null) return -1; if(bk[i]==null) return 1;
      if(ak[i]<bk[i]) return -1; if(ak[i]>bk[i]) return 1;
    } return 0; });
  return files;
}
async function moveFileToFD2ByName(name){
  try{
    const fd2 = await ensureDir(dirHandle, 'Funding Docs 2');
    const srcHandle = await dirHandle.getFileHandle(name);
    const srcFile = await srcHandle.getFile();
    const buf = await srcFile.arrayBuffer();
    await writeFile(fd2, name, new Blob([buf], {type:'application/pdf'}));
    await dirHandle.removeEntry(name);
    await sleep(50);
    log('Moved to Funding Docs 2: ' + name);
  }catch(e){ log('Move failed (skip if not in top-level): ' + name); }
}

// Read bytes from in-memory File, else fall back to disk (root, then FD2)
async function readBytesSmart(file){
  try { return await file.arrayBuffer(); } catch(e) {}
  if(!dirHandle) throw new Error('File not readable and no FS handle');
  try {
    const h = await dirHandle.getFileHandle(file.name);
    const f = await h.getFile();
    return await f.arrayBuffer();
  } catch(e) {}
  try {
    const fd2 = await ensureDir(dirHandle, 'Funding Docs 2');
    const h2 = await fd2.getFileHandle(file.name);
    const f2 = await h2.getFile();
    return await f2.arrayBuffer();
  } catch(e) {}
  throw new Error('Cannot read: ' + file.name);
}

// ======= Selection / Index =======
function indexFiles(files){
  pickedMap.clear();
  const pdfs = [...files].filter(f=> f.name.toLowerCase().endsWith('.pdf'));
  pdfs.sort((a,b)=>{ const ak=natKey(a.name), bk=natKey(b.name);
    for(let i=0;i<Math.max(ak.length,bk.length);i++){
      if(ak[i]==null) return -1; if(bk[i]==null) return 1;
      if(ak[i]<bk[i]) return -1; if(ak[i]>bk[i]) return 1;
    } return 0; });
  pickedFiles = pdfs;
  for(const f of pdfs){
    const n = norm(f.name);
    if(!pickedMap.has(n)) pickedMap.set(n,[]);
    pickedMap.get(n).push(f);
  }
  const sel = $('#selInfo'); if(sel) sel.textContent = `${pdfs.length} PDF(s) selected`;
  log(`Indexed ${pdfs.length} PDFs.`);
}

// ======= Matching =======
function listCandidates(){
  const all = Array.from(pickedMap.values()).flat();
  const lowerCache = new Map(all.map(f=>[f, norm(f.name)]));
  const resUsed = []; const summary = []; const slotDocs = new Map(CLOSING_ORDER.map(s=>[s,[]]));

  const uniqByName = (arr)=>{
    const seen = new Set(); const out=[];
    for(const f of arr){ if(!seen.has(f.name)){ seen.add(f.name); out.push(f); } }
    return out;
  };

  for(const slot of CLOSING_ORDER){
    const keys = [slot, ...(SYNONYMS[slot]||[])].map(norm);
    let matches = all.filter(f => keys.some(k => lowerCache.get(f).includes(k)));
    const nameNorm = (f)=> lowerCache.get(f);
    if(slot === 'cost of credit'){
      const cobRe = /(?<![a-z])cob(?![a-z])|\bcost[\s_-]*of[\s_-]*(?:borrowing|credit)\b/i;
      matches = matches.filter(f=> cobRe.test(nameNorm(f)) && !/(form 10|form10|compensation|to lender)/i.test(nameNorm(f)) );
    }
    if(slot === 'form 10 only') matches = matches.filter(f=> !/to lender/i.test(nameNorm(f)) );

    if(MULTI_INCLUDE_SLOTS.has(slot)){
      const chosen = uniqByName(matches.sort((a,b)=> a.lastModified - b.lastModified));
      slotDocs.set(slot, chosen);
      if(chosen.length){
        chosen.forEach(m=> { resUsed.push(m); summary.push(`${DISPLAY_NAMES[slot]||slot}: ${m.name}`); });
      } else {
        summary.push(`${DISPLAY_NAMES[slot]||slot}: MISSING`);
      }
    } else {
      const chosen = matches.sort((a,b)=> a.lastModified - b.lastModified).slice(-1)[0];
      if(chosen){ slotDocs.set(slot, [chosen]); resUsed.push(chosen); summary.push(`${DISPLAY_NAMES[slot]||slot}: ${chosen.name}`); }
      else { summary.push(`${DISPLAY_NAMES[slot]||slot}: MISSING`); }
    }
  }
  const seenUsed = new Set(); const resUsedUniq=[];
  for(const f of resUsed){ if(!seenUsed.has(f.name)){ seenUsed.add(f.name); resUsedUniq.push(f); } }
  return {slotDocs, resUsed: resUsedUniq, summary};
}

function renderSlots(view){
  const root = $('#slots'); if(!root) return;
  root.innerHTML = '';
  const ul = document.createElement('ul'); ul.className = 'inline';
  for(const slot of CLOSING_ORDER){
    const arr = view.slotDocs.get(slot) || [];
    const li = document.createElement('li'); li.className='pill';
    const ok = arr.length>0;
    li.innerHTML = `<span class="${ok? 'ok':'miss'}">${ok? '✔':'✖'}</span> <b>${DISPLAY_NAMES[slot]||slot}</b>${ ok? ` · ${arr.length} file(s)` : ''}`;
    ul.appendChild(li);
  }
  root.appendChild(ul);
}

// ======= PDF helpers (encryption tolerant) =======
async function mergePDFs(parts, summaryBlob, outName, writeToFolder=false){
  if (typeof window.PDFLib === 'undefined') {
    const msg = 'PDFLib not loaded. Check vendor path/order.';
    console.error(msg); log(msg); throw new Error(msg);
  }
  const { PDFDocument } = PDFLib;

  async function loadPdf(bytes){
    try { return await PDFDocument.load(bytes); }
    catch(e){ return await PDFDocument.load(bytes, { ignoreEncryption: true }); }
  }

  try {
    const out = await PDFDocument.create();

    if(summaryBlob){
      const sumBytes = await summaryBlob.arrayBuffer();
      const sumDoc = await loadPdf(sumBytes);
      const pages = await out.copyPages(sumDoc, sumDoc.getPageIndices());
      pages.forEach(p=> out.addPage(p));
    }

    const seen = new Set(); const uniqueParts=[];
    for(const f of parts){ if(!seen.has(f.name)){ seen.add(f.name); uniqueParts.push(f); } }

    for(const file of uniqueParts){
      try{
        const bytes = await readBytesSmart(file);
        const src = await loadPdf(bytes);
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(p=> out.addPage(p));
      }catch(e){ log(`Merge warning: skipped encrypted or unreadable file → ${file.name}`); }
    }

    const outBytes = await out.save();
    const outBlob = new Blob([outBytes], {type:'application/pdf'});
    if(writeToFolder && dirHandle){
      await writeFile(dirHandle, outName, outBlob);
      log('Saved directly: ' + outName);
    } else {
      downloadBlob(outBlob, outName);
    }
  } catch (e) {
    console.error(e); log('Merge error: ' + (e?.message || e));
    throw e;
  }
}

function downloadBlob(blob, name){
  const a = document.createElement('a');
  const url = URL.createObjectURL(blob);
  a.href = url; a.download = name; document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 1500);
}

async function makeSummaryPDF(title, lines){
  if (typeof window.PDFLib === 'undefined') {
    const msg = 'PDFLib not loaded. Check vendor path/order.';
    console.error(msg); log(msg); throw new Error(msg);
  }
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const doc = await PDFDocument.create();
  const page = doc.addPage([612, 792]); // Letter
  const font = await doc.embedFont(StandardFonts.Helvetica);

  page.drawText(title, { x:54, y:740, size:20, font, color: rgb(0.1,0.15,0.2) });
  let y = 710;
  for(const line of lines){
    const isMissing = /MISSING$/.test(line);
    page.drawText(line, { x:54, y, size:12, font, color: isMissing? rgb(0.75,0,0): rgb(0.2,0.25,0.3) });
    y -= 16; if(y<60){ y = 740; doc.addPage([612,792]); }
  }
  const bytes = await doc.save();
  return new Blob([bytes], {type:'application/pdf'});
}

// ======= Output naming =======
function folderPrefix(){
  let raw = (typeof dirHandle?.name === 'string' && dirHandle.name)
    ? dirHandle.name
    : (() => {
        const f = pickedFiles[0];
        if(!f) return '';
        const segs = (f.webkitRelativePath || f.name).split('/').filter(Boolean);
        return segs.length ? segs[0] : '';
      })();

  raw = (raw || '').trim();
  const parts = raw.split(/[^\p{L}\p{N}]+/u).filter(Boolean);
  return parts.length ? parts[0] : 'Client';
}

// ======= Build steps =======
async function buildClosing(){
  setStatus('Building Closing Docs…','running');
  log('Building Closing Docs...');
  const v = listCandidates();
  const summaryBlob = await makeSummaryPDF('Closing Docs', v.summary);

  const parts = (()=>{ 
    const seen=new Set(); const out=[];
    for(const slot of CLOSING_ORDER){
      const arr=v.slotDocs.get(slot)||[];
      for(const f of arr){ if(!seen.has(f.name)){ seen.add(f.name); out.push(f); } }
    }
    return out;
  })();
  const outName = folderPrefix()+" - Closing Docs.pdf";

  await mergePDFs(parts, summaryBlob, outName, !!dirHandle);

  lastClosingUsedNames = new Set(parts.map(f=>f.name));

  if(dirHandle){
    const fd2 = await ensureDir(dirHandle, 'Funding Docs 2');
    await writeFile(fd2, `_summary_closing_docs.pdf`, summaryBlob);
    for(const f of parts){ await moveFileToFD2ByName(f.name); }
    await sleep(300);
  }
  log('Closing Docs ready.');
  setStatus('Closing Docs complete.','ready');
}

async function buildFunding1(){
  try {
    setStatus('Building Funding Docs 1…', 'running');
    log('Building Funding Docs 1...');

    const signingRe = /(signing\s*(package|pkg))/i;
    let files;

    if (dirHandle) {
      const exclude = new Set([
        (folderPrefix()+" - Closing Docs.pdf").toLowerCase(),
        (folderPrefix()+" - Funding Docs 1.pdf").toLowerCase(),
        (folderPrefix()+" - Closing Package (for Client).pdf").toLowerCase()
      ]);
      const topLevel = await readTopLevelPDFsFromFS();

      const signersTop = topLevel.filter(f => signingRe.test(f.name));
      if (signersTop.length) for (const f of signersTop) await moveFileToFD2ByName(f.name);

      files = topLevel.filter(f =>
        !f.name.toLowerCase().startsWith('_summary') &&
        !exclude.has(f.name.toLowerCase()) &&
        !signingRe.test(f.name) &&
        !(lastClosingUsedNames && lastClosingUsedNames.has(f.name))
      );
    } else {
      const exclude = new Set([
        (folderPrefix()+" - Closing Docs.pdf").toLowerCase(),
        (folderPrefix()+" - Funding Docs 1.pdf").toLowerCase(),
        (folderPrefix()+" - Closing Package (for Client).pdf").toLowerCase()
      ]);
      files = pickedFiles.filter(f =>
        !/^_summary/i.test(f.name) &&
        !signingRe.test(f.name) &&
        !exclude.has(f.name.toLowerCase()) &&
        !(lastClosingUsedNames && lastClosingUsedNames.has(f.name))
      );
    }

    const seen = new Set();
    files = files.filter(f => !seen.has(f.name) && seen.add(f.name));

    log(`Funding1: ${files.length} file(s) to merge.`);
    const names = files.map(f => f.name).sort();
    const summaryBlob = await makeSummaryPDF('Funding Docs', names);
    const outName = folderPrefix()+" - Funding Docs 1.pdf";

    await mergePDFs(files, summaryBlob, outName, !!dirHandle);

    if (dirHandle) {
      const fd2 = await ensureDir(dirHandle, 'Funding Docs 2');
      await writeFile(fd2, `_summary_funding_docs.pdf`, summaryBlob);
      for (const f of files) await moveFileToFD2ByName(f.name);
    }

    log('Funding Docs 1 ready.');
    setStatus('Funding Docs 1 complete.', 'ready');
  } catch (e) {
    console.error(e);
    log('Error building Funding Docs 1: ' + (e?.message || e));
    setStatus('Funding Docs 1 failed.', 'error');
  }
}

async function buildClientPkg(){
  setStatus('Building Client Closing Package…','running');
  log('Building Client Closing Package...');
  if (typeof window.PDFLib === 'undefined') {
    const msg = 'PDFLib not loaded. Check vendor path/order.';
    console.error(msg); log(msg); setStatus('Client Package failed.','error'); return;
  }
  const all = Array.from(pickedMap.values()).flat();
  const lower = new Map(all.map(f=>[f, norm(f.name)]));
  const pickOne = (slot, matches)=> matches.sort((a,b)=> a.lastModified - b.lastModified).slice(-1)[0];

  const parts = []; const lines = [];
  const nameNorm = f => lower.get(f);

  let chosenThankYou = null; let chosenAmort = null;

  for(const slot of CLIENT_PKG_ORDER){
    const label = DISPLAY_NAMES[slot] || slot.replace(/\b\w/g,c=>c.toUpperCase());
    const keys = [slot, ...(SYNONYMS[slot]||[])].map(norm);
    let matches = all.filter(f => keys.some(k => nameNorm(f).includes(k)));
    if(slot==='cost of credit'){
      const cobRe = /(?<![a-z])cob(?![a-z])|\bcost[\s_-]*of[\s_-]*(?:borrowing|credit)\b/i;
      matches = matches.filter(f=> cobRe.test(nameNorm(f)) && !/(form 10|form10|compensation|to lender)/i.test(nameNorm(f)) );
    }
    if(slot==='form 10 only') matches = matches.filter(f=> !/to lender/i.test(nameNorm(f)) );

    const chosen = pickOne(slot, matches);
    if(chosen){
      if(!parts.find(p=>p.name===chosen.name)){ parts.push(chosen); lines.push(`${label}: ${chosen.name}`); }
      else { lines.push(`${label}: ${chosen.name} (deduped)`); }
      if(slot==='thank you') chosenThankYou = chosen;
      if(slot==='amortization') chosenAmort = chosen;
    } else { lines.push(`${label}: MISSING`); }
  }

  const cover = await makeSummaryPDF('Client Closing Package', lines);
  const outName = folderPrefix()+" - Closing Package (for Client).pdf";
  await mergePDFs(parts, cover, outName, !!dirHandle);

  if(dirHandle){
    const fd2 = await ensureDir(dirHandle, 'Funding Docs 2');
    await writeFile(fd2, `_summary_client_package.pdf`, cover);
    for(const f of [chosenThankYou, chosenAmort]){ if(f) await moveFileToFD2ByName(f.name); }
  }
  log('Client Closing Package ready.');
  setStatus('Client Package complete.','ready');
}

async function downloadFD2Zip(){
  try {
    setStatus('Preparing Funding Docs 2.zip…','running');
    log('Preparing Funding Docs 2.zip ...');
    if (typeof window.JSZip === 'undefined') {
      const msg = 'JSZip not loaded. Check vendor path/order.';
      console.error(msg); log(msg); throw new Error(msg);
    }
    const zip = new JSZip();
    if(dirHandle){
      const fd2 = await ensureDir(dirHandle, 'Funding Docs 2');
      const files = [];
      for await (const [name, handle] of fd2.entries()){
        if(handle.kind==='file'){
          const file = await handle.getFile();
          files.push(file);
        }
      }
      files.sort((a,b)=>{ const ak=natKey(a.name), bk=natKey(b.name);
        for(let i=0;i<Math.max(ak.length,bk.length);i++){
          if(ak[i]==null) return -1; if(bk[i]==null) return 1;
          if(ak[i]<bk[i]) return -1; if(ak[i]>bk[i]) return 1;
        } return 0; });
      for(const f of files){ const buf = await f.arrayBuffer(); zip.file(f.name, buf); }
      const blob = await zip.generateAsync({type:'blob'});
      const outName = folderPrefix()+" - Funding Docs 2.zip";
      await writeFile(dirHandle, outName, blob);
      log('Saved directly: ' + outName);
    } else {
      const v = listCandidates();
      const dir = zip.folder('Funding Docs 2');
      const usedSet = new Set(v.resUsed);
      for(const f of usedSet){ const buf = await f.arrayBuffer(); dir.file(f.name, buf); }
      const blob = await zip.generateAsync({type:'blob'});
      downloadBlob(blob, folderPrefix()+" - Funding Docs 2.zip");
      log('ZIP ready (download).');
    }
    setStatus('Funding Docs 2.zip ready.','ready');
  } catch(e){ 
    console.error(e); 
    log('ZIP error: ' + (e?.message || e)); 
    setStatus('ZIP creation failed.','error');
  }
}

async function runAll(){
  try {
    if(pickedFiles.length===0){ log('No PDFs selected.'); return; }
    if(!('showDirectoryPicker' in window)){
      const b = document.getElementById('grantFs');
      if(b){ b.disabled = true; b.title = 'Use Chrome/Edge for write mode'; }
    }
    setStatus('Running all steps…', 'running');
    log('Running all steps...');
    await buildClosing(); await sleep(800);
    await buildFunding1(); await sleep(800);
    await buildClientPkg(); await sleep(400);
    await downloadFD2Zip();
    log('All steps completed. If multiple files didn\'t appear, allow multiple downloads for this site.');
    setStatus('Success! All steps completed.','ready');
  } catch(e) { 
    console.error(e); 
    log('Run All error: ' + (e?.message || e)); 
    setStatus('Run All failed.','error');
  }
}

// ======= UI binding (defer-safe) =======
function setActionsEnabled(enabled){
  ['buildClosing','buildFunding1','buildClientPkg','zipFD2','runAll']
    .forEach(id=>{ const b=$('#'+id); if(b) b.disabled = !enabled; });
}

let _bound = false;
function bindUI() {
  if (_bound) return; _bound = true;

  const pick = document.getElementById('pick');
  const scanBtn = document.getElementById('scanBtn');
  const grantFs = document.getElementById('grantFs');
  if (!pick || !scanBtn || !grantFs) {
    console.error('Startup error: key UI elements not found');
    return;
  }

  pick.addEventListener('change', (e) => {
    indexFiles(e.target.files);
    const v = listCandidates();
    renderSlots(v);
    setActionsEnabled(pickedFiles.length > 0);
  });

  scanBtn.addEventListener('click', () => {
    if (pickedFiles.length === 0) return;
    const v = listCandidates();
    renderSlots(v);
    log('Analysis updated.');
  });

  grantFs.addEventListener('click', () => withBusy('grantFs','Granting…', grantFolderAccess));
  document.getElementById('buildClosing').addEventListener('click', () => withBusy('buildClosing','Running…', buildClosing));
  document.getElementById('buildFunding1').addEventListener('click', () => withBusy('buildFunding1','Running…', buildFunding1));
  document.getElementById('buildClientPkg').addEventListener('click', () => withBusy('buildClientPkg','Running…', buildClientPkg));
  document.getElementById('zipFD2').addEventListener('click', () => withBusy('zipFD2','Preparing…', downloadFD2Zip));
  document.getElementById('runAll').addEventListener('click', () => withBusy('runAll','Running…', runAll));

  setActionsEnabled(false);
  log('Ready. Select a folder to begin.');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', bindUI);
} else {
  bindUI();
}

// ======= Drag & drop folder support (Chrome-based, robust batching) =======
window.addEventListener('dragover', e => { e.preventDefault(); });

window.addEventListener('drop', e => {
  e.preventDefault();
  const items = e.dataTransfer.items;
  if (!items) return;

  const entries = [];
  for (const it of items) {
    const entry = it.webkitGetAsEntry?.();
    if (entry) entries.push(entry);
  }

  const readAllEntries = (dirReader) => new Promise(resolve => {
    const all = [];
    const readBatch = () => {
      dirReader.readEntries(batch => {
        if (!batch.length) return resolve(all);
        all.push(...batch); readBatch();
      }, () => resolve(all));
    };
    readBatch();
  });

  const walk = async (entry, path='') => {
    const files = [];
    if (entry.isFile) {
      await new Promise(res => entry.file(f => { f.webkitRelativePath = path + f.name; files.push(f); res(); }, () => res()));
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const children = await readAllEntries(reader);
      for (const ch of children) {
        const childFiles = await walk(ch, path + entry.name + '/');
        files.push(...childFiles);
      }
    }
    return files;
  };

  Promise.all(entries.map(ent => walk(ent))).then(groups => {
    const all = groups.flat();
    indexFiles(all);
    const v = listCandidates(); renderSlots(v);
    setActionsEnabled(pickedFiles.length > 0);
  });
});

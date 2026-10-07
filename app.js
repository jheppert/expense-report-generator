'use strict';
/* Expense Report Builder — invoice parsing, the editable form, and .xlsx generation.
   Depends on: TEMPLATE_B64 (template.js), ExcelJS (CDN), and pdf.js, which
   index.html imports in an inline module and hands over on window — see pdfjs(). */

const MONTHS = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11,
  january:0,february:1,march:2,april:3,june:5,july:6,august:7,september:8,october:9,november:10,december:11};
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

// Per-vendor constants. `item` and `purpose` mirror the detail-table text baked
// into the template, but the tool writes them per row: with a single invoice the
// item shifts up into row 43, which the template labels for Adobe.
const VENDORS = {
  adobe:  { label:'Adobe',  keyword:'adobe',  item:'Adobe Photoshop', purpose:'Photo editing for design work' },
  sketch: { label:'Sketch', keyword:'sketch', item:'Sketch',          purpose:'Interface design tool for work' },
};
const VENDOR_KEYS = ['adobe','sketch'];   // canonical order — detail rows and receipt columns follow it
const DETAIL_ROWS = [43, 44];             // the detail-table rows the tool owns

const state = { adobe:null, sketch:null, weekEnding:null, weekEditedByUser:false, lastBlob:null, lastName:null };

// ---------- pdf.js handover ----------
// index.html imports pdf.js in an inline module (see the comment there) and fires
// 'pdfjs-ready'. Module scripts are deferred, so this file may run first — hence
// the wait rather than reading window.pdfjsLib directly.
function pdfjs(){
  if(window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  return new Promise((resolve, reject) => {
    window.addEventListener('pdfjs-ready', () => resolve(window.pdfjsLib), {once:true});
    setTimeout(() => reject(new Error('pdf.js did not load')), 15000);
  });
}

// ---------- PDF parsing ----------
function parseInvoice(text){
  const t = text.replace(/\s+/g,' ');
  const low = t.toLowerCase();
  const vendor = VENDOR_KEYS.find(k => low.includes(VENDORS[k].keyword)) || null;
  let date=null, m;
  if((m=t.match(/Invoice Date\s*[:\-]?\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})/i))) date=mkDate(+m[3],MONTHS[m[2].toLowerCase()],+m[1]);
  if(!date && (m=t.match(/Date of issue\s*[:\-]?\s*([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/i))) date=mkDate(+m[3],MONTHS[m[1].toLowerCase()],+m[2]);
  if(!date && (m=t.match(/(\d{1,2})-([A-Za-z]{3})-(\d{4})/))) date=mkDate(+m[3],MONTHS[m[2].toLowerCase()],+m[1]);
  if(!date && (m=t.match(/([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/)) && MONTHS[m[1].toLowerCase()]!=null) date=mkDate(+m[3],MONTHS[m[1].toLowerCase()],+m[2]);
  let amount=null;
  if((m=t.match(/GRAND TOTAL\s*\(USD\)\s*\$?\s*([\d,]+\.\d{2})/i)) ||
     (m=t.match(/Amount due\s*\$?\s*([\d,]+\.\d{2})/i)) ||
     (m=t.match(/Total\s*\$?\s*([\d,]+\.\d{2})/i))) amount=parseFloat(m[1].replace(/,/g,''));
  return {vendor, date, amount};
}
function mkDate(y,mo,d){ return new Date(y, mo, d, 12, 0, 0); } // noon: timezone-safe

async function readPdf(file){
  const pdfjsLib = await pdfjs();
  const buf = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({data:new Uint8Array(buf)}).promise;
  let text='';
  const pages=[];
  for(let p=1;p<=doc.numPages;p++){
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    text += tc.items.map(i=>i.str).join(' ') + '\n';
    const viewport = page.getViewport({scale:2});
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width; canvas.height = viewport.height;
    await page.render({canvasContext:canvas.getContext('2d'), viewport}).promise;
    pages.push({ b64: canvas.toDataURL('image/png').split(',')[1], w:canvas.width, h:canvas.height });
  }
  return {text, pages};
}

// ---------- File handling ----------
async function handleFiles(fileList){
  hideError();
  const files = [...fileList].filter(f=>/pdf$/i.test(f.name) || f.type==='application/pdf');
  if(!files.length){ showError("Those don't look like PDFs. Drop the invoice PDFs and try again."); return; }
  try{ await pdfjs(); }
  catch{ showError('Could not load the PDF reader. Check your internet connection and reload.'); return; }
  for(const file of files){
    try{
      const {text, pages} = await readPdf(file);
      const parsed = parseInvoice(text);
      const slot = { vendor:parsed.vendor, date:parsed.date, amount:parsed.amount, pages, fileName:file.name };
      routeSlot(slot);
    }catch(e){
      console.error(e);
      showError(`Couldn't read ${file.name}. If it's a scanned image, you can still add it and type the date and amount by hand.`);
    }
  }
  refreshWeekDefault();
  render();
}

function routeSlot(slot){
  if(slot.vendor){ state[slot.vendor] = slot; return; }
  // unrecognized vendor — fill first open slot, flag for manual pick
  const open = VENDOR_KEYS.find(k => !state[k]);
  if(!open){ showError("Both slots are full. Remove one before adding another invoice."); return; }
  slot.vendor = open; slot.needsVendor = true; state[open] = slot;
}

// Present invoices in canonical order — 1 or 2 of them.
function presentSlots(){ return VENDOR_KEYS.map(k => state[k]).filter(Boolean); }
function readyToGenerate(){ const p = presentSlots(); return p.length > 0 && p.every(valid); }

function refreshWeekDefault(){
  if(state.weekEditedByUser) return;
  const dates = presentSlots().map(s=>s.date).filter(Boolean);
  if(!dates.length) return;
  const latest = new Date(Math.max(...dates.map(d=>d.getTime())));
  const wk = new Date(latest); wk.setDate(latest.getDate() + ((7 - latest.getDay()) % 7)); // Sunday ending Mon–Sun week
  wk.setHours(12,0,0,0);
  state.weekEnding = wk;
}

// ---------- Rendering ----------
const $ = id => document.getElementById(id);

function fmtDateInput(d){ if(!d) return ''; const p=n=>String(n).padStart(2,'0'); return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`; }
function fmtLong(d){ return d ? `${MONTH_NAMES[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}` : '—'; }
function money(n){ return '$'+(Number(n)||0).toFixed(2); }
function totalAmount(){ return presentSlots().reduce((sum,s)=>sum+(s.amount||0), 0); }

function slotMarkup(slot, key){
  const label = VENDORS[key].label;
  if(!slot){
    return `<div class="empty-hint">
        <span class="chip ghost">${label}</span>
        <span class="h">No ${label} invoice</span>
        <span class="s">Optional — drop one above if you have it this month</span>
      </div>`;
  }
  const needDate = !slot.date, needAmt = (slot.amount==null);
  const vendorSel = slot.needsVendor ? `
      <div class="field">
        <label>Vendor — couldn't tell, please confirm</label>
        <select class="vendor-select" data-vendor="${key}">
          ${VENDOR_KEYS.map(k=>`<option value="${k}" ${slot.vendor===k?'selected':''}>${VENDORS[k].item}</option>`).join('')}
        </select>
      </div>` : '';
  return `
    <div class="top">
      <span class="chip ${key}">${label}</span>
      <span class="fname" title="${slot.fileName}">${slot.fileName}</span>
      <button class="remove" data-remove="${key}" aria-label="Remove ${label} invoice">×</button>
    </div>
    ${vendorSel}
    <div class="field ${needDate?'needs':''}">
      <label>Purchase date</label>
      <input type="date" data-date="${key}" value="${fmtDateInput(slot.date)}" />
    </div>
    <div class="field amount ${needAmt?'needs':''}">
      <label>Amount</label>
      <span class="cur">$</span>
      <input type="text" inputmode="decimal" data-amount="${key}" value="${slot.amount!=null?slot.amount.toFixed(2):''}" placeholder="0.00" />
    </div>
    <div class="pages"><b>${slot.pages.length}</b> page${slot.pages.length>1?'s':''} → Receipts sheet</div>
  `;
}

function render(){
  VENDOR_KEYS.forEach(k=>{
    const el = $(`slot-${k}`);
    el.className = 'slot ' + (state[k]?'filled':'empty');
    el.innerHTML = slotMarkup(state[k], k);
  });

  const present = presentSlots();
  $('ledger').style.display = present.length ? 'block' : 'none';
  $('actions').style.display = present.length ? 'flex' : 'none';

  if(state.weekEnding) $('weekEnding').value = fmtDateInput(state.weekEnding);
  const friday = fridayFor(state.weekEnding);
  $('fridayText').textContent = fmtLong(friday);

  $('totalVal').textContent = money(totalAmount());

  const ready = readyToGenerate();
  $('generate').disabled = !ready;
  $('actionHint').textContent =
      !present.length ? 'Add at least one invoice to continue'
    : !ready ? 'Fill the highlighted date or amount to continue'
    : present.length === VENDOR_KEYS.length ? 'Adobe and Sketch look good.'
    : `${VENDORS[present[0].vendor].label} only — the report will have one line item.`;

  wireSlotInputs();
}

function fridayFor(weekEnding){ if(!weekEnding) return null; const f=new Date(weekEnding); f.setDate(weekEnding.getDate()-2); return f; }
function valid(slot){ return slot && slot.date instanceof Date && !isNaN(slot.date) && typeof slot.amount==='number' && slot.amount>0; }

function wireSlotInputs(){
  document.querySelectorAll('[data-date]').forEach(el=>el.onchange=e=>{
    const s=state[e.target.dataset.date]; if(!s) return;
    const [y,m,d]=e.target.value.split('-').map(Number);
    s.date = e.target.value ? mkDate(y,m-1,d) : null;
    refreshWeekDefault(); render();
  });
  document.querySelectorAll('[data-amount]').forEach(el=>el.oninput=e=>{
    const s=state[e.target.dataset.amount]; if(!s) return;
    const v=parseFloat(e.target.value.replace(/[^0-9.]/g,''));
    s.amount = isNaN(v)?null:v;
    // light re-render of total only, avoid caret jump
    $('totalVal').textContent = money(totalAmount());
    $('generate').disabled = !readyToGenerate();
  });
  document.querySelectorAll('[data-vendor]').forEach(el=>el.onchange=e=>{
    const cur=e.target.dataset.vendor, picked=e.target.value;
    const s=state[cur]; if(!s) return;
    if(picked!==cur && state[picked]){
      showError(`There's already a ${VENDORS[picked].label} invoice. Remove it first to move this one.`);
      render(); return;                       // render() resets the select to the slot it's in
    }
    s.vendor=picked; s.needsVendor=false;
    if(picked!==cur){ state[picked]=s; state[cur]=null; }   // re-route to the other slot
    render();
  });
  document.querySelectorAll('[data-remove]').forEach(el=>el.onclick=e=>{
    state[e.target.dataset.remove]=null; render();
  });
}

$('weekEnding').onchange = e=>{
  if(!e.target.value){ state.weekEditedByUser=false; refreshWeekDefault(); render(); return; }
  const [y,m,d]=e.target.value.split('-').map(Number);
  state.weekEnding=mkDate(y,m-1,d); state.weekEditedByUser=true; render();
};

// ---------- Generate ----------
function b64ToBytes(b64){ const bin=atob(b64); const len=bin.length; const bytes=new Uint8Array(len); for(let i=0;i<len;i++) bytes[i]=bin.charCodeAt(i); return bytes; }

async function generate(){
  hideError();
  try{
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(b64ToBytes(TEMPLATE_B64).buffer);
    const er = wb.getWorksheet('Expense Report');
    const rc = wb.getWorksheet('Receipts');

    const items = presentSlots();   // 1 or 2, in canonical Adobe-then-Sketch order

    // Detail table: items fill DETAIL_ROWS from the top, so a single invoice lands
    // on row 43 instead of leaving a blank line item above it. Item name and purpose
    // are written per row for the same reason — the template bakes Adobe's text into
    // row 43, which is wrong once Sketch alone occupies it. Unused rows are cleared.
    // Nothing on the sheet references these rows; F27 alone drives the totals.
    DETAIL_ROWS.forEach((row, i) => {
      const it = items[i], v = it && VENDORS[it.vendor];
      er.getCell(`A${row}`).value = it ? it.date    : null;
      er.getCell(`B${row}`).value = v  ? v.item     : null;
      er.getCell(`F${row}`).value = v  ? v.purpose  : null;
      er.getCell(`I${row}`).value = it ? it.amount  : null;
    });

    const combined = +(items.reduce((sum, it) => sum + it.amount, 0).toFixed(2));
    er.getCell('F27').value = combined;

    // weekly grid
    er.getCell('A10').value = state.weekEnding;
    er.getCell('F12').value = fridayFor(state.weekEnding);

    // receipts — first invoice in the left column, second in the right, side by side.
    // A lone receipt therefore sits on the left rather than stranded on the right.
    // Integer start columns are used on purpose: ExcelJS's fractional-column math
    // disagrees with Excel's real column widths, so a fractional offset would drift
    // and overlap. Col 7 clears the 408px-wide left image with a clean gap.
    const DISPLAY_W = 408;                 // ~4.25in, matching the template
    const START_COL = [0, 7];
    items.forEach((slot, i) => {
      let row = 1;                         // each column starts at the top
      for(const pg of slot.pages){         // multi-page invoices stack within their own column
        const H = Math.round(DISPLAY_W * pg.h / pg.w);
        const id = wb.addImage({ base64: pg.b64, extension:'png' });
        rc.addImage(id, { tl:{col:START_COL[i], row}, ext:{width:DISPLAY_W, height:H}, editAs:'oneCell' });
        row += Math.ceil(H/20) + 2;        // default row ~20px + small gap
      }
    });

    // force Excel to recalc totals on open
    wb.calcProperties = wb.calcProperties || {};
    wb.calcProperties.fullCalcOnLoad = true;

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const we = state.weekEnding;
    const name = `Consultant_Expense_Report_-_Jeff_Heppert_-_${MONTH_NAMES[we.getMonth()]}_${we.getFullYear()}.xlsx`;
    state.lastBlob = blob; state.lastName = name;
    triggerDownload(blob, name);
    showDone(name);
  }catch(e){
    console.error(e);
    showError('Something went wrong building the report. Reload the page and try again — your files never left your machine.');
  }
}

function triggerDownload(blob, name){
  const url = URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url; a.download=name; document.body.appendChild(a); a.click();
  a.remove(); setTimeout(()=>URL.revokeObjectURL(url), 1500);
}
function showDone(name){
  $('doneFile').textContent = name;
  $('done').style.display='block';
  $('done').scrollIntoView({behavior:'smooth', block:'center'});
}

// ---------- Events ----------
const drop=$('drop'), fileInput=$('file');
drop.onclick=()=>fileInput.click();
drop.onkeydown=e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); fileInput.click(); } };
fileInput.onchange=e=>{ handleFiles(e.target.files); fileInput.value=''; };
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault(); drop.classList.add('drag');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault(); if(ev==='dragleave' && drop.contains(e.relatedTarget)) return; drop.classList.remove('drag');}));
drop.addEventListener('drop',e=>{ if(e.dataTransfer?.files?.length) handleFiles(e.dataTransfer.files); });
// allow dropping anywhere on the page
window.addEventListener('dragover',e=>e.preventDefault());
window.addEventListener('drop',e=>{ e.preventDefault(); if(e.target.closest('#drop')) return; if(e.dataTransfer?.files?.length) handleFiles(e.dataTransfer.files); });

$('generate').onclick=generate;
$('downloadAgain').onclick=()=>{ if(state.lastBlob) triggerDownload(state.lastBlob, state.lastName); };
$('startOver').onclick=()=>{ VENDOR_KEYS.forEach(k=>state[k]=null); state.weekEnding=null; state.weekEditedByUser=false; state.lastBlob=null; $('done').style.display='none'; render(); window.scrollTo({top:0,behavior:'smooth'}); };

function showError(msg){ const e=$('err'); e.textContent=msg; e.classList.add('show'); }
function hideError(){ $('err').classList.remove('show'); }

if(typeof ExcelJS==='undefined'){ showError('Could not load the spreadsheet engine. Check your internet connection and reload.'); }
render();

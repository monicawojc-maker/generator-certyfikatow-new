
// Client-side certificate generator (clean new build)
// expects Excel headers: $name, $bday, $course, $hours, $score, $cert_number, $date
const pdfjsLib = window['pdfjs-dist/build/pdf'];
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@2.16.105/build/pdf.worker.min.js';
const { PDFDocument, rgb, StandardFonts } = PDFLib;

const fileTemplate = document.getElementById('fileTemplate');
const useDefault = document.getElementById('useDefault');
const fileExcel = document.getElementById('fileExcel');
const useDefaultExcel = document.getElementById('useDefaultExcel');
const pdfCanvas = document.getElementById('pdfCanvas');
const overlay = document.getElementById('overlay');
const fieldSelect = document.getElementById('fieldSelect');
const fontSizeInput = document.getElementById('fontSize');
const fontFamily = document.getElementById('fontFamily');
const fontColor = document.getElementById('fontColor');
const generateBtn = document.getElementById('generate');
const status = document.getElementById('status');
const excelPreview = document.getElementById('excelPreview');

let templateBytes = null;
let excelData = null;
let pdfPageViewport = null;
let scale = 1;
let fields = []; // {key,x,y,size,font,color}

// helper: safe get field value (check with/without $ and common names)
function getCellValue(row, key) {
  if(row == null) return '';
  if(key in row) return row[key];
  const plain = key.replace(/^\$/, '');
  if(plain in row) return row[plain];
  // common fallbacks
  if(('$'+plain) in row) return row['$'+plain];
  return '';
}

// try to convert Excel serial date to readable text if needed
function excelDateToText(val) {
  if(typeof val === 'number') {
    // Excel's epoch: 1900-01-01 with bug; use simple conversion
    const days = Math.floor(val) - 25569; // days since 1970-01-01
    const ms = days * 86400 * 1000;
    const d = new Date(ms);
    const dd = d.getUTCDate();
    const mm = d.toLocaleString('pl-PL', { month: 'long' });
    const yyyy = d.getUTCFullYear();
    return dd + ' ' + mm + ' ' + yyyy + ' r.';
  }
  return String(val);
}

function populateFieldSelect(headers) {
  fieldSelect.innerHTML = '';
  const uniq = Array.from(new Set(headers));
  uniq.forEach(h => {
    const opt = document.createElement('option');
    opt.value = h;
    opt.innerText = h;
    fieldSelect.appendChild(opt);
  });
}

useDefault.addEventListener('click', async () => {
  const resp = await fetch('template.pdf');
  templateBytes = new Uint8Array(await resp.arrayBuffer());
  await renderPDF(templateBytes);
  status.innerText = 'Użyto domyślnego szablonu.';
});

useDefaultExcel.addEventListener('click', async () => {
  const resp = await fetch('data.xlsx').catch(()=>null);
  if(!resp){ status.innerText='Brak domyślnego pliku data.xlsx'; return; }
  const arr = new Uint8Array(await resp.arrayBuffer());
  parseExcel(arr);
  status.innerText='Użyto domyślnego pliku Excel.';
});

fileTemplate.addEventListener('change', async (e)=>{
  const f = e.target.files[0]; if(!f) return;
  const arr = new Uint8Array(await f.arrayBuffer());
  templateBytes = arr;
  await renderPDF(arr);
});

fileExcel.addEventListener('change', async (e)=>{
  const f = e.target.files[0]; if(!f) return;
  const arr = new Uint8Array(await f.arrayBuffer());
  parseExcel(arr);
});

function parseExcel(arr) {
  const wb = XLSX.read(arr, {type:'array'});
  const first = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(first, {defval:''});
  // ensure keys are strings and keep $ if present
  excelData = json.map(r => {
    const obj = {};
    for(const k in r) {
      const kk = String(k).trim();
      obj[kk] = r[k];
    }
    return obj;
  });
  excelPreview.innerHTML = '<strong>Podgląd (pierwsze 10 wierszy):</strong><pre>' + JSON.stringify(excelData.slice(0,10), null, 2) + '</pre>';
  status.innerText = `Wczytano ${excelData.length} rekordów`;
  const headers = Object.keys(excelData[0] || {});
  // normalize headers: prefer with $ if present; ensure $bday exists as fallback
  const normalized = headers.map(h => h.startsWith('$') ? h : (h.includes('data') && h.toLowerCase().includes('urod') ? '$bday' : h));
  if(!normalized.includes('$bday') && headers.includes('bday')) normalized.push('$bday');
  populateFieldSelect(normalized);
}

async function renderPDF(bytes) {
  const loadingTask = pdfjsLib.getDocument({data: bytes});
  const pdf = await loadingTask.promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({scale: 1});
  const targetWidth = 800;
  scale = targetWidth / viewport.width;
  const scaledViewport = page.getViewport({scale});
  pdfPageViewport = {width: scaledViewport.width, height: scaledViewport.height, origWidth: viewport.width, origHeight: viewport.height};
  pdfCanvas.width = scaledViewport.width;
  pdfCanvas.height = scaledViewport.height;
  const ctx = pdfCanvas.getContext('2d');
  await page.render({canvasContext: ctx, viewport: scaledViewport}).promise;
  overlay.style.width = pdfCanvas.width + 'px';
  overlay.style.height = pdfCanvas.height + 'px';
  overlay.innerHTML = '';
  drawAllFields();
}

overlay.addEventListener('click', (e)=>{
  const rect = overlay.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;
  const key = fieldSelect.value || '$name';
  const size = parseInt(fontSizeInput.value || '18',10);
  const font = fontFamily.value || 'Lora';
  const color = fontColor.value || '#000000';
  fields.push({key,x,y,size,font,color});
  drawFieldMarker(fields.length-1);
});

function drawFieldMarker(i) {
  const f = fields[i];
  const el = document.createElement('div');
  el.className = 'marker';
  el.style.left = (f.x-30) + 'px';
  el.style.top = (f.y-12) + 'px';
  el.innerText = f.key;
  overlay.appendChild(el);
}
function drawAllFields(){ overlay.innerHTML=''; for(let i=0;i<fields.length;i++) drawFieldMarker(i); }

document.getElementById('clearFields').addEventListener('click', ()=>{ fields=[]; drawAllFields(); status.innerText='Pola wyczyszczone.'; });

// arrow keys move last selected marker
window.addEventListener('keydown', (e)=>{
  if(!fields.length) return;
  const last = fields[fields.length-1];
  let moved=false;
  if(e.key==='ArrowUp'){ last.y -=1; moved=true; }
  if(e.key==='ArrowDown'){ last.y +=1; moved=true; }
  if(e.key==='ArrowLeft'){ last.x -=1; moved=true; }
  if(e.key==='ArrowRight'){ last.x +=1; moved=true; }
  if(moved){ drawAllFields(); e.preventDefault(); }
});

generateBtn.addEventListener('click', async ()=>{
  if(!templateBytes){ alert('Wgraj szablon PDF'); return; }
  if(!excelData || !excelData.length){ alert('Wgraj plik Excel'); return; }
  status.innerText='Generowanie...';
  const zip = new JSZip();
  for(let i=0;i<excelData.length;i++){
    const row = excelData[i];
    const pdfDoc = await PDFDocument.load(templateBytes);
    const page = pdfDoc.getPages()[0];
    const { width, height } = page.getSize();
    // embed fonts (use standard fonts as fallback)
    let fontNormal = await pdfDoc.embedFont(StandardFonts.Helvetica);
    let fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    for(const f of fields){
      let raw = getCellValue(row, f.key);
      if(raw === '' || raw === null || raw === undefined) continue;
      // if it's date numeric, convert
      if(typeof raw === 'number') raw = excelDateToText(raw);
      const text = String(raw);
      const scaledX = f.x / scale;
      const scaledY = (pdfPageViewport.origHeight - (f.y / scale));
      const size = f.size || 18;
      const rgbColor = hexToRgb(f.color || '#000000');
      const fontToUse = (f.key === '$name') ? fontBold : fontNormal;
      page.drawText(text, { x: scaledX, y: scaledY, size, font: fontToUse, color: rgb(rgbColor.r/255, rgbColor.g/255, rgbColor.b/255) });
    }
    const filenameName = (getCellValue(row,'$name') || getCellValue(row,'name') || `uczestnik_${i+1}`);
    const pdfBytes = await pdfDoc.save();
    zip.file(`cert_${i+1}_${sanitizeFilename(filenameName)}.pdf`, pdfBytes);
    status.innerText = `Generowano ${i+1}/${excelData.length}`;
    await new Promise(res => setTimeout(res, 10)); // small yield
  }
  const content = await zip.generateAsync({type:'blob'});
  const link = document.createElement('a');
  link.href = URL.createObjectURL(content);
  link.download = 'certyfikaty.zip';
  link.click();
  status.innerText = 'Gotowe — pobieranie ZIP.';
});

function hexToRgb(hex){ const h=hex.replace('#',''); const bigint = parseInt(h,16); return {r:(bigint>>16)&255, g:(bigint>>8)&255, b:bigint&255}; }
function sanitizeFilename(s){ return String(s).replace(/[^\w\-\.\s]+/g, '_').slice(0,60); }

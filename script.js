/* ================= DATA (extracted from Finance_Tracker_-_Yashon.xlsx) ================= */

/* -- Master expense/budget categories (shared, extensible) -- */
let masterCategories = [
  {key:'food', label:'Food & Groceries'},
  {key:'transport', label:'Transportation'},
  {key:'electricity', label:'Electricity & Cable'},
  {key:'gas', label:'Gas Cylinder'},
  {key:'mobile', label:'Mobile Recharge & Internet'},
  {key:'rent', label:'House Rent'},
  {key:'labour', label:'House Labour'},
  {key:'others', label:'Others'},
  {key:'ammaMedical', label:'Amma Medical'},
  {key:'amma', label:'Amma'},
  {key:'ajithPLI', label:'Ajith - PLI'},
  {key:'ajithPPF', label:'Ajith - PPF'},
];

/* -- Asset allocation snapshots (start empty — add your own via "Add Monthly Snapshot" or Import) -- */
let assetRows = [];
assetRows.forEach((r,i)=> r._id = 'a'+i);
const assetFieldsMeta = [
  {key:'cash', label:'Cash'},
  {key:'debtFund', label:'Debt / Fixed Income'},
  {key:'mf', label:'Mutual Funds (MF)'},
  {key:'goldDig', label:'Gold (Digital)'},
  {key:'goldPhys', label:'Physical Gold'},
  {key:'stocks', label:'Stocks'},
];
function assetTotal(r){
  return (r.cash||0)+(r.debtFund||0)+(r.mf||0)+(r.goldDig||0)+(r.goldPhys||0)+(r.stocks||0);
}

/* -- Expense actuals (start empty — add your own via "Add Expence & income" or Import) -- */
let expenseRowsData = [];
expenseRowsData.forEach((r,i)=> r._id = 'e'+i);

/* -- Budget plans (start empty — add your own via "Add New Budget" or Import) -- */
let budgetRowsData = [];
budgetRowsData.forEach((r,i)=> r._id = 'b'+i);
let idCounter = 1000;
function newId(prefix){ return prefix+(idCounter++); }

function rowTotal(row){
  return masterCategories.reduce((sum,c)=> sum + (row.values[c.key]||0), 0);
}
/* find the budget that applies to a given expense month: latest budget dated on/before that month, else the earliest one */
const monthAbbrevs=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function monthIndex(label){
  const [mon,yr] = label.split(' ');
  return parseInt(yr)*12 + monthAbbrevs.indexOf(mon);
}
function budgetForMonth(monthLabel){
  if(budgetRowsData.length===0) return {month:monthLabel, values:{}};
  const targetIdx = monthIndex(monthLabel);
  let best = null;
  budgetRowsData.forEach(b=>{
    const bIdx = monthIndex(b.month);
    if(bIdx<=targetIdx && (best===null || bIdx>monthIndex(best.month))) best = b;
  });
  if(!best) best = budgetRowsData[budgetRowsData.length-1];
  return best;
}
function currentMonthLabel(){
  const d = new Date();
  return monthAbbrevs[d.getMonth()]+' '+d.getFullYear();
}
function defaultNextMonth(arr){
  return arr.length ? nextMonthLabel(arr[arr.length-1].month) : currentMonthLabel();
}

/* ================= HELPERS ================= */
function inr(n){
  if(n===null||n===undefined||isNaN(n)) return '—';
  const neg = n<0; n = Math.abs(Math.round(n));
  let s = n.toString();
  let last3 = s.substring(s.length-3);
  let rest = s.substring(0, s.length-3);
  if(rest !== '') last3 = ',' + last3;
  let formatted = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + last3;
  return (neg?'-':'') + '₹' + formatted;
}
function parseAmount(str){
  if(!str) return 0;
  return parseFloat(String(str).replace(/[₹,\s]/g,'')) || 0;
}
function toast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(window._toastTimer);
  window._toastTimer = setTimeout(()=>t.classList.remove('show'), 2600);
}
function nextMonthLabel(label){
  const [mon,yr] = label.split(' ');
  let idx = monthAbbrevs.indexOf(mon);
  let year = parseInt(yr);
  idx++; if(idx>11){idx=0; year++;}
  return monthAbbrevs[idx]+' '+year;
}
function slugKey(label, existingKeys){
  let key = label.trim().toLowerCase().replace(/[^a-z0-9]+(.)/g,(m,c)=>c.toUpperCase()).replace(/[^a-zA-Z0-9]/g,'');
  if(!key) key = 'category';
  let finalKey = key, i=2;
  while(existingKeys.includes(finalKey)){ finalKey = key+i; i++; }
  return finalKey;
}

/* ---- Month/Year split helpers (table shows month only; year is controlled via the year picker) ---- */
function getYearOf(label){
  const m = String(label).match(/(\d{4})/);
  return m ? parseInt(m[1]) : null;
}
function getMonthDisplay(label){
  return String(label).replace(/\s*\d{4}\s*/, ' ').trim().replace(/\s+/g,' ');
}
function distinctYears(rows){
  const set = new Set(rows.map(r=>getYearOf(r.month)).filter(y=>y!==null));
  return [...set].sort((a,b)=>a-b);
}

/* ---- Generic per-table state: year filter + pagination + bulk selection ---- */
const PAGE_SIZE = 12;
function makeTableState(){ return {year:null, page:1, selected:new Set()}; }
const tableStates = { asset: makeTableState(), exp: makeTableState(), bud: makeTableState(), snap: makeTableState() };

function initYear(state, rows){
  const years = distinctYears(rows);
  if(years.length===0){ state.year = null; return; }
  if(state.year===null || !years.includes(state.year)) state.year = years[years.length-1];
}
function yearRows(rows, state){
  if(state.year===null) return rows;
  return rows.filter(r=>getYearOf(r.month)===state.year);
}
function pageRows(rows, state){
  const totalPages = Math.max(1, Math.ceil(rows.length/PAGE_SIZE));
  if(state.page>totalPages) state.page = totalPages;
  if(state.page<1) state.page = 1;
  const start = (state.page-1)*PAGE_SIZE;
  return { slice: rows.slice(start, start+PAGE_SIZE), totalPages, total: rows.length, start };
}
function renderYearLabel(labelId, prevBtnId, nextBtnId, rowsGetter, state){
  const years = distinctYears(rowsGetter());
  document.getElementById(labelId).textContent = state.year!==null ? state.year : '—';
  document.getElementById(prevBtnId).disabled = years.length===0 || years.indexOf(state.year)<=0;
  document.getElementById(nextBtnId).disabled = years.length===0 || years.indexOf(state.year)>=years.length-1;
}
function renderPagination(infoId, controlsId, pg, onGo){
  document.getElementById(infoId).textContent = pg.total===0 ? 'No records' :
    `Showing ${pg.start+1}–${Math.min(pg.start+PAGE_SIZE,pg.total)} of ${pg.total}`;
  const controls = document.getElementById(controlsId);
  let html = `<button class="pg-btn" data-pg="prev" ${pg.page<=1?'disabled':''}>&larr;</button>`;
  for(let p=1;p<=pg.totalPages;p++){
    html += `<button class="pg-btn ${p===pg.page?'active':''}" data-pg="${p}">${p}</button>`;
  }
  html += `<button class="pg-btn" data-pg="next" ${pg.page>=pg.totalPages?'disabled':''}>&rarr;</button>`;
  controls.innerHTML = html;
  controls.querySelectorAll('[data-pg]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const v = btn.dataset.pg;
      if(v==='prev') onGo(pg.page-1);
      else if(v==='next') onGo(pg.page+1);
      else onGo(parseInt(v));
    });
  });
}
function updateBulkBar(barId, countId, state){
  const bar = document.getElementById(barId);
  document.getElementById(countId).textContent = state.selected.size;
  bar.classList.toggle('show', state.selected.size>0);
}
function actionIcons(){
  return `<div class="row-actions">
    <button class="act-icon edit-icon" title="Edit">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 013 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>
    </button>
    <button class="act-icon del del-icon" title="Delete">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg>
    </button>
  </div>`;
}

/* ---- JSON export / import (file download & file-picker upload) ---- */
function downloadJSON(filename, dataObj){
  const blob = new Blob([JSON.stringify(dataObj, null, 2)], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function readJSONFile(file, cb){
  if(!file){ toast('Import failed: no file selected'); return; }
  if(file.size===0){ toast('Import failed: '+file.name+' is empty (0 KB)'); return; }
  const reader = new FileReader();
  reader.onerror = ()=>{ toast('Import failed: could not read '+file.name); };
  reader.onload = ()=>{
    let text = String(reader.result);
    if(text.charCodeAt(0)===0xFEFF) text = text.slice(1); /* strip UTF-8 BOM some editors add on save */
    text = text.trim();
    try{
      const parsed = JSON.parse(text);
      cb(parsed);
    } catch(err){
      console.error('Import JSON parse error:', err, 'First 120 chars of file:', text.slice(0,120));
      toast('Import failed: '+file.name+' isn\'t valid JSON ('+err.message+')');
    }
  };
  reader.readAsText(file);
}
function fileExt(name){
  const m = String(name||'').toLowerCase().match(/\.([a-z0-9]+)$/);
  return m ? m[1] : '';
}
function downloadXLSX(filename, sheets){
  /* sheets: {SheetName: [{col:val,...}, ...]} */
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows])=>{
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0,31));
  });
  XLSX.writeFile(wb, filename);
}

/* ================= EXCEL / CSV IMPORT ENGINE ================= */
/* Understands two layouts:
   1) This app's own flat export: a "Month" column + one column per field.
   2) The original Finance Tracker workbook: "Snapshot" (category/item/amount/date rows),
      "Budget" and "Monthly Expence" (category rows x period columns). */
function normHeader(s){ return String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]/g,''); }
function isDateLike(v){ return v!=null && Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v.getTime ? v.getTime() : NaN); }

function matchFieldKey(header, metaList){
  const nh = normHeader(header);
  if(!nh) return null;
  for(const f of metaList){ if(normHeader(f.label)===nh || normHeader(f.key)===nh) return f.key; }
  for(const f of metaList){ if(nh.includes(normHeader(f.key)) || nh.includes(normHeader(f.label)) || normHeader(f.label).includes(nh)) return f.key; }
  return null;
}
function periodLabelFromHeader(v){
  if(isDateLike(v)) return monthAbbrevs[v.getMonth()]+' '+v.getFullYear();
  const s = String(v==null?'':v).trim();
  if(!s) return null;
  let m = s.match(/([A-Za-z]{3,9})[^A-Za-z0-9]{0,6}(\d{4})/);
  if(!m) m = s.match(/(\d{4})[^A-Za-z0-9]{0,6}([A-Za-z]{3,9})/);
  if(!m) return null;
  let monStr = /^\d+$/.test(m[1]) ? m[2] : m[1];
  let yr = /^\d+$/.test(m[1]) ? m[1] : m[2];
  const idx = monthAbbrevs.findIndex(a=> monStr.toLowerCase().startsWith(a.toLowerCase()));
  if(idx===-1) return null;
  return monthAbbrevs[idx]+' '+yr;
}
/* Sheet where rows = categories, columns = periods (matches "Budget" / "Monthly Expence" sheets) */
function parsePeriodColumnsSheet(sheet, metaList, autoCreateCategories){
  const raw = XLSX.utils.sheet_to_json(sheet, {header:1, raw:true, defval:null});
  if(!raw.length) return {rows:[], skippedCols:0};
  const header = raw[0];
  const colLabels = {};
  let skippedCols = 0;
  for(let c=1;c<header.length;c++){
    if(header[c]==null || header[c]==='') continue;
    const label = periodLabelFromHeader(header[c]);
    if(label) colLabels[c] = label; else skippedCols++;
  }
  const periods = {};
  Object.values(colLabels).forEach(l=> periods[l] = periods[l] || {});
  for(let r=1;r<raw.length;r++){
    const row = raw[r];
    if(!row || row[0]==null || String(row[0]).trim()==='') continue;
    const label = String(row[0]).trim();
    if(/^total$/i.test(label)) continue;
    let key = matchFieldKey(label, metaList);
    if(!key && autoCreateCategories) key = addMasterCategory(label);
    if(!key) continue;
    Object.entries(colLabels).forEach(([c,pLabel])=>{
      let v = row[c];
      if(v==null || v==='-') v = 0;
      if(typeof v!=='number') v = parseFloat(String(v).replace(/[^\d.\-]/g,'')) || 0;
      periods[pLabel][key] = v;
    });
  }
  const rows = Object.entries(periods).map(([month,values])=>({month, values}));
  return {rows, skippedCols};
}
/* Sheet with Category/Item/Amount/Date rows in blocks (matches the "Snapshot" sheet) */
function parseSnapshotBlockSheet(sheet){
  const raw = XLSX.utils.sheet_to_json(sheet, {header:1, raw:true, defval:null});
  const dataRows = raw.slice(1); /* skip header row */
  const monthly = {}; /* 'Mon YYYY' -> field values, later dates overwrite earlier within same month */
  const monthDate = {};
  let block = [];
  const flush = ()=>{
    if(!block.length) return;
    let date = null;
    for(const r of block){ if(isDateLike(r[3])){ date = r[3]; break; } }
    if(!date){ block=[]; return; }
    const label = monthAbbrevs[date.getMonth()]+' '+date.getFullYear();
    if(!monthDate[label] || date > monthDate[label]) monthDate[label] = date;
    const target = monthly[label] || (monthly[label] = {cash:0, debtFund:0, mf:0, goldDig:0, goldPhys:0, stocks:0});
    /* only take values from the block belonging to the *latest* date seen for this month */
    if(date.getTime() === monthDate[label].getTime()){
      block.forEach(r=>{
        const item = normHeader(r[1]); const amt = typeof r[2]==='number' ? r[2] : (parseFloat(r[2])||0);
        if(item.includes('cash')) target.cash = amt;
        else if(item.includes('debt')) target.debtFund = amt;
        else if(item.includes('equitymf') || (item.includes('equity')&&item.includes('mf'))) target.mf = amt;
        else if(item.includes('stock')) target.stocks = amt;
        else if(item.includes('physicalgold')) target.goldPhys = amt;
        else if(item.includes('gold')) target.goldDig = amt;
      });
    }
    block = [];
  };
  dataRows.forEach(r=>{
    const empty = !r || r.every(c=>c==null || c==='');
    if(empty){ flush(); return; }
    block.push(r);
  });
  flush();
  const rows = Object.keys(monthly).map(month=>({month, ...monthly[month]}));
  rows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
  return rows;
}
/* Flat table: a "Month" column + one column per field (this app's own export layout) */
function parseFlatSheet(sheet, metaList){
  const objRows = XLSX.utils.sheet_to_json(sheet, {defval:null});
  if(!objRows.length) return null;
  const headers = Object.keys(objRows[0]);
  const monthHeader = headers.find(h=> normHeader(h)==='month');
  if(!monthHeader) return null;
  const colMap = {};
  headers.forEach(h=>{ if(h!==monthHeader){ const k = matchFieldKey(h, metaList); if(k) colMap[h]=k; } });
  return objRows.map(r=>{
    const month = periodLabelFromHeader(r[monthHeader]) || String(r[monthHeader]||'').trim();
    const values = {};
    Object.entries(colMap).forEach(([h,k])=>{
      let v = r[h]; if(v==null || v==='-') v=0;
      if(typeof v!=='number') v = parseFloat(String(v).replace(/[^\d.\-]/g,'')) || 0;
      values[k]=v;
    });
    return {month, values};
  }).filter(r=>r.month);
}
function findSheet(wb, names){
  const target = names.map(normHeader);
  const found = wb.SheetNames.find(n=> target.includes(normHeader(n)));
  return found ? wb.Sheets[found] : null;
}
function readWorkbook(file, cb){
  if(!file){ toast('Import failed: no file selected'); return; }
  if(file.size===0){ toast('Import failed: '+file.name+' is empty (0 KB)'); return; }
  const reader = new FileReader();
  reader.onerror = ()=> toast('Import failed: could not read '+file.name);
  reader.onload = ()=>{
    try{
      const wb = XLSX.read(new Uint8Array(reader.result), {type:'array', cellDates:true});
      cb(wb);
    } catch(err){
      console.error('Workbook parse error:', err);
      toast('Import failed: '+file.name+' could not be read as Excel/CSV ('+err.message+')');
    }
  };
  reader.readAsArrayBuffer(file);
}
/* Generic router: pick the right file reader based on extension, then hand off to the
   right handler ('json' path keeps the existing JSON shape; 'sheet' path gets a workbook). */
function importAnyFile(file, {onJSON, onWorkbook}){
  const ext = fileExt(file && file.name);
  if(ext==='json') readJSONFile(file, onJSON);
  else if(ext==='xlsx' || ext==='xls' || ext==='csv') readWorkbook(file, onWorkbook);
  else toast('Import failed: unsupported file type — use .xlsx, .csv or .json');
}

/* ================= NAVIGATION ================= */
document.querySelectorAll('.navitem').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.navitem').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('page-'+btn.dataset.page).classList.add('active');
  });
});

/* ================= DASHBOARD ================= */
function renderChart(rangeSel){
  rangeSel = rangeSel || (document.querySelector('#rangeToggle button.active')?.dataset.range) || '6M';
  const svg = document.getElementById('netWorthChart');
  const labelsWrap = document.getElementById('chartLabels');
  const W = 900, H = 220, pad = 20;
  if(assetRows.length===0){
    svg.innerHTML = `<text x="450" y="115" text-anchor="middle" fill="#9497AC" font-size="14" font-family="Inter, sans-serif">No asset snapshots yet — add one to see your net worth grow</text>`;
    labelsWrap.innerHTML = '';
    document.getElementById('netWorthValue').textContent = inr(0);
    return;
  }
  const counts = {'6M':6, '1Y':12, '3Y':36, 'All':assetRows.length};
  const n = Math.min(counts[rangeSel]||6, assetRows.length);
  const sliced = assetRows.slice(assetRows.length-n);
  const series = sliced.map(r=>({label:r.month, value:assetTotal(r)/100000}));
  const vals = series.map(d=>d.value);
  const min = Math.min(...vals), max = Math.max(...vals);
  const spread = (max-min)||1;
  const stepX = series.length>1 ? (W-pad*2)/(series.length-1) : 0;
  const pts = series.map((d,i)=>{
    const x = pad + i*stepX;
    const y = H-pad - ((d.value-min)/spread)*(H-pad*2);
    return [x,y];
  });
  const linePath = pts.map((p,i)=> (i===0?'M':'L')+p[0]+','+p[1]).join(' ');
  const areaPath = linePath + ` L${pts[pts.length-1][0]},${H-pad} L${pts[0][0]},${H-pad} Z`;
  svg.innerHTML = `
    <defs><linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#6C63F7" stop-opacity="0.28"/>
      <stop offset="100%" stop-color="#6C63F7" stop-opacity="0"/>
    </linearGradient></defs>
    <path d="${areaPath}" fill="url(#areaFill)" />
    <path d="${linePath}" fill="none" stroke="#4F46E5" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
    ${pts.map((p,i)=>`<circle cx="${p[0]}" cy="${p[1]}" r="${i===pts.length-1?6:4.5}" fill="#fff" stroke="#4F46E5" stroke-width="3"/>`).join('')}
  `;
  const step = Math.ceil(series.length/8);
  labelsWrap.innerHTML = series.map((d,i)=>{
    const [mon,yr] = d.label.split(' ');
    const show = (i%step===0) || i===series.length-1;
    return `<span style="${show?'':'visibility:hidden;'}">${mon} '${yr.slice(2)}</span>`;
  }).join('');
  document.getElementById('netWorthValue').textContent = inr(assetTotal(assetRows[assetRows.length-1]));
  scheduleCloudSave();
}

document.getElementById('rangeToggle').addEventListener('click', (e)=>{
  const btn = e.target.closest('button');
  if(!btn) return;
  document.querySelectorAll('#rangeToggle button').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  renderChart(btn.dataset.range);
});

function renderSnapshots(){
  const state = tableStates.snap;
  const source = [...assetRows].reverse(); /* latest first */
  initYear(state, source);
  renderYearLabel('snapYearLabel','snapYearPrev','snapYearNext', ()=>source, state);

  const filtered = yearRows(source, state);
  const pg = pageRows(filtered, state);
  pg.page = state.page;

  const body = document.getElementById('snapshotBody');
  body.innerHTML = pg.slice.map((r)=>{
    const idxInOriginal = assetRows.indexOf(r);
    const prev = idxInOriginal>0 ? assetRows[idxInOriginal-1] : null;
    const total = assetTotal(r);
    const change = prev ? total-assetTotal(prev) : null;
    const expRow = expenseRowsData.find(e=>e.month===r.month);
    const expenses = expRow ? rowTotal(expRow) : null;
    return `
    <tr class="${r._new?'new-row':''}">
      <td>${getMonthDisplay(r.month)}</td>
      <td>${expenses!==null ? inr(expenses) : '<span class="muted">—</span>'}</td>
      <td style="font-weight:700;">${inr(total)}</td>
      <td style="text-align:right;">${change===null ? '<span class="muted">—</span>' : `<span class="pill ${change>=0?'under':'over'}">${change>=0?'↑':'↓'} ${inr(Math.abs(change))}</span>`}</td>
    </tr>`;
  }).join('') || `<tr><td colspan="4" style="text-align:center; color:var(--text-faint); padding:28px;">No snapshots yet — add one from Asset Allocation</td></tr>`;

  renderPagination('snapPgInfo','snapPgControls', pg, (p)=>{ state.page=p; renderSnapshots(); });
  scheduleCloudSave();
}
document.getElementById('snapYearPrev').addEventListener('click', ()=>{
  const source = [...assetRows].reverse();
  const years = distinctYears(source); const state = tableStates.snap;
  const i = years.indexOf(state.year); if(i>0){ state.year = years[i-1]; state.page=1; renderSnapshots(); }
});
document.getElementById('snapYearNext').addEventListener('click', ()=>{
  const source = [...assetRows].reverse();
  const years = distinctYears(source); const state = tableStates.snap;
  const i = years.indexOf(state.year); if(i<years.length-1){ state.year = years[i+1]; state.page=1; renderSnapshots(); }
});

/* ================= ASSET ALLOCATION ================= */
function renderAssets(){
  const state = tableStates.asset;
  initYear(state, assetRows);
  renderYearLabel('assetYearLabel','assetYearPrev','assetYearNext', ()=>assetRows, state);

  const filtered = yearRows(assetRows, state);
  const pg = pageRows(filtered, state);
  pg.page = state.page;

  const body = document.getElementById('assetBody');
  body.innerHTML = pg.slice.map((r)=>{
    const fullIdx = assetRows.indexOf(r);
    const prev = fullIdx>0 ? assetRows[fullIdx-1] : null;
    const total = assetTotal(r);
    const diff = prev ? total-assetTotal(prev) : null;
    const checked = state.selected.has(r._id) ? 'checked' : '';
    return `
    <tr class="${r._new?'new-row':''} ${fullIdx===assetRows.length-1 && !r._new?'highlight':''}" data-id="${r._id}">
      <td class="col-check"><input type="checkbox" class="rowchk" ${checked}></td>
      <td>${getMonthDisplay(r.month)}</td>
      <td>${inr(r.cash)}</td>
      <td>${inr(r.debtFund)}</td>
      <td>${inr(r.mf)}</td>
      <td>${inr(r.goldDig)}</td>
      <td>${inr(r.goldPhys)}</td>
      <td>${r.stocks!==null ? inr(r.stocks) : '<span class="muted">—</span>'}</td>
      <td style="font-weight:700;">${inr(total)}</td>
      <td style="text-align:right;">${diff===null ? '<span class="muted">First entry</span>' : `<span class="diff-amt ${diff>=0?'pos':'neg'}">${diff>=0?'+':'-'}${inr(Math.abs(diff))}</span>`}</td>
      <td class="col-actions">${actionIcons()}</td>
    </tr>`;
  }).join('') || `<tr><td colspan="11" style="text-align:center; color:var(--text-faint); padding:28px;">No records for this year</td></tr>`;

  document.getElementById('assetHeadChk').checked = pg.slice.length>0 && pg.slice.every(r=>state.selected.has(r._id));
  updateBulkBar('assetBulkBar','assetBulkCount', state);
  renderPagination('assetPgInfo','assetPgControls', pg, (p)=>{ state.page=p; renderAssets(); });

  body.querySelectorAll('tr[data-id]').forEach(tr=>{
    const id = tr.dataset.id;
    tr.querySelector('.rowchk').addEventListener('change', e=>{
      if(e.target.checked) state.selected.add(id); else state.selected.delete(id);
      updateBulkBar('assetBulkBar','assetBulkCount', state);
      document.getElementById('assetHeadChk').checked = pg.slice.every(r=>state.selected.has(r._id));
    });
    tr.querySelector('.edit-icon').addEventListener('click', ()=> openAssetEditModal(id));
    tr.querySelector('.del-icon').addEventListener('click', ()=> deleteAssetRow(id));
  });
  scheduleCloudSave();
}
function deleteAssetRow(id){
  const row = assetRows.find(r=>r._id===id);
  if(!row) return;
  if(!confirm(`Delete the ${row.month} asset snapshot? This can't be undone.`)) return;
  assetRows = assetRows.filter(r=>r._id!==id);
  tableStates.asset.selected.delete(id);
  renderAssets(); renderChart(); renderSnapshots();
  toast('Asset snapshot deleted');
}
document.getElementById('assetHeadChk').addEventListener('change', e=>{
  const state = tableStates.asset;
  const filtered = yearRows(assetRows, state);
  const pg = pageRows(filtered, state);
  pg.slice.forEach(r=> e.target.checked ? state.selected.add(r._id) : state.selected.delete(r._id));
  renderAssets();
});
document.getElementById('assetBulkClear').addEventListener('click', ()=>{ tableStates.asset.selected.clear(); renderAssets(); });
document.getElementById('assetBulkDelete').addEventListener('click', ()=>{
  const state = tableStates.asset;
  if(state.selected.size===0) return;
  if(!confirm(`Delete ${state.selected.size} selected snapshot(s)? This can't be undone.`)) return;
  assetRows = assetRows.filter(r=>!state.selected.has(r._id));
  state.selected.clear();
  renderAssets(); renderChart(); renderSnapshots();
  toast('Selected snapshots deleted');
});
document.getElementById('assetYearPrev').addEventListener('click', ()=>{
  const years = distinctYears(assetRows); const state = tableStates.asset;
  const i = years.indexOf(state.year); if(i>0){ state.year = years[i-1]; state.page=1; renderAssets(); }
});
document.getElementById('assetYearNext').addEventListener('click', ()=>{
  const years = distinctYears(assetRows); const state = tableStates.asset;
  const i = years.indexOf(state.year); if(i<years.length-1){ state.year = years[i+1]; state.page=1; renderAssets(); }
});

/* ================= EXPENSE & INCOME ================= */
function renderExpenseTableHead(){
  const head = document.getElementById('expenseTableHead');
  head.innerHTML = `<th class="col-check"><input type="checkbox" class="headchk" id="expHeadChk"></th><th style="text-align:left;">Month</th>` +
    masterCategories.map(c=>`<th>${c.label}</th>`).join('') +
    '<th>Budget Amount</th><th>Total Expense</th><th style="text-align:right;">Difference</th><th class="col-actions">Actions</th>';
}
function renderExpenses(){
  renderExpenseTableHead();
  const state = tableStates.exp;
  initYear(state, expenseRowsData);
  renderYearLabel('expYearLabel','expYearPrev','expYearNext', ()=>expenseRowsData, state);

  const filtered = yearRows(expenseRowsData, state);
  const pg = pageRows(filtered, state);
  pg.page = state.page;

  const body = document.getElementById('expenseBody');
  const colCount = masterCategories.length + 6;
  body.innerHTML = pg.slice.map(r=>{
    const budgetRow = budgetForMonth(r.month);
    const budget = rowTotal(budgetRow);
    const total = rowTotal(r);
    const diff = budget-total;
    const under = diff>=0;
    const checked = state.selected.has(r._id) ? 'checked' : '';
    return `
    <tr class="${r._new?'new-row':''}" data-id="${r._id}">
      <td class="col-check"><input type="checkbox" class="rowchk" ${checked}></td>
      <td>${getMonthDisplay(r.month)}</td>
      ${masterCategories.map(c=>`<td>${inr(r.values[c.key]||0)}</td>`).join('')}
      <td>${inr(budget)}</td>
      <td style="font-weight:700;">${inr(total)}</td>
      <td>
        <div class="diff-cell">
          <span class="diff-amt ${under?'pos':'neg'}">${under?'':'-'}${inr(Math.abs(diff))}</span>
          <span class="pill ${under?'under':'over'}">${under?'UNDER BUDGET':'OVER BUDGET'}</span>
        </div>
      </td>
      <td class="col-actions">${actionIcons()}</td>
    </tr>`;
  }).join('') || `<tr><td colspan="${colCount}" style="text-align:center; color:var(--text-faint); padding:28px;">No records for this year</td></tr>`;

  const headChk = document.getElementById('expHeadChk');
  if(headChk) headChk.checked = pg.slice.length>0 && pg.slice.every(r=>state.selected.has(r._id));
  updateBulkBar('expBulkBar','expBulkCount', state);
  renderPagination('expPgInfo','expPgControls', pg, (p)=>{ state.page=p; renderExpenses(); });

  body.querySelectorAll('tr[data-id]').forEach(tr=>{
    const id = tr.dataset.id;
    tr.querySelector('.rowchk').addEventListener('change', e=>{
      if(e.target.checked) state.selected.add(id); else state.selected.delete(id);
      updateBulkBar('expBulkBar','expBulkCount', state);
      document.getElementById('expHeadChk').checked = pg.slice.every(r=>state.selected.has(r._id));
    });
    tr.querySelector('.edit-icon').addEventListener('click', ()=> openExpenseEditModal(id));
    tr.querySelector('.del-icon').addEventListener('click', ()=> deleteExpenseRow(id));
  });
  if(headChk){
    headChk.addEventListener('change', e=>{
      const s = tableStates.exp;
      const flt = yearRows(expenseRowsData, s);
      const pgg = pageRows(flt, s);
      pgg.slice.forEach(r=> e.target.checked ? s.selected.add(r._id) : s.selected.delete(r._id));
      renderExpenses();
    });
  }
  scheduleCloudSave();
}
function deleteExpenseRow(id){
  const row = expenseRowsData.find(r=>r._id===id);
  if(!row) return;
  if(!confirm(`Delete the ${row.month} expense & income record? This can't be undone.`)) return;
  expenseRowsData = expenseRowsData.filter(r=>r._id!==id);
  tableStates.exp.selected.delete(id);
  renderExpenses(); renderSnapshots();
  toast('Expense & income record deleted');
}
document.getElementById('expBulkClear').addEventListener('click', ()=>{ tableStates.exp.selected.clear(); renderExpenses(); });
document.getElementById('expBulkDelete').addEventListener('click', ()=>{
  const state = tableStates.exp;
  if(state.selected.size===0) return;
  if(!confirm(`Delete ${state.selected.size} selected record(s)? This can't be undone.`)) return;
  expenseRowsData = expenseRowsData.filter(r=>!state.selected.has(r._id));
  state.selected.clear();
  renderExpenses(); renderSnapshots();
  toast('Selected records deleted');
});
document.getElementById('expYearPrev').addEventListener('click', ()=>{
  const years = distinctYears(expenseRowsData); const state = tableStates.exp;
  const i = years.indexOf(state.year); if(i>0){ state.year = years[i-1]; state.page=1; renderExpenses(); }
});
document.getElementById('expYearNext').addEventListener('click', ()=>{
  const years = distinctYears(expenseRowsData); const state = tableStates.exp;
  const i = years.indexOf(state.year); if(i<years.length-1){ state.year = years[i+1]; state.page=1; renderExpenses(); }
});

/* ================= BUDGET ================= */
function renderBudgetTableHead(){
  const head = document.getElementById('budgetTableHead');
  head.innerHTML = `<th class="col-check"><input type="checkbox" class="headchk" id="budHeadChk"></th><th style="text-align:left;">Month</th>` +
    masterCategories.map(c=>`<th>${c.label}</th>`).join('') + '<th>Total</th><th class="col-actions">Actions</th>';
}
function renderBudget(){
  renderBudgetTableHead();
  const state = tableStates.bud;
  initYear(state, budgetRowsData);
  renderYearLabel('budYearLabel','budYearPrev','budYearNext', ()=>budgetRowsData, state);

  const filtered = yearRows(budgetRowsData, state);
  const pg = pageRows(filtered, state);
  pg.page = state.page;

  const body = document.getElementById('budgetBody');
  const colCount = masterCategories.length + 4;
  body.innerHTML = pg.slice.map(r=>{
    const checked = state.selected.has(r._id) ? 'checked' : '';
    return `
    <tr class="${r._new?'new-row':''}" data-id="${r._id}">
      <td class="col-check"><input type="checkbox" class="rowchk" ${checked}></td>
      <td>${getMonthDisplay(r.month)}</td>
      ${masterCategories.map(c=>`<td>${inr(r.values[c.key]||0)}</td>`).join('')}
      <td style="font-weight:700;">${inr(rowTotal(r))}</td>
      <td class="col-actions">${actionIcons()}</td>
    </tr>`;
  }).join('') || `<tr><td colspan="${colCount}" style="text-align:center; color:var(--text-faint); padding:28px;">No records for this year</td></tr>`;

  const headChk = document.getElementById('budHeadChk');
  if(headChk){
    headChk.checked = pg.slice.length>0 && pg.slice.every(r=>state.selected.has(r._id));
    headChk.addEventListener('change', e=>{
      const s = tableStates.bud;
      const flt = yearRows(budgetRowsData, s);
      const pgg = pageRows(flt, s);
      pgg.slice.forEach(r=> e.target.checked ? s.selected.add(r._id) : s.selected.delete(r._id));
      renderBudget();
    });
  }
  updateBulkBar('budBulkBar','budBulkCount', state);
  renderPagination('budPgInfo','budPgControls', pg, (p)=>{ state.page=p; renderBudget(); });

  body.querySelectorAll('tr[data-id]').forEach(tr=>{
    const id = tr.dataset.id;
    tr.querySelector('.rowchk').addEventListener('change', e=>{
      if(e.target.checked) state.selected.add(id); else state.selected.delete(id);
      updateBulkBar('budBulkBar','budBulkCount', state);
      document.getElementById('budHeadChk').checked = pg.slice.every(r=>state.selected.has(r._id));
    });
    tr.querySelector('.edit-icon').addEventListener('click', ()=> openBudgetEditModal(id));
    tr.querySelector('.del-icon').addEventListener('click', ()=> deleteBudgetRow(id));
  });
  scheduleCloudSave();
}
function deleteBudgetRow(id){
  const row = budgetRowsData.find(r=>r._id===id);
  if(!row) return;
  if(!confirm(`Delete the ${row.month} budget? This can't be undone.`)) return;
  budgetRowsData = budgetRowsData.filter(r=>r._id!==id);
  tableStates.bud.selected.delete(id);
  renderBudget(); renderExpenses();
  toast('Budget deleted');
}
document.getElementById('budBulkClear').addEventListener('click', ()=>{ tableStates.bud.selected.clear(); renderBudget(); });
document.getElementById('budBulkDelete').addEventListener('click', ()=>{
  const state = tableStates.bud;
  if(state.selected.size===0) return;
  if(!confirm(`Delete ${state.selected.size} selected budget(s)? This can't be undone.`)) return;
  budgetRowsData = budgetRowsData.filter(r=>!state.selected.has(r._id));
  state.selected.clear();
  renderBudget(); renderExpenses();
  toast('Selected budgets deleted');
});
document.getElementById('budYearPrev').addEventListener('click', ()=>{
  const years = distinctYears(budgetRowsData); const state = tableStates.bud;
  const i = years.indexOf(state.year); if(i>0){ state.year = years[i-1]; state.page=1; renderBudget(); }
});
document.getElementById('budYearNext').addEventListener('click', ()=>{
  const years = distinctYears(budgetRowsData); const state = tableStates.bud;
  const i = years.indexOf(state.year); if(i<years.length-1){ state.year = years[i+1]; state.page=1; renderBudget(); }
});

/* ================= SYNC: new category added anywhere ================= */
function addMasterCategory(label){
  const existingKeys = masterCategories.map(c=>c.key);
  const key = slugKey(label, existingKeys);
  masterCategories.push({key, label});
  expenseRowsData.forEach(r=>{ if(!(key in r.values)) r.values[key]=0; });
  budgetRowsData.forEach(r=>{ if(!(key in r.values)) r.values[key]=0; });
  return key;
}

/* ================= MODAL: OPEN/CLOSE ================= */
function openModal(id){ document.getElementById(id).classList.add('open'); }
function closeModal(id){ document.getElementById(id).classList.remove('open'); }
document.querySelectorAll('[data-close]').forEach(b=> b.addEventListener('click', ()=>closeModal(b.dataset.close)));
document.querySelectorAll('.overlay').forEach(ov=> ov.addEventListener('click', (e)=>{ if(e.target===ov) ov.classList.remove('open'); }));
document.addEventListener('keydown', (e)=>{ if(e.key==='Escape') document.querySelectorAll('.overlay.open').forEach(o=>o.classList.remove('open')); });

/* ---- Add / Edit Asset Allocation modal (clean single-value fields per asset type) ---- */
let editingAssetId = null;
function renderAssetFields(values){
  const wrap = document.getElementById('assetFieldsWrap');
  wrap.innerHTML = assetFieldsMeta.map(f=>{
    const val = values[f.key] || 0;
    return `
    <div class="field-row asset-field-row" data-key="${f.key}">
      <div class="field"><label>Field</label><div class="locked-label">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="10" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>
        ${f.label}</div></div>
      <div class="field"><label>Amount</label><div class="amount-field"><span class="rupee">₹</span><input type="text" class="af-add" value="${val?val:''}" placeholder="0"></div></div>
      <button class="row-del" disabled></button>
    </div>`;
  }).join('');
}
document.getElementById('openAssetModal').addEventListener('click', ()=>{
  editingAssetId = null;
  document.getElementById('assetModalTitle').textContent = 'Add Asset Allocation';
  const month = defaultNextMonth(assetRows);
  document.getElementById('af_month').value = month;
  document.getElementById('assetModalMonthLabel').textContent = month;
  renderAssetFields({});
  openModal('assetOverlay');
});
function openAssetEditModal(id){
  const row = assetRows.find(r=>r._id===id);
  if(!row) return;
  editingAssetId = id;
  document.getElementById('assetModalTitle').textContent = 'Edit Asset Allocation';
  document.getElementById('af_month').value = row.month;
  document.getElementById('assetModalMonthLabel').textContent = row.month;
  renderAssetFields(row);
  openModal('assetOverlay');
}
document.getElementById('assetSaveBtn').addEventListener('click', ()=>{
  const month = document.getElementById('af_month').value.trim() || defaultNextMonth(assetRows);
  const values = {};
  document.querySelectorAll('#assetFieldsWrap .asset-field-row').forEach(row=>{
    values[row.dataset.key] = parseAmount(row.querySelector('.af-add').value);
  });
  if(editingAssetId){
    const row = assetRows.find(r=>r._id===editingAssetId);
    Object.assign(row, values, {month});
    assetRows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    toast('Asset allocation for '+month+' updated');
  } else {
    assetRows.forEach(r=>r._new=false);
    const newRow = Object.assign({month, _new:true, _id:newId('a')}, values);
    assetRows.push(newRow);
    assetRows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    toast('Asset allocation for '+month+' saved');
  }
  renderAssets();
  renderChart();
  renderSnapshots();
  closeModal('assetOverlay');
});

/* ---- Add Income & Expense modal ---- */
let incomeItems = [];
let modalExpenseValues = {}; /* key -> amount, for the month being added */
let modalNewCustomRows = []; /* {label, amount} not-yet-in-master rows added this session */

function renderIncomeRows(){
  const wrap = document.getElementById('incomeRows');
  wrap.innerHTML = incomeItems.map((it,i)=>`
    <div class="field-row">
      <div class="field"><label>Source</label><input type="text" data-inc-src="${i}" value="${it.source}"></div>
      <div class="field"><label>Amount</label><div class="amount-field"><span class="rupee">₹</span><input type="text" data-inc-amt="${i}" value="${it.amount.toLocaleString('en-IN')}"></div></div>
      <button class="row-del" data-inc-del="${i}"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg></button>
    </div>
  `).join('');
  wrap.querySelectorAll('[data-inc-src]').forEach(inp=>inp.addEventListener('input', e=>{ incomeItems[e.target.dataset.incSrc].source = e.target.value; }));
  wrap.querySelectorAll('[data-inc-amt]').forEach(inp=>inp.addEventListener('input', e=>{ incomeItems[e.target.dataset.incAmt].amount = parseAmount(e.target.value); updateIncomeExpenseTotals(); }));
  wrap.querySelectorAll('[data-inc-del]').forEach(btn=>btn.addEventListener('click', e=>{
    incomeItems.splice(e.currentTarget.dataset.incDel,1); renderIncomeRows(); updateIncomeExpenseTotals();
  }));
}
function renderExpenseRowsModal(){
  const wrap = document.getElementById('expenseRows');
  let html = '';
  /* existing master categories: label locked, amount editable */
  masterCategories.forEach(c=>{
    html += `
    <div class="field-row">
      <div class="field"><label>Category</label><div class="locked-label">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="10" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>
        ${c.label}</div></div>
      <div class="field"><label>Amount</label><div class="amount-field"><span class="rupee">₹</span><input type="text" data-mc-amt="${c.key}" value="${modalExpenseValues[c.key]||0}"></div></div>
      <button class="row-del" disabled></button>
    </div>`;
  });
  /* newly-added custom rows: fully editable */
  modalNewCustomRows.forEach((it,i)=>{
    html += `
    <div class="field-row">
      <div class="field"><label>New Category</label><input type="text" data-new-lbl="${i}" value="${it.label}" placeholder="Category name"></div>
      <div class="field"><label>Amount</label><div class="amount-field"><span class="rupee">₹</span><input type="text" data-new-amt="${i}" value="${it.amount}"></div></div>
      <button class="row-del" data-new-del="${i}"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M18 6L6 18M6 6l12 12"/></svg></button>
    </div>`;
  });
  wrap.innerHTML = html;
  wrap.querySelectorAll('[data-mc-amt]').forEach(inp=>inp.addEventListener('input', e=>{
    modalExpenseValues[e.target.dataset.mcAmt] = parseAmount(e.target.value); updateIncomeExpenseTotals();
  }));
  wrap.querySelectorAll('[data-new-lbl]').forEach(inp=>inp.addEventListener('input', e=>{ modalNewCustomRows[e.target.dataset.newLbl].label = e.target.value; }));
  wrap.querySelectorAll('[data-new-amt]').forEach(inp=>inp.addEventListener('input', e=>{ modalNewCustomRows[e.target.dataset.newAmt].amount = parseAmount(e.target.value); updateIncomeExpenseTotals(); }));
  wrap.querySelectorAll('[data-new-del]').forEach(btn=>btn.addEventListener('click', e=>{
    modalNewCustomRows.splice(e.currentTarget.dataset.newDel,1); renderExpenseRowsModal(); updateIncomeExpenseTotals();
  }));
}
function updateIncomeExpenseTotals(){
  const incTotal = incomeItems.reduce((a,b)=>a+b.amount,0);
  let expTotal = 0;
  masterCategories.forEach(c=> expTotal += (modalExpenseValues[c.key]||0));
  modalNewCustomRows.forEach(it=> expTotal += (it.amount||0));
  document.getElementById('incomeTotal').textContent = inr(incTotal);
  document.getElementById('expenseTotal').textContent = inr(expTotal);
}
document.getElementById('addIncomeRow').addEventListener('click', ()=>{
  incomeItems.push({source:'', amount:0}); renderIncomeRows(); updateIncomeExpenseTotals();
});
document.getElementById('addExpenseRow').addEventListener('click', ()=>{
  modalNewCustomRows.push({label:'', amount:0}); renderExpenseRowsModal(); updateIncomeExpenseTotals();
});
let editingExpenseId = null;
document.getElementById('openExpenseModal').addEventListener('click', ()=>{
  editingExpenseId = null;
  document.getElementById('expenseModalTitle').textContent = 'Add Income & expence';
  const nextMonth = defaultNextMonth(expenseRowsData);
  document.getElementById('exp_month').value = nextMonth;
  document.getElementById('expModalMonth').textContent = nextMonth;
  modalExpenseValues = {};
  masterCategories.forEach(c=> modalExpenseValues[c.key]=0);
  modalNewCustomRows = [];
  renderIncomeRows(); renderExpenseRowsModal(); updateIncomeExpenseTotals();
  openModal('expenseOverlay');
});
function openExpenseEditModal(id){
  const row = expenseRowsData.find(r=>r._id===id);
  if(!row) return;
  editingExpenseId = id;
  document.getElementById('expenseModalTitle').textContent = 'Edit Income & expence';
  document.getElementById('exp_month').value = row.month;
  document.getElementById('expModalMonth').textContent = row.month;
  modalExpenseValues = {};
  masterCategories.forEach(c=> modalExpenseValues[c.key] = row.values[c.key]||0);
  modalNewCustomRows = [];
  renderIncomeRows(); renderExpenseRowsModal(); updateIncomeExpenseTotals();
  openModal('expenseOverlay');
}
document.getElementById('exp_month').addEventListener('input', e=>{
  document.getElementById('expModalMonth').textContent = e.target.value || 'New month';
});
document.getElementById('expenseSaveBtn').addEventListener('click', ()=>{
  const monthKey = document.getElementById('exp_month').value.trim() || defaultNextMonth(expenseRowsData);

  /* fold any new custom categories into the shared master list */
  modalNewCustomRows.forEach(it=>{
    if(!it.label.trim()) return;
    const key = addMasterCategory(it.label.trim());
    modalExpenseValues[key] = it.amount||0;
  });

  const values = {};
  masterCategories.forEach(c=> values[c.key] = modalExpenseValues[c.key]||0);

  if(editingExpenseId){
    const row = expenseRowsData.find(r=>r._id===editingExpenseId);
    row.values = values; row.month = monthKey;
    expenseRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    toast('Income & expense for '+monthKey+' updated');
  } else {
    expenseRowsData.forEach(r=>r._new=false);
    const existingIdx = expenseRowsData.findIndex(r=>r.month===monthKey);
    const newRow = {month:monthKey, values, _new:true, _id:newId('e')};
    if(existingIdx>-1) expenseRowsData[existingIdx]=newRow; else { expenseRowsData.push(newRow); expenseRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month)); }
    toast('Income & expense for '+monthKey+' saved');
  }

  renderExpenses();
  renderBudget();
  renderSnapshots();
  closeModal('expenseOverlay');
});

/* ---- Add Budget modal ---- */
let modalBudgetValues = {};
let modalBudgetNewRows = [];
function renderBudgetRowsModal(){
  const wrap = document.getElementById('budgetRows');
  let html = '';
  masterCategories.forEach(c=>{
    html += `
    <div class="field-row">
      <div class="field"><label>Category</label><div class="locked-label">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="10" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>
        ${c.label}</div></div>
      <div class="field"><label>Amount</label><div class="amount-field"><span class="rupee">₹</span><input type="text" data-bmc-amt="${c.key}" value="${modalBudgetValues[c.key]||0}"></div></div>
      <button class="row-del" disabled></button>
    </div>`;
  });
  modalBudgetNewRows.forEach((it,i)=>{
    html += `
    <div class="field-row">
      <div class="field"><label>New Category</label><input type="text" data-bnew-lbl="${i}" value="${it.label}" placeholder="Category name"></div>
      <div class="field"><label>Amount</label><div class="amount-field"><span class="rupee">₹</span><input type="text" data-bnew-amt="${i}" value="${it.amount}"></div></div>
      <button class="row-del" data-bnew-del="${i}"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M18 6L6 18M6 6l12 12"/></svg></button>
    </div>`;
  });
  wrap.innerHTML = html;
  wrap.querySelectorAll('[data-bmc-amt]').forEach(inp=>inp.addEventListener('input', e=>{ modalBudgetValues[e.target.dataset.bmcAmt] = parseAmount(e.target.value); }));
  wrap.querySelectorAll('[data-bnew-lbl]').forEach(inp=>inp.addEventListener('input', e=>{ modalBudgetNewRows[e.target.dataset.bnewLbl].label = e.target.value; }));
  wrap.querySelectorAll('[data-bnew-amt]').forEach(inp=>inp.addEventListener('input', e=>{ modalBudgetNewRows[e.target.dataset.bnewAmt].amount = parseAmount(e.target.value); }));
  wrap.querySelectorAll('[data-bnew-del]').forEach(btn=>btn.addEventListener('click', e=>{
    modalBudgetNewRows.splice(e.currentTarget.dataset.bnewDel,1); renderBudgetRowsModal();
  }));
}
document.getElementById('addBudgetRow').addEventListener('click', ()=>{
  modalBudgetNewRows.push({label:'', amount:0}); renderBudgetRowsModal();
});
let editingBudgetId = null;
document.getElementById('openBudgetModal').addEventListener('click', ()=>{
  editingBudgetId = null;
  document.getElementById('budgetModalTitle').textContent = 'Add New Budget';
  document.getElementById('bf_month').value = defaultNextMonth(budgetRowsData);
  modalBudgetValues = {};
  masterCategories.forEach(c=> modalBudgetValues[c.key]=0);
  modalBudgetNewRows = [];
  renderBudgetRowsModal();
  openModal('budgetOverlay');
});
function openBudgetEditModal(id){
  const row = budgetRowsData.find(r=>r._id===id);
  if(!row) return;
  editingBudgetId = id;
  document.getElementById('budgetModalTitle').textContent = 'Edit Budget';
  document.getElementById('bf_month').value = row.month;
  modalBudgetValues = {};
  masterCategories.forEach(c=> modalBudgetValues[c.key] = row.values[c.key]||0);
  modalBudgetNewRows = [];
  renderBudgetRowsModal();
  openModal('budgetOverlay');
}
document.getElementById('budgetSaveBtn').addEventListener('click', ()=>{
  const month = document.getElementById('bf_month').value.trim() || defaultNextMonth(budgetRowsData);

  /* new categories added here sync into Expense & Income automatically */
  modalBudgetNewRows.forEach(it=>{
    if(!it.label.trim()) return;
    const key = addMasterCategory(it.label.trim());
    modalBudgetValues[key] = it.amount||0;
  });

  const values = {};
  masterCategories.forEach(c=> values[c.key] = modalBudgetValues[c.key]||0);

  if(editingBudgetId){
    const row = budgetRowsData.find(r=>r._id===editingBudgetId);
    row.values = values; row.month = month;
    budgetRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    toast('Budget for '+month+' updated');
  } else {
    budgetRowsData.forEach(r=>r._new=false);
    const existingIdx = budgetRowsData.findIndex(r=>r.month===month);
    const newRow = {month, values, _new:true, _id:newId('b')};
    if(existingIdx>-1) budgetRowsData[existingIdx]=newRow; else { budgetRowsData.push(newRow); budgetRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month)); }
    toast('Budget for '+month+' saved — new categories are now available in Expense & Income too');
  }

  renderBudget();
  renderExpenses();
  closeModal('budgetOverlay');
});

/* ---- Global search: filters the currently visible table ---- */
document.getElementById('globalSearch').addEventListener('input', (e)=>{
  const q = e.target.value.trim().toLowerCase();
  const activePage = document.querySelector('.page.active');
  if(!activePage) return;
  const rows = activePage.querySelectorAll('tbody tr');
  rows.forEach(row=>{ row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none'; });
});

/* ================= EXPORT / IMPORT ================= */
/* -- Asset Allocation (table-scoped) -- */
document.getElementById('assetExportBtn').addEventListener('click', ()=>{
  downloadJSON('asset-allocation.json', {type:'asset-allocation', rows: assetRows.map(({_id,_new,...rest})=>rest)});
  toast('Asset Allocation data exported');
});
document.getElementById('assetImportBtn').addEventListener('click', ()=> document.getElementById('assetImportInput').click());
document.getElementById('assetImportInput').addEventListener('change', e=>{
  const file = e.target.files[0]; if(!file) return;
  importAnyFile(file, {
    onJSON: data=>{
      if(data && data.type && data.type!=='asset-allocation'){
        toast('Import failed: that file is a "'+data.type+'" export, not Asset Allocation'); return;
      }
      const rows = Array.isArray(data) ? data : data.rows;
      if(!Array.isArray(rows)){ toast('Import failed: unexpected file format'); return; }
      assetRows = rows.map((r)=> Object.assign({_id:newId('a')}, r));
      assetRows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
      tableStates.asset = makeTableState(); tableStates.snap = makeTableState();
      renderAssets(); renderChart(); renderSnapshots();
      toast('Asset Allocation data imported');
    },
    onWorkbook: wb=>{
      const snapSheet = findSheet(wb, ['Snapshot','Snapshots']);
      let rows;
      if(snapSheet){
        rows = parseSnapshotBlockSheet(snapSheet);
      } else {
        const flat = parseFlatSheet(wb.Sheets[wb.SheetNames[0]], assetFieldsMeta);
        rows = flat ? flat.map(r=> Object.assign({month:r.month}, r.values)) : null;
      }
      if(!rows || !rows.length){ toast('Import failed: no "Snapshot" sheet and no "Month" column found in '+file.name); return; }
      assetRows = rows.map(r=> Object.assign({_id:newId('a')}, r));
      assetRows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
      tableStates.asset = makeTableState(); tableStates.snap = makeTableState();
      renderAssets(); renderChart(); renderSnapshots();
      toast('Asset Allocation imported from '+file.name+' — '+rows.length+' month(s)');
    }
  });
  e.target.value = '';
});

/* -- Expense & Income (table-scoped) -- */
document.getElementById('expExportBtn').addEventListener('click', ()=>{
  downloadJSON('expense-income.json', {type:'expense-income', rows: expenseRowsData.map(({_id,_new,...rest})=>rest)});
  toast('Expense & Income data exported');
});
document.getElementById('expImportBtn').addEventListener('click', ()=> document.getElementById('expImportInput').click());
document.getElementById('expImportInput').addEventListener('change', e=>{
  const file = e.target.files[0]; if(!file) return;
  importAnyFile(file, {
    onJSON: data=>{
      if(data && data.type && data.type!=='expense-income'){
        toast('Import failed: that file is a "'+data.type+'" export, not Expense & Income'); return;
      }
      const rows = Array.isArray(data) ? data : data.rows;
      if(!Array.isArray(rows)){ toast('Import failed: unexpected file format'); return; }
      expenseRowsData = rows.map(r=> Object.assign({_id:newId('e')}, r));
      expenseRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
      tableStates.exp = makeTableState();
      renderExpenses(); renderSnapshots();
      toast('Expense & Income data imported');
    },
    onWorkbook: wb=>{
      const expSheet = findSheet(wb, ['Monthly Expence','Monthly Expense','Expenses','Expense']);
      let rows, skippedCols=0;
      if(expSheet){
        const parsed = parsePeriodColumnsSheet(expSheet, masterCategories, true);
        rows = parsed.rows; skippedCols = parsed.skippedCols;
      } else {
        rows = parseFlatSheet(wb.Sheets[wb.SheetNames[0]], masterCategories);
      }
      if(!rows || !rows.length){ toast('Import failed: no "Monthly Expence" sheet and no "Month" column found in '+file.name); return; }
      expenseRowsData = rows.map(r=> Object.assign({_id:newId('e')}, r));
      expenseRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
      tableStates.exp = makeTableState();
      renderExpenses(); renderSnapshots();
      toast('Expense & Income imported from '+file.name+' — '+rows.length+' month(s)'+(skippedCols?'; '+skippedCols+' column(s) skipped (add a year to those headers, e.g. "Apr 2025")':''));
    }
  });
  e.target.value = '';
});

/* -- Budget (table-scoped) -- */
document.getElementById('budExportBtn').addEventListener('click', ()=>{
  downloadJSON('budget.json', {type:'budget', rows: budgetRowsData.map(({_id,_new,...rest})=>rest)});
  toast('Budget data exported');
});
document.getElementById('budImportBtn').addEventListener('click', ()=> document.getElementById('budImportInput').click());
document.getElementById('budImportInput').addEventListener('change', e=>{
  const file = e.target.files[0]; if(!file) return;
  importAnyFile(file, {
    onJSON: data=>{
      if(data && data.type && data.type!=='budget'){
        toast('Import failed: that file is a "'+data.type+'" export, not Budget'); return;
      }
      const rows = Array.isArray(data) ? data : data.rows;
      if(!Array.isArray(rows)){ toast('Import failed: unexpected file format'); return; }
      budgetRowsData = rows.map(r=> Object.assign({_id:newId('b')}, r));
      budgetRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
      tableStates.bud = makeTableState();
      renderBudget(); renderExpenses();
      toast('Budget data imported');
    },
    onWorkbook: wb=>{
      const budSheet = findSheet(wb, ['Budget','Budgets']);
      let rows, skippedCols=0;
      if(budSheet){
        const parsed = parsePeriodColumnsSheet(budSheet, masterCategories, true);
        rows = parsed.rows; skippedCols = parsed.skippedCols;
      } else {
        rows = parseFlatSheet(wb.Sheets[wb.SheetNames[0]], masterCategories);
      }
      if(!rows || !rows.length){ toast('Import failed: no "Budget" sheet and no "Month" column found in '+file.name); return; }
      budgetRowsData = rows.map(r=> Object.assign({_id:newId('b')}, r));
      budgetRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
      tableStates.bud = makeTableState();
      renderBudget(); renderExpenses();
      toast('Budget imported from '+file.name+' — '+rows.length+' period(s)'+(skippedCols?'; '+skippedCols+' column(s) skipped (add a year to those headers)':''));
    }
  });
  e.target.value = '';
});

/* -- Profile-level: full portal export / import (all 3 modules + categories) -- */
function buildPortalExportObject(){
  return {
    type:'yash-finance-full-export',
    exportedAt: new Date().toISOString(),
    masterCategories,
    incomeItems,
    assetRows: assetRows.map(({_id,_new,...rest})=>rest),
    expenseRowsData: expenseRowsData.map(({_id,_new,...rest})=>rest),
    budgetRowsData: budgetRowsData.map(({_id,_new,...rest})=>rest),
  };
}
function applyPortalJSONImport(data){
  if(!data || typeof data!=='object') return false;
  if(Array.isArray(data.masterCategories)) masterCategories = data.masterCategories;
  if(Array.isArray(data.incomeItems)) incomeItems = data.incomeItems;
  if(Array.isArray(data.assetRows)){
    assetRows = data.assetRows.map(r=> Object.assign({_id:newId('a')}, r));
    assetRows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    tableStates.asset = makeTableState();
    tableStates.snap = makeTableState();
  }
  if(Array.isArray(data.expenseRowsData)){
    expenseRowsData = data.expenseRowsData.map(r=> Object.assign({_id:newId('e')}, r));
    expenseRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    tableStates.exp = makeTableState();
  }
  if(Array.isArray(data.budgetRowsData)){
    budgetRowsData = data.budgetRowsData.map(r=> Object.assign({_id:newId('b')}, r));
    budgetRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
    tableStates.bud = makeTableState();
  }
  renderChart(); renderAssets(); renderSnapshots(); renderExpenses(); renderBudget();
  return true;
}
document.getElementById('profileExportBtn').addEventListener('click', ()=>{
  downloadJSON('yash-finance-portal-data.json', buildPortalExportObject());
  toast('Full portal data exported');
});
document.getElementById('profileImportBtn').addEventListener('click', ()=> document.getElementById('profileImportInput').click());
document.getElementById('profileImportInput').addEventListener('change', e=>{
  const file = e.target.files[0]; if(!file) return;
  importAnyFile(file, {
    onJSON: data=>{
      if(!applyPortalJSONImport(data)){ toast('Import failed: unexpected file format'); return; }
      toast('Full portal data imported');
    },
    onWorkbook: wb=>{
      const summary = [];
      const snapSheet = findSheet(wb, ['Snapshot','Snapshots']);
      if(snapSheet){
        const rows = parseSnapshotBlockSheet(snapSheet);
        if(rows.length){
          assetRows = rows.map(r=> Object.assign({_id:newId('a')}, r));
          assetRows.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
          tableStates.asset = makeTableState(); tableStates.snap = makeTableState();
          summary.push(rows.length+' asset snapshot(s)');
        }
      }
      const budSheet = findSheet(wb, ['Budget','Budgets']);
      if(budSheet){
        const {rows} = parsePeriodColumnsSheet(budSheet, masterCategories, true);
        if(rows.length){
          budgetRowsData = rows.map(r=> Object.assign({_id:newId('b')}, r));
          budgetRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
          tableStates.bud = makeTableState();
          summary.push(rows.length+' budget period(s)');
        }
      }
      const expSheet = findSheet(wb, ['Monthly Expence','Monthly Expense','Expenses','Expense']);
      if(expSheet){
        const {rows} = parsePeriodColumnsSheet(expSheet, masterCategories, true);
        if(rows.length){
          expenseRowsData = rows.map(r=> Object.assign({_id:newId('e')}, r));
          expenseRowsData.sort((a,b)=>monthIndex(a.month)-monthIndex(b.month));
          tableStates.exp = makeTableState();
          summary.push(rows.length+' expense month(s)');
        }
      }
      if(!summary.length){
        toast('Import failed: found no "Snapshot", "Budget" or "Monthly Expence" sheet in '+file.name);
        return;
      }
      renderChart(); renderAssets(); renderSnapshots(); renderExpenses(); renderBudget();
      toast('Imported from '+file.name+': '+summary.join(', '));
    }
  });
  e.target.value = '';
});

/* ================= FIREBASE SIGN-IN + FIRESTORE SAVE ================= */
/* To enable: paste your Firebase project's config below (see the "How to set up" (ⓘ) button
   in the header for step-by-step instructions). Must be served over http(s) — Firebase does
   not allow sign-in from a file opened directly off disk (file://). */
const firebaseConfig = {
  apiKey: "AIzaSyBVUkHjVaYecnVANUK543ukbwxCZyJLPX4",
  authDomain: "financial-tracker-70cec.firebaseapp.com",
  projectId: "financial-tracker-70cec",
  storageBucket: "financial-tracker-70cec.firebasestorage.app",
  messagingSenderId: "49549236214",
  appId: "1:49549236214:web:9e4d9709c0610d014519d7"
};
const FIRESTORE_COLLECTION = 'yashFinancePortals';

let firebaseUser = null;
let firestoreDb = null;
let cloudSaveTimer = null;
let isSyncingFromCloud = false;

/* Called at the end of every render*() function — i.e. after any data change.
   Debounced so rapid edits (typing, multiple field changes) collapse into one write. */
function scheduleCloudSave(){
  if(!firebaseUser || isSyncingFromCloud) return;
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = setTimeout(()=> saveToFirestore(true), 1200);
}

function isFirebaseConfigured(){
  return !!firebaseConfig.apiKey && firebaseConfig.apiKey.indexOf('PASTE_')!==0
      && !!firebaseConfig.projectId && firebaseConfig.projectId.indexOf('PASTE_')!==0;
}
function initFirebase(){
  if(!isFirebaseConfigured() || typeof firebase==='undefined') return;
  try{
    firebase.initializeApp(firebaseConfig);
    firestoreDb = firebase.firestore();
    firebase.auth().onAuthStateChanged(user=>{
      firebaseUser = user;
      updateFirebaseUI();
      if(user){
        toast('Signed in as '+(user.displayName || user.email));
        loadFromFirestore(true); /* auto-load this account's saved data right after sign-in */
      }
    });
  }catch(err){ console.error('Firebase init error:', err); }
}
window.addEventListener('load', ()=> setTimeout(initFirebase, 400));

function onFirebaseSignInClick(){
  if(!isFirebaseConfigured()){ openModal('firebaseSetupOverlay'); return; }
  if(typeof firebase==='undefined'){
    toast('Firebase script hasn\'t loaded — check your internet connection and reload');
    return;
  }
  const provider = new firebase.auth.GoogleAuthProvider();
  firebase.auth().signInWithPopup(provider).catch(err=>{
    console.error(err);
    toast('Sign-in failed: '+err.message);
  });
}
function signOutFirebase(){
  if(typeof firebase==='undefined') return;
  firebase.auth().signOut();
  toast('Signed out');
}
function updateFirebaseUI(){
  const nameEl = document.getElementById('chipName');
  const roleEl = document.getElementById('chipRole');
  const avatarEl = document.getElementById('chipAvatar');
  const signBtn = document.getElementById('firebaseSignInBtn');
  const signOutBtn = document.getElementById('firebaseSignOutBtn');
  if(firebaseUser){
    nameEl.textContent = firebaseUser.displayName || firebaseUser.email;
    roleEl.textContent = firebaseUser.email || '';
    avatarEl.innerHTML = firebaseUser.photoURL
      ? `<img src="${firebaseUser.photoURL}" alt="" style="width:100%;height:100%;border-radius:50%;object-fit:cover;">`
      : (firebaseUser.displayName||'?')[0].toUpperCase();
    signBtn.style.display = 'none';
    signOutBtn.style.display = 'flex';
  } else {
    nameEl.textContent = 'Anitha';
    roleEl.textContent = 'Admin';
    avatarEl.textContent = 'AN';
    signBtn.style.display = 'flex';
    signOutBtn.style.display = 'none';
  }
  const cloudReady = !!firebaseUser;
  document.getElementById('cloudSaveBtn').classList.toggle('disabled-icon', !cloudReady);
  document.getElementById('cloudLoadBtn').classList.toggle('disabled-icon', !cloudReady);
}

/* ---- Firestore save/load. Each signed-in user's data lives at
   yashFinancePortals/{their uid} — set Firestore security rules so only that
   uid can read/write its own document (see the setup guide for the rule). ---- */
async function saveToFirestore(silent){
  if(!firebaseUser){ if(!silent) toast('Sign in first to save to the cloud'); return; }
  try{
    await firestoreDb.collection(FIRESTORE_COLLECTION).doc(firebaseUser.uid).set(buildPortalExportObject());
    if(!silent) toast('Saved to your Firebase account');
  }catch(err){
    console.error(err);
    toast('Cloud save failed: '+err.message);
  }
}
async function loadFromFirestore(silent){
  if(!firebaseUser){ if(!silent) toast('Sign in first to load from the cloud'); return; }
  isSyncingFromCloud = true; /* pause auto-save while we apply loaded data, so we don't immediately re-save it or race a stale write */
  try{
    const snap = await firestoreDb.collection(FIRESTORE_COLLECTION).doc(firebaseUser.uid).get();
    if(!snap.exists){ if(!silent) toast('No saved data found yet — use the save icon first'); return; }
    if(!applyPortalJSONImport(snap.data())){ if(!silent) toast('Saved data format not recognized'); return; }
    if(!silent) toast('Loaded from your Firebase account');
  }catch(err){
    console.error(err);
    if(!silent) toast('Cloud load failed: '+err.message);
  } finally {
    isSyncingFromCloud = false;
  }
}
document.getElementById('firebaseSignInBtn').addEventListener('click', onFirebaseSignInClick);
document.getElementById('firebaseSignOutBtn').addEventListener('click', signOutFirebase);
document.getElementById('firebaseSetupInfoBtn').addEventListener('click', ()=> openModal('firebaseSetupOverlay'));
document.getElementById('cloudSaveBtn').addEventListener('click', ()=> saveToFirestore());
document.getElementById('cloudLoadBtn').addEventListener('click', ()=> loadFromFirestore());
updateFirebaseUI();

/* ================= INIT ================= */
renderChart();
renderAssets();
renderSnapshots();
renderExpenses();
renderBudget();

/* inicial.js (paginación “slide”, 1 tarjeta por fila, modal fullscreen)
   - Lee HOJA INICIAL (B..Q) y oculta “0” en Área/Puesto/Tarea
   - Une con hoja “Movimiento repetitivo” y muestra estado basado en columnas P y W
   - Card: estado (badge) + FACTORES (chips J..P). Modal: detalles P y W + preguntas/respuestas.
*/

const COLS = {
  B: "Área",
  C: "Puesto de trabajo",
  D: "Tareas del puesto de trabajo",
  E: "Horario de funcionamiento",
  F: "Horas extras POR DIA",
  G: "Horas extras POR SEMANA",
  H: "N° Trabajadores Expuestos HOMBRE",
  I: "N° Trabajadores Expuestos MUJER",
  J: "Trabajo repetitivo de miembros superiores.",
  K: "Postura de trabajo estática",
  L: "MMC Levantamiento/Descenso",
  M: "MMC Empuje/Arrastre",
  N: "Manejo manual de pacientes / personas",
  O: "Vibración de cuerpo completo",
  P: "Vibración segmento mano – brazo",
  Q: "Resultado identificación inicial",
};

const RISKS = [
  { key: 'J', label: COLS.J, css: 'f-rep' },
  { key: 'K', label: COLS.K, css: 'f-post' },
  { key: 'L', label: COLS.L, css: 'f-lev' },
  { key: 'M', label: COLS.M, css: 'f-push' },
  { key: 'N', label: COLS.N, css: 'f-pcts' },
  { key: 'O', label: COLS.O, css: 'f-vcc' },
  { key: 'P', label: COLS.P, css: 'f-vhb' },
];

let RAW_ROWS = [];
let MOVREP_MAP = Object.create(null); // key -> {P, W, rowObj, rowArr}
let MOVREP_HEADERS = [];

let POSTURA_MAP = Object.create(null); // key -> {rowObj, rowArr, condAceptable, condCritica}
let POSTURA_HEADERS = [];
let POSTURA_LOOKUP = Object.create(null); // normalized header -> [indexes]

let MMC_LEV_MAP = Object.create(null); // key -> {rowObj, rowArr, condAceptable, condCritica}
let MMC_LEV_HEADERS = [];
let MMC_LEV_STRUCTURE = null;

let MMC_EMP_MAP = Object.create(null); // key -> {rowObj, rowArr, condAceptable, condCritica}
let MMC_EMP_HEADERS = [];
let MMC_EMP_STRUCTURE = null;

let FILTERS = { area: "", puesto: "", tarea: "", factorKey: "", factorState: "" };
let STATE = { page:1, perPage:10, pageMax:1 };

const el = (id) => document.getElementById(id);

/* ======= Helpers ======= */
function escapeHtml(str){
  return String(str ?? "").replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
function toLowerNoAccents(s){
  return String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}
function isZeroish(v){
  if(v===0) return true;
  const s=String(v??"").trim();
  return !!s && /^0(\.0+)?$/.test(s);
}
function keyTriple(area, puesto, tarea){
  return `${toLowerNoAccents(area)}|${toLowerNoAccents(puesto)}|${toLowerNoAccents(tarea)}`;
}
function findIndexInsensitive(headers, keys){
  const norm = headers.map(h => toLowerNoAccents(String(h||"")));
  for(const k of keys){
    const idx = norm.indexOf(toLowerNoAccents(k));
    if(idx>=0) return idx;
  }
  return null;
}
function findHeaderIndex(headers, patterns){
  const norm = headers.map(h => toLowerNoAccents(String(h||"")));
  for (let i=0;i<norm.length;i++){
    const h = norm[i];
    if (patterns.some(p => h.includes(toLowerNoAccents(p)))) return i;
  }
  return null;
}
function objectFromRow(headers, row){
  const o={};
  const seen = Object.create(null);
  headers.forEach((h,i)=>{
    const base = h || `Col${i+1}`;
    const norm = base || `Col${i+1}`;
    const count = seen[norm] = (seen[norm] || 0) + 1;
    const key = count === 1 ? norm : `${norm} (${count})`;
    o[key] = row[i]??"";
  });
  return o;
}

/* ======= Bootstrap ======= */
document.addEventListener("DOMContentLoaded", () => {
  attemptFetchDefault();
  wireUI();
});

/* ======= UI ======= */
function wireUI(){
  el("fileInput").addEventListener("change", handleFile);

  el("filterArea").addEventListener("change", () => {
    FILTERS.area = el("filterArea").value || "";
    populatePuesto();
    FILTERS.puesto = "";
    el("filterPuesto").value = "";
    populateTarea();
    FILTERS.tarea = "";
    el("filterTarea").value = "";
    STATE.page = 1;
    render();
  });
  el("filterPuesto").addEventListener("change", () => {
    FILTERS.puesto = el("filterPuesto").value || "";
    populateTarea();
    FILTERS.tarea = "";
    el("filterTarea").value = "";
    STATE.page = 1;
    render();
  });
  el("filterTarea").addEventListener("change", () => {
    FILTERS.tarea = el("filterTarea").value || "";
    STATE.page = 1;
    render();
  });

  populateFactor();
  el("filterFactor").addEventListener("change", () => {
    FILTERS.factorKey = el("filterFactor").value || "";
    STATE.page = 1;
    render();
  });
  el("filterFactorState").addEventListener("change", () => {
    FILTERS.factorState = el("filterFactorState").value || "";
    STATE.page = 1;
    render();
  });

  el("btnReset").addEventListener("click", (e) => {
    e.preventDefault();
    FILTERS = { area: "", puesto: "", tarea: "", factorKey:"", factorState:"" };
    el("filterArea").value = "";
    el("filterPuesto").value = "";
    el("filterTarea").value = "";
    el("filterFactor").value = "";
    el("filterFactorState").value = "";
    populatePuesto(true);
    populateTarea(true);
    STATE.page = 1;
    render();
  });
  el("btnReload").addEventListener("click", attemptFetchDefault);

  el("btnPrev").addEventListener("click", ()=>{ if(STATE.page>1){ STATE.page--; render(); window.scrollTo({top:0,behavior:'smooth'});} });
  el("btnNext").addEventListener("click", ()=>{ if(STATE.page<STATE.pageMax){ STATE.page++; render(); window.scrollTo({top:0,behavior:'smooth'});} });
  el("perPage").addEventListener("change", ()=>{ STATE.perPage = parseInt(el("perPage").value,10)||10; STATE.page=1; render(); });
  el("btnTop").addEventListener("click", ()=> window.scrollTo({top:0,behavior:'smooth'}));

  // Click en tarjeta → modal fullscreen
  el("cardsWrap").addEventListener("click", (ev) => {
    const open = ev.target.closest("[data-open]");
    const card = ev.target.closest("[data-idx]");
    if(open && card){
      const idx = Number(card.dataset.idx);
      if(Number.isFinite(idx) && RAW_ROWS[idx]) openDetail(RAW_ROWS[idx]);
    }
  });
}

/* ======= Carga de archivo por defecto ======= */
async function attemptFetchDefault(){
  if(!window.DEFAULT_XLSX_PATH) return;
  try{
    const res = await fetch(window.DEFAULT_XLSX_PATH + `?v=${Date.now()}`, {cache:"no-store"});
    if(!res.ok){ throw new Error("Fetch failed"); }
    const buf = await res.arrayBuffer();
    processWorkbook(buf);
  }catch(e){
    console.warn("No se pudo cargar el Excel por defecto. Seleccione manualmente.", e);
  }
}
function handleFile(evt){
  const file = evt.target.files?.[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = (e) => processWorkbook(e.target.result);
  reader.readAsArrayBuffer(file);
}

/* ======= Parse libro ======= */
function pickInicialSheet(wb){
  const target = (wb.SheetNames || []).find(n => /inicial|inicio/i.test(String(n||"")));
  return target || wb.SheetNames[0];
}
function pickMovRepSheet(wb){
  const cand = wb.SheetNames.find(n => /mov|repet/i.test(n.toLowerCase()));
  return cand || null;
}
function pickPosturaSheet(wb){
  const cand = wb.SheetNames.find(n => /postura|estatic/i.test(n.toLowerCase()));
  return cand || null;
}
function pickMmcLevSheet(wb){
  const names = wb.SheetNames || [];
  return names.find(n => /mmc/i.test(n) && /lev|desc/i.test(n.toLowerCase())) || null;
}
function pickMmcEmpSheet(wb){
  const names = wb.SheetNames || [];
  return names.find(n => /mmc/i.test(n) && (/emp/i.test(n.toLowerCase()) || /arras/i.test(n.toLowerCase()))) || null;
}

function processWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  /* Hoja INICIAL */
  const initialSheetName = pickInicialSheet(wb);
  const ws = wb.Sheets[initialSheetName];

  RAW_ROWS = [];
  if(ws && ws['!ref']){
    const range = XLSX.utils.decode_range(ws['!ref']);
    for(let r = 2; r <= range.e.r; r++){ // fila 3 visible
      const vals = {};
      function getCell(c){
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        const value = cell ? (cell.w ?? cell.v) : "";
        return value == null ? "" : String(value).trim();
      }
      const colMap = { B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16 };
      for(const [k, idx] of Object.entries(colMap)){ vals[k] = getCell(idx); }

      if(!(vals.B || vals.C || vals.D)) continue;

      // Normaliza SI/NO
      ['J','K','L','M','N','O','P'].forEach(k => {
        if(vals[k]){
          const up = vals[k].toString().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
          if(["SI","YES","Y","S"].includes(up)) vals[k] = "SI";
          else if(["NO","N"].includes(up)) vals[k] = "NO";
        }
      });

      // Resultado por defecto si viene vacío
      const allNo = ['J','K','L','M','N','O','P'].every(k => (vals[k] || "").toUpperCase() === "NO");
      const anyPresent = ['J','K','L','M','N','O','P'].some(k => (vals[k] || "") !== "");
      const Qcalc = (allNo && anyPresent)
        ? "Ausencia total del riesgo, reevaluar cada 3 años con nueva identificación inicial"
        : "Aplicar identificación avanzada-condición aceptable para cada tipo de factor de riesgo identificado";
      vals.Q = vals.Q || Qcalc;

      // Filtra falsos positivos "0"
      if(isZeroish(vals.B) || isZeroish(vals.C) || isZeroish(vals.D)) continue;

      RAW_ROWS.push(vals);
    }
  }

  /* Hoja Movimiento repetitivo */
  MOVREP_MAP = Object.create(null);
  MOVREP_HEADERS = [];
  const movSheetName = pickMovRepSheet(wb);
  if(movSheetName){
    const wsMov = wb.Sheets[movSheetName];
    if(wsMov){
      const rows2d = XLSX.utils.sheet_to_json(wsMov, { header:1, defval:"" });
      if(rows2d.length){
        // En esta hoja la **fila 2** contiene los nombres reales de las columnas
        const headerRow =
          (rows2d[1] && rows2d[1].some(x => String(x||"").trim() !== "")) ? rows2d[1] :
          rows2d[0];
        const headers = headerRow.map(h => String(h||""));
        MOVREP_HEADERS = headers;

        // Índices de llaves (por texto y con fallback B,C,D)
        const idxArea   = findHeaderIndex(headers, ["área de trabajo","area de trabajo","área","area"]) ?? 1;
        const idxPuesto = findHeaderIndex(headers, ["puesto de trabajo","puesto"]) ?? 2;
        const idxTarea  = findHeaderIndex(headers, ["tareas del puesto","tareas del puesto de trabajo","tarea"]) ?? 3;

        // Detecta P y W por texto; fallback a índices correctos (A=0 → P=15, W=22)
        const idxP = findHeaderIndex(headers, ["condición aceptable","condicion aceptable"]) ?? 15;
        const idxW = findHeaderIndex(headers, ["condición crítica","condicion critica"]) ?? 22;

        for(let i=2;i<rows2d.length;i++){ // datos desde fila 3 (0-based: 2)
          const r = rows2d[i] || [];
          const area   = r[idxArea]   ?? "";
          const puesto = r[idxPuesto] ?? "";
          const tarea  = r[idxTarea]  ?? "";
          if(!(area||puesto||tarea)) continue;
          const k = keyTriple(area, puesto, tarea);

          const rec = objectFromRow(headers, r);
          MOVREP_MAP[k] = {
            P: r[idxP] ?? "",
            W: r[idxW] ?? "",
            rowObj: rec,
            rowArr: r.slice()
          };
        }
      }
    }
  }

  /* Hoja Postura estática */
  POSTURA_MAP = Object.create(null);
  POSTURA_HEADERS = [];
  POSTURA_LOOKUP = Object.create(null);
  const posturaSheetName = pickPosturaSheet(wb);
  if(posturaSheetName){
    const wsPost = wb.Sheets[posturaSheetName];
    if(wsPost){
      const rows2d = XLSX.utils.sheet_to_json(wsPost, { header:1, defval:"" });
      if(rows2d.length){
        let headerRow = null;
        let headerIndex = -1;
        for(let i=0;i<rows2d.length;i++){
          const row = rows2d[i] || [];
          if(row.some(cell => /área de trabajo|area de trabajo/i.test(String(cell || "")))){
            headerRow = row;
            headerIndex = i;
            break;
          }
        }
        if(headerRow){
          const headers = headerRow.map(h => String(h||""));
          POSTURA_HEADERS = headers;
          const lookup = Object.create(null);
          headers.forEach((h,i)=>{
            const norm = toLowerNoAccents(String(h||"")).replace(/\s+/g," ").trim();
            if(!norm) return;
            (lookup[norm] || (lookup[norm] = [])).push(i);
          });
          POSTURA_LOOKUP = lookup;

          const idxArea   = findHeaderIndex(headers, ["área de trabajo","area de trabajo"]) ?? 1;
          const idxPuesto = findHeaderIndex(headers, ["puesto de trabajo","puesto"]) ?? 2;
          const idxTarea  = findHeaderIndex(headers, ["tareas del puesto","tarea"]) ?? 3;
          const idxAcept  = findHeaderIndex(headers, ["condición aceptable","condicion aceptable"]) ?? 29;
          const idxCrit   = findHeaderIndex(headers, ["condición crítica","condicion critica"]) ?? 47;

          for(let i=headerIndex+1;i<rows2d.length;i++){
            const r = rows2d[i] || [];
            const area   = r[idxArea]   ?? "";
            const puesto = r[idxPuesto] ?? "";
            const tarea  = r[idxTarea]  ?? "";
            if(!(area||puesto||tarea)) continue;
            const key = keyTriple(area, puesto, tarea);
            const rowObj = objectFromRow(headers, r);
            POSTURA_MAP[key] = {
              rowArr: r.slice(),
              rowObj,
              condAceptable: r[idxAcept] ?? "",
              condCritica: r[idxCrit] ?? ""
            };
          }
        }
      }
    }
  }

  /* Hoja MMC Levantamiento/Descenso */
  MMC_LEV_MAP = Object.create(null);
  MMC_LEV_HEADERS = [];
  MMC_LEV_STRUCTURE = null;
  const mmcLevSheet = pickMmcLevSheet(wb);
  if(mmcLevSheet){
    const wsLev = wb.Sheets[mmcLevSheet];
    if(wsLev){
      const rows2d = XLSX.utils.sheet_to_json(wsLev, { header:1, defval:"" });
      if(rows2d.length){
        const headerRow =
          (rows2d[1] && rows2d[1].some(x => String(x||"").trim() !== "")) ? rows2d[1] :
          rows2d[0];
        const headers = headerRow.map(h => String(h||""));
        MMC_LEV_HEADERS = headers;

        const idxArea   = findHeaderIndex(headers, ["área de trabajo","area de trabajo"]) ?? 1;
        const idxPuesto = findHeaderIndex(headers, ["puesto de trabajo","puesto"]) ?? 2;
        const idxTarea  = findHeaderIndex(headers, ["tareas del puesto","tarea"]) ?? 3;
        const idxAcept  = findHeaderIndex(headers, ["condición aceptable","condicion aceptable"]);
        const idxCrit   = findHeaderIndex(headers, ["condición crítica","condicion critica"]);
        MMC_LEV_STRUCTURE = buildMmcStructure(headers, idxAcept, idxCrit);

        for(let i=2;i<rows2d.length;i++){
          const r = rows2d[i] || [];
          const area   = r[idxArea]   ?? "";
          const puesto = r[idxPuesto] ?? "";
          const tarea  = r[idxTarea]  ?? "";
          if(!(area||puesto||tarea)) continue;
          const key = keyTriple(area, puesto, tarea);

          MMC_LEV_MAP[key] = {
            rowArr: r.slice(),
            rowObj: objectFromRow(headers, r),
            condAceptable: (idxAcept != null) ? (r[idxAcept] ?? "") : "",
            condCritica: (idxCrit != null) ? (r[idxCrit] ?? "") : ""
          };
        }
      }
    }
  }

  /* Hoja MMC Empuje/Arrastre */
  MMC_EMP_MAP = Object.create(null);
  MMC_EMP_HEADERS = [];
  MMC_EMP_STRUCTURE = null;
  const mmcEmpSheet = pickMmcEmpSheet(wb);
  if(mmcEmpSheet){
    const wsEmp = wb.Sheets[mmcEmpSheet];
    if(wsEmp){
      const rows2d = XLSX.utils.sheet_to_json(wsEmp, { header:1, defval:"" });
      if(rows2d.length){
        const headerRow =
          (rows2d[1] && rows2d[1].some(x => String(x||"").trim() !== "")) ? rows2d[1] :
          rows2d[0];
        const headers = headerRow.map(h => String(h||""));
        MMC_EMP_HEADERS = headers;

        const idxArea   = findHeaderIndex(headers, ["área de trabajo","area de trabajo"]) ?? 1;
        const idxPuesto = findHeaderIndex(headers, ["puesto de trabajo","puesto"]) ?? 2;
        const idxTarea  = findHeaderIndex(headers, ["tareas del puesto","tarea"]) ?? 3;
        const idxAcept  = findHeaderIndex(headers, ["condición aceptable","condicion aceptable"]);
        const idxCrit   = findHeaderIndex(headers, ["condición crítica","condicion critica"]);
        MMC_EMP_STRUCTURE = buildMmcStructure(headers, idxAcept, idxCrit);

        for(let i=2;i<rows2d.length;i++){
          const r = rows2d[i] || [];
          const area   = r[idxArea]   ?? "";
          const puesto = r[idxPuesto] ?? "";
          const tarea  = r[idxTarea]  ?? "";
          if(!(area||puesto||tarea)) continue;
          const key = keyTriple(area, puesto, tarea);

          MMC_EMP_MAP[key] = {
            rowArr: r.slice(),
            rowObj: objectFromRow(headers, r),
            condAceptable: (idxAcept != null) ? (r[idxAcept] ?? "") : "",
            condCritica: (idxCrit != null) ? (r[idxCrit] ?? "") : ""
          };
        }
      }
    }
  }

  populateArea();
  populatePuesto(true);
  populateTarea(true);
  render();
}

/* ======= Filtros ======= */
function uniqueSorted(arr){
  return [...new Set(arr.filter(v => v && String(v).trim() !== ""))].sort((a,b)=> String(a).localeCompare(String(b)));
}
function populateArea(){
  const opts = uniqueSorted(RAW_ROWS.map(r => r.B));
  const sel = el("filterArea");
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = false;
}
function populatePuesto(){
  const sel = el("filterPuesto");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B === FILTERS.area);
  const opts = uniqueSorted(list.map(r => r.C));
  sel.innerHTML = `<option value="">(Todos)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length === 0;
}

function classifyCriticalState(value){
  const val = toLowerNoAccents(String(value || ""));
  if(!val) return "";
  if(/no\s*(es\s*)?crit/.test(val)) return "warn";
  if(/crit/.test(val)) return "bad";
  return "";
}

function classifyRowHighlight(label, value){
  const lab = toLowerNoAccents(String(label||""));
  const val = toLowerNoAccents(String(value||""));

  // Condición Aceptable: "Aceptable" => verde; "No aceptable" => rojo
  if(lab.includes("condicion aceptable") || lab.includes("condición aceptable")){
    if(val.includes("no acept")) return "is-aceptable-bad";
    if(val.includes("acept"))    return "is-aceptable-ok";
  }

  // Condición Crítica: "No crítico" => amarillo; "Crítico" => rojo
  if(lab.includes("condicion critica") || lab.includes("condición crítica")){
    const crit = classifyCriticalState(value);
    if(crit === "bad") return "is-critica-bad";
    if(crit === "warn") return "is-critica-warn";
    return "";
  }
  return "";
}

function isFactorPresent(row, key){
  if(!row) return false;
  const raw = String(row[key] ?? "").trim();
  if(!raw) return false;
  return toLowerNoAccents(raw) === "si";
}

function shouldSkipCritical(acceptableValue){
  const text = toLowerNoAccents(String(acceptableValue ?? ""));
  if(!text.trim()) return false;
  if(text.includes("no acept")) return false;
  return text.includes("acept");
}

function normalizedCriticalValue(acceptableValue, criticalValue){
  return shouldSkipCritical(acceptableValue) ? "N/A" : (criticalValue ?? "");
}

function postureIndex(label, occurrence=0){
  if(!label) return null;
  const norm = toLowerNoAccents(String(label||"")).replace(/\s+/g," ").trim();
  const arr = POSTURA_LOOKUP[norm];
  if(!arr || arr.length === 0) return null;
  return arr[Math.min(occurrence, arr.length - 1)];
}

function postureEntry(label, occurrence=0){
  const idx = postureIndex(label, occurrence);
  if(idx == null) return { label, index:null };
  return { label: POSTURA_HEADERS[idx] || label, index: idx };
}

function stateHighlightClass(label, value){
  const base = classifyRowHighlight(label, value);
  if(base === "is-aceptable-ok") return "hl-ok";
  if(base === "is-aceptable-bad" || base === "is-critica-bad") return "hl-risk";
  if(base === "is-critica-warn") return "hl-warn";
  return "";
}


function populateTarea(){
  const sel = el("filterTarea");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B === FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C === FILTERS.puesto);
  const opts = uniqueSorted(list.map(r => r.D));
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length === 0;
}
function populateFactor(){
  const sel = el("filterFactor");
  sel.innerHTML = `<option value="">(Todos)</option>` +
    RISKS.map(r => `<option value="${r.key}">${escapeHtml(r.label)}</option>`).join("");
}

function filteredRows(){
  let list = RAW_ROWS.slice();
  if(FILTERS.area)   list = list.filter(r => r.B === FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C === FILTERS.puesto);
  if(FILTERS.tarea)  list = list.filter(r => r.D === FILTERS.tarea);

  if(FILTERS.factorKey){
    const key = FILTERS.factorKey;
    const state = (FILTERS.factorState || "").toUpperCase();
    list = list.filter(r => {
      const v = (r[key] || "").toUpperCase();
      if(state === "SI") return v === "SI";
      if(state === "NO") return v === "NO";
      return v === "SI" || v === "NO";
    });
  }
  return list;
}

/* ======= Render + paginación ======= */
function render(){
  const target = el("cardsWrap");
  const all = filteredRows();

  const per = STATE.perPage = parseInt(el("perPage").value,10) || 10;
  STATE.pageMax = Math.max(1, Math.ceil(all.length / per));
  if(STATE.page > STATE.pageMax) STATE.page = STATE.pageMax;
  const start = (STATE.page - 1) * per;
  const pageData = all.slice(start, start + per);

  el("countRows").textContent = pageData.length;
  el("countRowsTotal").textContent = all.length;
  el("pageCur").textContent = STATE.page;
  el("pageMax").textContent = STATE.pageMax;

  if(pageData.length === 0){
    target.innerHTML = `<div class="col"><div class="alert alert-warning mb-0">
      <i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.
    </div></div>`;
    return;
  }

  target.innerHTML = pageData.map((r) => {
    const idx = RAW_ROWS.indexOf(r);
    return cardHtml(r, idx);
  }).join("");
}

/* ======= Estado Movimiento repetitivo (P/W) ======= */
function getMovRepFor(r){
  const k = keyTriple(r.B, r.C, r.D);
  return MOVREP_MAP[k] || null;
}
function getPosturaFor(r){
  const k = keyTriple(r.B, r.C, r.D);
  return POSTURA_MAP[k] || null;
}
function getMmcLevFor(r){
  const k = keyTriple(r.B, r.C, r.D);
  return MMC_LEV_MAP[k] || null;
}
function getMmcEmpFor(r){
  const k = keyTriple(r.B, r.C, r.D);
  return MMC_EMP_MAP[k] || null;
}
function classifyMovRep(p, w){
  const s = `${String(p||"")} ${String(w||"")}`;
  const t = toLowerNoAccents(s);
  if(!t.trim()) return { cls:"status-unk", label:"Sin dato" };

  if(t.includes("no acept") || t.includes("alto") || t.includes("riesgo alto") || t.includes("critico") || t.includes("crítico"))
    return { cls:"status-bad", label:"No aceptable" };
  if(t.includes("moderad") || t.includes("medio") || t.includes("precauc") || t.includes("mejorable"))
    return { cls:"status-warn", label:"Precaución" };
  if(t.includes("acept") || t.includes("bajo") || t.includes("sin riesgo"))
    return { cls:"status-ok", label:"Aceptable" };
  return { cls:"status-unk", label:"Revisar" };
}

/* ======= Chips de factores ======= */
function factorChips(r){
  const parts = [];
  for(const rf of RISKS){
    const raw = (r[rf.key]||"").toString().trim().toUpperCase();
    if(raw !== "SI" && raw !== "NO") continue;
    const isYes = raw === "SI";
    const cls = `factor-chip ${rf.css} ${isYes ? 'is-yes' : 'is-no'}`;
    const ico = isYes ? '<i class="bi bi-check-circle-fill"></i>' : '<i class="bi bi-dash-circle-fill"></i>';
    const st  = `<span class="state">${isYes ? 'S' : 'N'}</span>`;
    parts.push(`<span class="${cls}" title="${escapeHtml(rf.label)}">${st}${ico} <span>${escapeHtml(rf.label)}</span></span>`);
  }
  if(!parts.length){
    return `<span class="factor-chip is-no" title="Sin factores con SI/NO definidos">
      <span class="state">-</span><i class="bi bi-info-circle"></i> Sin factores definidos
    </span>`;
  }
  return parts.join("");
}

function renderStateCard(title, value, icon){
  const cls = stateHighlightClass(title, value) || "hl-neutral";
  const valText = String(value ?? "").trim();
  const display = valText ? escapeHtml(valText) : '<span class="text-muted">Sin dato</span>';
  const iconHtml = icon ? `<i class="bi ${icon}"></i>` : '';
  return `
    <div class="col-12 col-md-6">
      <div class="state-card ${cls}">
        <div class="sc-head">${iconHtml}<span>${escapeHtml(title)}</span></div>
        <div class="sc-body fw-bold">${display}</div>
      </div>
    </div>
  `;
}

function renderTabSummary(cards){
  if(!cards || !cards.length) return "";
  return `<div class="tab-summary row row-cols-1 row-cols-md-2 g-3">${cards.join("")}</div>`;
}

function getPosturaStructure(){
  if(!POSTURA_HEADERS.length) return null;
  const mk = (label, occurrence=0) => postureEntry(label, occurrence);

  return {
    acceptable: [
      {
        title: "Condición Aceptable (CABEZA Y TRONCO)",
        entries: [
          mk("¿Las posturas de tronco y cuello son simétricas?"),
          mk("En caso de haber flexión de tronco (inclinación hacia delante), ¿es menor a 20º?, o en caso de existir extensión (inclinación hacia atrás), ¿el tronco está totalmente apoyado?"),
          mk("Si existe flexión de tronco entre 20º y 60º, ¿este se encuentra totalmente apoyado?"),
          mk("¿Está ausente la extensión de cuello?"),
          mk("En caso de que exista flexión de cuello, ¿no supera los 25º?"),
          mk("Estando la cabeza inclinada hacia atrás (extensión), ¿se encuentra totalmente apoyada?, o en caso de inclinación anterior (flexión), ¿está a menos de 25º?"),
          mk("Si está sentado, ¿la curvatura de la espalda se mantiene no forzada?")
        ]
      },
      {
        title: "Condición Aceptable · Miembros superiores",
        entries: [
          mk("Miembro con mayor exigencia", 0),
          mk("¿Están ausentes las posturas de MMSS separadas del cuerpo, elevadas sobre nivel de hombro de manera sostenida y no apoyadas?"),
          mk("¿Los hombros no se encuentran levantados?"),
          mk("Estando el brazo sin apoyo completo, ¿La elevación del miembro superior es menor a 20°?"),
          mk("Con el brazo totalmente apoyado, ¿la elevación del miembro superior no supera los 60°?"),
          mk("¿Están ausentes la flexión / extensión extrema de codo y la rotación extrema de antebrazo?"),
          mk("¿Está ausente el giro extremo del antebrazo?")
        ]
      },
      {
        title: "Condición Aceptable · Miembros inferiores",
        entries: [
          mk("Miembro con mayor exigencia", 1),
          mk("¿Está ausente la flexión extrema de rodilla?"),
          mk("En postura de pie ¿La rodilla no está en flexión?"),
          mk("¿El tobillo está en posición neutra?"),
          mk("¿Están ausentes las posiciones cuclillas y arrodillado?"),
          mk("Cuando está sentado, ¿El ángulo de la rodilla está entre 90º y 135º?")
        ]
      }
    ],
    acceptableResult: mk("Condición Aceptable"),
    critical: [
      {
        title: "Condición Crítica (CABEZA Y TRONCO)",
        entries: [
          mk("¿La postura de tronco o la postura de cuello están en rangos extremos?"),
          mk("¿Existe flexión de tronco (inclinación hacia adelante) de 60° o más?"),
          mk("¿Existe flexión de tronco (aun cuando sea levemente) durante más de 4 minutos?"),
          mk("¿Está la cabeza extendida (inclinada hacia atrás) sin apoyo?"),
          mk("¿Está la cabeza en flexión (inclinación hacia adelante) de 85° o más?"),
          mk("¿Está la cabeza en flexión (aun cuando sea levemente) durante más de 8 minutos?"),
          mk("Si está sentado, ¿la espalda (región lumbar) está forzada y no logra mantener la curvatura natural?")
        ]
      },
      {
        title: "Condición Crítica · Miembros superiores",
        entries: [
          mk("¿Hay posturas de brazos que los trabajadores relaten como muy incomodas y que les impiden el normal desenvolvimiento?"),
          mk("¿Los hombros se encuentran levantados sobre los 60°?"),
          mk("¿Los hombros se encuentran levantados (aún cuando sea levemente) durante más de 3 minutos?"),
          mk("¿Hay flexión / extensión extrema de codo y rotación extrema de antebrazo?"),
          mk("¿Hay giro extremo de muñeca?")
        ]
      },
      {
        title: "Condición Crítica · Miembros inferiores",
        entries: [
          mk("¿Hay flexión extrema de rodilla (posición de cuclillas o sentado en los talones)?"),
          mk("Estando en postura de pie, ¿la rodilla se encuentra en flexión leve sostenida?"),
          mk("¿El tobillo se encuentra en una posición extrema de flexión o extensión?"),
          mk("¿Se adoptan posiciones en cuclillas y/o arrodillado de la rodilla?"),
          mk("Estando sentado, ¿la angulación de rodilla es menor de 90° y mayor de 135°?")
        ]
      }
    ],
    criticalResult: mk("Condición Crítica"),
    plan: [
      mk("Fecha aplicación identificación avanzada"),
      mk("Medidas de control (administrativas)"),
      mk("Medidas de control (Ingeniería)"),
      mk("Responsable aplicación medida"),
      mk("Fecha de implementación medidas (Max 90 días)"),
      mk("Evidencia")
    ]
  };
}

function renderPosturaSection(section, row){
  if(!section) return "";
  const body = (section.entries || []).map((entry) => {
    if(entry.index == null) return null;
    const raw = row[entry.index];
    const value = String(raw ?? "").trim();
    const cell = value ? escapeHtml(value) : '<span class="text-muted fst-italic">Sin dato</span>';
    const rowCls = classifyRowHighlight(entry.label, raw);
    return `<tr class="${rowCls}"><th>${escapeHtml(entry.label)}</th><td>${cell}</td></tr>`;
  }).filter(Boolean);
  if(!body.length) return "";
  return `
    <div class="detail-section">
      <h6 class="section-title">${escapeHtml(section.title)}</h6>
      <div class="table-like table-compact">
        <table>
          <thead><tr><th style="min-width:260px">Pregunta</th><th>Respuesta</th></tr></thead>
          <tbody>${body.join("")}</tbody>
        </table>
      </div>
    </div>
  `;
}

function renderPosturaPlan(entries, row){
  if(!entries || !entries.length) return "";
  const items = entries.map((entry) => {
    if(entry.index == null) return null;
    const raw = row[entry.index];
    const value = String(raw ?? "").trim();
    const cell = value ? escapeHtml(value) : '<span class="text-muted fst-italic">Sin dato</span>';
    return `
      <div class="plan-item">
        <div class="plan-label">${escapeHtml(entry.label)}</div>
        <div class="plan-value">${cell}</div>
      </div>
    `;
  }).filter(Boolean);
  if(!items.length) return "";
  return `
    <div class="detail-section">
      <h6 class="section-title">Plan de acción</h6>
      <div class="plan-grid">
        ${items.join("")}
      </div>
    </div>
  `;
}

function renderPosturaTab(post, summaryCards){
  const summaryBlock = renderTabSummary(summaryCards);
  if(!post){
    return `<div class="postura-tab">${summaryBlock}<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se encontraron coincidencias en la hoja “Postura estática”.</div></div>`;
  }
  const structure = getPosturaStructure();
  if(!structure){
    return `<div class="postura-tab">${summaryBlock}<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se pudo interpretar la hoja “Postura estática”.</div></div>`;
  }
  const row = post.rowArr || [];
  const planBlock = renderPosturaPlan(structure.plan, row);
  const skipCritical = shouldSkipCritical(post.condAceptable);
  const acceptable = structure.acceptable.map(sec => renderPosturaSection(sec, row)).filter(Boolean).join("") ||
    `<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> Sin respuestas para condición aceptable.</div>`;
  const critical = skipCritical
    ? `<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> No se evaluó la condición crítica (N/A).</div>`
    : (structure.critical.map(sec => renderPosturaSection(sec, row)).filter(Boolean).join("") ||
      `<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> Sin respuestas para condición crítica.</div>`);

  return `
    <div class="postura-tab">
      ${summaryBlock}
      ${planBlock ? `<div class="group-block">${planBlock}</div>` : ""}
      <div class="group-block${planBlock ? " mt-4" : ""}">
        <div class="group-title text-uppercase small text-muted fw-bold mb-2">Condición Aceptable</div>
        ${acceptable}
      </div>
      <div class="group-block mt-4">
        <div class="group-title text-uppercase small text-muted fw-bold mb-2">Condición Crítica</div>
        ${critical}
      </div>
    </div>
  `;
}

function buildMmcStructure(headers, idxAcept, idxCrit){
  if(!headers || idxAcept == null) return null;
  const list = headers.map(h => String(h || ""));
  const infoEnd = 8; // columnas A..I contienen datos base

  const acceptable = [];
  for(let i = infoEnd + 1; i < idxAcept; i++){
    const label = list[i];
    if(!label || !label.trim()) continue;
    acceptable.push({ label, index: i });
  }

  const critical = [];
  if(idxCrit != null){
    for(let i = idxAcept + 1; i < idxCrit; i++){
      const label = list[i];
      if(!label || !label.trim()) continue;
      critical.push({ label, index: i });
    }
  }

  const plan = [];
  const planStart = (idxCrit != null ? idxCrit + 1 : idxAcept + 1);
  for(let i = planStart; i < list.length; i++){
    const label = list[i];
    if(!label || !label.trim()) continue;
    plan.push({ label, index: i });
  }

  return {
    acceptable,
    acceptableResultIndex: idxAcept,
    critical,
    criticalResultIndex: idxCrit,
    plan
  };
}

function renderMmcSection(title, entries, row){
  if(!entries || !entries.length) return "";
  return renderPosturaSection({ title, entries }, row);
}

function renderMmcPlan(entries, row){
  if(!entries || !entries.length) return "";
  return renderPosturaPlan(entries, row);
}

function renderMmcTab(data, structure, sheetLabel, summaryCards){
  const summaryBlock = renderTabSummary(summaryCards);
  if(!structure){
    return `<div class="postura-tab">${summaryBlock}<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se pudo interpretar la hoja “${escapeHtml(sheetLabel)}”.</div></div>`;
  }
  if(!data){
    return `<div class="postura-tab">${summaryBlock}<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se encontraron coincidencias en la hoja “${escapeHtml(sheetLabel)}”.</div></div>`;
  }

  const row = data.rowArr || [];
  const skipCritical = shouldSkipCritical(data.condAceptable);
  const planBlock = renderMmcPlan(structure.plan, row);
  const acceptableBlock = renderMmcSection("Preguntas evaluadas", structure.acceptable, row) ||
    `<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> Sin respuestas para condición aceptable.</div>`;
  const criticalBlock = skipCritical
    ? `<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> No se evaluó la condición crítica (N/A).</div>`
    : (structure.critical && structure.critical.length
      ? (renderMmcSection("Preguntas evaluadas", structure.critical, row) ||
          `<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> Sin respuestas para condición crítica.</div>`)
      : "");

  return `
    <div class="postura-tab">
      ${summaryBlock}
      ${planBlock ? `<div class="group-block">${planBlock}</div>` : ""}
      <div class="group-block${planBlock ? " mt-4" : ""}">
        <div class="group-title text-uppercase small text-muted fw-bold mb-2">Condición Aceptable</div>
        ${acceptableBlock}
      </div>
      ${(structure.critical && structure.critical.length) || skipCritical ? `
        <div class="group-block mt-4">
          <div class="group-title text-uppercase small text-muted fw-bold mb-2">Condición Crítica</div>
          ${criticalBlock}
        </div>` : ""}
    </div>
  `;
}

function renderMmcLevTab(data, summaryCards){
  return renderMmcTab(data, MMC_LEV_STRUCTURE, "MMC Levantamiento/Descenso", summaryCards);
}

function renderMmcEmpTab(data, summaryCards){
  return renderMmcTab(data, MMC_EMP_STRUCTURE, "MMC Empuje/Arrastre", summaryCards);
}

/* ======= HTML Tarjeta + Modal ======= */
function cardHtml(r, idx){
  const mov = getMovRepFor(r);
  const movCritical = normalizedCriticalValue(mov?.P, mov?.W);
  const posturaRaw = POSTURA_HEADERS.length ? getPosturaFor(r) : null;
  const posturaCritical = normalizedCriticalValue(posturaRaw?.condAceptable, posturaRaw?.condCritica);
  const mmcLevRaw = MMC_LEV_STRUCTURE ? getMmcLevFor(r) : null;
  const mmcLevCritical = normalizedCriticalValue(mmcLevRaw?.condAceptable, mmcLevRaw?.condCritica);
  const mmcEmpRaw = MMC_EMP_STRUCTURE ? getMmcEmpFor(r) : null;
  const mmcEmpCritical = normalizedCriticalValue(mmcEmpRaw?.condAceptable, mmcEmpRaw?.condCritica);
  const status = classifyMovRep(mov?.P, movCritical);
  const cardState = status?.cls ? status.cls.replace("status-", "") : "";
  const factorPresent = {
    mov: isFactorPresent(r, 'J'),
    postura: isFactorPresent(r, 'K'),
    mmcLev: isFactorPresent(r, 'L'),
    mmcEmp: isFactorPresent(r, 'M')
  };
  const summaryCards = [];
  const addSummary = (html) => { if(html) summaryCards.push(html); };
  if(factorPresent.mov){
    addSummary(renderStateCard("Trabajo repetitivo de miembros superiores · Condición aceptable", mov?.P, "bi-activity"));
    addSummary(renderStateCard("Trabajo repetitivo de miembros superiores · Condición crítica", movCritical, "bi-exclamation-diamond-fill"));
  }
  if(factorPresent.postura && POSTURA_HEADERS.length){
    addSummary(renderStateCard("Postura estática · Condición aceptable", posturaRaw?.condAceptable, "bi-person-standing"));
    addSummary(renderStateCard("Postura estática · Condición crítica", posturaCritical, "bi-exclamation-octagon"));
  }
  if(factorPresent.mmcLev && MMC_LEV_STRUCTURE){
    addSummary(renderStateCard("MMC Levantamiento/Descenso · Condición aceptable", mmcLevRaw?.condAceptable, "bi-box-seam"));
    addSummary(renderStateCard("MMC Levantamiento/Descenso · Condición crítica", mmcLevCritical, "bi-exclamation-octagon-fill"));
  }
  if(factorPresent.mmcEmp && MMC_EMP_STRUCTURE){
    addSummary(renderStateCard("MMC Empuje/Arrastre · Condición aceptable", mmcEmpRaw?.condAceptable, "bi-cart-check"));
    addSummary(renderStateCard("MMC Empuje/Arrastre · Condición crítica", mmcEmpCritical, "bi-exclamation-triangle-fill"));
  }
  const hasAdvanced = summaryCards.length > 0;
  const summaryBlock = hasAdvanced
    ? `<div class="advanced-grid tab-summary row row-cols-1 row-cols-md-2 g-3">${summaryCards.join("")}</div>`
    : `<div class="alert alert-light border text-muted mb-0"><i class="bi bi-info-circle"></i> No hay información de identificación avanzada disponible.</div>`;

  return `
    <div class="col" data-idx="${idx}">
      <div class="card card-ficha h-100 shadow-sm"${cardState ? ` data-state="${cardState}"` : ""}>
        <div class="card-body">
          <div class="d-flex align-items-start justify-content-between">
            <div>
              <div class="small text-muted mb-1">Tarea</div>
              <h5 class="title mb-2">${escapeHtml(r.D || "-")}</h5>
            </div>
            <div class="text-end">
              <span class="chip"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
            </div>
          </div>

          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
          <div class="mb-2"><i class="bi bi-geo-alt"></i> <strong>Área:</strong> ${escapeHtml(r.B || "-")}</div>

          <!-- (1) Fila de horario/HE -->
          <div class="row g-2 small mb-2">
            <div class="col-6"><i class="bi bi-clock"></i> <strong>Horario:</strong> ${escapeHtml(r.E || "-")}</div>
            <div class="col-6"><i class="bi bi-plus-circle"></i> <strong>HE/Día:</strong> ${escapeHtml(r.F || "0")}</div>
            <div class="col-6"><i class="bi bi-plus-circle-dotted"></i> <strong>HE/Semana:</strong> ${escapeHtml(r.G || "0")}</div>
          </div>

          <!-- FACTORES (J..P) -->
          <div class="mb-2">
            <div class="small text-muted mb-1"><i class="bi bi-exclamation-octagon"></i> Identificación Inicial</div>
            <div class="factors-wrap">
              ${factorChips(r)}
            </div>
          </div>

          <div class="advanced-block mt-3">
            <div class="advanced-title small text-muted text-uppercase fw-bold">
              <i class="bi bi-clipboard2-pulse"></i> Identificación avanzada
            </div>
            ${summaryBlock}
          </div>

          <div class="d-flex justify-content-end mt-3">
            <button type="button" class="btn btn-primary btn-sm btn-open" data-open>
              <i class="bi bi-arrows-fullscreen"></i> Ver detalles
            </button>
          </div>
        </div>
      </div>
    </div>
  `;
}

/* Ignorar en modal: Col2..Col9 (1..8 zero-based) y títulos */
const SKIP_IDX = new Set([1,2,3,4,5,6,7,8]);
const SKIP_LABELS = new Set([
  "mujeres","col2","col3","col4","col5","col6","col7","col8","col9","n°","n."
]);

function renderMovRepTab(mov, summaryCards){
  const summaryBlock = renderTabSummary(summaryCards);
  if(!mov || !(mov.rowArr || mov.rowObj)){
    return `<div class="postura-tab">${summaryBlock}<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se encontraron detalles coincidentes en la hoja “Movimiento repetitivo”.</div></div>`;
  }

  const rows = [];
  const skipCritical = shouldSkipCritical(mov.P);
  if(mov.rowArr && Array.isArray(MOVREP_HEADERS) && MOVREP_HEADERS.length){
    for(let i=0;i<MOVREP_HEADERS.length;i++){
      if(SKIP_IDX.has(i)) continue;
      const label = MOVREP_HEADERS[i] || `Col${i+1}`;
      if(SKIP_LABELS.has(toLowerNoAccents(label))) continue;
      const normLabel = toLowerNoAccents(label);
      let val = mov.rowArr[i];
      if(skipCritical && normLabel.includes("condicion critica")){
        val = "N/A";
      }
      if(String(val ?? "").trim() === "") continue;
      rows.push([label, val]);
    }
  }else{
    for(const [k,v] of Object.entries(mov.rowObj || {})){
      if(String(v ?? "").trim() === "") continue;
      if(SKIP_LABELS.has(toLowerNoAccents(k))) continue;
      const normLabel = toLowerNoAccents(k);
      const value = (skipCritical && normLabel.includes("condicion critica")) ? "N/A" : v;
      rows.push([k, value]);
    }
  }

  if(!rows.length){
    return `<div class="postura-tab">${summaryBlock}<div class="alert alert-light border text-muted"><i class="bi bi-info-circle"></i> Sin respuestas registradas en esta hoja.</div></div>`;
  }

  const bodyHtml = rows.map(([k,v]) => {
    const rowCls = classifyRowHighlight(k, v);
    return `<tr class="${rowCls}"><th>${escapeHtml(k)}</th><td>${escapeHtml(String(v))}</td></tr>`;
  }).join("");

  return `
    <div class="postura-tab">
      ${summaryBlock}
      <div class="table-like">
        <table>
          <thead><tr><th style="min-width:260px">Pregunta</th><th>Respuesta</th></tr></thead>
          <tbody>${bodyHtml}</tbody>
        </table>
      </div>
    </div>
  `;
}

function openDetail(r){
  const movRaw = getMovRepFor(r);
  const posturaRaw = getPosturaFor(r);
  const mmcLevRaw = MMC_LEV_STRUCTURE ? getMmcLevFor(r) : null;
  const mmcEmpRaw = MMC_EMP_STRUCTURE ? getMmcEmpFor(r) : null;

  const movCritical = normalizedCriticalValue(movRaw?.P, movRaw?.W);
  const posturaCritical = normalizedCriticalValue(posturaRaw?.condAceptable, posturaRaw?.condCritica);
  const mmcLevCritical = normalizedCriticalValue(mmcLevRaw?.condAceptable, mmcLevRaw?.condCritica);
  const mmcEmpCritical = normalizedCriticalValue(mmcEmpRaw?.condAceptable, mmcEmpRaw?.condCritica);

  const mov = movRaw ? { ...movRaw, W: movCritical } : null;
  const postura = posturaRaw ? { ...posturaRaw, condCritica: posturaCritical } : null;
  const mmcLev = mmcLevRaw ? { ...mmcLevRaw, condCritica: mmcLevCritical } : null;
  const mmcEmp = mmcEmpRaw ? { ...mmcEmpRaw, condCritica: mmcEmpCritical } : null;

  const factorPresent = {
    mov: isFactorPresent(r, 'J'),
    postura: isFactorPresent(r, 'K'),
    mmcLev: isFactorPresent(r, 'L'),
    mmcEmp: isFactorPresent(r, 'M')
  };

  const tabSummaries = Object.create(null);
  const addSummary = (tabId, cardHtml) => {
    if(!cardHtml) return;
    (tabSummaries[tabId] ||= []).push(cardHtml);
  };
  if(factorPresent.mov){
    addSummary("mov", renderStateCard("Mov. repetitivo · Condición aceptable", mov?.P, "bi-activity"));
    addSummary("mov", renderStateCard("Mov. repetitivo · Condición crítica", mov?.W, "bi-exclamation-diamond-fill"));
  }
  if(factorPresent.postura){
    addSummary("postura", renderStateCard("Postura estática · Condición aceptable", postura?.condAceptable, "bi-person-standing"));
    addSummary("postura", renderStateCard("Postura estática · Condición crítica", postura?.condCritica, "bi-exclamation-octagon"));
  }
  if(factorPresent.mmcLev && MMC_LEV_STRUCTURE){
    addSummary("mmc-lev", renderStateCard("MMC Levantamiento/Descenso · Condición aceptable", mmcLev?.condAceptable, "bi-box-seam"));
    addSummary("mmc-lev", renderStateCard("MMC Levantamiento/Descenso · Condición crítica", mmcLev?.condCritica, "bi-exclamation-octagon-fill"));
  }
  if(factorPresent.mmcEmp && MMC_EMP_STRUCTURE){
    addSummary("mmc-emp", renderStateCard("MMC Empuje/Arrastre · Condición aceptable", mmcEmp?.condAceptable, "bi-cart-check"));
    addSummary("mmc-emp", renderStateCard("MMC Empuje/Arrastre · Condición crítica", mmcEmp?.condCritica, "bi-exclamation-triangle-fill"));
  }

  const header = `
    <div class="detail-card mb-3">
      <div class="d-flex flex-wrap justify-content-between align-items-start gap-2">
        <div>
          <div class="small text-muted">Tarea</div>
          <h5 class="mb-1">${escapeHtml(r.D || "-")}</h5>
          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
          <div class="mb-1"><i class="bi bi-geo-alt"></i> <strong>Área:</strong> ${escapeHtml(r.B || "-")}</div>
        </div>
      </div>
      <div class="mt-3">
        <div class="small text-muted mb-1"><i class="bi bi-exclamation-octagon"></i> Factores</div>
        <div class="factors-wrap">${factorChips(r)}</div>
      </div>
    </div>
  `;

  const tabs = [];
  const disabledTabs = [];
  const sections = [
    {
      id: "mov",
      title: "Movimiento repetitivo",
      present: factorPresent.mov,
      available: Boolean(mov || MOVREP_HEADERS.length),
      render: () => renderMovRepTab(mov, tabSummaries["mov"])
    },
    {
      id: "postura",
      title: "Postura estática",
      present: factorPresent.postura,
      available: Boolean(postura || POSTURA_HEADERS.length),
      render: () => renderPosturaTab(postura, tabSummaries["postura"])
    },
    {
      id: "mmc-lev",
      title: "MMC Levantamiento/Descenso",
      present: factorPresent.mmcLev,
      available: Boolean(MMC_LEV_STRUCTURE),
      render: () => renderMmcLevTab(mmcLev, tabSummaries["mmc-lev"])
    },
    {
      id: "mmc-emp",
      title: "MMC Empuje/Arrastre",
      present: factorPresent.mmcEmp,
      available: Boolean(MMC_EMP_STRUCTURE),
      render: () => renderMmcEmpTab(mmcEmp, tabSummaries["mmc-emp"])
    }
  ];

  for(const section of sections){
    if(section.present){
      const content = section.render();
      if(content){
        tabs.push({ id: section.id, title: section.title, content });
      }
    }else if(section.available){
      disabledTabs.push(section.title);
    }
  }

  const disabledBadgesList = disabledTabs.map((title) => `<span class="badge bg-secondary me-1">${escapeHtml(title)}</span>`).join(" ");

  let tabsHtml = "";
  if(!tabs.length){
    if(disabledTabs.length){
      const disabledNav = disabledTabs.map((title) => `
        <li class="nav-item" role="presentation">
          <button class="nav-link disabled" type="button" tabindex="-1" aria-disabled="true">
            ${escapeHtml(title)} <span class="badge bg-secondary ms-2">Inactiva</span>
          </button>
        </li>
      `).join("");
      tabsHtml = `
        <div class="detail-card detail-tabs-card">
          <h6 class="section-title mb-3">Identificaciones avanzadas</h6>
          <ul class="nav nav-tabs detail-tabs" role="tablist">
            ${disabledNav}
          </ul>
          <div class="alert alert-light border text-muted mt-3">
            <i class="bi bi-info-circle"></i> Factores inactivos: ${disabledBadgesList}. No se encuentran presentes en la identificación inicial.
          </div>
        </div>
      `;
    }else{
      tabsHtml = `<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No hay información avanzada disponible para esta tarea.</div>`;
    }
  }else if(tabs.length === 1 && !disabledTabs.length){
    const tab = tabs[0];
    tabsHtml = `
      <div class="detail-card detail-tabs-card">
        <h6 class="section-title mb-3">${escapeHtml(tab.title)}</h6>
        ${tab.content}
      </div>
    `;
  }else{
    const navActive = tabs.map((tab, idx) => `
      <li class="nav-item" role="presentation">
        <button class="nav-link${idx===0 ? " active" : ""}" id="detail-tab-${tab.id}" data-bs-toggle="tab" data-bs-target="#detail-pane-${tab.id}" type="button" role="tab" aria-controls="detail-pane-${tab.id}" aria-selected="${idx===0 ? "true" : "false"}">
          ${escapeHtml(tab.title)}
        </button>
      </li>
    `).join("");
    const disabledNav = disabledTabs.map((title) => `
      <li class="nav-item" role="presentation">
        <button class="nav-link disabled" type="button" tabindex="-1" aria-disabled="true">
          ${escapeHtml(title)} <span class="badge bg-secondary ms-2">Inactiva</span>
        </button>
      </li>
    `).join("");
    const panes = tabs.map((tab, idx) => `
      <div class="tab-pane fade${idx===0 ? " show active" : ""}" id="detail-pane-${tab.id}" role="tabpanel" aria-labelledby="detail-tab-${tab.id}">
        ${tab.content}
      </div>
    `).join("");
    const disabledNotice = disabledTabs.length
      ? `<div class="alert alert-light border text-muted mt-3"><i class="bi bi-info-circle"></i> Factores inactivos: ${disabledBadgesList}. No se encuentran presentes en la identificación inicial.</div>`
      : "";
    tabsHtml = `
      <div class="detail-card detail-tabs-card">
        <ul class="nav nav-tabs detail-tabs" role="tablist">
          ${navActive}${disabledNav}
        </ul>
        <div class="tab-content">
          ${panes}
        </div>
        ${disabledNotice}
      </div>
    `;
  }

  el("detailBody").innerHTML = `${header}${tabsHtml}`;
  el("detailTitle").textContent = `Detalle · Identificación avanzada`;
  const modal = bootstrap.Modal.getOrCreateInstance('#detailModal');
  modal.show();
}

/* ======= Utils ======= */
function escapeCSV(str){ return `"${String(str??"").replace(/"/g,'""')}"`; }

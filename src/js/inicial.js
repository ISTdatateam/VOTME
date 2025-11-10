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
  headers.forEach((h,i)=>{ o[h||`Col${i+1}`] = row[i]??""; });
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

/* ======= HTML Tarjeta + Modal ======= */
function cardHtml(r, idx){
  const mov = getMovRepFor(r);
  const status = classifyMovRep(mov?.P, mov?.W);
  const criticalTone = classifyCriticalState(mov?.W);
  const criticalClass = criticalTone ? ` critical-pill-${criticalTone}` : "";
  const criticalHtml = mov
    ? `<span class="pill critical-pill${criticalClass}"><strong>Condición Crítica:</strong> ${escapeHtml(mov.W ?? "")}</span>`
    : `<span class="pill">Hoja Mov. repetitivo: sin coincidencia</span>`;

  return `
    <div class="col" data-idx="${idx}">
      <div class="card card-ficha h-100 shadow-sm">
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

          <hr>

          <!-- Estado P/W -->
          <div class="d-flex flex-wrap align-items-center gap-2 mt-2 mb-2">
            <div class="w-100 small text-muted">
              IDENTIFICACIÓN AVANZADA – Trabajo Repetitivo de Miembros Superiores
            </div>
            <span class="status-pill ${status.cls}" title="Estado según hoja Movimiento repetitivo (P/W)">
              <i class="bi bi-activity"></i> Condición Aceptable: ${status.label}
            </span>
            ${criticalHtml}
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

function openDetail(r){
  const mov = getMovRepFor(r);
  const status = classifyMovRep(mov?.P, mov?.W);

  const pText = String(mov?.P ?? "").trim();
  const wText = String(mov?.W ?? "").trim();
  const pBad  = /no acept|criti|alto/i.test(pText);
  const wBad  = /no acept|criti|alto/i.test(wText);



  const statesBlock = `
    <div class="row g-3 mb-3">

    </div>
  `;

  const header = `
    <div class="detail-card mb-3">
      <div class="d-flex flex-wrap justify-content-between align-items-start gap-2">
        <div>
          <div class="small text-muted">Tarea</div>
          <h5 class="mb-1">${escapeHtml(r.D || "-")}</h5>
          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
          <div class="mb-1"><i class="bi bi-geo-alt"></i> <strong>Área:</strong> ${escapeHtml(r.B || "-")}</div>
        </div>
        <div class="d-flex flex-column align-items-end gap-2">
          <span class="status-pill ${status.cls}" style="font-size:1rem;">
            <i class="bi bi-activity"></i> ${status.label}
          </span>
        </div>
      </div>
      <div class="mt-3">
        <div class="small text-muted mb-1"><i class="bi bi-exclamation-octagon"></i> Factores</div>
        <div class="factors-wrap">${factorChips(r)}</div>
      </div>
      ${statesBlock}
    </div>
  `;

  // Tabla de preguntas/respuestas de la fila “Mov. Repetitivo”
  let qa = "";
  if(mov && (mov.rowArr || mov.rowObj)){
    const rows = [];
    if (mov.rowArr && Array.isArray(MOVREP_HEADERS) && MOVREP_HEADERS.length){
      for(let i=0;i<MOVREP_HEADERS.length;i++){
        if(SKIP_IDX.has(i)) continue;                                // ignora Col2..Col9
        const label = MOVREP_HEADERS[i] || `Col${i+1}`;
        if(SKIP_LABELS.has(toLowerNoAccents(label))) continue;       // ignora por título (ej. Mujeres)
        const val = mov.rowArr[i];
        if(String(val ?? "").trim() === "") continue;
        rows.push([label, val]);
      }
    }else{
      for(const [k,v] of Object.entries(mov.rowObj)){
        if(String(v ?? "").trim() === "") continue;
        if(SKIP_LABELS.has(toLowerNoAccents(k))) continue;
        rows.push([k, v]);
      }
    }

    const headHtml = `<thead><tr><th style="min-width:260px">Pregunta</th><th>Respuesta</th></tr></thead>`;
    const bodyHtml = `<tbody>${
      rows.map(([k,v]) => {
        const rowCls = classifyRowHighlight(k, v);
        return `<tr class="${rowCls}"><th>${escapeHtml(k)}</th><td>${escapeHtml(String(v))}</td></tr>`;
      }).join("")
    }</tbody>`;
    qa = `
      <div class="table-like">
        <table>${headHtml}${bodyHtml}</table>
      </div>
    `;
  }else{
    qa = `<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se encontraron detalles coincidentes en la hoja “Movimiento repetitivo”.</div>`;
  }

  el("detailBody").innerHTML = `
    ${header}
    <div class="detail-grid">
      ${qa}
    </div>
  `;
  el("detailTitle").textContent = `Detalle · Movimiento repetitivo`;
  const modal = bootstrap.Modal.getOrCreateInstance('#detailModal');
  modal.show();
}

/* ======= Utils ======= */
function escapeCSV(str){ return `"${String(str??"").replace(/"/g,'""')}"`; }

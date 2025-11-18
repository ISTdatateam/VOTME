
const XLSX_CACHE_KEY = "PORTADA_XLSX_CACHE_V1";

function arrayBufferToBase64(buffer){
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for(let i=0;i<bytes.length;i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

function loadCached(){
  try{
    const raw = localStorage.getItem(XLSX_CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  }catch(e){
    console.warn("No se pudo leer el cache local", e);
    return null;
  }
}

function saveCached(buffer, name){
  const payload = {
    name: name || "Excel cargado",
    ts: Date.now(),
    data: arrayBufferToBase64(buffer)
  };
  localStorage.setItem(XLSX_CACHE_KEY, JSON.stringify(payload));
  return payload;
}

function removeCached(){
  localStorage.removeItem(XLSX_CACHE_KEY);
}

function formatTimestamp(ts){
  try{
    return new Date(ts).toLocaleString();
  }catch(e){
    return "";
  }
}

function updateStatus(statusEl, clearBtn){
  const cache = loadCached();
  if(cache){
    statusEl.innerHTML = `<i class="bi bi-check-circle text-success"></i> Excel guardado: <strong>${cache.name}</strong> (${formatTimestamp(cache.ts)})`;
    clearBtn?.classList.remove("d-none");
  }else{
    statusEl.innerHTML = `<i class="bi bi-exclamation-circle text-muted"></i> Ningún archivo cargado aún.`;
    clearBtn?.classList.add("d-none");
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const y = document.querySelector('[data-year]');
  if(y) y.textContent = new Date().getFullYear();

  const input = document.getElementById('portadaFileInput');
  const status = document.getElementById('uploadStatus');
  const clearBtn = document.getElementById('btnClearExcel');

  if(input && status){
    input.addEventListener('change', (ev) => {
      const file = ev.target.files?.[0];
      if(!file){ return; }
      status.innerHTML = `<i class="bi bi-hourglass-split"></i> Procesando archivo…`;
      const reader = new FileReader();
      reader.onload = (e) => {
        const payload = saveCached(e.target.result, file.name);
        status.innerHTML = `<i class="bi bi-check-circle text-success"></i> Archivo cargado y guardado localmente: <strong>${payload.name}</strong>. Al abrir la matriz se usará este archivo.`;
        clearBtn?.classList.remove("d-none");
      };
      reader.onerror = () => {
        status.innerHTML = `<i class="bi bi-x-circle text-danger"></i> No se pudo leer el archivo. Intente nuevamente.`;
      };
      reader.readAsArrayBuffer(file);
    });
  }

  if(clearBtn){
    clearBtn.addEventListener('click', () => {
      removeCached();
      updateStatus(status, clearBtn);
    });
  }

  if(status){
    updateStatus(status, clearBtn);
  }
});

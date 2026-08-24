import { openFile, saveFile, supportsFileSystemAccess } from '../core/storage.js';

const XLSX_TYPES = [{
  description: 'Libro de Excel',
  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
}];

const ROWS = 30;
const COLS = 12;

let tableEl, statusEl, currentHandle = null, currentName = 'Hoja sin título.xlsx';

export async function mount(root) {
  root.innerHTML = `
    <div class="module module--sheet">
      <div class="toolbar">
        <button id="back-to-launcher-inline" class="btn btn--ghost" title="Volver (Esc)">←</button>
        <span class="toolbar__sep"></span>
        <button id="btn-open" class="btn">Abrir…</button>
        <button id="btn-save" class="btn btn--primary">Guardar</button>
        <span id="status" class="toolbar__status"></span>
      </div>
      <div class="grid-wrap">
        <table id="grid" class="grid"></table>
      </div>
    </div>
  `;

  tableEl = root.querySelector('#grid');
  statusEl = root.querySelector('#status');
  buildEmptyGrid();

  root.querySelector('#back-to-launcher-inline').addEventListener('click', () => {
    document.getElementById('back-to-launcher').click();
  });
  root.querySelector('#btn-open').addEventListener('click', handleOpen);
  root.querySelector('#btn-save').addEventListener('click', handleSave);

  return { unmount: () => { tableEl = null; } };
}

function colLabel(i) {
  let s = '';
  i += 1;
  while (i > 0) {
    const rem = (i - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

function buildEmptyGrid(data = []) {
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.appendChild(document.createElement('th'));
  for (let c = 0; c < COLS; c++) {
    const th = document.createElement('th');
    th.textContent = colLabel(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);

  const tbody = document.createElement('tbody');
  for (let r = 0; r < ROWS; r++) {
    const tr = document.createElement('tr');
    const th = document.createElement('th');
    th.textContent = String(r + 1);
    tr.appendChild(th);
    for (let c = 0; c < COLS; c++) {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      td.textContent = data[r]?.[c] ?? '';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  tableEl.innerHTML = '';
  tableEl.appendChild(thead);
  tableEl.appendChild(tbody);
}

function readGridData() {
  const rows = [...tableEl.querySelectorAll('tbody tr')];
  return rows.map((tr) => [...tr.querySelectorAll('td')].map((td) => td.textContent));
}

async function handleOpen() {
  const file = await openFile(XLSX_TYPES);
  if (!file) return;
  setStatus('Leyendo hoja de cálculo…');
  try {
    const XLSX = await loadXLSX();
    const arrayBuffer = await file.blob.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    buildEmptyGrid(data);
    currentHandle = file.handle;
    currentName = file.name;
    setStatus(`Abierto (hoja "${wb.SheetNames[0]}")`);
  } catch (err) {
    setStatus('No se pudo leer el archivo: ' + err.message);
  }
}

async function handleSave() {
  setStatus('Generando .xlsx…');
  try {
    const XLSX = await loadXLSX();
    const data = readGridData();
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const target = await saveFile({
      blob,
      handle: currentHandle,
      suggestedName: currentName,
      types: XLSX_TYPES,
    });
    if (target) currentHandle = target;
    setStatus(supportsFileSystemAccess ? 'Guardado' : 'Descargado');
  } catch (err) {
    setStatus('No se pudo guardar: ' + err.message);
  }
}

function setStatus(msg) {
  if (statusEl) statusEl.textContent = msg;
}

const scriptCache = new Map();
function loadXLSX() {
  const src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (scriptCache.has(src)) return scriptCache.get(src);
  const p = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = () => resolve(window.XLSX);
    s.onerror = () => reject(new Error('No se pudo cargar la librería de hojas de cálculo'));
    document.head.appendChild(s);
  });
  scriptCache.set(src, p);
  return p;
}

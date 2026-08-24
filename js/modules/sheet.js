import { openFile, saveFile, supportsFileSystemAccess } from '../core/storage.js';

const XLSX_TYPES = [{
  description: 'Libro de Excel',
  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
}];

const ROWS = 30;
const COLS = 12;

// Modelo de datos: guarda lo que el usuario escribió (texto, número o
// fórmula "=..."). Lo que se ve en pantalla (displayValue) se recalcula
// aparte, para no perder la fórmula original al reabrir la celda.
let rawData = [];
let tableEl, statusEl, currentHandle = null, currentName = 'Hoja sin título.xlsx';
let activeCell = null; // {r, c} de la celda con foco actualmente

export async function mount(root) {
  root.innerHTML = `
    <div class="module module--sheet">
      <div class="toolbar">
        <button id="back-to-launcher-inline" class="btn btn--ghost" title="Volver (Esc)">←</button>
        <span class="toolbar__sep"></span>
        <button id="btn-open" class="btn">Abrir…</button>
        <button id="btn-save" class="btn btn--primary">Guardar</button>
        <span class="toolbar__sep"></span>
        <span class="toolbar__hint">Fórmulas: =SUMA(A1:A5)  =PROMEDIO(B1:B5)  =A1+B2*3</span>
        <span id="status" class="toolbar__status"></span>
      </div>
      <div class="grid-wrap">
        <table id="grid" class="grid"></table>
      </div>
    </div>
  `;

  tableEl = root.querySelector('#grid');
  statusEl = root.querySelector('#status');
  rawData = emptyModel();
  buildGrid();

  root.querySelector('#back-to-launcher-inline').addEventListener('click', () => {
    document.getElementById('back-to-launcher').click();
  });
  root.querySelector('#btn-open').addEventListener('click', handleOpen);
  root.querySelector('#btn-save').addEventListener('click', handleSave);

  return { unmount: () => { tableEl = null; } };
}

function emptyModel() {
  return Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => ''));
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

function buildGrid() {
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
    const rowHead = document.createElement('th');
    rowHead.textContent = String(r + 1);
    tr.appendChild(rowHead);
    for (let c = 0; c < COLS; c++) {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      td.dataset.row = String(r);
      td.dataset.col = String(c);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  tableEl.innerHTML = '';
  tableEl.appendChild(thead);
  tableEl.appendChild(tbody);

  tableEl.addEventListener('focusin', onCellFocus);
  tableEl.addEventListener('focusout', onCellBlur);
  tableEl.addEventListener('keydown', onCellKeydown);

  recalcAndRender();
}

function cellAt(r, c) {
  return tableEl.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
}

// Al enfocar una celda, mostramos el contenido "crudo" (la fórmula tal cual
// se escribió), como en cualquier hoja de cálculo real.
function onCellFocus(e) {
  const td = e.target.closest('td');
  if (!td) return;
  const r = Number(td.dataset.row), c = Number(td.dataset.col);
  activeCell = { r, c };
  td.textContent = rawData[r][c];
}

function onCellBlur(e) {
  const td = e.target.closest('td');
  if (!td) return;
  const r = Number(td.dataset.row), c = Number(td.dataset.col);
  rawData[r][c] = td.textContent.trim();
  activeCell = null;
  recalcAndRender();
}

function onCellKeydown(e) {
  if (e.key !== 'Enter' && e.key !== 'Tab') return;
  const td = e.target.closest('td');
  if (!td) return;
  const r = Number(td.dataset.row), c = Number(td.dataset.col);
  let next = null;
  if (e.key === 'Enter') next = cellAt(Math.min(ROWS - 1, r + 1), c);
  else if (e.key === 'Tab') next = cellAt(r, Math.min(COLS - 1, c + (e.shiftKey ? -1 : 1)));
  if (next) {
    e.preventDefault();
    next.focus();
    placeCaretAtEnd(next);
  }
}

function placeCaretAtEnd(el) {
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(false);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

/** Recalcula todas las fórmulas y pinta los valores (excepto en la celda con foco). */
function recalcAndRender() {
  const cache = new Map();
  for (let r = 0; r < ROWS; r++) {
    for (let c = 0; c < COLS; c++) {
      if (activeCell && activeCell.r === r && activeCell.c === c) continue;
      const td = cellAt(r, c);
      td.textContent = displayValue(r, c, cache);
    }
  }
}

function displayValue(r, c, cache) {
  const value = evaluateCell(r, c, cache, new Set());
  return value === '' || value === undefined ? '' : String(value);
}

// ---------- Motor de fórmulas ----------
// Soporta: números y texto normales, referencias tipo A1, rangos A1:B3,
// SUMA()/SUM(), PROMEDIO()/AVERAGE(), y aritmética básica + - * / ().

function evaluateCell(r, c, cache, visiting) {
  const key = `${r},${c}`;
  if (cache.has(key)) return cache.get(key);
  if (r < 0 || r >= ROWS || c < 0 || c >= COLS) return 0;
  const raw = (rawData[r]?.[c] ?? '').trim();
  if (!raw.startsWith('=')) {
    cache.set(key, raw);
    return raw;
  }
  if (visiting.has(key)) return '#CIRC!';
  visiting.add(key);

  let result;
  try {
    let expr = raw.slice(1).toUpperCase();

    expr = expr.replace(/(SUMA|SUM|PROMEDIO|AVERAGE)\(([^()]*)\)/g, (_m, fn, argsStr) => {
      const values = collectValues(argsStr, cache, visiting);
      if (fn === 'SUMA' || fn === 'SUM') return String(values.reduce((a, b) => a + b, 0));
      const avg = values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
      return String(avg);
    });

    expr = expr.replace(/[A-Z]+\d+/g, (ref) => {
      const idx = cellRefToIndex(ref);
      if (!idx) return '0';
      const val = evaluateCell(idx.row, idx.col, cache, visiting);
      const num = parseFloat(val);
      return String(Number.isNaN(num) ? 0 : num);
    });

    result = safeEvalArithmetic(expr);
  } catch {
    result = '#ERROR!';
  }

  visiting.delete(key);
  cache.set(key, result);
  return result;
}

function collectValues(argsStr, cache, visiting) {
  const values = [];
  for (const part of argsStr.split(',')) {
    const token = part.trim();
    if (!token) continue;
    if (token.includes(':')) {
      const [fromRef, toRef] = token.split(':').map((t) => t.trim());
      const from = cellRefToIndex(fromRef), to = cellRefToIndex(toRef);
      if (!from || !to) continue;
      const r1 = Math.min(from.row, to.row), r2 = Math.max(from.row, to.row);
      const c1 = Math.min(from.col, to.col), c2 = Math.max(from.col, to.col);
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const num = parseFloat(evaluateCell(r, c, cache, visiting));
          if (!Number.isNaN(num)) values.push(num);
        }
      }
    } else {
      const idx = cellRefToIndex(token);
      if (idx) {
        const num = parseFloat(evaluateCell(idx.row, idx.col, cache, visiting));
        if (!Number.isNaN(num)) values.push(num);
      } else {
        const num = parseFloat(token);
        if (!Number.isNaN(num)) values.push(num);
      }
    }
  }
  return values;
}

function cellRefToIndex(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref.trim());
  if (!m) return null;
  const [, colLetters, rowStr] = m;
  let col = 0;
  for (const ch of colLetters) col = col * 26 + (ch.charCodeAt(0) - 64);
  col -= 1;
  const row = parseInt(rowStr, 10) - 1;
  if (row < 0 || col < 0) return null;
  return { row, col };
}

function safeEvalArithmetic(expr) {
  const cleaned = expr.trim();
  if (cleaned === '') return '';
  if (!/^[0-9+\-*/().\s]*$/.test(cleaned)) return '#ERROR!';
  // eslint-disable-next-line no-new-func
  const value = Function(`"use strict"; return (${cleaned});`)();
  if (typeof value !== 'number' || !Number.isFinite(value)) return '#ERROR!';
  return Math.round(value * 1e6) / 1e6;
}

// ---------- Abrir / Guardar .xlsx ----------

async function handleOpen() {
  const file = await openFile(XLSX_TYPES);
  if (!file) return;
  setStatus('Leyendo hoja de cálculo…');
  try {
    const XLSX = await loadXLSX();
    const arrayBuffer = await file.blob.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    rawData = emptyModel();
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let r = range.s.r; r <= Math.min(range.e.r, ROWS - 1); r++) {
      for (let c = range.s.c; c <= Math.min(range.e.c, COLS - 1); c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell) continue;
        rawData[r][c] = cell.f ? '=' + cell.f : (cell.v ?? '').toString();
      }
    }
    buildGrid();
    currentHandle = file.handle;
    currentName = file.name;
    setStatus(`Abierto (hoja "${sheetName}")`);
  } catch (err) {
    setStatus('No se pudo leer el archivo: ' + err.message);
  }
}

async function handleSave() {
  setStatus('Generando .xlsx…');
  try {
    const XLSX = await loadXLSX();
    const ws = {};
    let maxR = 0, maxC = 0;
    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        const raw = (rawData[r][c] ?? '').trim();
        if (!raw) continue;
        maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
        const addr = XLSX.utils.encode_cell({ r, c });
        if (raw.startsWith('=')) {
          const computed = evaluateCell(r, c, new Map(), new Set());
          const num = parseFloat(computed);
          ws[addr] = Number.isNaN(num)
            ? { t: 's', v: String(computed), f: raw.slice(1) }
            : { t: 'n', v: num, f: raw.slice(1) };
        } else {
          const num = parseFloat(raw);
          ws[addr] = (!Number.isNaN(num) && String(num) === raw)
            ? { t: 'n', v: num }
            : { t: 's', v: raw };
        }
      }
    }
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
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
    setStatus(supportsFileSystemAccess ? 'Guardado (con fórmulas)' : 'Descargado (con fórmulas)');
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

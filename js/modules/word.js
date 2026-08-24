import { openFile, saveFile, supportsFileSystemAccess } from '../core/storage.js';

const DOCX_TYPES = [{
  description: 'Documento Word',
  accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] },
}];

const AUTOSAVE_KEY = 'carpeta:word:autosave:v1';

let editor, statusEl, currentHandle = null, currentName = 'Documento sin título.docx';

export async function mount(root) {
  root.innerHTML = `
    <div class="module module--word">
      <div class="toolbar">
        <button id="back-to-launcher-inline" class="btn btn--ghost" title="Volver (Esc)">←</button>
        <span class="toolbar__sep"></span>
        <button data-cmd="bold" class="btn"><b>N</b></button>
        <button data-cmd="italic" class="btn"><i>K</i></button>
        <button data-cmd="underline" class="btn"><u>S</u></button>
        <span class="toolbar__sep"></span>
        <select id="block-style" class="btn">
          <option value="p">Párrafo</option>
          <option value="h1">Título 1</option>
          <option value="h2">Título 2</option>
        </select>
        <button data-cmd="insertUnorderedList" class="btn">• Lista</button>
        <button data-cmd="insertOrderedList" class="btn">1. Lista</button>
        <span class="toolbar__sep"></span>
        <button id="btn-image" class="btn">🖼 Imagen</button>
        <button id="btn-table" class="btn">▦ Tabla</button>
        <input id="image-input" type="file" accept="image/png,image/jpeg,image/gif" hidden>
        <span class="toolbar__sep"></span>
        <button id="btn-new" class="btn">Nuevo</button>
        <button id="btn-open" class="btn">Abrir…</button>
        <button id="btn-save" class="btn btn--primary">Guardar</button>
        <span id="status" class="toolbar__status"></span>
      </div>
      <div class="page-wrap">
        <div id="editor" class="page" contenteditable="true" spellcheck="true">
          <p>Empieza a escribir aquí…</p>
        </div>
      </div>
      <dialog id="table-dialog" class="dialog">
        <form method="dialog" class="dialog__form">
          <h3>Insertar tabla</h3>
          <label>Filas <input id="table-rows" type="number" min="1" max="30" value="3"></label>
          <label>Columnas <input id="table-cols" type="number" min="1" max="10" value="3"></label>
          <div class="dialog__actions">
            <button value="cancel" class="btn">Cancelar</button>
            <button id="table-confirm" value="ok" class="btn btn--primary">Insertar</button>
          </div>
        </form>
      </dialog>
    </div>
  `;

  editor = root.querySelector('#editor');
  statusEl = root.querySelector('#status');

  root.querySelector('#back-to-launcher-inline').addEventListener('click', () => {
    document.getElementById('back-to-launcher').click();
  });

  root.querySelectorAll('[data-cmd]').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.execCommand(btn.dataset.cmd, false, null);
      editor.focus();
    });
  });

  root.querySelector('#block-style').addEventListener('change', (e) => {
    document.execCommand('formatBlock', false, e.target.value);
    editor.focus();
  });

  root.querySelector('#btn-open').addEventListener('click', handleOpen);
  root.querySelector('#btn-save').addEventListener('click', handleSave);
  root.querySelector('#btn-new').addEventListener('click', handleNew);

  const imageInput = root.querySelector('#image-input');
  root.querySelector('#btn-image').addEventListener('click', () => {
    saveSelection();
    imageInput.click();
  });
  imageInput.addEventListener('change', handleInsertImage);

  const tableDialog = root.querySelector('#table-dialog');
  root.querySelector('#btn-table').addEventListener('click', () => {
    saveSelection();
    tableDialog.showModal();
  });
  root.querySelector('#table-confirm').addEventListener('click', (e) => {
    e.preventDefault();
    const rows = clampInt(root.querySelector('#table-rows').value, 1, 30, 3);
    const cols = clampInt(root.querySelector('#table-cols').value, 1, 10, 3);
    tableDialog.close();
    insertTable(rows, cols);
  });

  editor.addEventListener('keyup', saveSelection);
  editor.addEventListener('mouseup', saveSelection);
  editor.addEventListener('input', scheduleAutosave);

  restoreAutosave();

  return { unmount: () => { clearTimeout(autosaveTimer); editor = null; } };
}

function clampInt(value, min, max, fallback) {
  const n = parseInt(value, 10);
  if (Number.isNaN(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

// ---------- Autoguardado local (formato propio: el HTML tal cual) ----------
// Objetivo: no perder el trabajo si se cierra la pestaña antes de exportar a
// .docx. No sustituye a "Guardar" — es solo una red de seguridad local.
let autosaveTimer = null;
function scheduleAutosave() {
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    try {
      localStorage.setItem(AUTOSAVE_KEY, editor.innerHTML);
    } catch {
      // Almacenamiento lleno o no disponible (modo privado): no es crítico.
    }
  }, 800);
}

function restoreAutosave() {
  let saved = null;
  try {
    saved = localStorage.getItem(AUTOSAVE_KEY);
  } catch {
    return;
  }
  if (saved && saved.trim() && saved !== '<p>Empieza a escribir aquí…</p>') {
    editor.innerHTML = saved;
    setStatus('Recuperado de tu último autoguardado local');
  }
}

function handleNew() {
  if (!confirm('¿Empezar un documento nuevo? Se perderá lo que no hayas exportado a .docx.')) return;
  editor.innerHTML = '<p>Empieza a escribir aquí…</p>';
  currentHandle = null;
  currentName = 'Documento sin título.docx';
  try { localStorage.removeItem(AUTOSAVE_KEY); } catch { /* no crítico */ }
  setStatus('Documento nuevo');
}

// Los botones de la toolbar quitan el foco del editor; guardamos dónde
// estaba el cursor para poder insertar imágenes/tablas justo ahí.
let savedRange = null;
function saveSelection() {
  const sel = window.getSelection();
  if (sel.rangeCount && editor.contains(sel.anchorNode)) {
    savedRange = sel.getRangeAt(0).cloneRange();
  }
}

function insertAtSavedRange(node) {
  editor.focus();
  const sel = window.getSelection();
  sel.removeAllRanges();
  if (savedRange) sel.addRange(savedRange);

  if (savedRange) {
    savedRange.collapse(false);
    savedRange.insertNode(node);
    savedRange.setStartAfter(node);
    savedRange.collapse(true);
    sel.removeAllRanges();
    sel.addRange(savedRange);
  } else {
    editor.appendChild(node);
  }
  scheduleAutosave();
}

async function handleInsertImage(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  const img = document.createElement('img');
  img.src = dataUrl;
  img.alt = file.name;
  const p = document.createElement('p');
  p.appendChild(img);
  insertAtSavedRange(p);
  setStatus('Imagen insertada');
}

function insertTable(rows, cols) {
  const table = document.createElement('table');
  table.className = 'doc-table';
  const tbody = document.createElement('tbody');
  for (let r = 0; r < rows; r++) {
    const tr = document.createElement('tr');
    for (let c = 0; c < cols; c++) {
      const td = document.createElement('td');
      td.innerHTML = '<br>';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  const spacer = document.createElement('p');
  spacer.innerHTML = '<br>';
  insertAtSavedRange(table);
  table.after(spacer);
  setStatus('Tabla insertada');
}

async function handleOpen() {
  const file = await openFile(DOCX_TYPES);
  if (!file) return;
  setStatus('Convirtiendo documento…');
  try {
    // mammoth se carga solo cuando el usuario realmente abre un .docx.
    const mammoth = await loadMammoth();
    const arrayBuffer = await file.blob.arrayBuffer();
    const { value: html, messages } = await mammoth.convertToHtml({ arrayBuffer });
    editor.innerHTML = html || '<p></p>';
    currentHandle = file.handle;
    currentName = file.name;
    scheduleAutosave();
    setStatus(messages.length ? 'Abierto (algunos estilos avanzados se simplificaron)' : 'Abierto');
  } catch (err) {
    setStatus('No se pudo convertir el archivo: ' + err.message);
  }
}

async function handleSave() {
  setStatus('Generando .docx…');
  try {
    const docx = await loadDocx();
    const blocks = htmlToDocxBlocks(editor, docx);
    const doc = new docx.Document({ sections: [{ children: blocks }] });
    const blob = await docx.Packer.toBlob(doc);
    const target = await saveFile({
      blob,
      handle: currentHandle,
      suggestedName: currentName,
      types: DOCX_TYPES,
    });
    if (target) currentHandle = target;
    setStatus(supportsFileSystemAccess ? 'Guardado' : 'Descargado');
  } catch (err) {
    setStatus('No se pudo guardar: ' + err.message);
  }
}

// Ancho útil de la página (720px máx. de .page, menos 72px de margen a cada lado).
const PAGE_CONTENT_WIDTH = 576;

/** Traducción del HTML del editor a bloques (párrafos/tablas) de la librería docx. */
function htmlToDocxBlocks(container, docx) {
  const blocks = [];
  for (const node of container.children) {
    if (node.tagName === 'H1') {
      blocks.push(new docx.Paragraph({ children: runsFromInline(node, docx), heading: docx.HeadingLevel.HEADING_1 }));
    } else if (node.tagName === 'H2') {
      blocks.push(new docx.Paragraph({ children: runsFromInline(node, docx), heading: docx.HeadingLevel.HEADING_2 }));
    } else if (node.tagName === 'LI') {
      blocks.push(new docx.Paragraph({ children: runsFromInline(node, docx), bullet: { level: 0 } }));
    } else if (node.tagName === 'UL' || node.tagName === 'OL') {
      for (const li of node.children) {
        blocks.push(new docx.Paragraph({ children: runsFromInline(li, docx), bullet: { level: 0 } }));
      }
    } else if (node.tagName === 'IMG') {
      blocks.push(new docx.Paragraph({ children: [imageNodeToRun(node, docx)].filter(Boolean) }));
    } else if (node.tagName === 'TABLE') {
      blocks.push(tableNodeToDocxTable(node, docx));
    } else {
      blocks.push(new docx.Paragraph({ children: runsFromInline(node, docx) }));
    }
  }
  return blocks.length ? blocks : [new docx.Paragraph({ text: '' })];
}

function tableNodeToDocxTable(tableEl, docx) {
  const rows = [];
  for (const tr of tableEl.querySelectorAll('tr')) {
    const cells = [];
    for (const cell of tr.children) {
      cells.push(new docx.TableCell({
        children: [new docx.Paragraph({ children: runsFromInline(cell, docx) })],
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
      }));
    }
    rows.push(new docx.TableRow({ children: cells }));
  }
  return new docx.Table({
    rows,
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
  });
}

/**
 * Recorre los hijos de un elemento (párrafo, título, celda, li…) y produce
 * Runs de docx conservando negrita/cursiva/subrayado anidados y las
 * imágenes que encuentre en el camino.
 */
function runsFromInline(node, docx, style = {}) {
  const runs = [];
  for (const child of node.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      if (child.textContent) runs.push(makeTextRun(child.textContent, style, docx));
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      if (child.tagName === 'IMG') {
        const run = imageNodeToRun(child, docx);
        if (run) runs.push(run);
      } else if (child.tagName === 'BR') {
        runs.push(new docx.TextRun({ text: '', break: 1 }));
      } else {
        const nextStyle = { ...style, ...tagToStyle(child.tagName) };
        runs.push(...runsFromInline(child, docx, nextStyle));
      }
    }
  }
  return runs.length ? runs : [makeTextRun('', style, docx)];
}

function tagToStyle(tagName) {
  switch (tagName) {
    case 'STRONG': case 'B': return { bold: true };
    case 'EM': case 'I': return { italics: true };
    case 'U': return { underline: true };
    default: return {};
  }
}

function makeTextRun(text, style, docx) {
  return new docx.TextRun({
    text,
    bold: !!style.bold,
    italics: !!style.italics,
    underline: style.underline ? {} : undefined,
  });
}

function imageNodeToRun(imgEl, docx) {
  const src = imgEl.getAttribute('src') || '';
  if (!src.startsWith('data:image/')) return null; // imágenes remotas no soportadas offline
  try {
    const type = imageMimeToDocxType(src);
    const data = dataUriToUint8Array(src);
    const { width, height } = scaledImageSize(imgEl);
    return new docx.ImageRun({ type, data, transformation: { width, height } });
  } catch {
    return null;
  }
}

function imageMimeToDocxType(dataUri) {
  const match = /^data:image\/(png|jpe?g|gif|bmp);/i.exec(dataUri);
  const ext = (match ? match[1] : 'png').toLowerCase();
  return ext === 'jpg' ? 'jpeg' : ext;
}

function dataUriToUint8Array(dataUri) {
  const base64 = dataUri.slice(dataUri.indexOf(',') + 1);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** Escala la imagen al ancho de la página si es más ancha, conservando proporción. */
function scaledImageSize(imgEl) {
  const naturalW = imgEl.naturalWidth || 400;
  const naturalH = imgEl.naturalHeight || 300;
  if (naturalW <= PAGE_CONTENT_WIDTH) return { width: naturalW, height: naturalH };
  const ratio = PAGE_CONTENT_WIDTH / naturalW;
  return { width: PAGE_CONTENT_WIDTH, height: Math.round(naturalH * ratio) };
}

function setStatus(msg) {
  if (statusEl) statusEl.textContent = msg;
}

// Las librerías pesadas se traen por CDN solo la primera vez que se usan,
// y el navegador las cachea después (ideal para equipos con poca RAM/disco).
function loadMammoth() {
  return loadUMD('https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js', () => window.mammoth);
}
function loadDocx() {
  return loadUMD('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.js', () => window.docx);
}

const scriptCache = new Map();
function loadUMD(src, getGlobal) {
  if (getGlobal()) return Promise.resolve(getGlobal());
  if (scriptCache.has(src)) return scriptCache.get(src);
  const p = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = () => resolve(getGlobal());
    s.onerror = () => reject(new Error('No se pudo cargar ' + src));
    document.head.appendChild(s);
  });
  scriptCache.set(src, p);
  return p;
}

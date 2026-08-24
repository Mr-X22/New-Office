import { openFile, saveFile, supportsFileSystemAccess } from '../core/storage.js';

const DOCX_TYPES = [{
  description: 'Documento Word',
  accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] },
}];

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
        <button id="btn-open" class="btn">Abrir…</button>
        <button id="btn-save" class="btn btn--primary">Guardar</button>
        <span id="status" class="toolbar__status"></span>
      </div>
      <div class="page-wrap">
        <div id="editor" class="page" contenteditable="true" spellcheck="true">
          <p>Empieza a escribir aquí…</p>
        </div>
      </div>
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

  return { unmount: () => { editor = null; } };
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

/** Traducción mínima de HTML del editor a párrafos de la librería docx. */
function htmlToDocxBlocks(container, docx) {
  const blocks = [];
  for (const node of container.children) {
    const text = node.textContent || '';
    if (node.tagName === 'H1') {
      blocks.push(new docx.Paragraph({ text, heading: docx.HeadingLevel.HEADING_1 }));
    } else if (node.tagName === 'H2') {
      blocks.push(new docx.Paragraph({ text, heading: docx.HeadingLevel.HEADING_2 }));
    } else if (node.tagName === 'LI') {
      blocks.push(new docx.Paragraph({ text, bullet: { level: 0 } }));
    } else if (node.tagName === 'UL' || node.tagName === 'OL') {
      for (const li of node.children) {
        blocks.push(new docx.Paragraph({ text: li.textContent || '', bullet: { level: 0 } }));
      }
    } else {
      blocks.push(new docx.Paragraph({ text }));
    }
  }
  return blocks.length ? blocks : [new docx.Paragraph({ text: '' })];
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

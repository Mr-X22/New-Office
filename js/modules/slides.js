import { saveFile, supportsFileSystemAccess } from '../core/storage.js';

const PPTX_TYPES = [{
  description: 'Presentación PowerPoint',
  accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
}];

let slides = [{ title: 'Título de la diapositiva', body: 'Texto de apoyo…' }];
let activeIndex = 0;
let root, statusEl;

export async function mount(container) {
  root = container;
  root.innerHTML = `
    <div class="module module--slides">
      <div class="toolbar">
        <button id="back-to-launcher-inline" class="btn btn--ghost" title="Volver (Esc)">←</button>
        <span class="toolbar__sep"></span>
        <button id="btn-add" class="btn">+ Diapositiva</button>
        <button id="btn-del" class="btn">Eliminar</button>
        <span class="toolbar__sep"></span>
        <button id="btn-save" class="btn btn--primary">Exportar .pptx</button>
        <span id="status" class="toolbar__status"></span>
      </div>
      <div class="slides-layout">
        <div id="slide-list" class="slide-list"></div>
        <div class="slide-canvas-wrap">
          <div class="slide-canvas">
            <div id="slide-title" class="slide-canvas__title" contenteditable="true"></div>
            <div id="slide-body" class="slide-canvas__body" contenteditable="true"></div>
          </div>
        </div>
      </div>
    </div>
  `;

  statusEl = root.querySelector('#status');

  root.querySelector('#back-to-launcher-inline').addEventListener('click', () => {
    document.getElementById('back-to-launcher').click();
  });
  root.querySelector('#btn-add').addEventListener('click', addSlide);
  root.querySelector('#btn-del').addEventListener('click', deleteSlide);
  root.querySelector('#btn-save').addEventListener('click', handleSave);
  root.querySelector('#slide-title').addEventListener('input', syncActiveSlide);
  root.querySelector('#slide-body').addEventListener('input', syncActiveSlide);

  renderList();
  renderCanvas();

  return { unmount: () => { root = null; } };
}

function renderList() {
  const list = root.querySelector('#slide-list');
  list.innerHTML = '';
  slides.forEach((s, i) => {
    const item = document.createElement('button');
    item.className = 'slide-thumb' + (i === activeIndex ? ' slide-thumb--active' : '');
    item.textContent = `${i + 1}. ${s.title || 'Sin título'}`;
    item.addEventListener('click', () => {
      activeIndex = i;
      renderList();
      renderCanvas();
    });
    list.appendChild(item);
  });
}

function renderCanvas() {
  const s = slides[activeIndex];
  root.querySelector('#slide-title').textContent = s.title;
  root.querySelector('#slide-body').textContent = s.body;
}

function syncActiveSlide() {
  slides[activeIndex] = {
    title: root.querySelector('#slide-title').textContent,
    body: root.querySelector('#slide-body').textContent,
  };
  renderList();
}

function addSlide() {
  slides.push({ title: 'Nueva diapositiva', body: '' });
  activeIndex = slides.length - 1;
  renderList();
  renderCanvas();
}

function deleteSlide() {
  if (slides.length === 1) return;
  slides.splice(activeIndex, 1);
  activeIndex = Math.max(0, activeIndex - 1);
  renderList();
  renderCanvas();
}

async function handleSave() {
  setStatus('Generando .pptx…');
  try {
    const PptxGenJS = await loadPptxGen();
    const pres = new PptxGenJS();
    for (const s of slides) {
      const slide = pres.addSlide();
      slide.addText(s.title, { x: 0.5, y: 0.4, w: 9, h: 1, fontSize: 28, bold: true });
      slide.addText(s.body, { x: 0.5, y: 1.6, w: 9, h: 4, fontSize: 18 });
    }
    const blob = await pres.write({ outputType: 'blob' });
    const target = await saveFile({
      blob,
      handle: null,
      suggestedName: 'Presentación sin título.pptx',
      types: PPTX_TYPES,
    });
    setStatus(supportsFileSystemAccess ? 'Guardado' : 'Descargado');
  } catch (err) {
    setStatus('No se pudo exportar: ' + err.message);
  }
}

function setStatus(msg) {
  if (statusEl) statusEl.textContent = msg;
}

const scriptCache = new Map();
function loadPptxGen() {
  const src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
  if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
  if (scriptCache.has(src)) return scriptCache.get(src);
  const p = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = () => resolve(window.PptxGenJS);
    s.onerror = () => reject(new Error('No se pudo cargar la librería de presentaciones'));
    document.head.appendChild(s);
  });
  scriptCache.set(src, p);
  return p;
}

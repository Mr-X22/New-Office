/**
 * app-shell.js — Orquesta qué módulo está activo y lo carga bajo demanda.
 *
 * Clave para equipos con pocos recursos: ninguna librería pesada (docx, xlsx,
 * pptx) se descarga hasta que el usuario realmente abre ese módulo o hace un
 * import/export que la necesite. El shell en sí (esta pantalla) no depende de
 * ninguna de ellas.
 */

const MODULES = {
  word: {
    label: 'Documentos',
    color: '#2B4C7E',
    loader: () => import('../modules/word.js'),
  },
  sheet: {
    label: 'Hojas de cálculo',
    color: '#2F6E4E',
    loader: () => import('../modules/sheet.js'),
  },
  slides: {
    label: 'Presentaciones',
    color: '#B0552E',
    loader: () => import('../modules/slides.js'),
  },
};

const root = document.getElementById('app-root');
const launcher = document.getElementById('launcher');
let activeModule = null;

export function initShell() {
  document.querySelectorAll('[data-open-module]').forEach((btn) => {
    btn.addEventListener('click', () => openModule(btn.dataset.openModule));
  });
  document.getElementById('back-to-launcher').addEventListener('click', closeModule);
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && activeModule) closeModule();
  });
}

async function openModule(key) {
  const def = MODULES[key];
  if (!def) return;

  launcher.hidden = true;
  root.hidden = false;
  root.setAttribute('aria-busy', 'true');
  root.innerHTML = `<div class="loading">Cargando ${def.label.toLowerCase()}…</div>`;
  document.documentElement.style.setProperty('--module-accent', def.color);

  try {
    const mod = await def.loader();
    root.innerHTML = '';
    activeModule = await mod.mount(root);
  } catch (err) {
    root.innerHTML = `<div class="loading loading--error">
      No se pudo cargar este módulo. Revisa tu conexión si es la primera vez que lo abres.
      <br><small>${err.message}</small>
    </div>`;
  } finally {
    root.removeAttribute('aria-busy');
  }
}

function closeModule() {
  if (activeModule && typeof activeModule.unmount === 'function') {
    activeModule.unmount();
  }
  activeModule = null;
  root.hidden = true;
  launcher.hidden = false;
  root.innerHTML = '';
}

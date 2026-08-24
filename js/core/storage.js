/**
 * storage.js — Núcleo de acceso a archivos, compartido por los tres módulos.
 *
 * Estrategia pensada para equipos con pocos recursos:
 *  - Si el navegador soporta File System Access API, se usa directamente
 *    (edición "in place" del archivo real del usuario, sin duplicar datos en RAM).
 *  - Si no (navegadores viejos, Chromebooks limitados), se cae a un flujo
 *    clásico de <input type=file> para abrir y descarga de blob para guardar.
 *  - IndexedDB guarda solo metadatos + "recientes", nunca el archivo completo,
 *    para no inflar memoria en equipos limitados.
 */

const DB_NAME = 'carpeta-suite';
const DB_VERSION = 1;
const STORE_RECENTS = 'recientes';

let dbPromise = null;

function openDB() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_RECENTS)) {
        db.createObjectStore(STORE_RECENTS, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

export const supportsFileSystemAccess = 'showOpenFilePicker' in window;

/**
 * Abre un archivo. Devuelve { name, blob, handle } — handle es null si se usó
 * el flujo de <input type=file>, en cuyo caso "guardar" siempre pedirá ruta.
 */
export async function openFile(acceptTypes) {
  if (supportsFileSystemAccess) {
    try {
      const [handle] = await window.showOpenFilePicker({ types: acceptTypes });
      const blob = await handle.getFile();
      await registerRecent(blob.name, handle);
      return { name: blob.name, blob, handle };
    } catch (err) {
      if (err.name === 'AbortError') return null;
      throw err;
    }
  }
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = acceptTypes.flatMap(t => Object.values(t.accept)).flat().join(',');
    input.onchange = () => {
      const file = input.files[0];
      resolve(file ? { name: file.name, blob: file, handle: null } : null);
    };
    input.click();
  });
}

/**
 * Guarda un blob. Si hay handle previo, escribe directamente (rápido, sin
 * diálogo). Si no, abre el selector de "Guardar como" o dispara descarga.
 */
export async function saveFile({ blob, handle, suggestedName, types }) {
  if (supportsFileSystemAccess) {
    let target = handle;
    if (!target) {
      target = await window.showSaveFilePicker({ suggestedName, types });
    }
    const writable = await target.createWritable();
    await writable.write(blob);
    await writable.close();
    await registerRecent(suggestedName, target);
    return target;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = suggestedName;
  a.click();
  URL.revokeObjectURL(url);
  return null;
}

async function registerRecent(name, handle) {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_RECENTS, 'readwrite');
    tx.objectStore(STORE_RECENTS).put({
      id: name + ':' + Date.now(),
      name,
      handle,
      openedAt: Date.now(),
    });
  } catch {
    // IndexedDB no disponible (modo privado, etc.) — no es crítico, seguimos sin recientes.
  }
}

export async function getRecents(limit = 6) {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_RECENTS, 'readonly');
    const all = await new Promise((resolve, reject) => {
      const req = tx.objectStore(STORE_RECENTS).getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return all.sort((a, b) => b.openedAt - a.openedAt).slice(0, limit);
  } catch {
    return [];
  }
}

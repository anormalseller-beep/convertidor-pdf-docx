// Guarda los archivos convertidos en IndexedDB del navegador (nunca en un servidor),
// para poder volver a descargarlos desde el historial en este mismo navegador.
const BD = "convertidor-pdf";
const TIENDA = "archivos";
const MAX_ARCHIVOS = 20;

interface Entrada {
  id: string;
  nombre: string;
  blob: Blob;
  fecha: number;
}

function abrir(): Promise<IDBDatabase> {
  return new Promise((res, rej) => {
    const r = indexedDB.open(BD, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(TIENDA, { keyPath: "id" });
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}

function tx<T>(modo: IDBTransactionMode, f: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return abrir().then(
    (db) =>
      new Promise<T>((res, rej) => {
        const t = db.transaction(TIENDA, modo);
        const r = f(t.objectStore(TIENDA));
        t.oncomplete = () => { db.close(); res(r.result); };
        t.onerror = () => { db.close(); rej(t.error); };
      }),
  );
}

export async function guardarArchivo(id: string, nombre: string, blob: Blob) {
  try {
    await tx("readwrite", (s) => s.put({ id, nombre, blob, fecha: Date.now() } satisfies Entrada));
    const todas = await tx("readonly", (s) => s.getAll() as IDBRequest<Entrada[]>);
    const viejas = todas.sort((a, b) => b.fecha - a.fecha).slice(MAX_ARCHIVOS);
    for (const v of viejas) await tx("readwrite", (s) => s.delete(v.id));
  } catch {
    // Sin espacio o modo privado: el historial queda solo con metadatos
  }
}

export async function leerArchivo(id: string): Promise<Entrada | null> {
  try {
    return ((await tx("readonly", (s) => s.get(id) as IDBRequest<Entrada | undefined>)) ?? null);
  } catch {
    return null;
  }
}

export async function borrarArchivo(id: string) {
  try {
    await tx("readwrite", (s) => s.delete(id));
  } catch {
    /* nada */
  }
}

// QA interno: compara el resultado contra el PDF de origen releyendo el archivo generado.
// La relectura es en streaming (por partes), así funciona igual con 10.000 páginas:
//   DOCX: se localiza word/document.xml con el directorio central del ZIP y se descomprime por bloques.
//   DOC:  se recorre el XML por bloques.
import { Inflate } from "fflate";
import type { ResultadoConversion } from "./convertir-pdf";

export interface InformeQA {
  ok: boolean;
  lineas: string[];
  paginasOrigen: number;
  paginasSalida: number;
  imagenesOrigen: number;
  imagenesSalida: number;
  areaOrigen: number; // pt² totales
  areaSalida: number;
}

interface Conteo {
  paginas: { ancho: number; alto: number }[];
  imagenes: number; // colocaciones de imagen
  formas: number;
  archivosImagen: number;
}

/** Recorre texto XML por bloques; cada bloque se corta en el último ">" para no partir etiquetas. */
class Escaner {
  private resto = "";
  private dec = new TextDecoder();
  c: Conteo = { paginas: [], imagenes: 0, formas: 0, archivosImagen: 0 };
  constructor(private reImagen: RegExp, private reForma: RegExp | null, private reArchivo: RegExp | null) {}
  empujar(bytes: Uint8Array, fin: boolean) {
    const t = this.resto + this.dec.decode(bytes, { stream: !fin });
    const corte = fin ? t.length : t.lastIndexOf(">") + 1;
    this.procesar(t.slice(0, corte));
    this.resto = t.slice(corte);
  }
  private procesar(t: string) {
    for (const m of t.matchAll(/<w:pgSz w:w="(\d+)" w:h="(\d+)"/g)) this.c.paginas.push({ ancho: +m[1] / 20, alto: +m[2] / 20 });
    this.c.imagenes += (t.match(this.reImagen) || []).length;
    if (this.reForma) this.c.formas += (t.match(this.reForma) || []).length;
    if (this.reArchivo) this.c.archivosImagen += (t.match(this.reArchivo) || []).length;
  }
}

async function leerRango(b: Blob, ini: number, fin: number) {
  return new Uint8Array(await b.slice(ini, fin).arrayBuffer());
}

async function contarDocx(blob: Blob): Promise<Conteo> {
  // fin del directorio central (EOCD)
  const cola = await leerRango(blob, Math.max(0, blob.size - 66000), blob.size);
  const dvc = new DataView(cola.buffer);
  let eocd = -1;
  for (let i = cola.length - 22; i >= 0; i--) if (dvc.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("ZIP sin directorio central");
  const tamCd = dvc.getUint32(eocd + 12, true), iniCd = dvc.getUint32(eocd + 16, true);
  const cd = await leerRango(blob, iniCd, iniCd + tamCd);
  const dv = new DataView(cd.buffer);
  const dec = new TextDecoder();
  let doc: { metodo: number; comp: number; local: number } | null = null;
  let media = 0;
  for (let o = 0; o + 46 <= cd.length && dv.getUint32(o, true) === 0x02014b50; ) {
    const metodo = dv.getUint16(o + 10, true), comp = dv.getUint32(o + 20, true);
    const ln = dv.getUint16(o + 28, true), le = dv.getUint16(o + 30, true), lc = dv.getUint16(o + 32, true);
    const local = dv.getUint32(o + 42, true);
    const nombre = dec.decode(cd.subarray(o + 46, o + 46 + ln));
    if (nombre === "word/document.xml") doc = { metodo, comp, local };
    if (nombre.startsWith("word/media/")) media++;
    o += 46 + ln + le + lc;
  }
  if (!doc) throw new Error("Falta word/document.xml");
  const cab = await leerRango(blob, doc.local, doc.local + 30);
  const dvl = new DataView(cab.buffer);
  const inicio = doc.local + 30 + dvl.getUint16(26, true) + dvl.getUint16(28, true);
  const esc = new Escaner(/<a:blip r:embed=/g, /<wps:cNvSpPr\/>/g, null);
  const PASO = 4 * 1024 * 1024;
  if (doc.metodo === 0) {
    for (let p = 0; p < doc.comp; p += PASO) esc.empujar(await leerRango(blob, inicio + p, inicio + Math.min(doc.comp, p + PASO)), p + PASO >= doc.comp);
  } else {
    const inf = new Inflate((d, fin) => esc.empujar(d, fin));
    for (let p = 0; p < doc.comp; p += PASO) inf.push(await leerRango(blob, inicio + p, inicio + Math.min(doc.comp, p + PASO)), p + PASO >= doc.comp);
  }
  esc.c.archivosImagen = media;
  return esc.c;
}

async function contarDoc(blob: Blob): Promise<Conteo> {
  const esc = new Escaner(/<v:imagedata /g, /<v:(?:rect|line) /g, /<w:binData /g);
  const PASO = 4 * 1024 * 1024;
  for (let p = 0; p < blob.size; p += PASO) esc.empujar(await leerRango(blob, p, Math.min(blob.size, p + PASO)), p + PASO >= blob.size);
  return esc.c;
}

export async function verificar(r: ResultadoConversion, formato: "docx" | "doc", capasVectoriales: number): Promise<InformeQA> {
  const lineas: string[] = [];
  let ok = true;
  const falla = (t: string) => { ok = false; lineas.push("ERROR " + t); };
  const area = (l: { ancho: number; alto: number }[]) => l.reduce((a, p) => a + p.ancho * p.alto, 0);
  const areaOrigen = area(r.tamPaginasPdf);

  let c: Conteo;
  try {
    c = formato === "docx" ? await contarDocx(r.blob) : await contarDoc(r.blob);
  } catch (e) {
    falla(`No se pudo releer el archivo generado (${(e as Error).message}).`);
    c = { paginas: r.tamPaginas, imagenes: r.imagenes + capasVectoriales, formas: 0, archivosImagen: 0 };
  }
  const imagenesSalida = c.imagenes - capasVectoriales;
  lineas.push(
    `Relectura del ${formato.toUpperCase()}: ${c.paginas.length} secciones, ${c.imagenes} imágenes colocadas` +
      ` (${c.archivosImagen} archivos distintos, ${capasVectoriales} dibujos vectoriales como imagen), ${c.formas} líneas y rectángulos nativos.`,
  );

  const areaSalida = area(c.paginas);
  if (c.paginas.length !== r.paginasPdf) falla(`Páginas: PDF ${r.paginasPdf}, salida ${c.paginas.length}.`);
  else lineas.push(`Páginas: ${c.paginas.length} de ${r.paginasPdf}. Coinciden.`);

  // Word guarda el tamaño en twips (1/20 pt): tolerancia de 0,05 pt por lado
  const distintas = r.tamPaginasPdf.filter((p, i) => {
    const s = c.paginas[i];
    return !s || Math.abs(s.ancho - p.ancho) > 0.05 || Math.abs(s.alto - p.alto) > 0.05;
  }).length;
  if (distintas) falla(`${distintas} páginas con tamaño distinto al del PDF.`);
  else lineas.push(`Tamaño de página: área total ${areaOrigen.toFixed(0)} pt² en el PDF y ${areaSalida.toFixed(0)} pt² en la salida. Coinciden.`);

  if (imagenesSalida !== r.imagenesPdf) falla(`Imágenes: PDF ${r.imagenesPdf}, salida ${imagenesSalida}.`);
  else lineas.push(`Imágenes: ${imagenesSalida} de ${r.imagenesPdf}. Coinciden.`);

  if (r.paginasEscaneadas) lineas.push(`Páginas escaneadas (sin texto): ${r.paginasEscaneadas}.`);
  if (r.reemplazadasOcr) lineas.push(`Imágenes escaneadas reemplazadas por texto OCR: ${r.reemplazadasOcr} (no se cuentan en la salida).`);
  return { ok, lineas, paginasOrigen: r.paginasPdf, paginasSalida: c.paginas.length, imagenesOrigen: r.imagenesPdf, imagenesSalida, areaOrigen, areaSalida };
}

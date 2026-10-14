/// <reference lib="webworker" />
// Web Worker de conversión. Todo el trabajo pesado ocurre aquí, en el navegador del usuario.
import * as pdfjs from "pdfjs-dist";
import * as pdfjsWorker from "pdfjs-dist/build/pdf.worker.mjs";
import { convertirPdf, ConversionCancelada } from "@/lib/motor/convertir-pdf";
import { verificar, type InformeQA } from "@/lib/motor/qa";
import { codificarPng } from "@/lib/motor/png";
import { crearMedidor, dibujarVectores, ESCALA_VECTORES, type Ctx2D } from "@/lib/motor/lienzo";
import type { Entorno, ImagenCruda, PaginaInterpretada, SegmentoRot } from "@/lib/motor/interprete";
import { EscritorDocx } from "@/lib/motor/escritor-docx";
import { EscritorDoc2003 } from "@/lib/motor/escritor-doc";
import { mapearFuente } from "@/lib/motor/fuentes";
import type { FormatoSalida } from "@/lib/motor/tipos";
import { leerDoc, leerDocx, leerHtml, leerOdt, leerRtf, leerTxt } from "@/lib/flujo/lectores";
import { flujoADoc2003, flujoADocx } from "@/lib/flujo/escritor-flujo";
import type { DocumentoFlujo } from "@/lib/flujo/modelo";
import type { MensajeAlWorker, MensajeDelWorker } from "./protocolo";

declare const self: DedicatedWorkerGlobalScope;

// pdf.js usa su "worker" dentro de este mismo hilo (ya estamos fuera del hilo principal)
(globalThis as unknown as { pdfjsWorker: unknown }).pdfjsWorker = pdfjsWorker;

const enviar = (m: MensajeDelWorker, transferir: Transferable[] = []) => self.postMessage(m, transferir);
let cancelado = false;
void leerHtml;

const LIMITE_MEMORIA = 64 * 1024 * 1024; // por encima se lee el PDF por rangos desde el disco

class TransporteArchivo extends pdfjs.PDFDataRangeTransport {
  constructor(private archivo: File) {
    super(archivo.size, null);
  }
  requestDataRange(inicio: number, fin: number) {
    this.archivo
      .slice(inicio, fin)
      .arrayBuffer()
      .then((b) => this.onDataRange(inicio, new Uint8Array(b)));
  }
}

// Las fábricas por defecto de pdf.js usan `document.baseURI`, que no existe en un Worker.
class FabricaCMap {
  constructor(private o: { baseUrl: string; isCompressed: boolean }) {}
  async fetch({ name }: { name: string }) {
    const r = await fetch(this.o.baseUrl + name + (this.o.isCompressed ? ".bcmap" : ""));
    if (!r.ok) throw new Error(`No se pudo cargar el CMap ${name}.`);
    return { cMapData: new Uint8Array(await r.arrayBuffer()), compressionType: this.o.isCompressed ? 1 : 0 };
  }
}
class FabricaFuentes {
  constructor(private o: { baseUrl: string }) {}
  async fetch({ filename }: { filename: string }) {
    const r = await fetch(this.o.baseUrl + filename);
    if (!r.ok) throw new Error(`No se pudo cargar la fuente ${filename}.`);
    return new Uint8Array(await r.arrayBuffer());
  }
}

async function abrirPdf(archivo: File, contrasena?: string) {
  const base = new URL("/pdfjs/", self.location.origin).href;
  const comunes = {
    password: contrasena || undefined,
    isOffscreenCanvasSupported: false, // queremos los píxeles crudos de cada imagen
    isImageDecoderSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
    standardFontDataUrl: base + "standard_fonts/",
    cMapUrl: base + "cmaps/",
    cMapPacked: true,
    CMapReaderFactory: FabricaCMap,
    StandardFontDataFactory: FabricaFuentes,
    verbosity: 0,
  };
  const tarea =
    archivo.size > LIMITE_MEMORIA
      ? pdfjs.getDocument({ ...comunes, range: new TransporteArchivo(archivo), rangeChunkSize: 1 << 20, disableAutoFetch: true, disableStream: true, length: archivo.size })
      : pdfjs.getDocument({ ...comunes, data: new Uint8Array(await archivo.arrayBuffer()) });
  return tarea.promise;
}

// ---------------------------------------------------------------- entorno del navegador
function crearEntorno(): Entorno {
  const medir = crearMedidor(new OffscreenCanvas(8, 8).getContext("2d") as unknown as Ctx2D);
  return {
    medirTexto: medir,
    codificarImagen: async (img: ImagenCruda) => {
      // fotos opacas y grandes: JPEG de alta calidad (90); el resto PNG sin pérdida
      if (img.opaca && img.ancho * img.alto > 250_000) {
        const c = new OffscreenCanvas(img.ancho, img.alto);
        const ctx = c.getContext("2d")!;
        ctx.putImageData(new ImageData(new Uint8ClampedArray(img.rgba.buffer as ArrayBuffer, img.rgba.byteOffset, img.rgba.byteLength), img.ancho, img.alto), 0, 0);
        const b = await c.convertToBlob({ type: "image/jpeg", quality: 0.9 });
        return { datos: new Uint8Array(await b.arrayBuffer()), ext: "jpeg" };
      }
      return { datos: codificarPng(img.ancho, img.alto, img.rgba, !img.opaca), ext: "png" };
    },
    rasterizarVectores: async (ancho, alto, cmds) => {
      const w = Math.ceil(ancho * ESCALA_VECTORES), h = Math.ceil(alto * ESCALA_VECTORES);
      const c = new OffscreenCanvas(w, h);
      dibujarVectores(c.getContext("2d") as unknown as Ctx2D, cmds);
      const b = await c.convertToBlob({ type: "image/png" });
      return { datos: new Uint8Array(await b.arrayBuffer()), pixAncho: w, pixAlto: h };
    },
  };
}

// ---------------------------------------------------------------- OCR (Tesseract.js, en el navegador)
type TessWorker = Awaited<ReturnType<typeof import("tesseract.js")["createWorker"]>>;
let tess: TessWorker | null = null;

async function ocr(datos: Uint8Array, mime: string) {
  if (!tess) {
    const { createWorker } = await import("tesseract.js");
    enviar({ tipo: "estado", texto: "Cargando motor OCR (spa + eng)..." });
    tess = await createWorker(["spa", "eng"], 1);
  }
  const r = await tess.recognize(new Blob([datos as BlobPart], { type: mime }), {}, { blocks: true, text: false });
  const lineas: { texto: string; x0: number; y0: number; x1: number; y1: number; base: number }[] = [];
  for (const b of r.data.blocks ?? [])
    for (const p of b.paragraphs)
      for (const l of p.lines) {
        const t = l.text.replace(/\s+$/, "");
        if (!t || l.confidence < 30) continue;
        const base = l.baseline && l.baseline.y0 > l.bbox.y0 ? (l.baseline.y0 + l.baseline.y1) / 2 : l.bbox.y1 - (l.bbox.y1 - l.bbox.y0) * 0.2;
        lineas.push({ texto: t, x0: l.bbox.x0, y0: l.bbox.y0, x1: l.bbox.x1, y1: l.bbox.y1, base });
      }
  return lineas;
}

/** Reemplaza la imagen escaneada de la página por texto OCR posicionado. */
async function aplicarOcr(p: PaginaInterpretada): Promise<PaginaInterpretada> {
  const scan = [...p.imagenes].sort((a, b) => b.ancho * b.alto - a.ancho * a.alto)[0];
  if (!scan || scan.rot) return p;
  const lineas = await ocr(scan.datos, scan.ext === "jpeg" ? "image/jpeg" : "image/png");
  if (!lineas.length) {
    p.avisos.push(`Página ${p.numero}: OCR sin texto reconocible; se conserva la imagen.`);
    return p;
  }
  const ex = scan.ancho / scan.pixAncho, ey = scan.alto / scan.pixAlto;
  const estilo = mapearFuente({ name: "Arial" });
  const segmentos: SegmentoRot[] = lineas.map((l, i) => {
    const alto = (l.y1 - l.y0) * ey;
    const tam = Math.max(4, Math.min(72, alto * 0.82));
    const x = scan.x + (scan.volteadaV ? 0 : l.x0 * ex);
    const ancho = (l.x1 - l.x0) * ex;
    return {
      x, base: scan.y + l.base * ey, ancho, tam, orden: 100000 + i, rot: 0,
      corridas: [{ texto: l.texto, x, ancho, estilo: { fuente: estilo.fuente, fuenteOriginal: "OCR", tam, negrita: false, cursiva: false, color: "000000" } }],
    };
  });
  p.avisos.push(`Página ${p.numero}: escaneada; ${segmentos.length} líneas reconocidas con OCR reemplazan la imagen.`);
  return { ...p, segmentos, imagenes: p.imagenes.filter((im) => im !== scan), imagenesEnPdf: p.imagenesEnPdf - 1 };
}

// ---------------------------------------------------------------- imagen suelta (JPG/PNG)
async function convertirImagen(archivo: File, formato: FormatoSalida, conOcr: boolean, entorno: Entorno) {
  const datos = new Uint8Array(await archivo.arrayBuffer());
  const bmp = await createImageBitmap(archivo);
  const pixAncho = bmp.width, pixAlto = bmp.height;
  bmp.close();
  // 96 ppp: 1 px = 0,75 pt
  const ancho = pixAncho * 0.75, alto = pixAlto * 0.75;
  const ext: "png" | "jpeg" = archivo.type === "image/png" || /\.png$/i.test(archivo.name) ? "png" : "jpeg";
  let pagina: PaginaInterpretada = {
    numero: 1, ancho, alto, segmentos: [], formas: [], vectores: null, imagenesEnPdf: 1, escaneada: true, avisos: [], textoInvisible: 0,
    imagenes: [{ x: 0, y: 0, ancho, alto, pixAncho, pixAlto, datos, ext, orden: 0, transformada: false, rot: 0, volteadaV: false }],
  };
  if (conOcr) {
    enviar({ tipo: "progreso", fase: "ocr", pagina: 1, total: 1, porcentaje: 30 });
    pagina = await aplicarOcr(pagina);
  }
  const esc = formato === "doc" ? new EscritorDoc2003(entorno.medirTexto, archivo.name) : new EscritorDocx(entorno.medirTexto, archivo.name);
  esc.agregarPagina(pagina, true);
  const r = await esc.finalizar();
  return {
    ...r, paginasPdf: 1, imagenesPdf: pagina.imagenesEnPdf, paginasEscaneadas: 1, capasVectoriales: 0, reemplazadasOcr: conOcr ? 1 - pagina.imagenesEnPdf : 0,
    tamPaginasPdf: [{ ancho, alto }], avisos: pagina.avisos,
  };
}

// ---------------------------------------------------------------- formatos de flujo
async function convertirFlujo(archivo: File, ext: string, formato: FormatoSalida) {
  const datos = new Uint8Array(await archivo.arrayBuffer());
  if (ext === "docx" && formato === "docx") {
    return { blob: new Blob([datos], { type: archivo.type }), avisos: ["El archivo ya está en formato DOCX; se entrega sin cambios."], paginas: null as number | null };
  }
  let d: DocumentoFlujo;
  if (ext === "txt") d = leerTxt(datos);
  else if (ext === "rtf") d = leerRtf(datos);
  else if (ext === "odt") d = leerOdt(datos);
  else if (ext === "docx") d = await leerDocx(datos);
  else if (ext === "doc") d = leerDoc(datos);
  else throw new Error("Formato de archivo no admitido.");
  enviar({ tipo: "progreso", fase: "empaquetando", pagina: 1, total: 1, porcentaje: 80 });
  const r = formato === "doc" ? flujoADoc2003(d) : await flujoADocx(d);
  return { blob: r.blob, avisos: d.avisos, paginas: null };
}

// ---------------------------------------------------------------- mensajes
self.onmessage = async (ev: MessageEvent<MensajeAlWorker>) => {
  const m = ev.data;
  if (m.tipo === "cancelar") {
    cancelado = true;
    return;
  }
  try {
    if (m.tipo === "analizar") {
      const ext = m.archivo.name.split(".").pop()!.toLowerCase();
      if (ext !== "pdf") {
        enviar({ tipo: "analisis", paginas: null, escaneado: ["jpg", "jpeg", "png"].includes(ext), cifrado: false });
        return;
      }
      const doc = await abrirPdf(m.archivo, m.contrasena);
      // muestra: hasta 5 páginas repartidas; si ninguna tiene texto, es probable que sea escaneado
      const n = doc.numPages;
      const muestras = [...new Set([1, 2, Math.ceil(n / 2), n - 1, n].filter((k) => k >= 1 && k <= n))];
      let conTexto = 0;
      for (const k of muestras) {
        const tc = await (await doc.getPage(k)).getTextContent();
        if (tc.items.some((i) => "str" in i && i.str.trim())) conTexto++;
      }
      enviar({ tipo: "analisis", paginas: n, escaneado: conTexto === 0, cifrado: !!m.contrasena });
      await doc.destroy();
      return;
    }

    if (m.tipo === "convertir") {
      cancelado = false;
      const t0 = performance.now();
      const ext = m.archivo.name.split(".").pop()!.toLowerCase();
      const titulo = m.archivo.name.replace(/\.[^.]+$/, "");
      const entorno = crearEntorno();
      let qa: InformeQA | null = null;
      let blob: Blob;
      let avisos: string[] = [];
      let paginas: number | null = null;

      if (ext === "pdf") {
        enviar({ tipo: "progreso", fase: "leyendo", pagina: 0, total: 0, porcentaje: 0 });
        const doc = await abrirPdf(m.archivo, m.contrasena);
        const r = await convertirPdf(doc, entorno, {
          formato: m.formato,
          tamBloque: 50,
          titulo,
          ocr: m.ocr ? aplicarOcr : undefined,
          cancelado: () => cancelado,
          alProgreso: (p) => enviar({ tipo: "progreso", ...p }),
        });
        await doc.destroy();
        enviar({ tipo: "progreso", fase: "verificando", pagina: r.paginasPdf, total: r.paginasPdf, porcentaje: 98 });
        qa = await verificar(r, m.formato, r.capasVectoriales);
        blob = r.blob;
        avisos = r.avisos;
        paginas = r.paginasPdf;
        if (r.paginasEscaneadas && !m.ocr) avisos.unshift(`${r.paginasEscaneadas} páginas parecen escaneadas (sin texto). Activa OCR para obtener texto editable.`);
      } else if (["jpg", "jpeg", "png"].includes(ext)) {
        const r = await convertirImagen(m.archivo, m.formato, m.ocr, entorno);
        qa = await verificar(r, m.formato, 0);
        blob = r.blob;
        avisos = r.avisos;
        paginas = 1;
      } else {
        const r = await convertirFlujo(m.archivo, ext, m.formato);
        blob = r.blob;
        avisos = r.avisos;
        paginas = r.paginas;
      }
      if (cancelado) throw new ConversionCancelada();
      enviar({ tipo: "listo", blob, avisos, qa, paginas, segundos: (performance.now() - t0) / 1000 });
    }
  } catch (e) {
    const err = e as { name?: string; code?: number; message?: string };
    if (e instanceof ConversionCancelada) enviar({ tipo: "cancelado" });
    else if (err?.name === "PasswordException") enviar({ tipo: "contrasena", incorrecta: err.code === 2 });
    else if (err?.name === "InvalidPDFException") enviar({ tipo: "error", mensaje: "El archivo no es un PDF válido o está dañado." });
    else if (err instanceof RangeError || /memory|allocation/i.test(err?.message || "")) enviar({ tipo: "error", mensaje: "El archivo supera el tamaño que el navegador puede procesar. Intenta con uno más pequeño." });
    else enviar({ tipo: "error", mensaje: err?.message || "Error desconocido durante la conversión." });
  }
};

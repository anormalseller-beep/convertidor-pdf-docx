// Intérprete de la lista de operadores de pdf.js para UNA página.
// Reproduce el estado gráfico (CTM, matriz de texto, colores) igual que el renderizador
// de canvas de pdf.js, pero en lugar de pintar guarda objetos posicionados:
//   - texto real con fuente/tamaño/color/negrita/cursiva y coordenadas exactas
//   - imágenes con su caja real (ancho/alto/posición/rotación) en puntos
//   - trazos vectoriales (líneas de tablas, recuadros) para una capa aparte
import type { PDFPageProxy } from "pdfjs-dist";
import { OPS } from "./ops";
import { mapearFuente, type FuentePdf } from "./fuentes";
import type { Corrida, Estilo, ImagenPagina, PaginaModelo, Segmento } from "./tipos";

type Mat = [number, number, number, number, number, number];

const ID: Mat = [1, 0, 0, 1, 0, 0];

function mul(a: Mat, b: Mat): Mat {
  return [
    a[0] * b[0] + a[2] * b[1],
    a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3],
    a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4],
    a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}

function aplicar(m: Mat, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function hex2(n: number) {
  return Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");
}

function colorDeArgs(args: unknown): string {
  // pdf.js 4.x entrega [r,g,b] (0-255) en un Uint8ClampedArray; versiones nuevas, "#rrggbb"
  const a = args as ArrayLike<number> & { 0: unknown };
  if (typeof a[0] === "string") return (a[0] as string).replace("#", "").toUpperCase();
  return (hex2(a[0] as number) + hex2(a[1]) + hex2(a[2])).toUpperCase();
}

interface Estado {
  ctm: Mat;
  relleno: string;
  trazo: string;
  grosor: number;
  fuente: FuentePdf | null;
  fuenteId: string;
  tam: number;
  direccion: number;
  espLetra: number;
  espPalabra: number;
  escalaH: number;
  interlineado: number;
  subida: number;
  modoTexto: number;
  mt: Mat; // matriz de texto
  x: number;
  y: number;
  lineX: number;
  lineY: number;
}

/** Un trozo de texto continuo tal como lo dibuja el PDF. */
interface Pieza {
  texto: string;
  x: number;
  base: number;
  ancho: number;
  tam: number;
  rot: number; // grados, sentido horario
  estilo: Estilo;
  orden: number;
}

export interface ComandoVector {
  // segmentos en coordenadas de página (pt, origen arriba-izquierda)
  ruta: (["M", number, number] | ["L", number, number] | ["C", number, number, number, number, number, number] | ["Z"])[];
  relleno: string | null;
  trazo: string | null;
  grosor: number;
  parImpar: boolean;
  orden: number;
}

/** Rectángulos y líneas rectas: se escriben como formas nativas de Word (editables, nítidas). */
export type Forma =
  | { tipo: "rect"; x: number; y: number; ancho: number; alto: number; relleno: string | null; trazo: string | null; grosor: number; orden: number }
  | { tipo: "linea"; x1: number; y1: number; x2: number; y2: number; trazo: string; grosor: number; orden: number };

export interface ImagenCruda {
  ancho: number;
  alto: number;
  rgba: Uint8ClampedArray | Uint8Array;
  opaca: boolean;
}

export interface Entorno {
  /** Codifica píxeles RGBA a PNG/JPEG. */
  codificarImagen: (img: ImagenCruda) => Promise<{ datos: Uint8Array; ext: "png" | "jpeg" }>;
  /** Rasteriza la capa vectorial (opcional; sin él se omite). */
  rasterizarVectores?: (ancho: number, alto: number, cmds: ComandoVector[]) => Promise<{ datos: Uint8Array; pixAncho: number; pixAlto: number } | null>;
  /** Mide el ancho en pt de un texto con una fuente estándar (opcional). */
  medirTexto?: (texto: string, fuente: string, tam: number, negrita: boolean, cursiva: boolean) => number | null;
  /** Caché compartida entre páginas: la misma imagen (logo, fondo) se codifica una sola vez. */
  cacheImagenes?: Map<string, { datos: Uint8Array; ext: "png" | "jpeg"; pixAncho: number; pixAlto: number }>;
}

/** Huella rápida de los píxeles (dos hashes de 32 bits sobre palabras de 4 bytes). */
function huella(img: ImagenCruda): string {
  const d = img.rgba;
  const u32 = d.byteOffset % 4 === 0 ? new Uint32Array(d.buffer, d.byteOffset, d.byteLength >> 2) : new Uint32Array(Uint8Array.from(d).buffer);
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  for (let i = 0; i < u32.length; i++) {
    h1 = Math.imul(h1 ^ u32[i], 0x01000193);
    h2 = Math.imul(h2 ^ u32[i], 0x5bd1e995) ^ (h2 >>> 13);
  }
  return `${img.ancho}x${img.alto}:${(h1 >>> 0).toString(36)}${(h2 >>> 0).toString(36)}`;
}

const MAX_PIX = 36_000_000; // límite de píxeles por imagen para no agotar la RAM

/** Convierte los datos de imagen de pdf.js a RGBA. */
function aRgba(img: null | { width: number; height: number; kind?: number; data?: Uint8ClampedArray | Uint8Array }): ImagenCruda | null {
  if (!img) return null;
  const { width: w, height: h, kind, data } = img;
  if (!data || !w || !h) return null;
  const n = w * h;
  if (kind === 3 || data.length === n * 4) {
    let opaca = true;
    for (let i = 3; i < data.length; i += 4) if (data[i] !== 255) { opaca = false; break; }
    return { ancho: w, alto: h, rgba: data, opaca };
  }
  const out = new Uint8ClampedArray(n * 4);
  if (kind === 2 || data.length === n * 3) {
    for (let i = 0, j = 0; i < n; i++, j += 3) {
      out[i * 4] = data[j]; out[i * 4 + 1] = data[j + 1]; out[i * 4 + 2] = data[j + 2]; out[i * 4 + 3] = 255;
    }
    return { ancho: w, alto: h, rgba: out, opaca: true };
  }
  // GRAYSCALE_1BPP: bits empaquetados por fila, 1 = blanco
  const porFila = Math.ceil(w / 8);
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) {
      const bit = (data[y * porFila + (x >> 3)] >> (7 - (x & 7))) & 1;
      const v = bit ? 255 : 0;
      const o = (y * w + x) * 4;
      out[o] = out[o + 1] = out[o + 2] = v; out[o + 3] = 255;
    }
  return { ancho: w, alto: h, rgba: out, opaca: true };
}

/** Máscara de 1 bit pintada con el color de relleno (logos, texto escaneado). */
function mascaraARgba(img: { width: number; height: number; data?: Uint8Array | Uint8ClampedArray }, color: string): ImagenCruda | null {
  const { width: w, height: h, data } = img;
  if (!data || !w || !h) return null;
  const r = parseInt(color.slice(0, 2), 16), g = parseInt(color.slice(2, 4), 16), b = parseInt(color.slice(4, 6), 16);
  const out = new Uint8ClampedArray(w * h * 4);
  const porFila = Math.ceil(w / 8);
  if (data.length >= w * h * 4) {
    // ya viene expandida como RGBA: el alfa indica dónde pintar
    for (let i = 0; i < w * h; i++) {
      out[i * 4] = r; out[i * 4 + 1] = g; out[i * 4 + 2] = b; out[i * 4 + 3] = data[i * 4 + 3];
    }
  } else {
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const bit = (data[y * porFila + (x >> 3)] >> (7 - (x & 7))) & 1;
        const o = (y * w + x) * 4;
        out[o] = r; out[o + 1] = g; out[o + 2] = b; out[o + 3] = bit ? 0 : 255;
      }
  }
  return { ancho: w, alto: h, rgba: out, opaca: false };
}

function reducir(img: ImagenCruda): ImagenCruda {
  const n = img.ancho * img.alto;
  if (n <= MAX_PIX) return img;
  const f = Math.sqrt(MAX_PIX / n);
  const w = Math.max(1, Math.floor(img.ancho * f)), h = Math.max(1, Math.floor(img.alto * f));
  const out = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    const sy = Math.min(img.alto - 1, Math.floor(y / f));
    for (let x = 0; x < w; x++) {
      const sx = Math.min(img.ancho - 1, Math.floor(x / f));
      const s = (sy * img.ancho + sx) * 4, d = (y * w + x) * 4;
      out[d] = img.rgba[s]; out[d + 1] = img.rgba[s + 1]; out[d + 2] = img.rgba[s + 2]; out[d + 3] = img.rgba[s + 3];
    }
  }
  return { ancho: w, alto: h, rgba: out, opaca: img.opaca };
}

/** Caja de una imagen: el cuadrado unidad transformado por la CTM. */
function geometriaImagen(ctm: Mat) {
  const [ox, oy] = aplicar(ctm, 0, 0);
  const [ux, uy] = aplicar(ctm, 1, 0);
  const [vx, vy] = aplicar(ctm, 0, 1);
  const u = [ux - ox, uy - oy], v = [vx - ox, vy - oy];
  const ancho = Math.hypot(u[0], u[1]);
  const alto = Math.hypot(v[0], v[1]);
  const [cx, cy] = aplicar(ctm, 0.5, 0.5);
  let rot = (Math.atan2(u[1], u[0]) * 180) / Math.PI;
  // En pantalla (y hacia abajo) la parte superior de la imagen está a -90° de "u".
  const cruz = u[0] * v[1] - u[1] * v[0];
  const volteadaV = cruz > 0;
  const ortogonal = Math.abs(u[0] * v[0] + u[1] * v[1]) < 0.02 * ancho * alto;
  rot = ((rot % 360) + 360) % 360;
  if (Math.abs(rot - 360) < 0.05 || rot < 0.05) rot = 0;
  return { x: cx - ancho / 2, y: cy - alto / 2, ancho, alto, rot, volteadaV, ortogonal };
}

export interface ImagenColocada extends ImagenPagina {
  rot: number;
  volteadaV: boolean;
}

export interface SegmentoRot extends Segmento {
  rot: number;
}

export interface PaginaInterpretada extends PaginaModelo {
  formas: Forma[];
  imagenes: ImagenColocada[];
  segmentos: SegmentoRot[];
  textoInvisible: number;
}

export async function interpretarPagina(pagina: PDFPageProxy, numero: number, entorno: Entorno): Promise<PaginaInterpretada> {
  const vp = pagina.getViewport({ scale: 1 });
  const ancho = vp.width, alto = vp.height;
  const lista = await pagina.getOperatorList();
  const avisos: string[] = [];
  const piezas: Pieza[] = [];
  const imagenes: ImagenColocada[] = [];
  const vectores: ComandoVector[] = [];
  let imagenesEnPdf = 0;
  let textoInvisible = 0;
  let orden = 0;
  let rutaActual: ComandoVector["ruta"] = [];
  const cacheImg = new Map<string, { datos: Uint8Array; ext: "png" | "jpeg"; pixAncho: number; pixAlto: number }>();
  const avisados = new Set<string>();
  const avisar = (t: string) => { if (!avisados.has(t)) { avisados.add(t); avisos.push(t); } };

  let st: Estado = {
    ctm: vp.transform as Mat,
    relleno: "000000",
    trazo: "000000",
    grosor: 1,
    fuente: null,
    fuenteId: "",
    tam: 0,
    direccion: 1,
    espLetra: 0,
    espPalabra: 0,
    escalaH: 1,
    interlineado: 0,
    subida: 0,
    modoTexto: 0,
    mt: ID,
    x: 0, y: 0, lineX: 0, lineY: 0,
  };
  const pila: Estado[] = [];

  const obtenerObj = (id: string): unknown => {
    try {
      return id.startsWith("g_") ? pagina.commonObjs.get(id) : pagina.objs.get(id);
    } catch {
      return null;
    }
  };

  // Las imágenes se resuelven de forma asíncrona: se espera al objeto (máx. 15 s)
  const esperarObj = (id: string): Promise<unknown> =>
    new Promise((res) => {
      const almacen = id.startsWith("g_") ? pagina.commonObjs : pagina.objs;
      const t = setTimeout(() => res(null), 15000);
      try {
        almacen.get(id, (d: unknown) => { clearTimeout(t); res(d); });
      } catch {
        clearTimeout(t);
        res(null);
      }
    });

  const estiloActual = (): Estilo => {
    const m = mapearFuente(st.fuente ?? undefined);
    const color = (st.modoTexto & 3) === 1 ? st.trazo : st.relleno;
    return { fuente: m.fuente, fuenteOriginal: m.original, tam: 0, negrita: m.negrita, cursiva: m.cursiva, color };
  };

  const mostrarTexto = (glifos: unknown[]) => {
    const f = st.fuente;
    if (!f || st.tam === 0) return;
    const invisible = (st.modoTexto & 3) === 3;
    const fm = f.fontMatrix || [0.001, 0, 0, 0.001, 0, 0];
    const escalaAvance = st.tam * fm[0];
    const escalaH = st.escalaH * st.direccion;
    const M = mul(st.ctm, mul(st.mt, [1, 0, 0, 1, st.x, st.y + st.subida]));
    const tamDisp = st.tam * Math.hypot(M[2], M[3]);
    const rot = (((Math.atan2(M[1], M[0]) * 180) / Math.PI) % 360 + 360) % 360;
    const estilo = estiloActual();
    let x = 0;
    let piezaIni = 0;
    let texto = "";
    const cerrar = (fin: number) => {
      const limpio = texto.replace(/\s+$/u, "");
      if (limpio.trim() && !invisible) {
        // quitar espacios iniciales moviendo el origen
        const [x0, y0] = aplicar(M, piezaIni * escalaH, 0);
        const [x1, y1] = aplicar(M, fin * escalaH, 0);
        piezas.push({
          texto,
          x: x0,
          base: y0,
          ancho: Math.hypot(x1 - x0, y1 - y0),
          tam: tamDisp,
          rot: Math.abs(rot) < 0.5 || Math.abs(rot - 360) < 0.5 ? 0 : rot,
          estilo: { ...estilo, tam: tamDisp },
          orden: orden++,
        });
      } else if (limpio.trim() && invisible) textoInvisible++;
      texto = "";
    };
    for (const g of glifos) {
      if (typeof g === "number") {
        const salto = (-g * st.tam) / 1000;
        if (salto * Math.abs(escalaH) > st.tam * 0.9 && texto.trim()) {
          // hueco grande (columna/tabulación): se corta la pieza
          cerrar(x);
          x += salto;
          piezaIni = x;
          continue;
        }
        if (salto > st.tam * 0.15 && texto && !texto.endsWith(" ")) texto += " ";
        x += salto;
        continue;
      }
      const gl = g as { unicode?: string; width?: number; isSpace?: boolean };
      if (!texto) piezaIni = x;
      const esp = (gl.isSpace ? st.espPalabra : 0) + st.espLetra;
      texto += gl.unicode ?? "";
      x += (gl.width ?? 0) * escalaAvance + esp * st.direccion;
    }
    cerrar(x);
    st.x += x * escalaH;
  };

  const colocarImagen = async (clave: string | null, cruda: () => ImagenCruda | null | Promise<ImagenCruda | null>, ctm: Mat) => {
    imagenesEnPdf++;
    const geo = geometriaImagen(ctm);
    if (geo.ancho < 0.5 || geo.alto < 0.5) {
      imagenesEnPdf--; // invisible (menos de medio punto): no cuenta para el QA
      return;
    }
    if (!geo.ortogonal) avisar(`Página ${numero}: imagen inclinada; se coloca con su caja sin inclinación.`);
    let cod = clave ? cacheImg.get(clave) : undefined;
    if (!cod) {
      const raw = await cruda();
      if (!raw) {
        avisar(`Página ${numero}: una imagen no se pudo leer.`);
        return;
      }
      const red = reducir(raw);
      if (red !== raw) avisar(`Página ${numero}: imagen muy grande, se redujo su resolución (el tamaño en página se mantiene).`);
      const k = entorno.cacheImagenes ? huella(red) : "";
      cod = entorno.cacheImagenes?.get(k);
      if (!cod) {
        const c = await entorno.codificarImagen(red);
        cod = { ...c, pixAncho: red.ancho, pixAlto: red.alto };
        if (entorno.cacheImagenes) {
          if (entorno.cacheImagenes.size > 500) entorno.cacheImagenes.clear();
          entorno.cacheImagenes.set(k, cod);
        }
      }
      if (clave) cacheImg.set(clave, cod);
    }
    imagenes.push({
      x: geo.x, y: geo.y, ancho: geo.ancho, alto: geo.alto,
      pixAncho: cod.pixAncho, pixAlto: cod.pixAlto, datos: cod.datos, ext: cod.ext,
      orden: orden++, transformada: geo.rot !== 0 || !geo.ortogonal, rot: geo.rot, volteadaV: geo.volteadaV,
    });
  };

  const pintarRuta = (relleno: boolean, trazo: boolean, parImpar = false) => {
    if (rutaActual.length) {
      const escala = Math.sqrt(Math.abs(st.ctm[0] * st.ctm[3] - st.ctm[1] * st.ctm[2]));
      vectores.push({
        ruta: rutaActual,
        relleno: relleno ? st.relleno : null,
        trazo: trazo ? st.trazo : null,
        grosor: Math.max(0.35, st.grosor * escala),
        parImpar,
        orden: orden++,
      });
    }
    rutaActual = [];
  };

  const { fnArray, argsArray } = lista;
  for (let i = 0; i < fnArray.length; i++) {
    const fn = fnArray[i];
    const a = argsArray[i] as unknown[];
    switch (fn) {
      case OPS.save:
        pila.push({ ...st });
        break;
      case OPS.restore:
        if (pila.length) st = pila.pop()!;
        break;
      case OPS.transform:
        st.ctm = mul(st.ctm, a as unknown as Mat);
        break;
      case OPS.paintFormXObjectBegin: {
        pila.push({ ...st });
        const m = a[0] as Mat | null;
        if (m && m.length === 6) st.ctm = mul(st.ctm, m);
        break;
      }
      case OPS.paintFormXObjectEnd:
        if (pila.length) st = pila.pop()!;
        break;
      case OPS.beginGroup: {
        pila.push({ ...st });
        const g = a[0] as { matrix?: Mat } | undefined;
        if (g?.matrix) st.ctm = mul(st.ctm, g.matrix);
        break;
      }
      case OPS.endGroup:
        if (pila.length) st = pila.pop()!;
        break;
      case OPS.setFillRGBColor:
        st.relleno = colorDeArgs(a);
        break;
      case OPS.setStrokeRGBColor:
        st.trazo = colorDeArgs(a);
        break;
      case OPS.setLineWidth:
        st.grosor = a[0] as number;
        break;
      // ---- texto ----
      case OPS.beginText:
        st.mt = ID; st.x = st.lineX = 0; st.y = st.lineY = 0;
        break;
      case OPS.setFont: {
        const id = a[0] as string;
        let tam = a[1] as number;
        st.fuente = obtenerObj(id) as FuentePdf | null;
        st.fuenteId = id;
        if (tam < 0) { tam = -tam; st.direccion = -1; } else st.direccion = 1;
        st.tam = tam;
        if (st.fuente?.vertical) avisar(`Página ${numero}: texto vertical (CJK) colocado en horizontal.`);
        break;
      }
      case OPS.setCharSpacing: st.espLetra = a[0] as number; break;
      case OPS.setWordSpacing: st.espPalabra = a[0] as number; break;
      case OPS.setHScale: st.escalaH = (a[0] as number) / 100; break;
      case OPS.setLeading: st.interlineado = -(a[0] as number); break;
      case OPS.setTextRise: st.subida = a[0] as number; break;
      case OPS.setTextRenderingMode: st.modoTexto = a[0] as number; break;
      case OPS.moveText:
        st.x = st.lineX += a[0] as number;
        st.y = st.lineY += a[1] as number;
        break;
      case OPS.setLeadingMoveText:
        st.interlineado = a[1] as number;
        st.x = st.lineX += a[0] as number;
        st.y = st.lineY += a[1] as number;
        break;
      case OPS.setTextMatrix: {
        const m = (Array.isArray(a[0]) || ArrayBuffer.isView(a[0]) ? a[0] : a) as unknown as Mat;
        st.mt = [m[0], m[1], m[2], m[3], m[4], m[5]];
        st.x = st.lineX = 0; st.y = st.lineY = 0;
        break;
      }
      case OPS.nextLine:
        st.x = st.lineX += 0;
        st.y = st.lineY += st.interlineado;
        st.x = st.lineX;
        break;
      case OPS.showText:
      case OPS.showSpacedText:
        mostrarTexto(a[0] as unknown[]);
        break;
      case OPS.nextLineShowText:
        st.y = st.lineY += st.interlineado; st.x = st.lineX;
        mostrarTexto(a[0] as unknown[]);
        break;
      case OPS.nextLineSetSpacingShowText:
        st.espPalabra = a[0] as number; st.espLetra = a[1] as number;
        st.y = st.lineY += st.interlineado; st.x = st.lineX;
        mostrarTexto(a[2] as unknown[]);
        break;
      // ---- imágenes ----
      case OPS.paintImageXObject: {
        const id = a[0] as string;
        await colocarImagen(id, async () => aRgba((await esperarObj(id)) as never), st.ctm);
        break;
      }
      case OPS.paintInlineImageXObject:
        await colocarImagen(null, () => aRgba(a[0] as never), st.ctm);
        break;
      case OPS.paintImageXObjectRepeat: {
        const [id, sx, sy, pos] = a as [string, number, number, number[]];
        for (let k = 0; k < pos.length; k += 2)
          await colocarImagen(id, async () => aRgba((await esperarObj(id)) as never), mul(st.ctm, [sx, 0, 0, sy, pos[k], pos[k + 1]]));
        break;
      }
      case OPS.paintInlineImageXObjectGroup: {
        const [img, mapa] = a as [unknown, { transform: Mat }[]];
        for (const e of mapa) await colocarImagen(null, () => aRgba(img as never), mul(st.ctm, e.transform));
        break;
      }
      case OPS.paintImageMaskXObject: {
        const img = a[0] as never;
        const color = st.relleno;
        await colocarImagen(null, () => mascaraARgba(img, color), st.ctm);
        break;
      }
      case OPS.paintImageMaskXObjectGroup: {
        const grupo = a[0] as { data: Uint8Array; width: number; height: number; transform: Mat }[];
        for (const m of grupo) {
          const color = st.relleno;
          await colocarImagen(null, () => mascaraARgba(m, color), mul(st.ctm, m.transform));
        }
        break;
      }
      case OPS.paintImageMaskXObjectRepeat: {
        const [img, sx, , , sy, pos] = a as [never, number, number, number, number, number[]];
        for (let k = 0; k < pos.length; k += 2) {
          const color = st.relleno;
          await colocarImagen(null, () => mascaraARgba(img, color), mul(st.ctm, [sx, 0, 0, sy, pos[k], pos[k + 1]]));
        }
        break;
      }
      case OPS.shadingFill:
        avisar(`Página ${numero}: degradados del PDF no se reproducen.`);
        break;
      // ---- trazos ----
      case OPS.constructPath: {
        const [ops, coords] = a as [number[], number[]];
        let j = 0;
        let cx = 0, cy = 0;
        const P = (x: number, y: number) => aplicar(st.ctm, x, y);
        for (const op of ops) {
          switch (op) {
            case OPS.rectangle: {
              const x = coords[j++], y = coords[j++], w = coords[j++], h = coords[j++];
              const p1 = P(x, y), p2 = P(x + w, y), p3 = P(x + w, y + h), p4 = P(x, y + h);
              rutaActual.push(["M", ...p1], ["L", ...p2], ["L", ...p3], ["L", ...p4], ["Z"]);
              cx = x; cy = y;
              break;
            }
            case OPS.moveTo: cx = coords[j++]; cy = coords[j++]; rutaActual.push(["M", ...P(cx, cy)]); break;
            case OPS.lineTo: cx = coords[j++]; cy = coords[j++]; rutaActual.push(["L", ...P(cx, cy)]); break;
            case OPS.curveTo: {
              const c1 = P(coords[j], coords[j + 1]), c2 = P(coords[j + 2], coords[j + 3]);
              cx = coords[j + 4]; cy = coords[j + 5]; j += 6;
              rutaActual.push(["C", ...c1, ...c2, ...P(cx, cy)]);
              break;
            }
            case OPS.curveTo2: {
              const c1 = P(cx, cy), c2 = P(coords[j], coords[j + 1]);
              cx = coords[j + 2]; cy = coords[j + 3]; j += 4;
              rutaActual.push(["C", ...c1, ...c2, ...P(cx, cy)]);
              break;
            }
            case OPS.curveTo3: {
              const c1 = P(coords[j], coords[j + 1]);
              cx = coords[j + 2]; cy = coords[j + 3]; j += 4;
              const fin = P(cx, cy);
              rutaActual.push(["C", ...c1, ...fin, ...fin]);
              break;
            }
            case OPS.closePath: rutaActual.push(["Z"]); break;
          }
        }
        break;
      }
      case OPS.stroke: pintarRuta(false, true); break;
      case OPS.closeStroke: rutaActual.push(["Z"]); pintarRuta(false, true); break;
      case OPS.fill: pintarRuta(true, false); break;
      case OPS.eoFill: pintarRuta(true, false, true); break;
      case OPS.fillStroke: pintarRuta(true, true); break;
      case OPS.eoFillStroke: pintarRuta(true, true, true); break;
      case OPS.closeFillStroke: rutaActual.push(["Z"]); pintarRuta(true, true); break;
      case OPS.closeEOFillStroke: rutaActual.push(["Z"]); pintarRuta(true, true, true); break;
      case OPS.endPath: rutaActual = []; break;
    }
  }

  const segmentos = agruparLineas(piezas, entorno);

  // Vectores: rectángulos y líneas → formas nativas; el resto (curvas) → imagen recortada a su caja.
  // Los rellenos blancos que cubren casi toda la página (fondos) se omiten.
  let formas: Forma[] = [];
  let complejos: ComandoVector[] = [];
  const dibujables: ComandoVector[] = [];
  for (const v of vectores) {
    if (!v.trazo && v.relleno === "FFFFFF") {
      const c = cajaRuta(v);
      if (c.ancho * c.alto > ancho * alto * 0.6) continue;
    }
    dibujables.push(v);
    const f = aForma(v);
    if (f) formas.push(...f);
    else complejos.push(v);
  }
  if (formas.length > 3000) {
    // dibujos con miles de piezas (gráficos, mapas): mejor como una sola imagen
    complejos = dibujables;
    formas = [];
    avisar(`Página ${numero}: dibujo vectorial muy complejo; se coloca como imagen.`);
  }
  let capa: ImagenPagina | null = null;
  if (complejos.length && entorno.rasterizarVectores) {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const c of complejos) {
      const b = cajaRuta(c);
      x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y); x1 = Math.max(x1, b.x + b.ancho); y1 = Math.max(y1, b.y + b.alto);
    }
    x0 = Math.max(0, Math.floor(x0 - 2)); y0 = Math.max(0, Math.floor(y0 - 2));
    x1 = Math.min(ancho, Math.ceil(x1 + 2)); y1 = Math.min(alto, Math.ceil(y1 + 2));
    if (x1 > x0 && y1 > y0) {
      const desplazados = complejos.map((c) => ({ ...c, ruta: c.ruta.map((sg) => desplazar(sg, -x0, -y0)) }));
      const r = await entorno.rasterizarVectores(x1 - x0, y1 - y0, desplazados);
      if (r) capa = { x: x0, y: y0, ancho: x1 - x0, alto: y1 - y0, pixAncho: r.pixAncho, pixAlto: r.pixAlto, datos: r.datos, ext: "png", orden: -1, transformada: false };
    }
  }

  const area = ancho * alto;
  const escaneada = segmentos.length === 0 && imagenes.some((im) => im.ancho * im.alto > area * 0.5);
  if (textoInvisible > 0) avisos.push(`Página ${numero}: ${textoInvisible} fragmentos de texto invisible (capa OCR del PDF) no se copian.`);

  return {
    numero, ancho, alto, segmentos, imagenes, vectores: capa, formas, imagenesEnPdf, escaneada, avisos, textoInvisible,
  };
}

type Seg = ComandoVector["ruta"][number];

function desplazar(sg: Seg, dx: number, dy: number): Seg {
  if (sg[0] === "M" || sg[0] === "L") return [sg[0], sg[1] + dx, sg[2] + dy];
  if (sg[0] === "C") return ["C", sg[1] + dx, sg[2] + dy, sg[3] + dx, sg[4] + dy, sg[5] + dx, sg[6] + dy];
  return sg;
}

function cajaRuta(v: ComandoVector) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const sg of v.ruta) {
    for (let i = 1; i + 1 < sg.length; i += 2) {
      const x = sg[i] as number, y = sg[i + 1] as number;
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
  }
  if (!Number.isFinite(x0)) return { x: 0, y: 0, ancho: 0, alto: 0 };
  const m = v.trazo ? v.grosor / 2 : 0;
  return { x: x0 - m, y: y0 - m, ancho: x1 - x0 + 2 * m, alto: y1 - y0 + 2 * m };
}

/** Si la ruta son rectángulos alineados y/o segmentos rectos, devuelve formas nativas; si no, null. */
function aForma(v: ComandoVector): Forma[] | null {
  const subrutas: { pts: [number, number][]; cerrada: boolean }[] = [];
  let actual: [number, number][] = [];
  for (const sg of v.ruta) {
    if (sg[0] === "C") return null;
    if (sg[0] === "M") {
      if (actual.length) subrutas.push({ pts: actual, cerrada: false });
      actual = [[sg[1], sg[2]]];
    } else if (sg[0] === "L") actual.push([sg[1], sg[2]]);
    else if (actual.length) {
      subrutas.push({ pts: actual, cerrada: true });
      actual = [];
    }
  }
  if (actual.length) subrutas.push({ pts: actual, cerrada: false });
  // un moveTo suelto (sin segmentos) no dibuja nada
  for (let i = subrutas.length - 1; i >= 0; i--) if (subrutas[i].pts.length < 2) subrutas.splice(i, 1);
  if (!subrutas.length || subrutas.length > 400) return null;
  const tol = 0.05;
  const out: Forma[] = [];
  for (const { pts: original, cerrada } of subrutas) {
    let pts = original;
    if (pts.length === 5 && Math.abs(pts[0][0] - pts[4][0]) < tol && Math.abs(pts[0][1] - pts[4][1]) < tol) pts = pts.slice(0, 4);
    const esRect =
      pts.length === 4 && (cerrada || !!v.relleno) &&
      pts.every((q, k) => {
        const r = pts[(k + 1) % 4];
        return Math.abs(q[0] - r[0]) < tol || Math.abs(q[1] - r[1]) < tol;
      });
    if (esRect) {
      const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]);
      const x = Math.min(...xs), y = Math.min(...ys);
      const w = Math.max(...xs) - x, h = Math.max(...ys) - y;
      if (v.relleno && !v.trazo && (w < 1.2 || h < 1.2)) {
        // rectángulo rellenado muy fino = línea (bordes de tabla dibujados como rellenos)
        if (w >= h) out.push({ tipo: "linea", x1: x, y1: y + h / 2, x2: x + w, y2: y + h / 2, trazo: v.relleno, grosor: Math.max(h, 0.25), orden: v.orden });
        else out.push({ tipo: "linea", x1: x + w / 2, y1: y, x2: x + w / 2, y2: y + h, trazo: v.relleno, grosor: Math.max(w, 0.25), orden: v.orden });
      } else out.push({ tipo: "rect", x, y, ancho: w, alto: h, relleno: v.relleno, trazo: v.trazo, grosor: v.grosor, orden: v.orden });
      continue;
    }
    // polilínea de segmentos rectos sin relleno
    if (!v.trazo || (v.relleno && pts.length > 2) || pts.length < 2) return null;
    const lista = cerrada ? [...pts, pts[0]] : pts;
    for (let k = 0; k + 1 < lista.length; k++)
      out.push({ tipo: "linea", x1: lista[k][0], y1: lista[k][1], x2: lista[k + 1][0], y2: lista[k + 1][1], trazo: v.trazo, grosor: v.grosor, orden: v.orden });
  }
  return out.length ? out : null;
}

/** Agrupa piezas en segmentos de línea; respeta columnas y huecos grandes. */
export function agruparLineas(piezas: Pieza[], entorno: Pick<Entorno, "medirTexto">): SegmentoRot[] {
  // Elimina duplicados exactos (negrita simulada dibujando dos veces)
  const vistos = new Set<string>();
  const unicas = piezas.filter((p) => {
    const k = `${p.texto}|${Math.round(p.x)}|${Math.round(p.base)}|${p.rot}`;
    if (vistos.has(k)) return false;
    vistos.add(k);
    return true;
  });
  const rectas = unicas.filter((p) => p.rot === 0).sort((a, b) => a.base - b.base || a.x - b.x);
  const out: SegmentoRot[] = [];

  // líneas por línea base
  const lineas: Pieza[][] = [];
  for (const p of rectas) {
    const ult = lineas[lineas.length - 1];
    if (ult && Math.abs(ult[0].base - p.base) < Math.min(ult[0].tam, p.tam) * 0.3) ult.push(p);
    else lineas.push([p]);
  }
  for (const linea of lineas) {
    linea.sort((a, b) => a.x - b.x);
    let seg: Pieza[] = [];
    const volcar = () => {
      if (seg.length) out.push(construirSegmento(seg, 0, entorno));
      seg = [];
    };
    for (const p of linea) {
      const ult = seg[seg.length - 1];
      if (ult) {
        const fin = ult.x + ult.ancho;
        const hueco = p.x - fin;
        const tam = Math.max(ult.tam, p.tam);
        if (hueco > tam * 1.2 || hueco < -tam * 0.5) volcar();
      }
      seg.push(p);
    }
    volcar();
  }
  for (const p of unicas.filter((q) => q.rot !== 0)) out.push(construirSegmento([p], p.rot, entorno));
  return out;
}

function construirSegmento(piezas: Pieza[], rot: number, entorno: Pick<Entorno, "medirTexto">): SegmentoRot {
  const corridas: Corrida[] = [];
  for (let i = 0; i < piezas.length; i++) {
    const p = piezas[i];
    const sig = piezas[i + 1];
    let texto = p.texto;
    let ancho = p.ancho;
    if (sig) {
      const hueco = sig.x - (p.x + p.ancho);
      if (hueco > Math.min(p.tam, sig.tam) * 0.12 && !texto.endsWith(" ") && !sig.texto.startsWith(" ")) texto += " ";
      ancho = Math.max(p.ancho, sig.x - p.x); // la corrida "absorbe" el hueco hasta la siguiente
    }
    const prev = corridas[corridas.length - 1];
    if (prev && mismoEstilo(prev.estilo, p.estilo)) {
      prev.texto += texto;
      prev.ancho = p.x + ancho - prev.x;
    } else corridas.push({ texto, x: p.x, ancho, estilo: p.estilo });
  }
  const primera = piezas[0];
  const ultima = piezas[piezas.length - 1];
  const tam = Math.max(...piezas.map((p) => p.tam));
  void entorno;
  return {
    x: primera.x,
    base: primera.base,
    ancho: ultima.x + ultima.ancho - primera.x,
    tam,
    corridas,
    orden: primera.orden,
    rot,
  };
}

function mismoEstilo(a: Estilo, b: Estilo) {
  return a.fuente === b.fuente && a.negrita === b.negrita && a.cursiva === b.cursiva && a.color === b.color && Math.abs(a.tam - b.tam) < 0.1;
}

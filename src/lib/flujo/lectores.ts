// Lectores de formatos de flujo → DocumentoFlujo. Todo corre en el navegador.
import { unzipSync, strFromU8 } from "fflate";
import * as CFB from "cfb";
import { dimensionesImagen, tokenizar, type BloqueF, type CorridaF, type DocumentoFlujo, type Alineacion } from "./modelo";

const PX_A_PT = 0.75;

// ---------------------------------------------------------------- TXT
export function leerTxt(datos: Uint8Array): DocumentoFlujo {
  let texto: string;
  try {
    texto = new TextDecoder("utf-8", { fatal: true }).decode(datos);
  } catch {
    texto = new TextDecoder("windows-1252").decode(datos);
  }
  texto = texto.replace(/^﻿/, "");
  const bloques: BloqueF[] = texto.split(/\r\n|\r|\n/).map((l) => ({ tipo: "p", corridas: l ? [{ texto: l.replace(/\t/g, "    ") }] : [] }));
  return { bloques, avisos: [] };
}

// ---------------------------------------------------------------- HTML (salida de mammoth)
export function leerHtml(html: string): DocumentoFlujo {
  const bloques: BloqueF[] = [];
  const avisos: string[] = [];
  const fmt = { negrita: 0, cursiva: 0, subrayado: 0 };
  let actual: { corridas: CorridaF[]; encabezado?: number; lista?: boolean } | null = null;
  let tabla: CorridaF[][][] | null = null;
  let fila: CorridaF[][] | null = null;
  let celda: CorridaF[] | null = null;
  let enLista = 0;
  const destino = () => celda ?? (actual ??= { corridas: [], lista: enLista > 0 }).corridas;
  const cerrarParrafo = () => {
    if (actual && !celda) bloques.push({ tipo: "p", corridas: actual.corridas, encabezado: actual.encabezado, lista: actual.lista });
    actual = null;
  };
  for (const tk of tokenizar(html)) {
    if (tk.t === "texto") {
      if (!tk.texto) continue;
      destino().push({ texto: tk.texto, negrita: fmt.negrita > 0, cursiva: fmt.cursiva > 0, subrayado: fmt.subrayado > 0 });
    } else if (tk.t === "abre") {
      const n = tk.nombre.toLowerCase();
      if (n === "p" || n === "li") { if (!celda) { cerrarParrafo(); actual = { corridas: [], lista: n === "li" || enLista > 0 }; } else if (celda.length) celda.push({ texto: "", salto: true }); }
      else if (/^h[1-6]$/.test(n)) { cerrarParrafo(); actual = { corridas: [], encabezado: +n[1] }; }
      else if (n === "strong" || n === "b") fmt.negrita++;
      else if (n === "em" || n === "i") fmt.cursiva++;
      else if (n === "u") fmt.subrayado++;
      else if (n === "br") destino().push({ texto: "", salto: true });
      else if (n === "ul" || n === "ol") enLista++;
      else if (n === "table") { cerrarParrafo(); tabla = []; }
      else if (n === "tr" && tabla) { fila = []; tabla.push(fila); }
      else if ((n === "td" || n === "th") && fila) { celda = []; fila.push(celda); }
      else if (n === "img") {
        const m = /^data:image\/(\w+);base64,(.*)$/.exec(tk.attrs.src || "");
        if (m) {
          const bin = Uint8Array.from(atob(m[2]), (c) => c.charCodeAt(0));
          const dim = dimensionesImagen(bin);
          if (dim) {
            cerrarParrafo();
            bloques.push({ tipo: "img", datos: bin, ext: dim.ext, ancho: dim.ancho * PX_A_PT, alto: dim.alto * PX_A_PT });
          }
        } else avisos.push("Una imagen en formato no compatible (por ejemplo EMF/WMF) se omitió.");
      }
    } else {
      const n = tk.nombre.toLowerCase();
      if (n === "p" || n === "li" || /^h[1-6]$/.test(n)) cerrarParrafo();
      else if (n === "strong" || n === "b") fmt.negrita--;
      else if (n === "em" || n === "i") fmt.cursiva--;
      else if (n === "u") fmt.subrayado--;
      else if (n === "ul" || n === "ol") enLista--;
      else if (n === "td" || n === "th") celda = null;
      else if (n === "tr") fila = null;
      else if (n === "table" && tabla) { bloques.push({ tipo: "tabla", filas: tabla }); tabla = null; }
    }
  }
  cerrarParrafo();
  return { bloques, avisos };
}

export async function leerDocx(datos: Uint8Array): Promise<DocumentoFlujo> {
  const mammoth = await import("mammoth");
  const ab = datos.buffer.slice(datos.byteOffset, datos.byteOffset + datos.byteLength) as ArrayBuffer;
  // navegador: arrayBuffer; Node (pruebas): buffer
  const entrada = typeof window === "undefined" && typeof self === "undefined" ? { buffer: Buffer.from(ab) } : { arrayBuffer: ab };
  const r = await mammoth.convertToHtml(entrada as { arrayBuffer: ArrayBuffer });
  const d = leerHtml(r.value);
  if (r.messages.some((m) => m.type === "warning")) d.avisos.push("Algunos estilos del DOCX no se pudieron reproducir.");
  return d;
}

// ---------------------------------------------------------------- ODT
export function leerOdt(datos: Uint8Array): DocumentoFlujo {
  const zip = unzipSync(datos);
  if (!zip["content.xml"]) throw new Error("El archivo ODT no contiene content.xml.");
  const xml = strFromU8(zip["content.xml"]);
  const avisos: string[] = [];
  // estilos automáticos: nombre → formato
  const estilos = new Map<string, Partial<CorridaF> & { alin?: Alineacion; padre?: string }>();
  let estiloActual: string | null = null;
  for (const tk of tokenizar(xml)) {
    if (tk.t === "abre" && tk.nombre === "style:style") {
      estiloActual = tk.attrs["style:name"];
      estilos.set(estiloActual, { padre: tk.attrs["style:parent-style-name"] });
    } else if (tk.t === "abre" && estiloActual && tk.nombre === "style:text-properties") {
      const e = estilos.get(estiloActual)!;
      if (tk.attrs["fo:font-weight"] === "bold") e.negrita = true;
      if (tk.attrs["fo:font-style"] === "italic") e.cursiva = true;
      if (tk.attrs["style:text-underline-style"] && tk.attrs["style:text-underline-style"] !== "none") e.subrayado = true;
      if (tk.attrs["fo:font-size"]?.endsWith("pt")) e.tam = parseFloat(tk.attrs["fo:font-size"]);
      if (tk.attrs["fo:color"]?.startsWith("#")) e.color = tk.attrs["fo:color"].slice(1).toUpperCase();
      if (tk.attrs["style:font-name"]) e.fuente = tk.attrs["style:font-name"];
    } else if (tk.t === "abre" && estiloActual && tk.nombre === "style:paragraph-properties") {
      const a = tk.attrs["fo:text-align"];
      estilos.get(estiloActual)!.alin = a === "center" ? "center" : a === "end" || a === "right" ? "right" : a === "justify" ? "both" : undefined;
    } else if (tk.t === "cierra" && tk.nombre === "style:style") estiloActual = null;
  }
  const bloques: BloqueF[] = [];
  const pila: (Partial<CorridaF> | undefined)[] = [];
  let parrafo: { corridas: CorridaF[]; encabezado?: number; alineacion?: Alineacion; lista?: boolean } | null = null;
  let tabla: CorridaF[][][] | null = null, fila: CorridaF[][] | null = null, celda: CorridaF[] | null = null;
  let lista = 0;
  let marco: { ancho: number; alto: number } | null = null;
  const formato = () => Object.assign({}, ...pila.filter(Boolean)) as Partial<CorridaF>;
  const medida = (v?: string) => {
    if (!v) return 0;
    const n = parseFloat(v);
    return v.endsWith("cm") ? n * 28.3465 : v.endsWith("mm") ? n * 2.83465 : v.endsWith("in") ? n * 72 : v.endsWith("px") ? n * 0.75 : n;
  };
  const destino = () => celda ?? parrafo?.corridas;
  for (const tk of tokenizar(xml)) {
    if (tk.t === "abre") {
      const n = tk.nombre;
      if (n === "text:p" || n === "text:h") {
        const est = estilos.get(tk.attrs["text:style-name"]);
        if (celda) { if (celda.length) celda.push({ texto: "", salto: true }); pila.push(est); if (tk.vacio) pila.pop(); continue; }
        parrafo = { corridas: [], encabezado: n === "text:h" ? +(tk.attrs["text:outline-level"] || 1) : undefined, alineacion: est?.alin, lista: lista > 0 };
        pila.push(est);
        if (tk.vacio) { bloques.push({ tipo: "p", corridas: [] }); parrafo = null; pila.pop(); }
      } else if (n === "text:span") { pila.push(estilos.get(tk.attrs["text:style-name"])); if (tk.vacio) pila.pop(); }
      else if (n === "text:s") destino()?.push({ texto: " ".repeat(+(tk.attrs["text:c"] || 1)), ...formato() });
      else if (n === "text:tab") destino()?.push({ texto: "\t", ...formato() });
      else if (n === "text:line-break") destino()?.push({ texto: "", salto: true });
      else if (n === "text:list") lista++;
      else if (n === "table:table") tabla = [];
      else if (n === "table:table-row" && tabla) { fila = []; tabla.push(fila); }
      else if (n === "table:table-cell" && fila) { celda = []; fila.push(celda); if (tk.vacio) celda = null; }
      else if (n === "draw:frame") marco = { ancho: medida(tk.attrs["svg:width"]), alto: medida(tk.attrs["svg:height"]) };
      else if (n === "draw:image") {
        const ruta = tk.attrs["xlink:href"];
        const bin = ruta ? zip[ruta] : undefined;
        const dim = bin ? dimensionesImagen(bin) : null;
        if (bin && dim) {
          const ancho = marco?.ancho || dim.ancho * PX_A_PT, alto = marco?.alto || dim.alto * PX_A_PT;
          if (parrafo && parrafo.corridas.length) { bloques.push({ tipo: "p", ...parrafo }); parrafo = { corridas: [] }; }
          bloques.push({ tipo: "img", datos: bin, ext: dim.ext, ancho, alto });
        } else if (ruta) avisos.push("Una imagen del ODT en formato no compatible se omitió.");
      }
    } else if (tk.t === "cierra") {
      const n = tk.nombre;
      if (n === "text:p" || n === "text:h") {
        pila.pop();
        if (!celda && parrafo) { bloques.push({ tipo: "p", ...parrafo }); parrafo = null; }
      } else if (n === "text:span") pila.pop();
      else if (n === "text:list") lista--;
      else if (n === "table:table-cell") celda = null;
      else if (n === "table:table-row") fila = null;
      else if (n === "table:table" && tabla) { bloques.push({ tipo: "tabla", filas: tabla }); tabla = null; }
      else if (n === "draw:frame") marco = null;
    } else if (tk.texto && destino()) destino()!.push({ texto: tk.texto, ...formato() });
  }
  return { bloques, avisos };
}

// ---------------------------------------------------------------- RTF
export function leerRtf(datos: Uint8Array): DocumentoFlujo {
  const s = new TextDecoder("latin1").decode(datos);
  if (!s.startsWith("{\\rtf")) throw new Error("El archivo no es un RTF válido.");
  const bloques: BloqueF[] = [];
  const avisos: string[] = [];
  const cp1252 = new TextDecoder("windows-1252");
  interface Est { negrita: boolean; cursiva: boolean; subrayado: boolean; tam: number; color: number; saltar: boolean; uc: number }
  let est: Est = { negrita: false, cursiva: false, subrayado: false, tam: 12, color: 0, saltar: false, uc: 1 };
  const pila: Est[] = [];
  const colores: string[] = []; // la entrada 0 es el color automático (";" inicial)
  let enColores = false;
  let colorParcial = { r: 0, g: 0, b: 0 };
  let corridas: CorridaF[] = [];
  let alin: Alineacion | undefined;
  let texto = "";
  let pict: { hex: string; ancho: number; alto: number; tipo: "png" | "jpeg" | null } | null = null;
  let profundidadPict = -1;
  let saltarBytes = 0;
  let sinSoporte = 0; // imágenes WMF/EMF (Word guarda además una copia PNG en \shppict)
  const volcar = () => {
    if (texto) corridas.push({ texto, negrita: est.negrita, cursiva: est.cursiva, subrayado: est.subrayado, tam: est.tam, color: colores[est.color] && colores[est.color] !== "000000" ? colores[est.color] : undefined });
    texto = "";
  };
  const finParrafo = () => {
    volcar();
    bloques.push({ tipo: "p", corridas, alineacion: alin });
    corridas = [];
  };
  const DESTINOS_IGNORADOS = new Set(["fonttbl", "stylesheet", "info", "header", "footer", "headerl", "headerr", "footerl", "footerr", "listtable", "listoverridetable", "rsidtbl", "generator", "xmlnstbl", "themedata", "colorschememapping", "latentstyles", "datastore", "fldinst", "object", "nonshppict", "pgdsctbl"]);
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (c === "{") {
      volcar(); pila.push({ ...est }); i++;
      continue;
    }
    if (c === "}") {
      volcar();
      if (pict && pila.length === profundidadPict) {
        if (pict.tipo && pict.hex.length > 16) {
          const bin = new Uint8Array(pict.hex.length >> 1);
          for (let k = 0; k < bin.length; k++) bin[k] = parseInt(pict.hex.substr(k * 2, 2), 16);
          const dim = dimensionesImagen(bin);
          if (dim) {
            if (corridas.length || texto) finParrafo();
            bloques.push({ tipo: "img", datos: bin, ext: dim.ext, ancho: pict.ancho || dim.ancho * PX_A_PT, alto: pict.alto || dim.alto * PX_A_PT });
          }
        } else sinSoporte++;
        pict = null; profundidadPict = -1;
      }
      if (enColores && pila.length) enColores = false;
      est = pila.pop() ?? est; i++;
      continue;
    }
    if (c === "\\") {
      const m = /^\\([a-z]+)(-?\d+)? ?|^\\'([0-9a-f]{2})|^\\(.)/i.exec(s.slice(i, i + 40));
      if (!m) { i++; continue; }
      i += m[0].length;
      if (m[3]) {
        if (saltarBytes > 0) { saltarBytes--; continue; }
        if (!est.saltar) texto += cp1252.decode(new Uint8Array([parseInt(m[3], 16)]));
        continue;
      }
      if (m[4]) {
        const ch = m[4];
        if (ch === "*") est.saltar = true;
        else if (ch === "~" && !est.saltar) texto += " ";
        else if (ch === "\n" || ch === "\r") { if (!est.saltar) finParrafo(); }
        else if (!est.saltar && "\\{}".includes(ch)) texto += ch;
        continue;
      }
      const pal = m[1], n = m[2] !== undefined ? +m[2] : undefined;
      if (DESTINOS_IGNORADOS.has(pal)) { est.saltar = true; continue; }
      if (pal === "colortbl") { enColores = true; est.saltar = true; continue; }
      if (enColores) {
        if (pal === "red") colorParcial.r = n ?? 0;
        else if (pal === "green") colorParcial.g = n ?? 0;
        else if (pal === "blue") colorParcial.b = n ?? 0;
        continue;
      }
      if (pal === "pict") { volcar(); pict = { hex: "", ancho: 0, alto: 0, tipo: null }; profundidadPict = pila.length; est.saltar = true; continue; }
      if (pict) {
        if (pal === "pngblip") pict.tipo = "png";
        else if (pal === "jpegblip") pict.tipo = "jpeg";
        else if (pal === "picwgoal") pict.ancho = (n ?? 0) / 20;
        else if (pal === "pichgoal") pict.alto = (n ?? 0) / 20;
        continue;
      }
      if (est.saltar) continue;
      switch (pal) {
        case "par": finParrafo(); break;
        case "line": volcar(); corridas.push({ texto: "", salto: true }); break;
        case "tab": texto += "\t"; break;
        case "cell": texto += "\t"; break;
        case "row": finParrafo(); break;
        case "page": finParrafo(); bloques.push({ tipo: "salto" }); break;
        case "pard": volcar(); alin = undefined; break;
        case "qc": alin = "center"; break;
        case "qr": alin = "right"; break;
        case "qj": alin = "both"; break;
        case "ql": alin = "left"; break;
        case "plain": volcar(); est = { ...est, negrita: false, cursiva: false, subrayado: false, tam: 12, color: 0 }; break;
        case "b": volcar(); est.negrita = n !== 0; break;
        case "i": volcar(); est.cursiva = n !== 0; break;
        case "ul": volcar(); est.subrayado = n !== 0; break;
        case "ulnone": volcar(); est.subrayado = false; break;
        case "fs": volcar(); est.tam = (n ?? 24) / 2; break;
        case "cf": volcar(); est.color = n ?? 0; break;
        case "uc": est.uc = n ?? 1; break;
        case "u": texto += String.fromCharCode(n! < 0 ? n! + 65536 : n!); saltarBytes = est.uc; break;
        case "emdash": texto += "—"; break;
        case "endash": texto += "–"; break;
        case "bullet": texto += "•"; break;
        case "lquote": texto += "‘"; break;
        case "rquote": texto += "’"; break;
        case "ldblquote": texto += "“"; break;
        case "rdblquote": texto += "”"; break;
      }
      continue;
    }
    if (enColores && c === ";") {
      colores.push([colorParcial.r, colorParcial.g, colorParcial.b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase());
      colorParcial = { r: 0, g: 0, b: 0 };
      i++;
      continue;
    }
    if (pict && pila.length === profundidadPict && /[0-9a-f]/i.test(c)) { pict.hex += c; i++; continue; }
    if (pict) { i++; continue; }
    if (c === "\r" || c === "\n") { i++; continue; }
    if (saltarBytes > 0) { saltarBytes--; i++; continue; }
    if (!est.saltar) texto += c;
    i++;
  }
  if (texto || corridas.length) finParrafo();
  const extraidas = bloques.filter((b) => b.tipo === "img").length;
  if (sinSoporte > extraidas) avisos.push(`${sinSoporte - extraidas} imágenes del RTF en formato no compatible (WMF/EMF) se omitieron.`);
  return { bloques, avisos };
}

// ---------------------------------------------------------------- DOC (Word 97-2003 binario): solo texto
export function leerDoc(datos: Uint8Array): DocumentoFlujo {
  const cfb = CFB.read(datos, { type: "array" });
  const wd = CFB.find(cfb, "WordDocument");
  if (!wd?.content) throw new Error("El archivo .doc no tiene el flujo WordDocument.");
  const w = Uint8Array.from(wd.content as ArrayLike<number>);
  const dvw = new DataView(w.buffer);
  const flags = dvw.getUint16(0x0a, true);
  if (flags & 0x0100) throw new Error("El archivo .doc está protegido con contraseña.");
  const tabla = CFB.find(cfb, flags & 0x0200 ? "1Table" : "0Table");
  if (!tabla?.content) throw new Error("El archivo .doc no tiene tabla de piezas.");
  const t = Uint8Array.from(tabla.content as ArrayLike<number>);
  const dvt = new DataView(t.buffer);
  const fcClx = dvw.getUint32(0x01a2, true), lcbClx = dvw.getUint32(0x01a6, true);
  let o = fcClx;
  const fin = fcClx + lcbClx;
  let texto = "";
  while (o < fin) {
    const tipo = t[o];
    if (tipo === 0x01) { o += 3 + dvt.getInt16(o + 1, true); continue; }
    if (tipo !== 0x02) break;
    const lcb = dvt.getUint32(o + 1, true);
    const plc = o + 5;
    const n = (lcb - 4) / 12;
    for (let k = 0; k < n; k++) {
      const cpIni = dvt.getUint32(plc + k * 4, true), cpFin = dvt.getUint32(plc + (k + 1) * 4, true);
      const pcd = plc + (n + 1) * 4 + k * 8;
      const fc = dvt.getUint32(pcd + 2, true);
      const comprimido = (fc & 0x40000000) !== 0;
      const largo = cpFin - cpIni;
      if (comprimido) {
        const ini = (fc & 0x3fffffff) / 2;
        texto += new TextDecoder("windows-1252").decode(w.subarray(ini, ini + largo));
      } else {
        texto += new TextDecoder("utf-16le").decode(w.subarray(fc, fc + largo * 2));
      }
    }
    break;
  }
  // quitar códigos de campo (\x13 instrucción \x14 resultado \x15) conservando el resultado
  texto = texto.replace(/\x13[^\x13\x14\x15]*\x14([^\x15]*)\x15/g, "$1").replace(/\x13[^\x15]*\x15/g, "");
  const bloques: BloqueF[] = [];
  for (const parte of texto.split(/\r/)) {
    if (parte.includes("\x0c")) {
      parte.split("\x0c").forEach((p, idx) => {
        if (idx) bloques.push({ tipo: "salto" });
        bloques.push({ tipo: "p", corridas: [{ texto: limpiar(p) }] });
      });
    } else bloques.push({ tipo: "p", corridas: [{ texto: limpiar(parte) }] });
  }
  return { bloques, avisos: ["Archivo .doc binario: se convierte el texto; formato, tablas e imágenes no se conservan."] };
}

function limpiar(s: string) {
  return s.replace(/\x07/g, "\t").replace(/\x0b/g, " ").replace(/[\x00-\x08\x0e-\x1f]/g, "");
}

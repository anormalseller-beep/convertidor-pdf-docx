// Modelo de documento "de flujo" (sin posiciones fijas) para entradas que no son PDF:
// TXT, RTF, ODT, DOCX y DOC. Se escribe con la librería `docx` (DOCX) o en Word 2003 XML (DOC).

export interface CorridaF {
  texto: string;
  negrita?: boolean;
  cursiva?: boolean;
  subrayado?: boolean;
  tam?: number; // pt
  color?: string; // RRGGBB
  fuente?: string;
  salto?: boolean; // salto de línea dentro del párrafo
}

export type Alineacion = "left" | "center" | "right" | "both";

export type BloqueF =
  | { tipo: "p"; corridas: CorridaF[]; encabezado?: number; alineacion?: Alineacion; lista?: boolean }
  | { tipo: "img"; datos: Uint8Array; ext: "png" | "jpeg" | "gif" | "bmp"; ancho: number; alto: number } // pt
  | { tipo: "tabla"; filas: CorridaF[][][] }
  | { tipo: "salto" };

export interface DocumentoFlujo {
  bloques: BloqueF[];
  avisos: string[];
}

/** Lee ancho/alto en píxeles de PNG, JPEG, GIF o BMP. */
export function dimensionesImagen(d: Uint8Array): { ancho: number; alto: number; ext: "png" | "jpeg" | "gif" | "bmp" } | null {
  const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
  if (d[0] === 0x89 && d[1] === 0x50) return { ancho: dv.getUint32(16), alto: dv.getUint32(20), ext: "png" };
  if (d[0] === 0x47 && d[1] === 0x49) return { ancho: dv.getUint16(6, true), alto: dv.getUint16(8, true), ext: "gif" };
  if (d[0] === 0x42 && d[1] === 0x4d) return { ancho: dv.getInt32(18, true), alto: Math.abs(dv.getInt32(22, true)), ext: "bmp" };
  if (d[0] === 0xff && d[1] === 0xd8) {
    let o = 2;
    while (o < d.length - 9) {
      if (d[o] !== 0xff) { o++; continue; }
      const m = d[o + 1];
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc)
        return { alto: dv.getUint16(o + 5), ancho: dv.getUint16(o + 7), ext: "jpeg" };
      o += 2 + dv.getUint16(o + 2);
    }
  }
  return null;
}

/** Tokenizador XML/HTML mínimo (sin DOM: funciona dentro de un Web Worker). */
export type Token =
  | { t: "abre"; nombre: string; attrs: Record<string, string>; vacio: boolean }
  | { t: "cierra"; nombre: string }
  | { t: "texto"; texto: string };

const ENTIDADES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };

export function decodificarEntidades(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e: string) => {
    if (e[0] === "#") {
      const n = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(n) ? String.fromCodePoint(n) : m;
    }
    return ENTIDADES[e.toLowerCase()] ?? m;
  });
}

export function* tokenizar(xml: string): Generator<Token> {
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<[?!][^>]*>|<\/([\w:.-]+)\s*>|<([\w:.-]+)((?:\s+[\w:.-]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*(\/?)>|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    if (m[1] !== undefined) yield { t: "texto", texto: m[1] };
    else if (m[2]) yield { t: "cierra", nombre: m[2] };
    else if (m[3]) {
      const attrs: Record<string, string> = {};
      const ra = /([\w:.-]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g;
      let a: RegExpExecArray | null;
      while ((a = ra.exec(m[4] || ""))) attrs[a[1]] = decodificarEntidades(a[2] ?? a[3] ?? a[4] ?? "");
      yield { t: "abre", nombre: m[3], attrs, vacio: m[5] === "/" };
    } else if (m[6] !== undefined) yield { t: "texto", texto: decodificarEntidades(m[6]) };
  }
}

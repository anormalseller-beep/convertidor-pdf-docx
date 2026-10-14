// Utilidades compartidas por los escritores DOCX (2007+) y DOC (Word 2003 XML).
import type { Corrida, PaginaModelo } from "./tipos";
import type { SegmentoRot } from "./interprete";

export type Medidor = (texto: string, fuente: string, tam: number, negrita: boolean, cursiva: boolean) => number | null;

export const EMU_POR_PT = 12700;
export const emu = (pt: number) => Math.round(pt * EMU_POR_PT);
export const twip = (pt: number) => Math.round(pt * 20);

export function esc(s: string): string {
  return s
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "")
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Descendente (fracción del em) que Word usa para cada fuente; fija la línea base. */
const DESCENDENTE: Record<string, number> = {
  Arial: 0.212, "Arial Narrow": 0.212, "Times New Roman": 0.216, Calibri: 0.25, Cambria: 0.222,
  Georgia: 0.219, "Courier New": 0.3, Verdana: 0.21, Tahoma: 0.207, "Trebuchet MS": 0.222,
  "Segoe UI": 0.25, Symbol: 0.216, Wingdings: 0.212,
};
// +0.037 em: corrección medida contra Microsoft Word 16 (desvío mediano de 0,41 pt en 11 pt)
export const descendente = (f: string) => (DESCENDENTE[f] ?? 0.22) + 0.037;

export const FACTOR_LINEA = 1.25;

export interface CajaTexto {
  x: number;
  y: number;
  ancho: number;
  alto: number;
  rot: number; // grados, horario
}

/** Caja (sin rotar) que Word debe dibujar para que la línea base caiga donde estaba en el PDF. */
export function cajaSegmento(s: SegmentoRot): CajaTexto {
  const f = s.corridas[0]?.estilo.fuente ?? "Arial";
  const alto = s.tam * FACTOR_LINEA;
  const ancho = s.ancho + Math.max(2, s.tam * 0.6);
  const d = descendente(f) * s.tam;
  if (!s.rot) return { x: s.x, y: s.base - (alto - d), ancho, alto, rot: 0 };
  // vector desde el inicio de la línea base hasta el centro de la caja, rotado
  const vx = ancho / 2, vy = alto / 2 - (alto - d);
  const t = (s.rot * Math.PI) / 180;
  const cx = s.x + vx * Math.cos(t) - vy * Math.sin(t);
  const cy = s.base + vx * Math.sin(t) + vy * Math.cos(t);
  return { x: cx - ancho / 2, y: cy - alto / 2, ancho, alto, rot: s.rot };
}

/** Un cuadro de texto: una o varias líneas consecutivas de un mismo párrafo. */
export interface BloqueTexto {
  caja: CajaTexto;
  lineas: { seg: SegmentoRot; altoTw: number }[];
  orden: number;
}

/**
 * Agrupa líneas consecutivas alineadas a la misma X, con el mismo tamaño y un interlineado
 * regular, en un solo cuadro de texto (una línea = un párrafo de interlineado exacto).
 * Reduce unas 10 veces la cantidad de objetos y Word puede abrir documentos de miles de páginas,
 * sin mover ninguna línea: el alto de cada párrafo se reparte en twips sin acumular redondeo.
 */
export function bloquesDeTexto(segs: SegmentoRot[]): BloqueTexto[] {
  const out: BloqueTexto[] = [];
  const abiertos: { lineas: SegmentoRot[]; paso: number | null }[] = [];
  const cerrar = (b: { lineas: SegmentoRot[]; paso: number | null }) => {
    if (b.lineas.length === 1 || b.paso === null) {
      for (const s of b.lineas) out.push({ caja: cajaSegmento(s), lineas: [{ seg: s, altoTw: twip(s.tam * FACTOR_LINEA) }], orden: s.orden });
      return;
    }
    const paso = b.paso;
    const p0 = b.lineas[0];
    const d = descendente(p0.corridas[0]?.estilo.fuente ?? "Arial") * p0.tam;
    const x = Math.min(...b.lineas.map((s) => s.x));
    const fin = Math.max(...b.lineas.map((s) => s.x + s.ancho));
    const n = b.lineas.length;
    out.push({
      caja: { x, y: p0.base - (paso - d), ancho: fin - x + Math.max(2, p0.tam * 0.6), alto: paso * n, rot: 0 },
      lineas: b.lineas.map((s, k) => ({ seg: s, altoTw: twip((k + 1) * paso) - twip(k * paso) })),
      orden: Math.min(...b.lineas.map((s) => s.orden)),
    });
  };
  const rectos = segs.filter((s) => !s.rot).sort((a, b) => a.base - b.base || a.x - b.x);
  for (const s of rectos) {
    let elegido: (typeof abiertos)[number] | null = null;
    for (let i = abiertos.length - 1; i >= 0; i--) {
      const b = abiertos[i];
      const u = b.lineas[b.lineas.length - 1];
      const d = s.base - u.base;
      // bloque que ya quedó atrás: se cierra
      if (d > Math.max(u.tam, s.tam) * 2.5) {
        cerrar(b);
        abiertos.splice(i, 1);
        continue;
      }
      if (elegido || Math.abs(u.x - s.x) > 0.5 || Math.abs(u.tam - s.tam) > 0.3) continue;
      const ok = b.paso === null ? d >= 1.05 * s.tam && d <= 2.2 * s.tam : Math.abs(d - b.paso) < 0.3;
      if (ok) elegido = b;
    }
    if (elegido) {
      if (elegido.paso === null) elegido.paso = s.base - elegido.lineas[0].base;
      elegido.lineas.push(s);
    } else abiertos.push({ lineas: [s], paso: null });
  }
  abiertos.forEach(cerrar);
  for (const s of segs.filter((q) => q.rot)) out.push({ caja: cajaSegmento(s), lineas: [{ seg: s, altoTw: twip(s.tam * FACTOR_LINEA) }], orden: s.orden });
  return out;
}

/** Ajuste de espaciado y escala horizontal para que el texto ocupe el mismo ancho que en el PDF. */
export function ajusteAncho(c: Corrida, medir?: Medidor): { espTwips: number; escala: number } {
  const n = [...c.texto].length;
  if (!medir || n === 0 || c.ancho <= 0) return { espTwips: 0, escala: 100 };
  const tamWord = Math.max(1, Math.round(c.estilo.tam * 2)) / 2;
  const real = medir(c.texto, c.estilo.fuente, tamWord, c.estilo.negrita, c.estilo.cursiva);
  if (!real || real <= 0) return { espTwips: 0, escala: 100 };
  const razon = c.ancho / real;
  let escala = 100;
  let restante = c.ancho - real;
  if (Math.abs(razon - 1) > 0.15) {
    escala = Math.max(50, Math.min(200, Math.round(razon * 100)));
    restante = c.ancho - (real * escala) / 100;
  }
  let espPt = restante / n;
  espPt = Math.max(-c.estilo.tam * 0.25, Math.min(c.estilo.tam * 0.5, espPt));
  return { espTwips: Math.round(espPt * 20), escala };
}

/** Márgenes según la caja que ocupa el contenido (los objetos van anclados a la página). */
export function margenes(p: PaginaModelo): { sup: number; inf: number; izq: number; der: number } {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const s of p.segmentos) {
    x0 = Math.min(x0, s.x); x1 = Math.max(x1, s.x + s.ancho);
    y0 = Math.min(y0, s.base - s.tam); y1 = Math.max(y1, s.base + s.tam * 0.3);
  }
  for (const im of p.imagenes) {
    x0 = Math.min(x0, im.x); x1 = Math.max(x1, im.x + im.ancho);
    y0 = Math.min(y0, im.y); y1 = Math.max(y1, im.y + im.alto);
  }
  const lim = (v: number, max: number) => (Number.isFinite(v) ? Math.max(0, Math.min(max, v)) : 72);
  return {
    sup: lim(y0, p.alto * 0.35),
    inf: lim(p.alto - y1, p.alto * 0.35),
    izq: lim(x0, p.ancho * 0.35),
    der: lim(p.ancho - x1, p.ancho * 0.35),
  };
}

// Funciones que necesitan un canvas 2D (OffscreenCanvas en el Worker; @napi-rs/canvas en pruebas).
import type { ComandoVector } from "./interprete";
import type { Medidor } from "./comun";

// Subconjunto del contexto 2D que usamos (compatible con OffscreenCanvas y @napi-rs/canvas)
export interface Ctx2D {
  font: string;
  fillStyle: unknown;
  strokeStyle: unknown;
  lineWidth: number;
  lineJoin: string;
  lineCap: string;
  measureText(t: string): { width: number };
  beginPath(): void;
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  bezierCurveTo(a: number, b: number, c: number, d: number, e: number, f: number): void;
  closePath(): void;
  fill(regla?: "nonzero" | "evenodd"): void;
  stroke(): void;
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void;
}

/** Resolución de la capa de líneas: 2 px por punto (144 ppp). */
export const ESCALA_VECTORES = 2;

export function dibujarVectores(ctx: Ctx2D, cmds: ComandoVector[], escala = ESCALA_VECTORES) {
  ctx.setTransform(escala, 0, 0, escala, 0, 0);
  ctx.lineJoin = "miter";
  ctx.lineCap = "butt";
  for (const c of cmds) {
    ctx.beginPath();
    for (const s of c.ruta) {
      if (s[0] === "M") ctx.moveTo(s[1], s[2]);
      else if (s[0] === "L") ctx.lineTo(s[1], s[2]);
      else if (s[0] === "C") ctx.bezierCurveTo(s[1], s[2], s[3], s[4], s[5], s[6]);
      else ctx.closePath();
    }
    if (c.relleno) {
      ctx.fillStyle = "#" + c.relleno;
      ctx.fill(c.parImpar ? "evenodd" : "nonzero");
    }
    if (c.trazo) {
      ctx.strokeStyle = "#" + c.trazo;
      // mínimo 1 px de la capa para que las líneas finas no desaparezcan
      ctx.lineWidth = Math.max(c.grosor, 0.5);
      ctx.stroke();
    }
  }
}

/** Medidor de texto con detección de fuentes ausentes (devuelve null si no está instalada). */
export function crearMedidor(ctx: Ctx2D): Medidor {
  const disponibles = new Map<string, boolean>();
  const prueba = "mmmmmmmmmlliWWQ@#0123456789abcdefghij";
  const disponible = (f: string) => {
    let v = disponibles.get(f);
    if (v === undefined) {
      v = false;
      for (const gen of ["monospace", "serif", "sans-serif"]) {
        ctx.font = `40px ${gen}`;
        const a = ctx.measureText(prueba).width;
        ctx.font = `40px "${f}", ${gen}`;
        const b = ctx.measureText(prueba).width;
        if (Math.abs(a - b) > 0.01) { v = true; break; }
      }
      disponibles.set(f, v);
    }
    return v;
  };
  return (texto, fuente, tam, negrita, cursiva) => {
    if (!disponible(fuente)) return null;
    // se mide a 100 px y se escala: más precisión que medir a tamaños chicos
    ctx.font = `${cursiva ? "italic " : ""}${negrita ? "bold " : ""}100px "${fuente}"`;
    return (ctx.measureText(texto).width * tam) / 100;
  };
}

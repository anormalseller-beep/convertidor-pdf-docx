// Mapeo de fuentes del PDF a fuentes estándar seguras con métricas parecidas.
// Cada familia del PDF conserva su propia sustituta (no hay una fuente global).

export interface FuentePdf {
  name?: string;
  loadedName?: string;
  fallbackName?: string;
  bold?: boolean;
  black?: boolean;
  italic?: boolean;
  fontMatrix?: number[];
  isType3Font?: boolean;
  vertical?: boolean;
}

const REGLAS: [RegExp, string][] = [
  [/calibri|carlito/i, "Calibri"],
  [/cambria|caladea/i, "Cambria"],
  [/georgia|gelasio/i, "Georgia"],
  [/garamond|minion|palatino|book ?antiqua|baskerville|caslon|bodoni|century|charter|libertin|cmr\d|cmti|lmroman|utopia|dejavuserif|liberationserif|noto ?serif|source ?serif/i, "Times New Roman"],
  [/times|tinos|nimbusrom|roman|serif/i, "Times New Roman"],
  [/courier|cousine|nimbusmon|mono|consolas|menlo|cmtt|inconsolata|lucida ?console/i, "Courier New"],
  [/verdana/i, "Verdana"],
  [/tahoma/i, "Tahoma"],
  [/trebuchet/i, "Trebuchet MS"],
  [/segoe/i, "Segoe UI"],
  [/arial ?narrow|helvetica ?condensed|helvetica-?narrow/i, "Arial Narrow"],
  [/arial|helvetica|arimo|liberationsans|nimbussan|frutiger|univers|myriad|roboto|open ?sans|lato|source ?sans|dejavusans|noto ?sans|calibri|sans/i, "Arial"],
  [/symbol/i, "Symbol"],
  [/wingding|dingbat/i, "Wingdings"],
];

/** Quita el prefijo de subconjunto "ABCDEF+" y separadores. */
export function nombreLimpio(nombre: string): string {
  return nombre.replace(/^[A-Z]{6}\+/, "").replace(/[_,]/g, "-");
}

export function mapearFuente(f: FuentePdf | undefined): {
  fuente: string;
  original: string;
  negrita: boolean;
  cursiva: boolean;
} {
  const original = nombreLimpio(f?.name || f?.loadedName || "desconocida");
  let fuente = "";
  for (const [re, destino] of REGLAS) {
    if (re.test(original)) {
      fuente = destino;
      break;
    }
  }
  if (!fuente) {
    // Fuente rara o embebida sin nombre reconocible: se usa la familia genérica que
    // reporta pdf.js (serif / sans-serif / monospace).
    const fb = f?.fallbackName || "sans-serif";
    fuente = fb === "serif" ? "Times New Roman" : fb === "monospace" ? "Courier New" : "Arial";
  }
  const negrita = !!(f?.bold || f?.black) || /bold|black|heavy|semibold|demi|-bd\b|,bd/i.test(original);
  const cursiva = !!f?.italic || /italic|oblique|-it\b|kursiv/i.test(original);
  return { fuente, original, negrita, cursiva };
}

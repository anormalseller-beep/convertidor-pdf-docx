// Orquestador: recorre el PDF en bloques de páginas, interpreta cada página, la escribe
// en el DOCX/DOC incrementalmente y libera memoria de pdf.js entre bloques.
import type { PDFDocumentProxy } from "pdfjs-dist";
import { interpretarPagina, type Entorno, type PaginaInterpretada } from "./interprete";
import { EscritorDocx, type ResultadoEscritor } from "./escritor-docx";
import { EscritorDoc2003 } from "./escritor-doc";
import type { FormatoSalida } from "./tipos";
import type { Medidor } from "./comun";

export interface Progreso {
  fase: "leyendo" | "convirtiendo" | "ocr" | "empaquetando" | "verificando";
  pagina: number;
  total: number;
  porcentaje: number;
}

export interface OpcionesConversion {
  formato: FormatoSalida;
  tamBloque?: number; // páginas por bloque (50-100)
  ocr?: (pagina: PaginaInterpretada) => Promise<PaginaInterpretada>; // solo páginas escaneadas
  cancelado?: () => boolean;
  alProgreso?: (p: Progreso) => void;
  alAviso?: (t: string) => void;
  titulo?: string;
}

export interface ResultadoConversion extends ResultadoEscritor {
  paginasPdf: number;
  imagenesPdf: number;
  paginasEscaneadas: number;
  capasVectoriales: number;
  /** imágenes escaneadas que el OCR reemplazó por texto (no se esperan en la salida) */
  reemplazadasOcr: number;
  tamPaginasPdf: { ancho: number; alto: number }[];
  avisos: string[];
}

export class ConversionCancelada extends Error {
  constructor() {
    super("Conversión cancelada.");
  }
}

export async function convertirPdf(doc: PDFDocumentProxy, entorno: Entorno & { medirTexto?: Medidor }, op: OpcionesConversion): Promise<ResultadoConversion> {
  const total = doc.numPages;
  const bloque = Math.max(1, Math.min(100, op.tamBloque ?? 50));
  entorno.cacheImagenes ??= new Map();
  const escritor = op.formato === "doc" ? new EscritorDoc2003(entorno.medirTexto, op.titulo) : new EscritorDocx(entorno.medirTexto, op.titulo);
  const avisos: string[] = [];
  const tamPaginasPdf: { ancho: number; alto: number }[] = [];
  let imagenesPdf = 0;
  let escaneadas = 0;
  let capas = 0;
  let reemplazadasOcr = 0;
  const aviso = (t: string) => {
    avisos.push(t);
    op.alAviso?.(t);
  };

  for (let ini = 1; ini <= total; ini += bloque) {
    const fin = Math.min(total, ini + bloque - 1);
    for (let n = ini; n <= fin; n++) {
      if (op.cancelado?.()) throw new ConversionCancelada();
      const pagina = await doc.getPage(n);
      let modelo = await interpretarPagina(pagina, n, entorno);
      tamPaginasPdf.push({ ancho: modelo.ancho, alto: modelo.alto });
      if (modelo.escaneada) {
        escaneadas++;
        if (op.ocr) {
          op.alProgreso?.({ fase: "ocr", pagina: n, total, porcentaje: ((n - 0.5) / total) * 95 });
          const antes = modelo.imagenesEnPdf;
          modelo = await op.ocr(modelo);
          reemplazadasOcr += antes - modelo.imagenesEnPdf;
        }
      }
      imagenesPdf += modelo.imagenesEnPdf;
      if (modelo.vectores) capas++;
      modelo.avisos.forEach(aviso);
      escritor.agregarPagina(modelo, n === total);
      pagina.cleanup();
      op.alProgreso?.({ fase: "convirtiendo", pagina: n, total, porcentaje: (n / total) * 95 });
    }
    // libera fuentes e imágenes decodificadas del bloque
    await doc.cleanup();
  }
  if (op.cancelado?.()) throw new ConversionCancelada();
  op.alProgreso?.({ fase: "empaquetando", pagina: total, total, porcentaje: 96 });
  const r = await escritor.finalizar();
  return { ...r, paginasPdf: total, imagenesPdf, paginasEscaneadas: escaneadas, capasVectoriales: capas, reemplazadasOcr, tamPaginasPdf, avisos };
}

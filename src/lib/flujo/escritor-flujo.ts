// Escribe un DocumentoFlujo como DOCX (librería `docx`) o como DOC (Word 2003 XML).
// Página A4 con márgenes de 2,5 cm. Las imágenes conservan su relación de aspecto y solo
// se reducen (nunca se estiran) si son más anchas que el área de texto.
import {
  AlignmentType, Document, HeadingLevel, ImageRun, Packer, PageBreak, Paragraph, Table, TableCell, TableRow, TextRun, WidthType, BorderStyle,
} from "docx";
import type { Alineacion, BloqueF, CorridaF, DocumentoFlujo } from "./modelo";

const A4 = { ancho: 595.28, alto: 841.89 };
const MARGEN = 70.87; // 2,5 cm
const ANCHO_UTIL = A4.ancho - 2 * MARGEN;

function ajustarImagen(ancho: number, alto: number) {
  const f = ancho > ANCHO_UTIL ? ANCHO_UTIL / ancho : 1;
  const altoMax = A4.alto - 2 * MARGEN;
  const g = alto * f > altoMax ? altoMax / (alto * f) : 1;
  return { ancho: ancho * f * g, alto: alto * f * g };
}

const ALIN: Record<Alineacion, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, both: AlignmentType.JUSTIFIED,
};
const NIVELES = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];

function corridas(cs: CorridaF[]) {
  return cs.map((c) =>
    new TextRun({
      text: c.texto,
      break: c.salto ? 1 : undefined,
      bold: c.negrita, italics: c.cursiva, underline: c.subrayado ? {} : undefined,
      size: c.tam ? Math.round(c.tam * 2) : undefined, color: c.color, font: c.fuente,
    }),
  );
}

export async function flujoADocx(doc: DocumentoFlujo): Promise<{ blob: Blob; imagenes: number }> {
  const hijos: (Paragraph | Table)[] = [];
  let imagenes = 0;
  for (const b of doc.bloques) {
    if (b.tipo === "p") {
      hijos.push(new Paragraph({ children: corridas(b.corridas), heading: b.encabezado ? NIVELES[Math.min(6, b.encabezado) - 1] : undefined, alignment: b.alineacion ? ALIN[b.alineacion] : undefined, bullet: b.lista ? { level: 0 } : undefined }));
    } else if (b.tipo === "img") {
      const { ancho, alto } = ajustarImagen(b.ancho, b.alto);
      imagenes++;
      hijos.push(new Paragraph({ children: [new ImageRun({ type: b.ext === "jpeg" ? "jpg" : b.ext, data: b.datos, transformation: { width: (ancho * 96) / 72, height: (alto * 96) / 72 } })] }));
    } else if (b.tipo === "tabla") {
      const cols = Math.max(1, ...b.filas.map((f) => f.length));
      const borde = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
      hijos.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: { top: borde, bottom: borde, left: borde, right: borde, insideHorizontal: borde, insideVertical: borde },
        rows: b.filas.map((f) => new TableRow({
          children: Array.from({ length: cols }, (_, i) => new TableCell({ children: [new Paragraph({ children: corridas(f[i] ?? []) })] })),
        })),
      }));
    } else hijos.push(new Paragraph({ children: [new PageBreak()] }));
  }
  if (!hijos.length) hijos.push(new Paragraph({}));
  const d = new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: { page: { size: { width: Math.round(A4.ancho * 20), height: Math.round(A4.alto * 20) }, margin: { top: Math.round(MARGEN * 20), bottom: Math.round(MARGEN * 20), left: Math.round(MARGEN * 20), right: Math.round(MARGEN * 20) } } },
      children: hijos,
    }],
  });
  return { blob: await Packer.toBlob(d), imagenes };
}

const e = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");

function base64(d: Uint8Array) {
  let s = "";
  for (let i = 0; i < d.length; i += 0x8000) s += String.fromCharCode(...d.subarray(i, i + 0x8000));
  return btoa(s);
}

function corridas2003(cs: CorridaF[]) {
  return cs.map((c) => {
    if (c.salto) return `<w:r><w:br/></w:r>`;
    const rpr =
      (c.fuente ? `<w:rFonts w:ascii="${e(c.fuente)}" w:h-ansi="${e(c.fuente)}"/>` : "") +
      (c.negrita ? "<w:b/>" : "") + (c.cursiva ? "<w:i/>" : "") + (c.subrayado ? '<w:u w:val="single"/>' : "") +
      (c.color ? `<w:color w:val="${c.color}"/>` : "") + (c.tam ? `<w:sz w:val="${Math.round(c.tam * 2)}"/>` : "");
    return `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t>${e(c.texto)}</w:t></w:r>`;
  }).join("");
}

export function flujoADoc2003(doc: DocumentoFlujo): { blob: Blob; imagenes: number } {
  const partes: string[] = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n` +
      `<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xml:space="preserve">` +
      `<w:fonts><w:defaultFonts w:ascii="Calibri" w:fareast="Calibri" w:h-ansi="Calibri" w:cs="Calibri"/></w:fonts>` +
      `<w:styles><w:style w:type="paragraph" w:default="on" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style></w:styles><w:body>`,
  ];
  let n = 0;
  const p = (b: Extract<BloqueF, { tipo: "p" }>) => {
    const ppr =
      (b.alineacion ? `<w:jc w:val="${b.alineacion}"/>` : "") +
      (b.lista ? '<w:ind w:left="360" w:hanging="360"/>' : "");
    const titulo = b.encabezado ? b.corridas.map((c) => ({ ...c, negrita: true, tam: c.tam ?? [20, 16, 14, 12, 11, 11][b.encabezado! - 1] })) : b.corridas;
    return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${b.lista ? "<w:r><w:t>• </w:t></w:r>" : ""}${corridas2003(titulo)}</w:p>`;
  };
  for (const b of doc.bloques) {
    if (b.tipo === "p") partes.push(p(b));
    else if (b.tipo === "img") {
      const { ancho, alto } = ajustarImagen(b.ancho, b.alto);
      const nombre = `wordml://${String(++n).padStart(8, "0")}.${b.ext === "jpeg" ? "jpg" : b.ext}`;
      partes.push(`<w:p><w:r><w:pict><w:binData w:name="${nombre}">${base64(b.datos)}</w:binData><v:shape id="i${n}" style="width:${ancho.toFixed(2)}pt;height:${alto.toFixed(2)}pt"><v:imagedata src="${nombre}" o:title=""/></v:shape></w:pict></w:r></w:p>`);
    } else if (b.tipo === "tabla") {
      const borde = '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="999999"/><w:left w:val="single" w:sz="4" w:color="999999"/><w:bottom w:val="single" w:sz="4" w:color="999999"/><w:right w:val="single" w:sz="4" w:color="999999"/><w:insideH w:val="single" w:sz="4" w:color="999999"/><w:insideV w:val="single" w:sz="4" w:color="999999"/></w:tblBorders>';
      partes.push(`<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>${borde}</w:tblPr>` +
        b.filas.map((f) => `<w:tr>${f.map((c) => `<w:tc><w:p>${corridas2003(c)}</w:p></w:tc>`).join("")}</w:tr>`).join("") + `</w:tbl>`);
    } else partes.push(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
  }
  partes.push(
    `<w:sectPr><w:pgSz w:w="${Math.round(A4.ancho * 20)}" w:h="${Math.round(A4.alto * 20)}"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:wordDocument>`,
  );
  return { blob: new Blob(partes, { type: "application/msword" }), imagenes: n };
}

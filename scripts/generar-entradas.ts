// Crea pruebas/entradas/fuente.docx (encabezado, negrita, cursiva, tabla, imagen) y fuente.txt.
// Luego scripts/word-guardar-como.ps1 produce .doc, .rtf y .odt con Microsoft Word.
import { writeFileSync, mkdirSync } from "node:fs";
import { Document, HeadingLevel, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";
import { codificarPng } from "../src/lib/motor/png";

const px = new Uint8Array(200 * 100 * 4);
for (let i = 0; i < 200 * 100; i++) { px[i * 4] = 0; px[i * 4 + 1] = 102; px[i * 4 + 2] = (i % 200) + 50; px[i * 4 + 3] = 255; }

async function main() {
  mkdirSync("pruebas/entradas", { recursive: true });
  const d = new Document({
    sections: [{
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Documento de entrada")] }),
        new Paragraph({ children: [new TextRun("Texto normal, "), new TextRun({ text: "negrita", bold: true }), new TextRun(", "), new TextRun({ text: "cursiva", italics: true }), new TextRun(" y acentos: canción, ñandú.")] }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [["Mes", "Valor"], ["Enero", "120"], ["Febrero", "95"]].map((f) => new TableRow({ children: f.map((c) => new TableCell({ children: [new Paragraph(c)] })) })),
        }),
        new Paragraph({ children: [new ImageRun({ type: "png", data: codificarPng(200, 100, px, false), transformation: { width: 200, height: 100 } })] }),
        new Paragraph("Último párrafo."),
      ],
    }],
  });
  writeFileSync("pruebas/entradas/fuente.docx", new Uint8Array(await (await Packer.toBlob(d)).arrayBuffer()));
  writeFileSync("pruebas/entradas/fuente.txt", "Primera línea de texto plano.\n\nTercera línea con acentos: canción, ñandú.\n");
  console.log("OK");
}
main();

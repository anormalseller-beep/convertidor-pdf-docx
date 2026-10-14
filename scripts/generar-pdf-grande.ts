// Genera un PDF de N páginas (por defecto 10.000) con texto, una imagen y una tabla por página.
// Uso: npx tsx scripts/generar-pdf-grande.ts [paginas] [salida]
import { writeFileSync } from "node:fs";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { codificarPng } from "../src/lib/motor/png";

const N = Number(process.argv[2] || 10000);
const salida = process.argv[3] || "pruebas/grande-10000-paginas.pdf";

async function main() {
  const doc = await PDFDocument.create();
  const f = await doc.embedFont(StandardFonts.TimesRoman);
  const fb = await doc.embedFont(StandardFonts.HelveticaBold);
  const px = new Uint8Array(160 * 80 * 4);
  for (let i = 0; i < 160 * 80; i++) { px[i * 4] = (i % 160) + 60; px[i * 4 + 1] = 120; px[i * 4 + 2] = 200 - Math.floor(i / 160); px[i * 4 + 3] = 255; }
  const img = await doc.embedPng(codificarPng(160, 80, px, true)); // la misma imagen se reutiliza (como un logo)
  for (let n = 1; n <= N; n++) {
    const pg = doc.addPage([595.28, 841.89]);
    pg.drawText(`Capítulo ${n}`, { x: 72, y: 770, size: 16, font: fb });
    pg.drawImage(img, { x: 595.28 - 72 - 120, y: 700, width: 120, height: 60 });
    for (let l = 0; l < 30; l++)
      pg.drawText(`Línea ${l + 1} de la página ${n}: texto de relleno para medir el rendimiento del conversor.`, { x: 72, y: 680 - l * 15, size: 10.5, font: f });
    pg.drawRectangle({ x: 72, y: 150, width: 451, height: 60, borderColor: rgb(0, 0, 0), borderWidth: 0.8 });
    pg.drawText(`Total página ${n}`, { x: 80, y: 175, size: 11, font: fb });
    pg.drawText(`${n}`, { x: 500, y: 40, size: 9, font: f });
  }
  writeFileSync(salida, await doc.save({ objectsPerTick: 2000 }));
  console.log("OK", salida);
}
main();

// Compara las posiciones de texto (línea base y x) entre el PDF original y el exportado por Word.
// Uso: node scripts/medir-desvio.mjs original.pdf word.pdf
import { readFileSync } from "node:fs";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";

const abrir = (f) => pdfjs.getDocument({ data: new Uint8Array(readFileSync(f)), verbosity: 0, standardFontDataUrl: "node_modules/pdfjs-dist/standard_fonts/" }).promise;
const [a, b] = await Promise.all([abrir(process.argv[2]), abrir(process.argv[3])]);
const dy = [], dx = [], dw = [];
let sinPar = 0;
for (let n = 1; n <= Math.min(a.numPages, b.numPages); n++) {
  const ta = (await (await a.getPage(n)).getTextContent()).items.filter((i) => i.str.trim().length > 3);
  const tb = (await (await b.getPage(n)).getTextContent()).items.filter((i) => i.str.trim().length > 3);
  for (const i of ta) {
    const j = tb.find((k) => k.str.trim() === i.str.trim());
    if (!j) { sinPar++; continue; }
    dy.push(j.transform[5] - i.transform[5]);
    dx.push(j.transform[4] - i.transform[4]);
    dw.push(j.width - i.width);
  }
}
const est = (v) => { const s = [...v].sort((x, y) => x - y); return `mediana ${s[s.length >> 1]?.toFixed(2)} min ${s[0]?.toFixed(2)} max ${s[s.length - 1]?.toFixed(2)}`; };
console.log(`líneas comparadas ${dy.length}, sin par ${sinPar}`);
console.log("desvío vertical (pt, + = más arriba en Word):", est(dy));
console.log("desvío horizontal (pt):", est(dx));
console.log("diferencia de ancho (pt):", est(dw));

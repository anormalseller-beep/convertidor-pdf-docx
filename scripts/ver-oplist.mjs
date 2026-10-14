// Diagnóstico: imprime los operadores de una página con pdfjs (Node).
import { readFileSync } from "node:fs";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
const datos = new Uint8Array(readFileSync(process.argv[2]));
const doc = await pdfjs.getDocument({ data: datos, isOffscreenCanvasSupported: false, disableFontFace: true, useSystemFonts: false }).promise;
const pg = await doc.getPage(Number(process.argv[3] || 1));
const ops = await pg.getOperatorList();
const nombres = Object.fromEntries(Object.entries(pdfjs.OPS).map(([k, v]) => [v, k]));
ops.fnArray.slice(0, 60).forEach((fn, i) => {
  const a = ops.argsArray[i];
  console.log(nombres[fn], JSON.stringify(a, (k, v) => (typeof v === "object" && v && v.fontChar !== undefined ? v.unicode : v))?.slice(0, 160));
});
const tc = await pg.getTextContent();
console.log(tc.items.slice(0, 3), tc.styles);
for (const k of Object.keys(tc.styles)) { const f = pg.commonObjs.get(k); console.log(k, f.name, f.bold, f.italic, f.fallbackName, f.fontMatrix); }

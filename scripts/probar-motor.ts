// Prueba del motor en Node con el mismo código que usa el Web Worker.
// Uso: npx tsx scripts/probar-motor.ts [archivo.pdf] [docx|doc]
import { readFileSync, writeFileSync } from "node:fs";
import { basename } from "node:path";
import { createCanvas } from "@napi-rs/canvas";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { convertirPdf } from "../src/lib/motor/convertir-pdf";
import { verificar } from "../src/lib/motor/qa";
import { codificarPng } from "../src/lib/motor/png";
import { crearMedidor, dibujarVectores, ESCALA_VECTORES, type Ctx2D } from "../src/lib/motor/lienzo";

const archivo = process.argv[2] || "pruebas/muestra-12-paginas.pdf";
const formato = (process.argv[3] || "docx") as "docx" | "doc";

async function main() {
  const t0 = Date.now();
  const datos = new Uint8Array(readFileSync(archivo));
  const doc = await pdfjs.getDocument({
    data: datos,
    isOffscreenCanvasSupported: false,
    standardFontDataUrl: "node_modules/pdfjs-dist/standard_fonts/",
    cMapUrl: "node_modules/pdfjs-dist/cmaps/",
    cMapPacked: true,
    disableFontFace: true,
    verbosity: 0,
  }).promise;
  const medidor = crearMedidor(createCanvas(10, 10).getContext("2d") as unknown as Ctx2D);
  const r = await convertirPdf(doc, {
    codificarImagen: async (img) => ({ datos: codificarPng(img.ancho, img.alto, img.rgba, !img.opaca), ext: "png" }),
    rasterizarVectores: async (ancho, alto, cmds) => {
      const w = Math.ceil(ancho * ESCALA_VECTORES), h = Math.ceil(alto * ESCALA_VECTORES);
      const c = createCanvas(w, h);
      dibujarVectores(c.getContext("2d") as unknown as Ctx2D, cmds);
      return { datos: new Uint8Array(c.toBuffer("image/png")), pixAncho: w, pixAlto: h };
    },
    medirTexto: medidor,
  }, {
    formato,
    tamBloque: 50,
    titulo: basename(archivo),
    alProgreso: (p) => { if (p.pagina % 10 === 0 || p.pagina === p.total) process.stdout.write(`\r${p.fase} ${p.pagina}/${p.total} ${p.porcentaje.toFixed(0)}%   `); },
  });
  console.log();
  const qa = await verificar(r, formato, r.capasVectoriales);
  const salida = archivo.replace(/\.pdf$/i, "") + "." + formato;
  writeFileSync(salida, new Uint8Array(await r.blob.arrayBuffer()));
  console.log(`Salida: ${salida} (${(r.blob.size / 1024).toFixed(0)} KB) en ${((Date.now() - t0) / 1000).toFixed(1)} s`);
  console.log(qa.ok ? "QA OK" : "QA CON ERRORES");
  qa.lineas.forEach((l) => console.log("  " + l));
  r.avisos.forEach((a) => console.log("  aviso: " + a));
  console.log(`Memoria: ${(process.memoryUsage().rss / 1048576).toFixed(0)} MB`);
}
main().catch((e) => { console.error(e); process.exit(1); });

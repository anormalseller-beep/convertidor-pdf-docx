// Genera pruebas/muestra-12-paginas.pdf: texto, tablas, imágenes de distintas proporciones,
// colores, varias fuentes, página horizontal y tamaños Letter + A4.
// Uso: npx tsx scripts/generar-pdf-prueba.ts
import { writeFileSync, mkdirSync } from "node:fs";
import { PDFDocument, StandardFonts, rgb, PDFFont, PDFPage } from "pdf-lib";
import { codificarPng } from "../src/lib/motor/png";

const LETTER: [number, number] = [612, 792];
const A4: [number, number] = [595.28, 841.89];

function imagen(ancho: number, alto: number, pintar: (x: number, y: number) => [number, number, number, number]) {
  const px = new Uint8Array(ancho * alto * 4);
  for (let y = 0; y < alto; y++)
    for (let x = 0; x < ancho; x++) {
      const [r, g, b, a] = pintar(x, y);
      const i = (y * ancho + x) * 4;
      px[i] = r; px[i + 1] = g; px[i + 2] = b; px[i + 3] = a;
    }
  return codificarPng(ancho, alto, px, true);
}

// Paisaje 2:1 (cielo, sol, colinas)
const paisaje = imagen(600, 300, (x, y) => {
  const colina = 200 + 40 * Math.sin(x / 60) + 15 * Math.sin(x / 17);
  if (y > colina) return [60, 110 + ((x + y) % 30), 50, 255];
  const dx = x - 460, dy = y - 80;
  if (dx * dx + dy * dy < 45 * 45) return [240, 200, 60, 255];
  return [120 + y / 4, 170 + y / 6, 230, 255];
});
// Retrato 1:2.2 (franjas verticales + círculo) para detectar estiramientos
const retrato = imagen(200, 440, (x, y) => {
  const dx = x - 100, dy = y - 220;
  if (dx * dx + dy * dy < 80 * 80) return [200, 40, 40, 255];
  return Math.floor(x / 25) % 2 ? [30, 30, 30, 255] : [235, 235, 235, 255];
});
// Gráfico de barras 2:1
const grafico = imagen(800, 400, (x, y) => {
  if (y > 370 && y < 374) return [0, 0, 0, 255];
  const barra = Math.floor((x - 60) / 120);
  const alturas = [220, 300, 160, 330, 260, 120];
  if (x > 60 && barra < 6 && (x - 60) % 120 < 80 && y < 370 && y > 370 - alturas[barra])
    return barra % 2 ? [0, 102, 204, 255] : [46, 125, 50, 255];
  if (y % 74 === 0) return [220, 220, 220, 255];
  return [255, 255, 255, 255];
});
// Icono cuadrado con transparencia
const icono = imagen(64, 64, (x, y) => {
  const dx = x - 32, dy = y - 32;
  const d = Math.sqrt(dx * dx + dy * dy);
  if (d > 30) return [0, 0, 0, 0];
  return d > 24 ? [0, 102, 204, 255] : [255, 255, 255, 255];
});

const LOREM =
  "La conversión de documentos requiere conservar la posición exacta de cada elemento. " +
  "Este párrafo de prueba contiene acentos (canción, árbol, pingüino, ñandú) y signos de " +
  "puntuación para verificar que el texto se extrae como texto real y no como imagen. " +
  "Cada línea debe quedar en el mismo lugar que en el PDF original, sin reflujo ni saltos.";

function envolver(texto: string, fuente: PDFFont, tam: number, ancho: number): string[] {
  const palabras = texto.split(" ");
  const lineas: string[] = [];
  let actual = "";
  for (const p of palabras) {
    const prueba = actual ? actual + " " + p : p;
    if (fuente.widthOfTextAtSize(prueba, tam) > ancho && actual) {
      lineas.push(actual);
      actual = p;
    } else actual = prueba;
  }
  if (actual) lineas.push(actual);
  return lineas;
}

function parrafo(pg: PDFPage, texto: string, x: number, y: number, ancho: number, fuente: PDFFont, tam: number, color = rgb(0, 0, 0)) {
  const lineas = envolver(texto, fuente, tam, ancho);
  lineas.forEach((l, i) => pg.drawText(l, { x, y: y - i * tam * 1.3, size: tam, font: fuente, color }));
  return y - lineas.length * tam * 1.3;
}

async function main() {
  const doc = await PDFDocument.create();
  doc.setTitle("Documento de prueba");
  const times = await doc.embedFont(StandardFonts.TimesRoman);
  const timesB = await doc.embedFont(StandardFonts.TimesRomanBold);
  const helv = await doc.embedFont(StandardFonts.Helvetica);
  const helvB = await doc.embedFont(StandardFonts.HelveticaBold);
  const helvI = await doc.embedFont(StandardFonts.HelveticaOblique);
  const courier = await doc.embedFont(StandardFonts.Courier);
  const imgPaisaje = await doc.embedPng(paisaje);
  const imgRetrato = await doc.embedPng(retrato);
  const imgGrafico = await doc.embedPng(grafico);
  const imgIcono = await doc.embedPng(icono);
  const TOTAL = 12;

  const pie = (pg: PDFPage, n: number) => {
    const [w] = [pg.getWidth()];
    const t = `Página ${n} de ${TOTAL}`;
    pg.drawText(t, { x: w / 2 - helv.widthOfTextAtSize(t, 9) / 2, y: 30, size: 9, font: helv, color: rgb(0.4, 0.4, 0.4) });
  };

  const tabla = (pg: PDFPage, x: number, y: number, cols: number[], filas: string[][]) => {
    const alto = 22;
    const anchoTotal = cols.reduce((a, b) => a + b, 0);
    pg.drawRectangle({ x, y: y - alto, width: anchoTotal, height: alto, color: rgb(0.9, 0.9, 0.9) });
    filas.forEach((fila, i) => {
      let cx = x;
      fila.forEach((celda, j) => {
        pg.drawText(celda, { x: cx + 5, y: y - alto * (i + 1) + 7, size: 10, font: i === 0 ? helvB : helv });
        cx += cols[j];
      });
    });
    for (let i = 0; i <= filas.length; i++)
      pg.drawLine({ start: { x, y: y - i * alto }, end: { x: x + anchoTotal, y: y - i * alto }, thickness: 0.8 });
    let cx = x;
    for (let j = 0; j <= cols.length; j++) {
      pg.drawLine({ start: { x: cx, y }, end: { x: cx, y: y - filas.length * alto }, thickness: 0.8 });
      cx += cols[j] ?? 0;
    }
    return y - filas.length * alto;
  };

  for (let n = 1; n <= TOTAL; n++) {
    const tam = n === 6 ? ([LETTER[1], LETTER[0]] as [number, number]) : n <= 6 ? LETTER : A4;
    const pg = doc.addPage(tam);
    const [W, H] = tam;
    let y = H - 72;
    switch (n) {
      case 1: {
        pg.drawText("Informe de prueba de conversión", { x: 72, y, size: 24, font: timesB });
        y -= 40;
        pg.drawImage(imgPaisaje, { x: W - 72 - 240, y: y - 120, width: 240, height: 120 });
        y = parrafo(pg, LOREM + " " + LOREM, 72, y - 10, W - 144 - 255, times, 11);
        y = parrafo(pg, LOREM, 72, Math.min(y, H - 72 - 40 - 140) - 14, W - 144, times, 11);
        pg.drawText("Texto en rojo, 14 pt", { x: 72, y: y - 20, size: 14, font: helvB, color: rgb(0.8, 0.1, 0.1) });
        pg.drawText("Texto azul en cursiva, 12 pt", { x: 300, y: y - 20, size: 12, font: helvI, color: rgb(0, 0.4, 0.8) });
        break;
      }
      case 2: {
        pg.drawText("Tabla de resultados", { x: 72, y, size: 18, font: helvB });
        y = tabla(pg, 72, y - 20, [140, 110, 110, 108], [
          ["Producto", "Enero", "Febrero", "Marzo"],
          ["Papel A4", "1.200", "1.350", "1.410"],
          ["Tóner negro", "85", "92", "101"],
          ["Carpetas", "430", "415", "460"],
          ["Bolígrafos", "2.300", "2.150", "2.480"],
          ["Total", "4.015", "4.007", "4.451"],
        ]);
        parrafo(pg, "Nota: las cifras de la tabla están alineadas en celdas dibujadas con líneas vectoriales. " + LOREM, 72, y - 30, W - 144, times, 11);
        break;
      }
      case 3: {
        pg.drawText("Dos columnas con imagen vertical", { x: 72, y, size: 16, font: timesB });
        const col = (W - 144 - 24) / 2;
        parrafo(pg, LOREM + " " + LOREM + " " + LOREM, 72, y - 30, col, times, 10.5);
        pg.drawImage(imgRetrato, { x: 72 + col + 24, y: y - 30 - 220, width: 100, height: 220 });
        parrafo(pg, "Imagen de 100 x 220 pt a la derecha. " + LOREM, 72 + col + 24 + 110, y - 30, col - 110, times, 10.5);
        break;
      }
      case 4: {
        pg.drawText("Tipografías y tamaños", { x: 72, y, size: 18, font: helvB });
        const filas: [PDFFont, number, string][] = [
          [times, 8, "Times 8 pt"], [times, 12, "Times 12 pt"], [timesB, 16, "Times negrita 16 pt"],
          [helv, 10, "Helvetica 10 pt"], [helvI, 14, "Helvetica cursiva 14 pt"], [courier, 11, "Courier 11 pt: const x = 42;"],
        ];
        filas.forEach(([f, s, t], i) => pg.drawText(t, { x: 72, y: y - 40 - i * 30, size: s, font: f }));
        pg.drawImage(imgIcono, { x: 400, y: y - 120, width: 64, height: 64 });
        pg.drawText("Icono 64 x 64 con transparencia", { x: 370, y: y - 140, size: 9, font: helv });
        break;
      }
      case 5: {
        pg.drawRectangle({ x: 72, y: y - 200, width: W - 144, height: 200, color: rgb(0.93, 0.93, 0.93), borderColor: rgb(0.6, 0.6, 0.6), borderWidth: 1 });
        pg.drawText("Recuadro con fondo gris", { x: 86, y: y - 24, size: 14, font: helvB });
        parrafo(pg, LOREM, 86, y - 46, W - 172 - 130, helv, 10);
        pg.drawImage(imgIcono, { x: W - 72 - 110, y: y - 120, width: 90, height: 90 });
        parrafo(pg, LOREM + " " + LOREM, 72, y - 240, W - 144, times, 11);
        break;
      }
      case 6: {
        pg.drawText("Página horizontal (Letter apaisado)", { x: 72, y, size: 18, font: helvB });
        pg.drawImage(imgGrafico, { x: 72, y: y - 30 - 240, width: 480, height: 240 });
        tabla(pg, 72 + 480 + 20, y - 30, [80, 68], [["Mes", "Valor"], ["Ene", "220"], ["Feb", "300"], ["Mar", "160"], ["Abr", "330"], ["May", "260"], ["Jun", "120"]]);
        break;
      }
      default: {
        pg.drawText(`Sección ${n - 6} (A4)`, { x: 72, y, size: 18, font: timesB });
        y = parrafo(pg, LOREM + " " + LOREM, 72, y - 30, W - 144, times, 11);
        const imgs = [imgPaisaje, imgGrafico, imgRetrato];
        const im = imgs[n % 3];
        const escala = im === imgRetrato ? 0.45 : 0.4;
        const w = im.width * escala, h = im.height * escala;
        pg.drawImage(im, { x: n % 2 ? 72 : W - 72 - w, y: y - 20 - h, width: w, height: h });
        parrafo(pg, LOREM, n % 2 ? 72 + w + 15 : 72, y - 20, W - 144 - w - 15, times, 11);
        parrafo(pg, LOREM + " " + LOREM, 72, y - 40 - h, W - 144, times, 11);
      }
    }
    pie(pg, n);
  }
  mkdirSync("pruebas", { recursive: true });
  writeFileSync("pruebas/muestra-12-paginas.pdf", await doc.save());
  console.log("OK pruebas/muestra-12-paginas.pdf");
}
main();

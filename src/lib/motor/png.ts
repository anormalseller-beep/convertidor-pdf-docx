// Codificador PNG mínimo (RGBA 8 bits) basado en fflate. Funciona en Web Worker y en Node.
import { zlibSync } from "fflate";

const TABLA_CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(datos: Uint8Array, inicio: number, fin: number): number {
  let c = 0xffffffff;
  for (let i = inicio; i < fin; i++) c = TABLA_CRC[(c ^ datos[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function bloque(tipo: string, datos: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + datos.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, datos.length);
  for (let i = 0; i < 4; i++) out[4 + i] = tipo.charCodeAt(i);
  out.set(datos, 8);
  dv.setUint32(8 + datos.length, crc32(out, 4, 8 + datos.length));
  return out;
}

/** Codifica píxeles RGBA (o RGB si conAlfa=false) como PNG. */
export function codificarPng(
  ancho: number,
  alto: number,
  rgba: Uint8Array | Uint8ClampedArray,
  conAlfa = true,
): Uint8Array {
  const canales = conAlfa ? 4 : 3;
  const fila = ancho * canales;
  const crudo = new Uint8Array((fila + 1) * alto);
  for (let y = 0; y < alto; y++) {
    const o = y * (fila + 1);
    crudo[o] = 0; // filtro "None"
    if (conAlfa) {
      crudo.set(rgba.subarray(y * ancho * 4, (y + 1) * ancho * 4), o + 1);
    } else {
      let s = y * ancho * 4;
      let d = o + 1;
      for (let x = 0; x < ancho; x++, s += 4) {
        crudo[d++] = rgba[s];
        crudo[d++] = rgba[s + 1];
        crudo[d++] = rgba[s + 2];
      }
    }
  }
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, ancho);
  dv.setUint32(4, alto);
  ihdr[8] = 8;
  ihdr[9] = conAlfa ? 6 : 2;
  const partes = [
    new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    bloque("IHDR", ihdr),
    bloque("IDAT", zlibSync(crudo, { level: 6 })),
    bloque("IEND", new Uint8Array(0)),
  ];
  const total = partes.reduce((a, p) => a + p.length, 0);
  const png = new Uint8Array(total);
  let o = 0;
  for (const p of partes) {
    png.set(p, o);
    o += p.length;
  }
  return png;
}

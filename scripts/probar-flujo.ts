// Convierte pruebas/entradas/fuente.{txt,rtf,odt,docx,doc} a DOCX y DOC con los lectores del worker.
// Uso: npx tsx scripts/probar-flujo.ts
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { leerDoc, leerDocx, leerOdt, leerRtf, leerTxt } from "../src/lib/flujo/lectores";
import { flujoADoc2003, flujoADocx } from "../src/lib/flujo/escritor-flujo";
import type { DocumentoFlujo } from "../src/lib/flujo/modelo";

async function main() {
  mkdirSync("pruebas/salidas-flujo", { recursive: true });
  for (const ext of ["txt", "rtf", "odt", "docx", "doc"]) {
    const datos = new Uint8Array(readFileSync(`pruebas/entradas/fuente.${ext}`));
    const d: DocumentoFlujo =
      ext === "txt" ? leerTxt(datos) : ext === "rtf" ? leerRtf(datos) : ext === "odt" ? leerOdt(datos) : ext === "docx" ? await leerDocx(datos) : leerDoc(datos);
    const tipos = d.bloques.reduce<Record<string, number>>((a, b) => ((a[b.tipo] = (a[b.tipo] || 0) + 1), a), {});
    const texto = d.bloques.flatMap((b) => (b.tipo === "p" ? b.corridas.map((c) => c.texto) : b.tipo === "tabla" ? b.filas.flat(2).map((c) => c.texto) : [])).join(" ").replace(/\s+/g, " ").slice(0, 110);
    const x = await flujoADocx(d);
    writeFileSync(`pruebas/salidas-flujo/${ext}-a.docx`, new Uint8Array(await x.blob.arrayBuffer()));
    const y = flujoADoc2003(d);
    writeFileSync(`pruebas/salidas-flujo/${ext}-a.doc`, new Uint8Array(await y.blob.arrayBuffer()));
    console.log(`${ext.padEnd(5)} bloques ${JSON.stringify(tipos)} | ${texto}`);
    d.avisos.forEach((a) => console.log("      aviso: " + a));
  }
}
main().catch((e) => { console.error(e); process.exit(1); });

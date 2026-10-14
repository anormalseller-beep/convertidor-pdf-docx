// Escritor .doc: Word 2003 XML (WordprocessingML 2003) guardado con extensión .doc.
// No existe un generador gratuito del formato binario Word 97 que funcione en el navegador
// (haría falta LibreOffice en un servidor). Word 2003+ y LibreOffice abren este archivo
// como documento de Word. Usa cuadros de texto e imágenes VML con posición absoluta.
import type { PaginaInterpretada } from "./interprete";
import { ajusteAncho, bloquesDeTexto, esc, margenes, twip, type Medidor } from "./comun";
import type { ResultadoEscritor } from "./escritor-docx";

function base64(datos: Uint8Array): string {
  let s = "";
  const paso = 0x8000;
  for (let i = 0; i < datos.length; i += paso) s += String.fromCharCode(...datos.subarray(i, i + paso));
  // btoa existe en navegador, Worker y Node 16+
  return btoa(s).replace(/.{76}/g, "$&\n");
}

const pt = (v: number) => `${v.toFixed(2)}pt`;

export class EscritorDoc2003 {
  private partes: BlobPart[] = [];
  private bloques: Blob[] = [];
  private bytes = 0;
  private nImg = 0;
  private nForma = 0;
  private paginas = 0;
  private imagenes = 0;
  private tamPaginas: { ancho: number; alto: number }[] = [];
  private ultimaSect = "";
  private nombrePorDatos = new WeakMap<Uint8Array, string>();

  constructor(private medir?: Medidor, titulo = "Documento") {
    this.escribir(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n` +
        `<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xmlns:v="urn:schemas-microsoft-com:vml" ` +
        `xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
        `xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xml:space="preserve">` +
        `<o:DocumentProperties><o:Title>${esc(titulo)}</o:Title></o:DocumentProperties>` +
        `<w:fonts><w:defaultFonts w:ascii="Arial" w:fareast="Arial" w:h-ansi="Arial" w:cs="Arial"/></w:fonts>` +
        `<w:styles><w:style w:type="paragraph" w:default="on" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:h-ansi="Arial"/><w:sz w:val="22"/><w:lang w:val="ES"/></w:rPr></w:style></w:styles>` +
        `<w:docPr><w:view w:val="print"/><w:defaultTabStop w:val="708"/></w:docPr><w:body>` +
        `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
        `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>`,
    );
  }

  private escribir(s: string) {
    this.partes.push(s);
    this.bytes += s.length;
    if (this.bytes > 16 * 1024 * 1024) {
      this.bloques.push(new Blob(this.partes));
      this.partes = [];
      this.bytes = 0;
    }
  }

  private imagen(datos: Uint8Array, ext: "png" | "jpeg", x: number, y: number, ancho: number, alto: number, z: number, rot: number, volteadaV: boolean) {
    let nombre = this.nombrePorDatos.get(datos);
    const nueva = !nombre;
    if (!nombre) {
      nombre = `wordml://${String(++this.nImg).padStart(8, "0")}.${ext === "jpeg" ? "jpg" : "png"}`;
      this.nombrePorDatos.set(datos, nombre);
    }
    const idForma = ++this.nForma;
    const estilo =
      `position:absolute;margin-left:${pt(x)};margin-top:${pt(y)};width:${pt(ancho)};height:${pt(alto)};z-index:${z};` +
      `mso-position-horizontal-relative:page;mso-position-vertical-relative:page` +
      (rot ? `;rotation:${rot.toFixed(2)}` : "") +
      (volteadaV ? ";flip:y" : "");
    return (
      `<w:r><w:pict>${nueva ? `<w:binData w:name="${nombre}">${base64(datos)}</w:binData>` : ""}` +
      `<v:shape id="img${idForma}" type="#_x0000_t75" style="${estilo}"><v:imagedata src="${nombre}" o:title=""/><o:lock v:ext="edit" aspectratio="t"/><w10:wrap type="none"/><w10:anchorlock/></v:shape></w:pict></w:r>`
    );
  }

  agregarPagina(p: PaginaInterpretada, esUltima: boolean) {
    this.paginas++;
    this.tamPaginas.push({ ancho: p.ancho, alto: p.alto });
    const r: string[] = [];
    if (p.vectores) r.push(this.imagen(p.vectores.datos, "png", p.vectores.x, p.vectores.y, p.vectores.ancho, p.vectores.alto, -251658240, 0, false));
    for (const f of p.formas) {
      const id = ++this.nForma;
      const pos = "mso-position-horizontal-relative:page;mso-position-vertical-relative:page";
      if (f.tipo === "rect") {
        const estilo = `position:absolute;margin-left:${pt(f.x)};margin-top:${pt(f.y)};width:${pt(Math.max(f.ancho, 0.01))};height:${pt(Math.max(f.alto, 0.01))};z-index:${1000 + f.orden};${pos}`;
        r.push(
          `<w:r><w:pict><v:rect id="f${id}" style="${estilo}" filled="${f.relleno ? "t" : "f"}"${f.relleno ? ` fillcolor="#${f.relleno}"` : ""} stroked="${f.trazo ? "t" : "f"}"${f.trazo ? ` strokecolor="#${f.trazo}" strokeweight="${pt(f.grosor)}"` : ""}><w10:wrap type="none"/><w10:anchorlock/></v:rect></w:pict></w:r>`,
        );
      } else {
        r.push(
          `<w:r><w:pict><v:line id="f${id}" from="${pt(f.x1)},${pt(f.y1)}" to="${pt(f.x2)},${pt(f.y2)}" style="position:absolute;z-index:${1000 + f.orden};${pos}" strokecolor="#${f.trazo}" strokeweight="${pt(f.grosor)}"><w10:wrap type="none"/><w10:anchorlock/></v:line></w:pict></w:r>`,
        );
      }
    }
    for (const im of p.imagenes) {
      this.imagenes++;
      r.push(this.imagen(im.datos, im.ext, im.x, im.y, im.ancho, im.alto, 1000 + im.orden, im.rot, im.volteadaV));
    }
    for (const b of bloquesDeTexto(p.segmentos)) {
      const caja = b.caja;
      const id = ++this.nForma;
      const parrafos = b.lineas
        .map(({ seg, altoTw }) => {
          const corridas = seg.corridas
            .map((c) => {
              const { espTwips, escala } = ajusteAncho(c, this.medir);
              const e = c.estilo;
              const f = esc(e.fuente);
              const sz = Math.max(2, Math.round(e.tam * 2));
              return (
                `<w:r><w:rPr><w:rFonts w:ascii="${f}" w:fareast="${f}" w:h-ansi="${f}" w:cs="${f}"/>` +
                (e.negrita ? "<w:b/>" : "") + (e.cursiva ? "<w:i/>" : "") +
                (e.color !== "000000" ? `<w:color w:val="${e.color}"/>` : "") +
                (espTwips ? `<w:spacing w:val="${espTwips}"/>` : "") +
                (escala !== 100 ? `<w:w w:val="${escala}"/>` : "") +
                `<w:sz w:val="${sz}"/><w:sz-cs w:val="${sz}"/></w:rPr><w:t>${esc(c.texto)}</w:t></w:r>`
              );
            })
            .join("");
          const sangria = twip(seg.x - caja.x);
          return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${altoTw}" w:line-rule="exact"/>${sangria ? `<w:ind w:left="${sangria}"/>` : ""}</w:pPr>${corridas}</w:p>`;
        })
        .join("");
      const estilo =
        `position:absolute;margin-left:${pt(caja.x)};margin-top:${pt(caja.y)};width:${pt(caja.ancho)};height:${pt(caja.alto)};z-index:${100000 + b.orden};` +
        `mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap-style:none` +
        (caja.rot ? `;rotation:${caja.rot.toFixed(2)}` : "");
      r.push(
        `<w:r><w:pict><v:shape id="txt${id}" type="#_x0000_t202" style="${estilo}" filled="f" stroked="f">` +
          `<v:textbox inset="0,0,0,0"><w:txbxContent>${parrafos}</w:txbxContent></v:textbox>` +
          `<w10:wrap type="none"/><w10:anchorlock/></v:shape></w:pict></w:r>`,
      );
    }
    const m = margenes(p);
    const sect =
      `<w:sectPr><w:type w:val="next-page"/><w:pgSz w:w="${twip(p.ancho)}" w:h="${twip(p.alto)}"${p.ancho > p.alto ? ' w:orient="landscape"' : ""}/>` +
      `<w:pgMar w:top="${twip(m.sup)}" w:right="${twip(m.der)}" w:bottom="${twip(m.inf)}" w:left="${twip(m.izq)}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
    this.escribir(
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:line-rule="exact"/>${esUltima ? "" : sect}<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr>${r.join("")}</w:p>`,
    );
    if (esUltima) this.ultimaSect = sect;
  }

  async finalizar(): Promise<ResultadoEscritor> {
    this.escribir(`${this.ultimaSect}</w:body></w:wordDocument>`);
    if (this.partes.length) this.bloques.push(new Blob(this.partes));
    this.partes = [];
    return {
      blob: new Blob(this.bloques, { type: "application/msword" }),
      paginas: this.paginas,
      imagenes: this.imagenes,
      tamPaginas: this.tamPaginas,
    };
  }
}

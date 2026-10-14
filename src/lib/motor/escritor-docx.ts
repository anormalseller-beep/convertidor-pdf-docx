// Escritor DOCX en streaming. Cada página del PDF = una sección con su tamaño y orientación.
// Texto: un cuadro de texto (wps) por segmento de línea, anclado a la página, sin ajuste
// de línea (wrap="none"), así Word nunca reacomoda el texto. Imágenes: wp:anchor con el
// tamaño exacto en EMU y relación de aspecto bloqueada. El ZIP se genera con fflate en
// streaming: las imágenes se comprimen y se sueltan apenas se agregan.
import { Zip, ZipDeflate, ZipPassThrough, strToU8 } from "fflate";
import type { Forma, PaginaInterpretada } from "./interprete";
import { ajusteAncho, bloquesDeTexto, emu, esc, margenes, twip, type Medidor } from "./comun";
import type { Corrida } from "./tipos";

const NS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'mc:Ignorable="w14 wp14"';

const Z_VECTORES = 251658240;
const Z_IMAGENES = 251700000;
const Z_TEXTO = 252700000;

export interface ResultadoEscritor {
  blob: Blob;
  paginas: number;
  imagenes: number;
  tamPaginas: { ancho: number; alto: number }[];
}

export class EscritorDocx {
  private zip: Zip;
  private salida: BlobPart[] = [];
  private bytesPendientes = 0;
  private bloques: Blob[] = [];
  private cuerpo: string[] = [];
  private cuerpoBytes = 0;
  private cuerpoBlobs: Blob[] = [];
  private rels: string[] = [];
  private nImg = 0;
  private nDoc = 0;
  private terminado: Promise<void>;
  private resolverFin!: () => void;
  private rechazarFin!: (e: unknown) => void;
  private paginas = 0;
  private imagenesColocadas = 0;
  private tamPaginas: { ancho: number; alto: number }[] = [];
  private ultimaSect = "";
  private mediaPorDatos = new WeakMap<Uint8Array, string>();

  constructor(private medir?: Medidor, private titulo = "Documento") {
    this.terminado = new Promise((res, rej) => { this.resolverFin = res; this.rechazarFin = rej; });
    this.zip = new Zip((err, datos, final) => {
      if (err) return this.rechazarFin(err);
      this.salida.push(datos);
      this.bytesPendientes += datos.length;
      // consolidar cada ~32 MB para que el navegador pueda pasar los Blobs a disco
      if (this.bytesPendientes > 32 * 1024 * 1024) this.consolidar();
      if (final) { this.consolidar(); this.resolverFin(); }
    });
    this.archivo("[Content_Types].xml", CONTENT_TYPES);
    this.archivo("_rels/.rels", RELS_RAIZ);
    this.archivo("word/styles.xml", ESTILOS);
    this.archivo("word/settings.xml", AJUSTES);
    this.archivo("word/fontTable.xml", FUENTES);
  }

  private consolidar() {
    if (this.salida.length) this.bloques.push(new Blob(this.salida as BlobPart[]));
    this.salida = [];
    this.bytesPendientes = 0;
  }

  private archivo(nombre: string, xml: string) {
    const f = new ZipDeflate(nombre, { level: 6 });
    this.zip.add(f);
    f.push(strToU8(xml), true);
  }

  private agregarMedia(datos: Uint8Array, ext: "png" | "jpeg"): string {
    // la misma imagen en varias páginas (logos) se guarda una sola vez
    const previa = this.mediaPorDatos.get(datos);
    if (previa) return previa;
    this.nImg++;
    const nombre = `imagen${this.nImg}.${ext === "jpeg" ? "jpeg" : "png"}`;
    const f = new ZipPassThrough(`word/media/${nombre}`);
    this.zip.add(f);
    f.push(datos, true);
    const rid = `rIdImg${this.nImg}`;
    this.rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${nombre}"/>`);
    this.mediaPorDatos.set(datos, rid);
    return rid;
  }

  private escribirCuerpo(xml: string) {
    this.cuerpo.push(xml);
    this.cuerpoBytes += xml.length;
    if (this.cuerpoBytes > 8 * 1024 * 1024) {
      this.cuerpoBlobs.push(new Blob(this.cuerpo));
      this.cuerpo = [];
      this.cuerpoBytes = 0;
    }
  }

  private ancla(o: { x: number; y: number; ancho: number; alto: number; z: number; detras: boolean; grafico: string; bloquearAspecto: boolean; nombre: string }) {
    const id = ++this.nDoc;
    return (
      `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${o.z}" behindDoc="${o.detras ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="page"><wp:posOffset>${emu(o.x)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(o.y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${Math.max(1, emu(o.ancho))}" cy="${Math.max(1, emu(o.alto))}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
      `<wp:docPr id="${id}" name="${o.nombre} ${id}"/>` +
      `<wp:cNvGraphicFramePr>${o.bloquearAspecto ? '<a:graphicFrameLocks noChangeAspect="1"/>' : ""}</wp:cNvGraphicFramePr>` +
      o.grafico.replace("__ID__", String(id)) +
      `</wp:anchor>`
    );
  }

  private imagenXml(rid: string, ancho: number, alto: number, rot: number, volteadaV: boolean) {
    const attrs = `${rot ? ` rot="${Math.round(rot * 60000)}"` : ""}${volteadaV ? ' flipV="1"' : ""}`;
    return (
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="__ID__" name="Imagen"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr bwMode="auto"><a:xfrm${attrs}><a:off x="0" y="0"/><a:ext cx="${Math.max(1, emu(ancho))}" cy="${Math.max(1, emu(alto))}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr></pic:pic></a:graphicData></a:graphic>`
    );
  }

  private formaXml(f: Forma) {
    const color = (c: string) => `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
    let x: number, y: number, ancho: number, alto: number, geom: string, attrs = "", relleno = "<a:noFill/>", linea: string;
    if (f.tipo === "rect") {
      ({ x, y, ancho, alto } = f);
      geom = "rect";
      if (f.relleno) relleno = color(f.relleno);
      linea = f.trazo ? `<a:ln w="${emu(f.grosor)}" cap="flat">${color(f.trazo)}<a:miter lim="800000"/></a:ln>` : "<a:ln><a:noFill/></a:ln>";
    } else {
      x = Math.min(f.x1, f.x2); y = Math.min(f.y1, f.y2);
      ancho = Math.abs(f.x2 - f.x1); alto = Math.abs(f.y2 - f.y1);
      geom = "line";
      // por defecto la línea va de arriba-izquierda a abajo-derecha de su caja
      if ((f.x2 - f.x1) * (f.y2 - f.y1) < 0) attrs = ' flipV="1"';
      linea = `<a:ln w="${emu(f.grosor)}" cap="flat">${color(f.trazo)}</a:ln>`;
    }
    const grafico =
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>` +
      `<wps:spPr><a:xfrm${attrs}><a:off x="0" y="0"/><a:ext cx="${emu(ancho)}" cy="${emu(alto)}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${geom === "line" ? "" : relleno}${linea}</wps:spPr>` +
      `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>`;
    return { x, y, ancho, alto, grafico };
  }

  private corridaXml(c: Corrida) {
    const { espTwips, escala } = ajusteAncho(c, this.medir);
    const e = c.estilo;
    const f = esc(e.fuente);
    const sz = Math.max(2, Math.round(e.tam * 2));
    return (
      `<w:r><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}" w:eastAsia="${f}"/>` +
      (e.negrita ? "<w:b/><w:bCs/>" : "") +
      (e.cursiva ? "<w:i/><w:iCs/>" : "") +
      (e.color !== "000000" ? `<w:color w:val="${e.color}"/>` : "") +
      (espTwips ? `<w:spacing w:val="${espTwips}"/>` : "") +
      (escala !== 100 ? `<w:w w:val="${escala}"/>` : "") +
      `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
      `<w:t xml:space="preserve">${esc(c.texto)}</w:t></w:r>`
    );
  }

  agregarPagina(p: PaginaInterpretada, esUltima: boolean) {
    this.paginas++;
    this.tamPaginas.push({ ancho: p.ancho, alto: p.alto });
    const anclas: string[] = [];
    const envolver = (anclaXml: string, forma: boolean) =>
      forma
        ? `<w:r><w:rPr><w:noProof/><w:sz w:val="2"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${anclaXml}</w:drawing></mc:Choice></mc:AlternateContent></w:r>`
        : `<w:r><w:rPr><w:noProof/><w:sz w:val="2"/></w:rPr><w:drawing>${anclaXml}</w:drawing></w:r>`;

    if (p.vectores) {
      const v = p.vectores;
      const rid = this.agregarMedia(v.datos, v.ext);
      anclas.push(envolver(this.ancla({ x: v.x, y: v.y, ancho: v.ancho, alto: v.alto, z: Z_VECTORES, detras: true, grafico: this.imagenXml(rid, v.ancho, v.alto, 0, false), bloquearAspecto: true, nombre: "Dibujo" }), false));
    }
    for (const im of p.imagenes) {
      const rid = this.agregarMedia(im.datos, im.ext);
      this.imagenesColocadas++;
      anclas.push(envolver(this.ancla({ x: im.x, y: im.y, ancho: im.ancho, alto: im.alto, z: Z_IMAGENES + im.orden, detras: false, grafico: this.imagenXml(rid, im.ancho, im.alto, im.rot, im.volteadaV), bloquearAspecto: true, nombre: "Imagen" }), false));
    }
    for (const f of p.formas) {
      const g = this.formaXml(f);
      anclas.push(envolver(this.ancla({ x: g.x, y: g.y, ancho: g.ancho, alto: g.alto, z: Z_IMAGENES + f.orden, detras: false, grafico: g.grafico, bloquearAspecto: false, nombre: "Forma" }), true));
    }
    for (const b of bloquesDeTexto(p.segmentos)) {
      const caja = b.caja;
      const parrafo = b.lineas
        .map(
          ({ seg, altoTw }) =>
            `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${altoTw}" w:lineRule="exact"/><w:ind w:left="${twip(seg.x - caja.x) || 0}" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>` +
            seg.corridas.map((c) => this.corridaXml(c)).join("") +
            `</w:p>`,
        )
        .join("");
      const grafico =
        `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
        `<wps:spPr><a:xfrm${caja.rot ? ` rot="${Math.round(caja.rot * 60000)}"` : ""}><a:off x="0" y="0"/><a:ext cx="${emu(caja.ancho)}" cy="${emu(caja.alto)}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
        `<wps:txbx><w:txbxContent>${parrafo}</w:txbxContent></wps:txbx>` +
        `<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" spcCol="0" rtlCol="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>` +
        `</wps:wsp></a:graphicData></a:graphic>`;
      anclas.push(envolver(this.ancla({ x: caja.x, y: caja.y, ancho: caja.ancho, alto: caja.alto, z: Z_TEXTO + b.orden, detras: false, grafico, bloquearAspecto: false, nombre: "Texto" }), true));
    }

    const m = margenes(p);
    const apaisada = p.ancho > p.alto;
    const sect =
      `<w:sectPr><w:type w:val="nextPage"/>` +
      `<w:pgSz w:w="${twip(p.ancho)}" w:h="${twip(p.alto)}"${apaisada ? ' w:orient="landscape"' : ""}/>` +
      `<w:pgMar w:top="${twip(m.sup)}" w:right="${twip(m.der)}" w:bottom="${twip(m.inf)}" w:left="${twip(m.izq)}" w:header="0" w:footer="0" w:gutter="0"/>` +
      `<w:cols w:space="708"/></w:sectPr>`;
    const pPr = `<w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr>${esUltima ? "" : sect}</w:pPr>`;
    this.escribirCuerpo(`<w:p>${pPr}${anclas.join("")}</w:p>`);
    if (esUltima) this.ultimaSect = sect;
  }

  async finalizar(): Promise<ResultadoEscritor> {
    if (!this.ultimaSect) this.ultimaSect = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
    this.archivo(
      "word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdEstilos" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        `<Relationship Id="rIdAjustes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
        `<Relationship Id="rIdFuentes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>` +
        this.rels.join("") +
        `</Relationships>`,
    );
    const ahora = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    this.archivo(
      "docProps/core.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
        `<dc:title>${esc(this.titulo)}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${ahora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${ahora}</dcterms:modified></cp:coreProperties>`,
    );
    this.archivo(
      "docProps/app.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Convertidor de PDF a DOC y DOCX</Application><Pages>${this.paginas}</Pages></Properties>`,
    );
    // document.xml al final: se comprime en streaming desde los bloques guardados
    if (this.cuerpo.length) this.cuerpoBlobs.push(new Blob(this.cuerpo));
    this.cuerpo = [];
    const doc = new ZipDeflate("word/document.xml", { level: 6 });
    this.zip.add(doc);
    doc.push(strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>`));
    for (const b of this.cuerpoBlobs) {
      const buf = new Uint8Array(await b.arrayBuffer());
      for (let o = 0; o < buf.length; o += 4 * 1024 * 1024) doc.push(buf.subarray(o, o + 4 * 1024 * 1024));
    }
    this.cuerpoBlobs = [];
    doc.push(strToU8(`${this.ultimaSect}</w:body></w:document>`), true);
    this.zip.end();
    await this.terminado;
    return {
      blob: new Blob(this.bloques, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
      paginas: this.paginas,
      imagenes: this.imagenesColocadas,
      tamPaginas: this.tamPaginas,
    };
  }
}

const CONTENT_TYPES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="png" ContentType="image/png"/>` +
  `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
  `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
  `<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>` +
  `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
  `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
  `</Types>`;

const RELS_RAIZ =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
  `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
  `</Relationships>`;

const ESTILOS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr></w:rPrDefault>` +
  `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>` +
  `</w:styles>`;

const AJUSTES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>` +
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>` +
  `</w:settings>`;

const FUENTES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  ["Arial", "Times New Roman", "Calibri", "Cambria", "Georgia", "Courier New", "Verdana", "Tahoma"]
    .map((f) => `<w:font w:name="${f}"/>`)
    .join("") +
  `</w:fonts>`;

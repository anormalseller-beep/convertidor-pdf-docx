// Modelo intermedio: una página del PDF ya interpretada, en puntos (1/72 pulgada),
// con origen arriba a la izquierda (como Word).

export interface Estilo {
  fuente: string; // fuente estándar segura (Arial, Times New Roman, ...)
  fuenteOriginal: string; // nombre en el PDF, para el registro
  tam: number; // pt
  negrita: boolean;
  cursiva: boolean;
  color: string; // RRGGBB
}

export interface Corrida {
  texto: string;
  x: number; // inicio de la corrida
  ancho: number; // ancho real en el PDF (pt)
  estilo: Estilo;
}

/** Un segmento = trozo de una línea de texto que se coloca como cuadro absoluto. */
export interface Segmento {
  x: number;
  base: number; // línea base (desde arriba)
  ancho: number;
  tam: number; // tamaño dominante
  corridas: Corrida[];
  orden: number; // orden de pintado (z)
}

export interface ImagenPagina {
  x: number;
  y: number;
  ancho: number;
  alto: number;
  pixAncho: number;
  pixAlto: number;
  datos: Uint8Array; // archivo ya codificado
  ext: "png" | "jpeg";
  orden: number;
  /** true si el PDF la dibujaba rotada/inclinada (se usa su caja envolvente) */
  transformada: boolean;
}

export interface PaginaModelo {
  numero: number;
  ancho: number;
  alto: number;
  segmentos: Segmento[];
  imagenes: ImagenPagina[];
  /** capa de líneas/rellenos vectoriales (tablas, recuadros), PNG transparente o null */
  vectores: ImagenPagina | null;
  imagenesEnPdf: number; // cuántas imágenes pintaba el PDF (para QA)
  escaneada: boolean; // sin texto y cubierta por una imagen grande
  avisos: string[];
}

export type FormatoSalida = "docx" | "doc";

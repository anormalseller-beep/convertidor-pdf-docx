// Mensajes entre la interfaz y el Web Worker de conversión.
import type { InformeQA } from "@/lib/motor/qa";
import type { FormatoSalida } from "@/lib/motor/tipos";

export type MensajeAlWorker =
  | { tipo: "analizar"; archivo: File; contrasena?: string }
  | { tipo: "convertir"; archivo: File; formato: FormatoSalida; contrasena?: string; ocr: boolean }
  | { tipo: "cancelar" };

export type FaseProgreso = "leyendo" | "convirtiendo" | "ocr" | "empaquetando" | "verificando";

export type MensajeDelWorker =
  | { tipo: "analisis"; paginas: number | null; escaneado: boolean; cifrado: boolean }
  | { tipo: "progreso"; fase: FaseProgreso; pagina: number; total: number; porcentaje: number }
  | { tipo: "estado"; texto: string }
  | { tipo: "contrasena"; incorrecta: boolean }
  | { tipo: "listo"; blob: Blob; avisos: string[]; qa: InformeQA | null; paginas: number | null; segundos: number }
  | { tipo: "cancelado" }
  | { tipo: "error"; mensaje: string };

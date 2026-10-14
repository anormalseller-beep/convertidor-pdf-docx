"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import type { MensajeAlWorker, MensajeDelWorker, FaseProgreso } from "@/workers/protocolo";
import type { InformeQA } from "@/lib/motor/qa";
import { guardarArchivo } from "@/lib/almacen-local";
import { Historial, useSesion } from "./Historial";
import { guardarEnHistorial } from "@/lib/firebase";

const ACEPTADOS = [".pdf", ".docx", ".doc", ".odt", ".rtf", ".txt", ".jpg", ".jpeg", ".png"];
const UMBRAL_GRANDE = 200 * 1024 * 1024;

const FASES: Record<FaseProgreso, string> = {
  leyendo: "Leyendo el archivo",
  convirtiendo: "Convirtiendo",
  ocr: "Reconociendo texto (OCR)",
  empaquetando: "Generando el archivo",
  verificando: "Verificando el resultado",
};

function tamano(b: number) {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1024 / 1024).toFixed(1)} MB`;
}

interface Resultado {
  url: string;
  nombre: string;
  bytes: number;
  qa: InformeQA | null;
  avisos: string[];
  segundos: number;
  paginas: number | null;
}

type Estado = "inactivo" | "analizando" | "listo-para-convertir" | "convirtiendo" | "terminado" | "error";

function crearWorker() {
  return new Worker(new URL("../workers/conversor.worker.ts", import.meta.url), { type: "module" });
}

export function Convertidor() {
  const [archivo, setArchivo] = useState<File | null>(null);
  const [paginas, setPaginas] = useState<number | null>(null);
  const [escaneado, setEscaneado] = useState(false);
  const [pideContrasena, setPideContrasena] = useState(false);
  const [contrasenaMala, setContrasenaMala] = useState(false);
  const [contrasena, setContrasena] = useState("");
  const [formato, setFormato] = useState<"docx" | "doc">("docx");
  const [ocr, setOcr] = useState(false);
  const [aceptaGrande, setAceptaGrande] = useState(false);
  const [estado, setEstado] = useState<Estado>("inactivo");
  const [progreso, setProgreso] = useState<{ fase: FaseProgreso; pagina: number; total: number; porcentaje: number } | null>(null);
  const [estadoTexto, setEstadoTexto] = useState("");
  const [resultado, setResultado] = useState<Resultado | null>(null);
  const [error, setError] = useState("");
  const [arrastrando, setArrastrando] = useState(false);
  const [historialVersion, setHistorialVersion] = useState(0);
  const worker = useRef<Worker | null>(null);
  const entrada = useRef<HTMLInputElement>(null);
  const usuario = useSesion();

  const terminarWorker = () => {
    worker.current?.terminate();
    worker.current = null;
  };
  useEffect(() => () => terminarWorker(), []);
  useEffect(() => () => { if (resultado) URL.revokeObjectURL(resultado.url); }, [resultado]);

  const analizar = useCallback((f: File, clave?: string) => {
    terminarWorker();
    setEstado("analizando");
    setError("");
    const w = crearWorker();
    worker.current = w;
    w.onmessage = (ev: MessageEvent<MensajeDelWorker>) => {
      const m = ev.data;
      if (m.tipo === "analisis") {
        setPaginas(m.paginas);
        setEscaneado(m.escaneado);
        if (m.escaneado) setOcr(true);
        setPideContrasena(false);
        setEstado("listo-para-convertir");
        terminarWorker();
      } else if (m.tipo === "contrasena") {
        setPideContrasena(true);
        setContrasenaMala(m.incorrecta);
        setEstado("listo-para-convertir");
        terminarWorker();
      } else if (m.tipo === "error") {
        setError(m.mensaje);
        setEstado("error");
        terminarWorker();
      }
    };
    w.onerror = () => {
      setError("No se pudo iniciar el proceso de conversión en este navegador.");
      setEstado("error");
    };
    w.postMessage({ tipo: "analizar", archivo: f, contrasena: clave } satisfies MensajeAlWorker);
  }, []);

  const elegir = (f: File | undefined) => {
    if (!f) return;
    const ext = "." + (f.name.split(".").pop() || "").toLowerCase();
    setResultado(null);
    setProgreso(null);
    setPaginas(null);
    setEscaneado(false);
    setOcr(false);
    setContrasena("");
    setPideContrasena(false);
    setContrasenaMala(false);
    setAceptaGrande(false);
    if (!ACEPTADOS.includes(ext)) {
      setArchivo(null);
      setError("Formato no admitido. Usa PDF, DOCX, DOC, ODT, RTF, TXT, JPG o PNG.");
      setEstado("error");
      return;
    }
    setArchivo(f);
    analizar(f);
  };

  const convertir = () => {
    if (!archivo) return;
    terminarWorker();
    setEstado("convirtiendo");
    setError("");
    setResultado(null);
    setEstadoTexto("");
    setProgreso({ fase: "leyendo", pagina: 0, total: paginas ?? 0, porcentaje: 0 });
    const w = crearWorker();
    worker.current = w;
    const nombreSalida = archivo.name.replace(/\.[^.]+$/, "") + "." + formato;
    w.onmessage = async (ev: MessageEvent<MensajeDelWorker>) => {
      const m = ev.data;
      if (m.tipo === "progreso") setProgreso({ fase: m.fase, pagina: m.pagina, total: m.total, porcentaje: m.porcentaje });
      else if (m.tipo === "estado") setEstadoTexto(m.texto);
      else if (m.tipo === "contrasena") {
        setPideContrasena(true);
        setContrasenaMala(m.incorrecta);
        setEstado("listo-para-convertir");
        terminarWorker();
      } else if (m.tipo === "listo") {
        terminarWorker();
        setProgreso((p) => (p ? { ...p, porcentaje: 100 } : p));
        setResultado({ url: URL.createObjectURL(m.blob), nombre: nombreSalida, bytes: m.blob.size, qa: m.qa, avisos: m.avisos, segundos: m.segundos, paginas: m.paginas });
        setEstado("terminado");
        const idLocal = crypto.randomUUID();
        await guardarArchivo(idLocal, nombreSalida, m.blob);
        if (usuario) {
          try {
            await guardarEnHistorial(usuario.uid, { nombre: nombreSalida, formato, paginas: m.paginas, bytes: m.blob.size, qaOk: m.qa?.ok ?? null, idLocal });
            setHistorialVersion((v) => v + 1);
          } catch {
            /* el historial es opcional */
          }
        }
      } else if (m.tipo === "cancelado") {
        terminarWorker();
        setEstado("listo-para-convertir");
        setProgreso(null);
        setError("Conversión cancelada.");
      } else if (m.tipo === "error") {
        terminarWorker();
        setError(m.mensaje);
        setEstado("error");
      }
    };
    w.onerror = () => {
      terminarWorker();
      setError("El archivo supera el tamaño que el navegador puede procesar. Intenta con uno más pequeño.");
      setEstado("error");
    };
    w.postMessage({ tipo: "convertir", archivo, formato, contrasena: contrasena || undefined, ocr } satisfies MensajeAlWorker);
  };

  const cancelar = () => {
    worker.current?.postMessage({ tipo: "cancelar" } satisfies MensajeAlWorker);
    // si el worker está ocupado en una página muy pesada, se corta igual
    const w = worker.current;
    setTimeout(() => {
      if (worker.current === w && w) {
        terminarWorker();
        setEstado("listo-para-convertir");
        setProgreso(null);
        setError("Conversión cancelada.");
      }
    }, 1500);
  };

  const grande = !!archivo && archivo.size > UMBRAL_GRANDE;
  const ext = archivo?.name.split(".").pop()?.toLowerCase();
  const puedeConvertir =
    !!archivo && estado !== "convirtiendo" && estado !== "analizando" && !pideContrasena && (!grande || aceptaGrande) && !(estado === "error" && !archivo);
  const ocrDisponible = ext === "pdf" || ext === "jpg" || ext === "jpeg" || ext === "png";

  return (
    <main className="mx-auto w-full max-w-[640px] px-4 py-10">
      <h1 className="mb-8 text-[22px] font-semibold tracking-tight">Convertidor de PDF a DOC y DOCX</h1>

      <section className="mb-6">
        <span className="etiqueta">Archivo</span>
        <div
          role="button"
          tabIndex={0}
          className="zona"
          data-activa={arrastrando}
          onClick={() => entrada.current?.click()}
          onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && entrada.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setArrastrando(true); }}
          onDragLeave={() => setArrastrando(false)}
          onDrop={(e) => { e.preventDefault(); setArrastrando(false); elegir(e.dataTransfer.files[0]); }}
        >
          {archivo ? (
            <span className="mono" style={{ color: "var(--texto)" }}>{archivo.name}</span>
          ) : (
            <span>Arrastra un archivo aquí o haz clic para elegirlo.</span>
          )}
          <div className="nota mono mt-1">PDF, DOCX, DOC, ODT, RTF, TXT, JPG, PNG</div>
        </div>
        <input
          ref={entrada}
          type="file"
          accept={ACEPTADOS.join(",")}
          className="hidden"
          data-prueba="entrada-archivo"
          onChange={(e) => { elegir(e.target.files?.[0]); e.target.value = ""; }}
        />
        {archivo && (
          <div className="mono mt-2 flex flex-wrap gap-x-5 text-[var(--suave)]">
            <span>{tamano(archivo.size)}</span>
            {estado === "analizando" && <span>analizando...</span>}
            {paginas !== null && <span>{paginas} {paginas === 1 ? "página" : "páginas"}</span>}
          </div>
        )}
        {grande && (
          <div className="aviso mt-3">
            <p>El archivo pesa {tamano(archivo!.size)}. Por encima de 200 MB el navegador puede quedarse sin memoria.</p>
            <label className="mt-1 flex items-center gap-2">
              <input type="checkbox" checked={aceptaGrande} onChange={(e) => setAceptaGrande(e.target.checked)} /> Continuar de todos modos
            </label>
          </div>
        )}
        {pideContrasena && (
          <form
            className="mt-3 flex flex-wrap items-center gap-2"
            onSubmit={(e) => { e.preventDefault(); if (archivo) analizar(archivo, contrasena); }}
          >
            <label htmlFor="clave" className="text-[13px]">{contrasenaMala ? "Contraseña incorrecta. Contraseña del PDF" : "El PDF está protegido. Contraseña del PDF"}</label>
            <input id="clave" type="password" value={contrasena} onChange={(e) => setContrasena(e.target.value)} autoComplete="off" />
            <button type="submit" className="boton-sec">Abrir</button>
          </form>
        )}
        {escaneado && ext === "pdf" && (
          <p className="aviso mt-3">El PDF no contiene texto seleccionable (parece escaneado). Con OCR se obtiene texto editable.</p>
        )}
      </section>

      <section className="mb-6">
        <span className="etiqueta">Formato de salida</span>
        <div className="flex gap-6">
          <label className="flex items-center gap-2">
            <input type="radio" name="formato" value="docx" checked={formato === "docx"} onChange={() => setFormato("docx")} /> DOCX
          </label>
          <label className="flex items-center gap-2">
            <input type="radio" name="formato" value="doc" checked={formato === "doc"} onChange={() => setFormato("doc")} /> DOC (Word 97-2003)
          </label>
        </div>
        {formato === "doc" && (
          <p className="nota mt-2">
            DOC se genera como documento XML de Word 2003, que Word y LibreOffice abren como .doc. Puede perder algo de fidelidad; DOCX es la opción recomendada.
          </p>
        )}
        {ocrDisponible && (
          <label className="mt-3 flex items-start gap-2">
            <input type="checkbox" className="mt-1" checked={ocr} onChange={(e) => setOcr(e.target.checked)} />
            <span>
              OCR en páginas escaneadas
              <span className="nota block">Más lento. Descarga el modelo de idioma (español e inglés, unos 15 MB) la primera vez.</span>
            </span>
          </label>
        )}
      </section>

      <section className="mb-6 flex gap-3">
        <button type="button" className="boton" disabled={!puedeConvertir} onClick={convertir}>Convertir</button>
        {estado === "convirtiendo" && (
          <button type="button" className="boton-sec" onClick={cancelar}>Cancelar</button>
        )}
      </section>

      {(estado === "convirtiendo" || progreso) && (
        <section className="mb-6" aria-live="polite">
          <span className="etiqueta">Progreso</span>
          <div className="barra" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(progreso?.porcentaje ?? 0)}>
            <div style={{ width: `${progreso?.porcentaje ?? 0}%` }} />
          </div>
          <div className="mono mt-2 flex justify-between text-[var(--suave)]">
            <span>{progreso ? FASES[progreso.fase] : ""}{progreso && progreso.total > 0 ? ` · página ${progreso.pagina} / ${progreso.total}` : ""}</span>
            <span>{Math.round(progreso?.porcentaje ?? 0)} %</span>
          </div>
          {estadoTexto && estado === "convirtiendo" && <div className="nota mt-1">{estadoTexto}</div>}
        </section>
      )}

      {error && (
        <p className="aviso mb-6" role="alert">{error}</p>
      )}

      {resultado && (
        <section className="mb-6">
          <span className="etiqueta">Resultado</span>
          <div className="caja p-4">
            <div className="mono break-all">{resultado.nombre}</div>
            <div className="mono mt-1 text-[var(--suave)]">
              {tamano(resultado.bytes)}
              {resultado.paginas !== null ? ` · ${resultado.paginas} páginas` : ""} · {resultado.segundos.toFixed(1)} s
            </div>
            <a className="boton mt-3 inline-block no-underline" href={resultado.url} download={resultado.nombre} data-prueba="descargar">Descargar</a>
          </div>
          {resultado.qa && (
            <div className="mt-4">
              <span className="etiqueta">Verificación</span>
              <ul className="mono space-y-0.5 text-[12px]">
                <li>{resultado.qa.ok ? "Sin diferencias con el original." : "Se encontraron diferencias con el original:"}</li>
                {resultado.qa.lineas.map((l, i) => <li key={i}>{l}</li>)}
              </ul>
            </div>
          )}
          {resultado.avisos.length > 0 && (
            <div className="mt-4">
              <span className="etiqueta">Registro</span>
              <ul className="mono max-h-48 space-y-0.5 overflow-auto text-[12px] text-[var(--suave)]">
                {resultado.avisos.map((a, i) => <li key={i}>{a}</li>)}
              </ul>
            </div>
          )}
        </section>
      )}

      <Historial usuario={usuario} version={historialVersion} />
    </main>
  );
}

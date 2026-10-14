"use client";

import { useEffect, useState } from "react";
import type { User } from "firebase/auth";
import { borrarDelHistorial, cerrarSesion, escucharSesion, firebaseConfigurado, iniciarSesion, leerHistorial, type RegistroHistorial } from "@/lib/firebase";
import { borrarArchivo, leerArchivo } from "@/lib/almacen-local";

export function useSesion() {
  const [u, setU] = useState<User | null>(null);
  useEffect(() => escucharSesion(setU), []);
  return u;
}

function fecha(d: Date | null) {
  if (!d) return "";
  return d.toLocaleString("es", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
}

export function Historial({ usuario, version }: { usuario: User | null; version: number }) {
  const [filas, setFilas] = useState<RegistroHistorial[] | null>(null);
  const [error, setError] = useState("");
  const [nota, setNota] = useState("");

  useEffect(() => {
    if (!usuario) return;
    let vivo = true;
    leerHistorial(usuario.uid)
      .then((f) => vivo && setFilas(f))
      .catch(() => vivo && setError("No se pudo leer el historial."));
    return () => { vivo = false; };
  }, [usuario, version]);

  if (!firebaseConfigurado) return null;

  if (!usuario)
    return (
      <section className="mt-10 border-t border-[var(--linea)] pt-4 text-[13px] text-[var(--suave)]">
        Historial de conversiones:{" "}
        <a href="#" onClick={(e) => { e.preventDefault(); iniciarSesion().catch(() => setError("No se pudo iniciar sesión.")); }}>
          iniciar sesión con Google
        </a>
        {error && <span className="ml-2">{error}</span>}
      </section>
    );

  const descargar = async (r: RegistroHistorial) => {
    setNota("");
    const a = await leerArchivo(r.idLocal);
    if (!a) {
      setNota("El archivo no está guardado en este navegador. Vuelve a convertirlo.");
      return;
    }
    const url = URL.createObjectURL(a.blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = r.nombre;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  };

  const borrar = async (r: RegistroHistorial) => {
    await borrarDelHistorial(usuario.uid, r.id);
    await borrarArchivo(r.idLocal);
    setFilas((f) => f?.filter((x) => x.id !== r.id) ?? null);
  };

  return (
    <section className="mt-10 border-t border-[var(--linea)] pt-4">
      <div className="mb-2 flex items-baseline justify-between">
        <span className="etiqueta">Historial</span>
        <span className="text-[13px] text-[var(--suave)]">
          <span className="mono">{usuario.email}</span> ·{" "}
          <a href="#" onClick={(e) => { e.preventDefault(); cerrarSesion(); setFilas(null); }}>Salir</a>
        </span>
      </div>
      {error && <p className="aviso">{error}</p>}
      {nota && <p className="aviso mb-2">{nota}</p>}
      {filas && filas.length === 0 && <p className="nota">Sin conversiones.</p>}
      {filas && filas.length > 0 && (
        <table className="historial mono text-[12px]">
          <thead>
            <tr><th>Archivo</th><th>Fecha</th><th>Páginas</th><th>Tamaño</th><th></th></tr>
          </thead>
          <tbody>
            {filas.map((r) => (
              <tr key={r.id}>
                <td className="break-all">{r.nombre}</td>
                <td>{fecha(r.creado)}</td>
                <td>{r.paginas ?? "-"}</td>
                <td>{(r.bytes / 1024 / 1024).toFixed(2)} MB</td>
                <td className="whitespace-nowrap">
                  <a href="#" onClick={(e) => { e.preventDefault(); descargar(r); }}>Descargar</a>{" "}
                  <a href="#" onClick={(e) => { e.preventDefault(); borrar(r); }}>Borrar</a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      <p className="nota mt-2">Los archivos se guardan solo en este navegador. En la cuenta se guarda únicamente el nombre, la fecha y el tamaño.</p>
    </section>
  );
}

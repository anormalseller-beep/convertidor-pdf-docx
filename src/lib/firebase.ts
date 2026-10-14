// Firebase (plan Spark, gratis): solo inicio de sesión con Google y metadatos del historial.
// Los archivos convertidos NUNCA se suben: se guardan en IndexedDB del propio navegador.
import { initializeApp, getApps, type FirebaseApp } from "firebase/app";
import { getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithPopup, signOut, type User } from "firebase/auth";
import {
  addDoc, collection, getFirestore, limit, orderBy, query, getDocs, serverTimestamp, deleteDoc, doc, type Timestamp,
} from "firebase/firestore";

const config = {
  apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY,
  authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN,
  projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID,
  storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID,
  appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID,
};

export const firebaseConfigurado = !!(config.apiKey && config.projectId && config.appId);

let app: FirebaseApp | null = null;
function obtenerApp() {
  if (!firebaseConfigurado) return null;
  app ??= getApps()[0] ?? initializeApp(config);
  return app;
}

export function escucharSesion(cb: (u: User | null) => void) {
  const a = obtenerApp();
  if (!a) return () => {};
  return onAuthStateChanged(getAuth(a), cb);
}

export async function iniciarSesion() {
  const a = obtenerApp();
  if (!a) throw new Error("El inicio de sesión no está configurado.");
  await signInWithPopup(getAuth(a), new GoogleAuthProvider());
}

export async function cerrarSesion() {
  const a = obtenerApp();
  if (a) await signOut(getAuth(a));
}

export interface RegistroHistorial {
  id: string;
  nombre: string;
  formato: "docx" | "doc";
  paginas: number | null;
  bytes: number;
  qaOk: boolean | null;
  creado: Date | null;
  idLocal: string; // clave del archivo en IndexedDB de este navegador
}

export async function guardarEnHistorial(uid: string, r: Omit<RegistroHistorial, "id" | "creado">) {
  const a = obtenerApp();
  if (!a) return;
  await addDoc(collection(getFirestore(a), "usuarios", uid, "conversiones"), { ...r, creado: serverTimestamp() });
}

export async function leerHistorial(uid: string): Promise<RegistroHistorial[]> {
  const a = obtenerApp();
  if (!a) return [];
  const q = query(collection(getFirestore(a), "usuarios", uid, "conversiones"), orderBy("creado", "desc"), limit(50));
  const s = await getDocs(q);
  return s.docs.map((d) => {
    const v = d.data();
    return { id: d.id, ...(v as Omit<RegistroHistorial, "id" | "creado">), creado: (v.creado as Timestamp | null)?.toDate() ?? null };
  });
}

export async function borrarDelHistorial(uid: string, id: string) {
  const a = obtenerApp();
  if (a) await deleteDoc(doc(getFirestore(a), "usuarios", uid, "conversiones", id));
}

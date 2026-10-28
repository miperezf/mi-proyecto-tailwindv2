// ── CÓDIGO DE ACCESO (v55.0) ─────────────────────────────────────────────────
// La app entra con UNA cuenta compartida de Firebase (correo + contraseña). El
// correo va en la variable VITE_ACCESS_EMAIL; la contraseña es el "código de
// acceso" que se escribe una sola vez por equipo. Sin la variable, la app sigue
// entrando como antes (anónima).
// Para cambiar el código: Firebase → Authentication → eliminar la cuenta y crearla
// de nuevo con el mismo correo y el código nuevo. Todos los equipos lo vuelven a pedir.
import React, { useState } from "react";
import { Icon, Spinner } from "./ui";

export const ACCESS_EMAIL = String(import.meta.env.VITE_ACCESS_EMAIL || "").trim().toLowerCase();
export const CODIGO_ACTIVO = ACCESS_EMAIL !== "";
/** El correo configurado tiene forma de correo (evita pegar "VITE_ACCESS_EMAIL = ..."). */
export const CORREO_VALIDO = /^[^\s@=]+@[^\s@=]+\.[^\s@=]+$/.test(ACCESS_EMAIL);

/** ¿El usuario de Firebase es la cuenta compartida? */
export const esCuentaCompartida = (user: any): boolean =>
  !!user && !user.isAnonymous && String(user.email || "").toLowerCase() === ACCESS_EMAIL;

// ── Identificador de ESTE equipo (navegador). Distingue equipos en el aviso
// "también está editando" ahora que todos comparten la misma cuenta.
const EQUIPO_KEY = "frutam_equipo";
const nuevoId = () => {
  const abc = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let s = "";
  const r = new Uint32Array(6);
  try { crypto.getRandomValues(r); } catch { for (let i = 0; i < 6; i++) r[i] = Math.floor(Math.random() * 1e9); }
  r.forEach((n) => { s += abc[n % abc.length]; });
  return s;
};
export const EQUIPO_ID: string = (() => {
  try {
    const v = localStorage.getItem(EQUIPO_KEY);
    if (v && /^[A-Z0-9]{6}$/.test(v)) return v;
    const n = nuevoId();
    localStorage.setItem(EQUIPO_KEY, n);
    return n;
  } catch {
    return nuevoId();
  }
})();

/** Id del documento de presencia: con código, "uid_EQUIPO"; sin código, el uid (como antes). */
export const idPresencia = (uid: string | null): string | null =>
  !uid ? null : CODIGO_ACTIVO ? `${uid}_${EQUIPO_ID}` : uid;

/** Nombre que se muestra en el aviso de presencia. */
export const etiquetaPresencia = (id: unknown): string => {
  const s = String(id ?? "");
  const i = s.lastIndexOf("_");
  return i > 0 ? `EQUIPO ${s.slice(i + 1)}` : s.substring(0, 8).toUpperCase();
};

// ── Historial por equipo: Mail ID enviados desde este navegador. Con la cuenta
// compartida, lastModifiedBy es igual para todos; esta lista mantiene el historial
// "mis envíos" de cada equipo. También se recuerda el uid anónimo anterior para no
// perder el historial previo a la activación del código.
const ENVIADOS_KEY = "frutam_enviados_equipo";
const UID_ANTERIOR_KEY = "frutam_uid_anonimo";
const leer = (k: string): string[] => {
  try { const a = JSON.parse(localStorage.getItem(k) || "[]"); return Array.isArray(a) ? a.map(String) : []; } catch { return []; }
};
export const registrarEnviado = (mailId: string) => {
  if (!mailId) return;
  try { localStorage.setItem(ENVIADOS_KEY, JSON.stringify([mailId, ...leer(ENVIADOS_KEY).filter((x) => x !== mailId)].slice(0, 30))); } catch {}
};
export const recordarUidAnterior = (uid: string) => {
  if (!uid) return;
  try { localStorage.setItem(UID_ANTERIOR_KEY, JSON.stringify([uid, ...leer(UID_ANTERIOR_KEY).filter((x) => x !== uid)].slice(0, 3))); } catch {}
};
/** ¿Este pedido enviado es "mío" (de este equipo)? */
export const esEnvioDeEsteEquipo = (o: any, uid: string | null): boolean => {
  const h = o?.header || {};
  if (!CODIGO_ACTIVO) return h.lastModifiedBy === uid;
  return leer(ENVIADOS_KEY).includes(h.mailId) || leer(UID_ANTERIOR_KEY).includes(h.lastModifiedBy);
};

/** Mensaje para el usuario según el error de Firebase Auth. */
export const mensajeErrorCodigo = (code: string): string => {
  if (code === "auth/invalid-email" || code === "config/correo")
    return "La app tiene mal configurado el correo de acceso (variable VITE_ACCESS_EMAIL: debe ser solo el correo). Avisa al responsable de la app.";
  if (code === "auth/operation-not-allowed")
    return "El acceso con código no está habilitado en Firebase (Authentication → Correo electrónico/contraseña). Avisa al responsable de la app.";
  if (["auth/invalid-credential", "auth/wrong-password", "auth/user-not-found", "auth/invalid-login-credentials"].includes(code))
    return "Código incorrecto. Revisa mayúsculas y vuelve a intentarlo.";
  if (code === "auth/too-many-requests") return "Demasiados intentos seguidos. Espera unos minutos y vuelve a intentarlo.";
  if (code === "auth/network-request-failed") return "Sin conexión a internet. Revisa la red y vuelve a intentarlo.";
  if (code === "auth/user-disabled") return "El acceso está deshabilitado. Avisa al responsable de la app.";
  return `No se pudo ingresar (${code || "error desconocido"}).`;
};

type Props = { banner?: React.ReactNode; onIngresar: (codigo: string) => Promise<string | null> };

/** Pantalla C1: un campo de código, oculto con opción de ver. */
export function PantallaCodigo({ banner, onIngresar }: Props) {
  const [codigo, setCodigo] = useState("");
  const [ver, setVer] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [enviando, setEnviando] = useState(false);
  const logo = `${import.meta.env.BASE_URL}logo4.png`;

  const enviar = async (e: React.FormEvent) => {
    e.preventDefault();
    const c = codigo.trim();
    if (!c || enviando) return;
    setEnviando(true);
    setError(null);
    const err = await onIngresar(c);
    setEnviando(false);
    if (err) setError(err);
  };

  return (
    <div className="flex min-h-screen flex-col bg-bg text-fg">
      {banner}
      <main className="flex flex-1 items-center justify-center px-4 py-10">
        <form onSubmit={enviar} data-testid="pantalla-codigo"
          className="w-full max-w-sm rounded-2xl border border-line bg-surface p-7 shadow-xl">
          <div className="mb-6 flex flex-col items-center gap-3 text-center">
            <img src={logo} alt="" width="48" height="48" className="h-12 w-12" />
            <div>
              <h1 className="text-lg font-semibold text-fg">Pedidos Comercial Frutam</h1>
              <p className="text-sm text-muted">Ingresa el código de acceso</p>
            </div>
          </div>
          <div className="space-y-4">
            <div>
              <label className="label" htmlFor="codigo-acceso">Código de acceso</label>
              <div className="relative">
                <span className="pointer-events-none absolute top-1/2 left-3 -translate-y-1/2 text-subtle"><Icon name="key" className="h-4 w-4" /></span>
                <input id="codigo-acceso" name="codigo-acceso" type={ver ? "text" : "password"} autoComplete="current-password"
                  autoFocus value={codigo} onChange={(e) => { setCodigo(e.target.value); if (error) setError(null); }}
                  aria-invalid={!!error || undefined} aria-describedby={error ? "codigo-error" : undefined}
                  className={`field pr-10 pl-9 ${error ? "border-danger" : ""}`} />
                <button type="button" onClick={() => setVer((v) => !v)} aria-label={ver ? "Ocultar código" : "Mostrar código"}
                  className="absolute top-1/2 right-2 -translate-y-1/2 rounded-md p-1 text-muted hover:text-fg">
                  <Icon name={ver ? "eyeOff" : "eye"} className="h-4 w-4" />
                </button>
              </div>
            </div>
            {error && (
              <p id="codigo-error" role="alert" className="flex items-start gap-2 rounded-lg bg-danger-soft px-3 py-2 text-xs font-medium text-danger">
                <Icon name="warning" className="h-4 w-4 shrink-0" />{error}
              </p>
            )}
            <button type="submit" disabled={!codigo.trim() || enviando}
              className="flex h-10 w-full items-center justify-center gap-2 rounded-lg bg-primary text-sm font-semibold text-primary-fg transition-colors hover:bg-primary-hover disabled:cursor-not-allowed disabled:opacity-50">
              {enviando ? <Spinner /> : null}Ingresar
            </button>
            <p className="text-center text-xs leading-snug text-muted">Se pide solo la primera vez que abres la app en este equipo.</p>
          </div>
        </form>
      </main>
    </div>
  );
}

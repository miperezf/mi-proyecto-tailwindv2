// ── SUGERENCIAS DE LA APP (v56.0) ───────────────────────────────────────────
// Reemplaza el autocompletado del navegador (desactivado desde la v52 porque su
// lista capturaba ↑ ↓ y rompía la navegación entre celdas).
//
// Origen: pedidos ENVIADOS del equipo en los últimos 90 días (una consulta al día
// por navegador, guardada en localStorage) + lo enviado desde este equipo.
// Solo lectura: no agrega campos ni cambia reglas.
//
// Teclado:
//   · La lista aparece solo al ESCRIBIR (no al entrar a una celda con flechas).
//   · Nada viene preseleccionado: ↓ entra a la lista; Enter o Tab aceptan la opción
//     marcada; Esc cierra. ↑ sin opción marcada cierra y navega como siempre.
//   · Con la lista cerrada, las flechas se comportan exactamente como antes.
import React, { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { collection, getDocs, query, where } from "firebase/firestore";
import { normalizarTexto, SEPARADOR_DIRECTO } from "./oficinas";

export const CAMPOS_CABECERA = ["reDestinatarios", "deNombrePais", "nave", "exporta"];
export const CAMPOS_ITEM = ["especie", "variedad", "formato", "calibre", "categoria"];
const CAMPOS = new Set([...CAMPOS_CABECERA, ...CAMPOS_ITEM]);
const DIAS = 90;
const MAX_OPCIONES = 6;
const MAX_LARGO = 60;

type Conteo = { t: number; e: Record<string, number> }; // total y por especie
type Tabla = Record<string, Record<string, Conteo>>; // campo → valor → conteo

const hoy = () => new Date().toISOString().slice(0, 10);
const claveCache = (appId: string) => `frutam_sugerencias_${appId}`;
const claveLocales = (appId: string) => `frutam_sugerencias_locales_${appId}`;
const leerJSON = (k: string) => { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch { return null; } };
const guardarJSON = (k: string, v: unknown) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* sin espacio o bloqueado */ } };

let TABLA: Tabla = {};

const limpiarValor = (campo: string, v: unknown): string => {
  const s = String(v ?? "").trim().toUpperCase().replace(/\s+/g, " ");
  if (!s || s === "-" || s.length > MAX_LARGO) return "";
  if (campo === "exporta" && normalizarTexto(s).includes(SEPARADOR_DIRECTO.trim())) return ""; // Directo va por su botón
  return s;
};

const sumar = (tabla: Tabla, campo: string, valor: string, especie: string) => {
  if (!valor) return;
  const c = ((tabla[campo] ||= {})[valor] ||= { t: 0, e: {} });
  c.t += 1;
  if (especie) c.e[especie] = (c.e[especie] || 0) + 1;
};

const itemsDe = (items: unknown): any[] => {
  if (Array.isArray(items)) return items;
  if (typeof items === "string") { try { const v = JSON.parse(items); return Array.isArray(v) ? v : []; } catch { return []; } }
  return [];
};

/** Suma a la tabla los valores de un pedido (cabecera + filas). */
const acumular = (tabla: Tabla, header: any, items: unknown) => {
  CAMPOS_CABECERA.forEach((k) => sumar(tabla, k, limpiarValor(k, header?.[k]), ""));
  itemsDe(items).forEach((it) => {
    const especie = normalizarTexto(it?.especie);
    CAMPOS_ITEM.forEach((k) => sumar(tabla, k, limpiarValor(k, it?.[k]), k === "especie" ? "" : especie));
  });
};

const combinar = (a: Tabla, b: Tabla): Tabla => {
  const r: Tabla = JSON.parse(JSON.stringify(a || {}));
  Object.entries(b || {}).forEach(([campo, vals]) =>
    Object.entries(vals).forEach(([v, c]) => {
      const d = ((r[campo] ||= {})[v] ||= { t: 0, e: {} });
      d.t += c.t;
      Object.entries(c.e || {}).forEach(([es, n]) => { d.e[es] = (d.e[es] || 0) + n; });
    }));
  return r;
};

/** Carga las sugerencias: desde el caché si es de hoy; si no, consulta la base. */
export const actualizarSugerencias = async (db: any, appId: string) => {
  const cache = leerJSON(claveCache(appId));
  const locales = leerJSON(claveLocales(appId))?.tabla || {};
  if (cache?.v === 1 && cache.dia === hoy() && cache.tabla) {
    TABLA = combinar(cache.tabla, locales);
    return "cache";
  }
  if (cache?.tabla) TABLA = combinar(cache.tabla, locales); // mientras consulta, usa lo anterior
  try {
    const desde = Date.now() - DIAS * 86400000;
    const snap = await getDocs(query(collection(db, `artifacts/${appId}/public/data/pedidos`), where("header.updatedAt", ">=", desde)));
    const tabla: Tabla = {};
    snap.docs.forEach((d: any) => {
      const data = d.data() || {};
      if (data.header?.status === "sent") acumular(tabla, data.header, data.items);
    });
    guardarJSON(claveCache(appId), { v: 1, dia: hoy(), tabla });
    guardarJSON(claveLocales(appId), { tabla: {} }); // lo enviado aquí ya viene en la consulta
    TABLA = tabla;
    return "base";
  } catch (e) {
    console.warn("[FRUTAM:SUGERENCIAS] no se pudieron actualizar:", (e as any)?.code || e);
    return "error";
  }
};

/** Al enviar: suma de inmediato los valores enviados desde este equipo. */
export const registrarEnviados = (appId: string, pedidos: any[]) => {
  const locales: Tabla = leerJSON(claveLocales(appId))?.tabla || {};
  const nuevo: Tabla = {};
  (pedidos || []).forEach((p) => acumular(nuevo, p?.header, p?.items));
  guardarJSON(claveLocales(appId), { tabla: combinar(locales, nuevo) });
  TABLA = combinar(TABLA, nuevo);
};

export type Opcion = { valor: string; n: number; prefijo: number };

/** Opciones para un campo según lo escrito y la especie de la fila (si aplica). */
export const opcionesPara = (campo: string, texto: string, especie = "", base: string[] = []): Opcion[] => {
  const q = normalizarTexto(texto);
  if (!q || !CAMPOS.has(campo)) return [];
  const es = normalizarTexto(especie);
  const fuente: Record<string, Conteo> = { ...(TABLA[campo] || {}) };
  base.forEach((v) => { const s = limpiarValor(campo, v); if (s && !fuente[s]) fuente[s] = { t: 0, e: {} }; });
  const lista = Object.entries(fuente).flatMap(([valor, c]) => {
    const n = normalizarTexto(valor);
    if (n === q) return [];
    const inicio = n.startsWith(q);
    if (!inicio && !n.split(/[\s\-/(),.]+/).some((w) => w.startsWith(q))) return [];
    const ce = es && campo !== "especie" ? c.e?.[es] || 0 : 0;
    return [{ valor, inicio, ce, t: c.t }];
  });
  lista.sort((a, b) => Number(b.inicio) - Number(a.inicio) || b.ce - a.ce || b.t - a.t || a.valor.localeCompare(b.valor));
  return lista.slice(0, MAX_OPCIONES).map((x) => ({ valor: x.valor, n: x.ce || x.t, prefijo: x.inicio ? texto.trim().length : 0 }));
};

/** Escribe un valor en un input controlado por React (dispara su onChange). */
const escribirEnInput = (el: HTMLInputElement, valor: string) => {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
  setter?.call(el, valor);
  el.dispatchEvent(new Event("input", { bubbles: true }));
};

type Estado = { el: HTMLInputElement; opciones: Opcion[]; sel: number; rect: DOMRect } | null;

/** Capa única de sugerencias para todos los campos de la app (escritorio y móvil). */
export function CapaSugerencias() {
  const [estado, setEstado] = useState<Estado>(null);
  const estadoRef = useRef<Estado>(null);
  const suprimir = useRef(false);
  estadoRef.current = estado;

  useEffect(() => {
    const cerrar = () => setEstado(null);
    const abrirPara = (el: HTMLInputElement) => {
      const opciones = opcionesPara(el.name, el.value, el.dataset.especie || "", (el.dataset.sugBase || "").split("|").filter(Boolean));
      if (!opciones.length) return cerrar();
      setEstado({ el, opciones, sel: -1, rect: el.getBoundingClientRect() });
    };
    const aceptar = (s: NonNullable<Estado>, i: number) => {
      const op = s.opciones[i];
      if (!op) return;
      suprimir.current = true;
      escribirEnInput(s.el, op.valor);
      cerrar();
    };

    const onInput = (e: Event) => {
      const el = e.target as HTMLInputElement;
      if (!(el instanceof HTMLInputElement) || !CAMPOS.has(el.name) || el.readOnly) return;
      if (suprimir.current) { suprimir.current = false; return; }
      abrirPara(el);
    };
    const onKey = (e: KeyboardEvent) => {
      const s = estadoRef.current;
      if (!s || e.target !== s.el) return;
      const consumir = () => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); };
      if (e.key === "ArrowDown") { consumir(); setEstado({ ...s, sel: Math.min(s.sel + 1, s.opciones.length - 1) }); }
      else if (e.key === "ArrowUp") {
        if (s.sel < 0) cerrar(); // sin selección: cierra y navega como siempre
        else { consumir(); setEstado({ ...s, sel: s.sel - 1 }); } // desde la 1.ª opción vuelve al campo
      }
      else if (e.key === "Enter") { if (s.sel >= 0) { consumir(); aceptar(s, s.sel); } else cerrar(); }
      else if (e.key === "Tab") { if (s.sel >= 0) aceptar(s, s.sel); else cerrar(); } // Tab sigue su curso
      else if (e.key === "Escape") { consumir(); cerrar(); }
    };
    const onFocusOut = (e: FocusEvent) => { if (estadoRef.current && e.target === estadoRef.current.el) cerrar(); };
    const reubicar = () => { const s = estadoRef.current; if (s) setEstado({ ...s, rect: s.el.getBoundingClientRect() }); };

    document.addEventListener("input", onInput, true);
    document.addEventListener("keydown", onKey, true);
    document.addEventListener("focusout", onFocusOut, true);
    window.addEventListener("scroll", reubicar, true);
    window.addEventListener("resize", reubicar);
    return () => {
      document.removeEventListener("input", onInput, true);
      document.removeEventListener("keydown", onKey, true);
      document.removeEventListener("focusout", onFocusOut, true);
      window.removeEventListener("scroll", reubicar, true);
      window.removeEventListener("resize", reubicar);
    };
  }, []);

  if (!estado) return null;
  const { rect, opciones, sel } = estado;
  const ancho = Math.min(Math.max(rect.width, 220), 360);
  const izq = Math.max(8, Math.min(rect.left, window.innerWidth - ancho - 8));
  const abajo = rect.bottom + 4;
  const caben = window.innerHeight - abajo > 60 + opciones.length * 36;
  const estilo: React.CSSProperties = caben
    ? { position: "fixed", left: izq, top: abajo, width: ancho, zIndex: 90 }
    : { position: "fixed", left: izq, bottom: window.innerHeight - rect.top + 4, width: ancho, zIndex: 90 };

  return createPortal(
    <div style={estilo} role="listbox" aria-label="Sugerencias" data-testid="sugerencias"
      className="overflow-hidden rounded-lg border border-line bg-surface text-sm shadow-xl">
      <div className="px-3 pt-2 pb-1 text-[11px] font-semibold tracking-wide text-muted uppercase">Usados antes</div>
      {opciones.map((o, i) => (
        <div key={o.valor} role="option" aria-selected={i === sel}
          onMouseDown={(e) => { e.preventDefault(); const s = estadoRef.current; if (s) { suprimir.current = true; escribirEnInput(s.el, o.valor); setEstado(null); } }}
          className={`flex cursor-pointer items-center justify-between gap-4 px-3 py-1.5 ${i === sel ? "bg-primary-soft text-primary" : "text-fg hover:bg-surface-2"}`}>
          <span className="truncate">
            {o.prefijo ? <><b>{o.valor.slice(0, o.prefijo)}</b>{o.valor.slice(o.prefijo)}</> : o.valor}
          </span>
          <span className="shrink-0 text-xs text-muted tabular-nums">{o.n > 0 ? `${o.n} ${o.n === 1 ? "vez" : "veces"}` : ""}</span>
        </div>
      ))}
      <div className="border-t border-line px-3 py-1.5 text-[11px] text-muted">↓ elegir · Enter aceptar · Esc cerrar</div>
    </div>,
    document.body
  );
}

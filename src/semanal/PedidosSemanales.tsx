// ── Sección "Pedidos semanales" (v58.0) ──────────────────────────────────────
// Una hoja por semana (W01–W52). Vista Pedido: una sola tabla Especie × Oficina
// donde se arma todo: cada especie es una fila con el pedido por oficina y, al
// abrirla, sus proveedores con los pallets asignados y "+ Agregar proveedor". Las celdas
// no se pintan: el estado de la asignación va en texto al final de la fila y el
// detalle por oficina en la fila "Falta" de la especie abierta. Al pie, los
// contenedores completos y los pallets sueltos de cada oficina. Vista
// Proveedores (consulta) y Ajustes para el catálogo.
// Todo se guarda al momento (a diferencia de los pedidos por Mail ID).
import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import ReactDOM from "react-dom";
import { Icon, Button, IconButton, Spinner } from "../ui";
import {
  type Asig, type Catalogo, type Semana, type SemanalStore,
  type Cuadratura, asigId, clave, cuadrar, currentWeek, familia, FAMILIAS, fmt, fmt0, gruposDuplicados, mismoQue, parecidos,
  parseCajas, parseWeek, shiftWeek, weekId, weekRange, weekShort,
} from "./datos";
import { CATALOGO_INICIAL } from "./catalogoInicial";
import Proveedores from "./Proveedores";

type Notify = (tone: string, message: React.ReactNode, ms?: number) => void;
type Vista = "pedido" | "proveedores" | "ajustes";
type Kind = "mk" | "esp" | "prov";
type Opt = { v: string; nuevo?: boolean; ya?: boolean; hint?: string };
type Celda = { cajas: number; ids: string[] };
type Bloque = {
  esp: string; ped: Record<string, number>; filas: Record<string, Record<string, Celda>>; provs: string[];
  pt: number; at: number; falta: Record<string, number>; ft: number; fp: number; fn: number; diff: boolean;
};
type Uso = Record<Kind, Record<string, Set<string>>>;

const SIN_PERMISO = "Firebase no permite leer ni guardar Pedidos semanales en esta base. Falta publicar las reglas de esta sección (Firestore → Reglas).";
const motivo = (e: unknown, txt: string) => ((e as { code?: string })?.code === "permission-denied" ? SIN_PERMISO : txt);
const VACIO: Catalogo = { especies: [], mercados: [], proveedores: [] };
const SEMANA_KEY = "frutam_semanal_semana";
const ABIERTAS_KEY = "frutam_semanal_abiertas";
// Filas sin cantidades todavía (especies y proveedores agregados o copiados de otra semana).
// Firestore solo guarda lo que tiene pallets, así que estas filas se recuerdan en este equipo.
const FILAS_KEY = "frutam_semanal_filas";
type FilasEsp = Record<string, string[]>;
type FilasProv = Record<string, Record<string, string[]>>;
const leerFilas = (): { esp: FilasEsp; prov: FilasProv } => {
  try {
    const x = JSON.parse(safeGet(FILAS_KEY) || "{}") || {};
    const esp: FilasEsp = {}; const prov: FilasProv = {};
    for (const [w, l] of Object.entries(x.esp || {})) if (Array.isArray(l)) esp[w] = l.filter((e) => typeof e === "string");
    for (const [w, o] of Object.entries(x.prov || {})) if (o && typeof o === "object") {
      prov[w] = {};
      for (const [e, l] of Object.entries(o as Record<string, unknown>)) if (Array.isArray(l)) prov[w][e] = l.filter((n) => typeof n === "string");
    }
    return { esp, prov };
  } catch { return { esp: {}, prov: {} }; }
};
const safeGet = (k: string) => { try { return localStorage.getItem(k); } catch { return null; } };
const safeSet = (k: string, v: string) => { try { localStorage.setItem(k, v); } catch { /* sin almacenamiento */ } };
const clonar = <T,>(x: T): T => JSON.parse(JSON.stringify(x));

// ── Celda numérica: edita local, guarda al salir ─────────────────────────────
// Sin caja a la vista: el borde aparece al pasar el mouse y al escribir.
// Como en una planilla: al llegar a la celda se selecciona todo y las flechas pasan a
// la celda vecina. Un segundo clic o F2 entra a "editar" (data-edit): ahí ← → mueven
// el cursor dentro del número. Esc deshace lo digitado.
function CeldaInput({ value, onCommit, label, mk, tono }: { value: number; onCommit: (v: number) => void; label: string; mk: string; tono: "pedido" | "prov" }) {
  const [txt, setTxt] = useState(fmt(value));
  const foco = useRef(false);
  const recienEnfocada = useRef(false);
  useEffect(() => { if (!foco.current) setTxt(fmt(value)); }, [value]);
  return (
    <input
      inputMode="numeric"
      aria-label={label}
      data-nav="cell"
      data-mk={mk}
      value={txt}
      className={`cell-input scroll-mt-16 scroll-mb-6 tabular-nums md:scroll-mt-44 ${tono === "prov" ? "text-[13px] text-muted" : "text-fg-2"}`}
      onFocus={(e) => { foco.current = true; recienEnfocada.current = true; e.currentTarget.select(); }}
      onMouseUp={(e) => { if (recienEnfocada.current) { e.preventDefault(); recienEnfocada.current = false; } else e.currentTarget.dataset.edit = "1"; }}
      onChange={(e) => { recienEnfocada.current = false; setTxt(e.target.value); }}
      onKeyDown={(e) => {
        recienEnfocada.current = false;
        const el = e.currentTarget;
        if (e.key === "F2") { e.preventDefault(); el.dataset.edit = "1"; el.setSelectionRange(el.value.length, el.value.length); }
        else if (e.key === "Escape" && txt !== fmt(value)) {
          e.preventDefault(); setTxt(fmt(value)); delete el.dataset.edit;
          requestAnimationFrame(() => { if (document.activeElement === el) el.select(); });
        }
      }}
      onBlur={(e) => { foco.current = false; recienEnfocada.current = false; delete e.currentTarget.dataset.edit; const v = parseCajas(txt); if (v !== (value || 0)) onCommit(v); setTxt(fmt(v)); }}
    />
  );
}

// ── Buscador con lista (agregar especie / proveedor) ─────────────────────────
// La lista se abre al hacer clic, al escribir o con Alt+↓. Si se llega con las flechas
// queda cerrada y ↑ ↓ siguen de largo por la tabla; con la lista abierta, ↑ ↓ eligen.
function Combo({ placeholder, getOptions, onPick, inputClassName, ariaLabel, nav }: {
  placeholder: string; getOptions: (q: string) => Opt[]; onPick: (o: Opt) => void; inputClassName: string;
  ariaLabel: string; nav?: string;
}) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const [idx, setIdx] = useState(0);
  const ref = useRef<HTMLInputElement>(null);
  const porPuntero = useRef(false);   // el foco llegó con clic o toque (no con el teclado)
  const movio = useRef(false);        // se movió por la lista con las flechas
  const [pos, setPos] = useState<{ left: number; top: number; bottom: number; width: number; up: boolean } | null>(null);
  const opts = open ? getOptions(q) : [];
  const ubicar = useCallback(() => {
    const el = ref.current; if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({ left: r.left, top: r.bottom + 4, bottom: window.innerHeight - r.top + 4, width: Math.max(r.width, 280), up: window.innerHeight - r.bottom < 300 });
  }, []);
  useEffect(() => {
    if (!open) return;
    ubicar();
    window.addEventListener("scroll", ubicar, true); window.addEventListener("resize", ubicar);
    return () => { window.removeEventListener("scroll", ubicar, true); window.removeEventListener("resize", ubicar); };
  }, [open, ubicar]);
  const elegir = (o: Opt) => { setQ(""); setIdx(0); setOpen(false); onPick(o); };
  const onKey = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Escape") { if (open || q) { e.preventDefault(); setQ(""); setOpen(false); } return; }
    if (e.key === "ArrowDown" && e.altKey && !open) { e.preventDefault(); movio.current = false; setIdx(0); setOpen(true); return; }
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      if (!open || !opts.length) return;           // lista cerrada: las flechas siguen por la tabla
      e.preventDefault(); movio.current = true;
      setIdx((i) => (i + (e.key === "ArrowDown" ? 1 : -1) + opts.length) % opts.length);
    } else if (e.key === "Enter") {
      if (!open || !opts.length) return;
      e.preventDefault(); if (!q.trim() && !movio.current) return;   // nada escrito ni elegido: no agrega por accidente
      const o = opts[Math.min(idx, opts.length - 1)]; if (o) elegir(o);
    }
  };
  return (
    <>
      <input
        ref={ref} value={q} placeholder={placeholder} aria-label={ariaLabel} autoComplete="off" role="combobox"
        aria-expanded={open && opts.length > 0} data-nav={nav} className={inputClassName}
        onMouseDown={() => { porPuntero.current = true; }}
        onFocus={() => { movio.current = false; setIdx(0); if (porPuntero.current) setOpen(true); porPuntero.current = false; }}
        onClick={() => setOpen(true)}
        onBlur={() => { setOpen(false); setQ(""); porPuntero.current = false; }}
        onChange={(e) => { setQ(e.target.value); setOpen(true); setIdx(0); movio.current = false; }}
        onKeyDown={onKey}
      />
      {open && opts.length > 0 && pos && ReactDOM.createPortal(
        <div role="listbox" className="fixed z-[70] max-h-80 overflow-y-auto rounded-xl border border-line bg-surface p-1.5 shadow-2xl"
          style={{ left: pos.left, width: pos.width, ...(pos.up ? { bottom: pos.bottom } : { top: pos.top }) }}>
          {opts.map((o, j) => (
            <button key={o.v + (o.nuevo ? "+" : "")} type="button" role="option" aria-selected={j === idx}
              onMouseDown={(e) => e.preventDefault()} onClick={() => elegir(o)}
              className={`flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-left text-sm ${j === idx ? "bg-primary-soft text-primary" : "text-fg hover:bg-primary-soft hover:text-primary"}`}>
              {o.nuevo ? <><Icon name="plus" className="h-4 w-4" strokeWidth={2.5} />Crear “{o.v}”</> : <span className="truncate">{o.v}</span>}
              {o.hint && <span className="ml-auto shrink-0 text-xs text-muted">{o.hint}</span>}
            </button>
          ))}
        </div>, document.body)}
    </>
  );
}

// Controles que solo se ven al pasar el mouse por la fila (siempre visibles en pantallas táctiles)
const AL_PASAR = "opacity-0 transition-opacity group-hover:opacity-100 focus-visible:opacity-100 [@media(hover:none)]:opacity-100";
// Filas de totales al pie de la tabla: se eligen en el menú "Totales" y se recuerdan en el equipo
const TOTALES = [
  { id: "pedido", label: "Pedido", desc: "Pallets pedidos por oficina" },
  { id: "asignado", label: "Asignado", desc: "Pallets asignados a proveedores" },
  { id: "contenedores", label: "Contenedores completos", desc: "21 pallets de manzana o 20 de otra especie" },
  { id: "sueltos", label: "Pallets sueltos", desc: "Los que no alcanzan a llenar otro contenedor" },
] as const;
type TotalId = typeof TOTALES[number]["id"];
type TotalesVis = Record<TotalId, boolean>;
const TOTALES_KEY = "frutam_semanal_totales";
const leerTotales = (): TotalesVis => {
  const v = { pedido: true, asignado: true, contenedores: true, sueltos: true };
  try { const x = JSON.parse(safeGet(TOTALES_KEY) || "{}") || {}; for (const t of TOTALES) if (typeof x[t.id] === "boolean") v[t.id] = x[t.id]; } catch { /* valores por defecto */ }
  return v;
};

function MenuTotales({ valor, onChange }: { valor: TotalesVis; onChange: (v: TotalesVis) => void }) {
  const [open, setOpen] = useState(false);
  const btn = useRef<HTMLButtonElement>(null);
  const pop = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; width: number; top: number; bottom: number; up: boolean } | null>(null);
  useEffect(() => {
    if (!open) return;
    const ubicar = () => {
      const r = btn.current?.getBoundingClientRect(); if (!r) return;
      const width = Math.min(320, window.innerWidth - 16);
      const left = Math.max(8, Math.min(r.right - width, window.innerWidth - width - 8));
      setPos({ left, width, top: r.bottom + 4, bottom: window.innerHeight - r.top + 4, up: window.innerHeight - r.bottom < 300 });
    };
    ubicar();
    const fuera = (e: MouseEvent) => { const t = e.target as Node; if (!pop.current?.contains(t) && !btn.current?.contains(t)) setOpen(false); };
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") { setOpen(false); btn.current?.focus(); } };
    window.addEventListener("scroll", ubicar, true); window.addEventListener("resize", ubicar);
    document.addEventListener("mousedown", fuera); document.addEventListener("keydown", esc);
    return () => {
      window.removeEventListener("scroll", ubicar, true); window.removeEventListener("resize", ubicar);
      document.removeEventListener("mousedown", fuera); document.removeEventListener("keydown", esc);
    };
  }, [open]);
  const n = TOTALES.filter((t) => valor[t.id]).length;
  const todas = (on: boolean) => onChange({ pedido: on, asignado: on, contenedores: on, sueltos: on });
  const enlace = "rounded px-1.5 py-0.5 text-xs font-semibold text-primary hover:bg-primary-soft";
  return (
    <>
      <button ref={btn} type="button" onClick={() => setOpen((o) => !o)} aria-haspopup="true" aria-expanded={open} data-menu-totales
        title="Elegir qué filas de totales se ven al pie de la tabla"
        className={`inline-flex h-8 shrink-0 items-center gap-1.5 rounded-lg border px-2.5 text-xs font-semibold transition-colors ${open ? "border-primary-line bg-primary-soft text-primary" : "border-line-strong bg-surface text-fg-2 hover:bg-surface-2"}`}>
        Totales <span className="font-normal text-muted">{n === TOTALES.length ? "todos" : n === 0 ? "ocultos" : `${n} de ${TOTALES.length}`}</span>
        <Icon name={open ? "chevronUp" : "chevronDown"} className="h-3.5 w-3.5" strokeWidth={2.5} />
      </button>
      {open && pos && ReactDOM.createPortal(
        <div ref={pop} role="group" aria-label="Filas de totales" className="fixed z-[70] rounded-xl border border-line bg-surface p-1.5 shadow-2xl"
          style={{ left: pos.left, width: pos.width, ...(pos.up ? { bottom: pos.bottom } : { top: pos.top }) }}>
          <p className="px-2.5 pt-1.5 pb-1 text-[11px] font-semibold tracking-wide text-muted uppercase">Filas al pie de la tabla</p>
          {TOTALES.map((t) => (
            <div key={t.id} className="flex items-center gap-1 rounded-lg hover:bg-surface-2">
              <label className="flex min-w-0 flex-1 cursor-pointer items-center gap-2.5 px-2.5 py-1.5">
                <input type="checkbox" checked={valor[t.id]} onChange={(e) => onChange({ ...valor, [t.id]: e.target.checked })} className="h-4 w-4 shrink-0 accent-[var(--color-primary)]" />
                <span className="min-w-0">
                  <span className="block text-sm font-medium text-fg">{t.label}</span>
                  <span className="block truncate text-[11px] text-muted">{t.desc}</span>
                </span>
              </label>
              <button type="button" onClick={() => onChange({ pedido: false, asignado: false, contenedores: false, sueltos: false, [t.id]: true })}
                title={`Ver solo ${t.label}`} className={`mr-1.5 ${enlace}`}>solo</button>
            </div>
          ))}
          <div className="mt-1 flex items-center justify-between border-t border-line px-1 pt-1.5 pb-0.5">
            <button type="button" onClick={() => todas(true)} className={enlace}>Ver todas</button>
            <button type="button" onClick={() => todas(false)} className={enlace}>Ocultar todas</button>
          </div>
        </div>, document.body)}
    </>
  );
}

// "+ Agregar especie" y "+ Agregar proveedor": borde punteado y texto oscuro para que se vean sin recargar
const COMBO_AGREGAR = "h-8 w-full rounded-lg border border-dashed border-line-strong bg-surface px-2.5 text-[13px] font-medium text-fg placeholder:font-semibold placeholder:text-fg-2 scroll-mt-16 md:scroll-mt-44 hover:border-primary hover:placeholder:text-primary focus:border-solid focus:border-primary focus:ring-2 focus:ring-primary/20 focus:outline-none focus:placeholder:font-normal focus:placeholder:text-subtle";
const mayusInicial = (x: string) => x.charAt(0).toLocaleUpperCase("es") + x.slice(1);

// ══════════════════════════════════════════════════════════════════════════════
export default function PedidosSemanales({ store, notify, onPantallaCompleta }: {
  store: SemanalStore; notify: Notify; onPantallaCompleta?: (activa: boolean) => void;
}) {
  const [semana, setSemana] = useState(() => {
    const v = safeGet(SEMANA_KEY);
    return v && /^\d{4}-W\d{2}$/.test(v) && parseWeek(v).w <= 52 ? v : currentWeek();
  });
  const [vista, setVista] = useState<Vista>("pedido");
  // Pantalla completa: oculta la barra de la app, usa todo el ancho y pide pantalla completa al navegador
  const [full, setFull] = useState(false);
  const alternarFull = () => {
    const on = !full; setFull(on);
    try {
      if (on) document.documentElement.requestFullscreen?.().catch(() => { /* el navegador no lo permite: queda el modo ampliado */ });
      else if (document.fullscreenElement) document.exitFullscreen?.().catch(() => { /* ya salió */ });
    } catch { /* sin API de pantalla completa */ }
  };
  useEffect(() => { onPantallaCompleta?.(full); }, [full, onPantallaCompleta]);
  useEffect(() => {
    if (!full) return;
    let entro = !!document.fullscreenElement;
    const cambio = () => { if (document.fullscreenElement) entro = true; else if (entro) setFull(false); }; // salió con Esc
    document.addEventListener("fullscreenchange", cambio);
    return () => document.removeEventListener("fullscreenchange", cambio);
  }, [full]);
  const [abiertas, setAbiertas] = useState<Record<string, boolean>>(() => { try { return JSON.parse(safeGet(ABIERTAS_KEY) || "{}") || {}; } catch { return {}; } });
  useEffect(() => { safeSet(ABIERTAS_KEY, JSON.stringify(abiertas)); }, [abiertas]);
  const [totales, setTotales] = useState<TotalesVis>(leerTotales);
  useEffect(() => { safeSet(TOTALES_KEY, JSON.stringify(totales)); }, [totales]);
  const [barH, setBarH] = useState(64);
  const barraRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = barraRef.current; if (!el) return;
    const medir = () => setBarH(el.offsetHeight);
    medir();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(medir); ro.observe(el);
    return () => ro.disconnect();
  }, []);
  const [cat, setCat] = useState<Catalogo | null>(null);
  const [catOk, setCatOk] = useState(false);
  const [semanas, setSemanas] = useState<Record<string, Semana>>({});
  const [semOk, setSemOk] = useState(false);
  const [asig, setAsig] = useState<Asig[]>([]);
  const [asigOk, setAsigOk] = useState(false);
  const [extraEsp, setExtraEsp] = useState<FilasEsp>(() => leerFilas().esp);
  const [extraProv, setExtraProv] = useState<FilasProv>(() => leerFilas().prov);
  useEffect(() => {
    const ws = [...new Set([...Object.keys(extraEsp), ...Object.keys(extraProv)])].sort().slice(-12); // últimas 12 semanas
    safeSet(FILAS_KEY, JSON.stringify({
      esp: Object.fromEntries(ws.filter((w) => extraEsp[w]?.length).map((w) => [w, extraEsp[w]])),
      prov: Object.fromEntries(ws.filter((w) => Object.values(extraProv[w] || {}).some((l) => l.length)).map((w) => [w, extraProv[w]])),
    }));
  }, [extraEsp, extraProv]);
  const [copiando, setCopiando] = useState(false);
  const [pendientes, setPendientes] = useState(0);
  const [errorGuardar, setErrorGuardar] = useState(false);
  const [aviso, setAviso] = useState("");
  const [uso, setUso] = useState<Uso | null>(null);

  const catRef = useRef(cat); catRef.current = cat;
  const asigRef = useRef(asig); asigRef.current = asig;
  const semanasRef = useRef(semanas); semanasRef.current = semanas;
  const cola = useRef<Promise<unknown>>(Promise.resolve());
  const gridRef = useRef<HTMLDivElement>(null);

  const enqueue = useCallback((fn: () => Promise<unknown>) => {
    setPendientes((p) => p + 1);
    cola.current = cola.current.then(fn)
      .then(() => setErrorGuardar(false))
      .catch((e) => { console.error("[FRUTAM:SEMANAL]", e); setErrorGuardar(true); notify("error", motivo(e, "No se pudo guardar. Revisa la conexión e inténtalo de nuevo."), 8000); })
      .finally(() => setPendientes((p) => p - 1));
    return cola.current;
  }, [notify]);

  // ── Suscripciones ───────────────────────────────────────────────────────────
  useEffect(() => {
    let okC = false, okS = false;
    const t = setTimeout(() => { if (!okC || !okS) setAviso("No se pudo conectar todavía; lo que ves puede no estar al día."); }, 10000);
    const fin = () => { if (okC && okS) { clearTimeout(t); setAviso(""); } };
    const u1 = store.subCatalogo((c) => { setCat(c); setCatOk(true); okC = true; fin(); },
      (e) => { console.error(e); setCatOk(true); okC = true; setAviso(motivo(e, "No se pudo leer el catálogo.")); });
    const u2 = store.subSemanas((s) => { setSemanas(s); setSemOk(true); okS = true; fin(); },
      (e) => { console.error(e); setSemOk(true); okS = true; setAviso(motivo(e, "No se pudieron leer las semanas.")); });
    return () => { clearTimeout(t); u1(); u2(); };
  }, [store]);
  useEffect(() => {
    setAsig([]); setAsigOk(false); safeSet(SEMANA_KEY, semana);
    return store.subAsig(semana, (a) => { setAsig(a); setAsigOk(true); },
      (e) => { console.error(e); setAsigOk(true); setAviso(motivo(e, "No se pudieron leer las asignaciones de esta semana.")); });
  }, [store, semana]);

  // ── Modelo ──────────────────────────────────────────────────────────────────
  const C = cat || VACIO;
  const mks = C.mercados;
  const provBy = (n: string) => C.proveedores.find((p) => p.nombre === n);
  const eidx = (e: string) => { const i = C.especies.indexOf(e); return i < 0 ? 999 : i; };
  const ordenar = (arr: string[]) => [...new Set(arr)].sort((a, b) => eidx(a) - eidx(b) || a.localeCompare(b));
  const pedido = semanas[semana]?.pedido || {};

  const M = useMemo(() => {
    const ex = extraEsp[semana] || []; const xp = extraProv[semana] || {};
    const esps = ordenar([...Object.keys(pedido).filter((e) => Object.values(pedido[e] || {}).some((v) => +v > 0)), ...asig.map((a) => a.especie), ...ex]);
    let pedTot = 0, asgTot = 0;
    const bloques: Bloque[] = esps.map((e) => {
      const p = pedido[e] || {}; const filas: Bloque["filas"] = {};
      for (const a of asig.filter((x) => x.especie === e)) {
        const f = (filas[a.proveedor] ||= {}); const c = f[a.mercado];
        f[a.mercado] = c ? { cajas: c.cajas + (+a.cajas || 0), ids: [...c.ids, a.id] } : { cajas: +a.cajas || 0, ids: [a.id] };
      }
      for (const n of xp[e] || []) filas[n] ||= {};
      const codes = new Set([...mks.map((m) => m.code), ...Object.keys(p), ...Object.values(filas).flatMap((r) => Object.keys(r))]);
      let pt = 0, at = 0; const falta: Record<string, number> = {};
      for (const m of codes) {
        const pv = +p[m] || 0; const av = Object.values(filas).reduce((s, r) => s + (r[m]?.cajas || 0), 0);
        pt += pv; at += av; falta[m] = pv - av;
      }
      pedTot += pt; asgTot += at;
      const fp = Object.values(falta).reduce((s, v) => s + (v > 0 ? v : 0), 0);
      const fn = Object.values(falta).reduce((s, v) => s + (v < 0 ? -v : 0), 0);
      return { esp: e, ped: p, filas, provs: Object.keys(filas).sort((a, b) => a.localeCompare(b)), pt, at, falta, ft: pt - at, fp, fn, diff: fp > 0 || fn > 0 };
    });
    return { bloques, pedTot, asgTot, falta: bloques.reduce((s, b) => s + b.fp, 0), demas: bloques.reduce((s, b) => s + b.fn, 0) };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pedido, asig, extraEsp, extraProv, semana, cat]);

  // Oficinas (columnas): las del catálogo + las que aparezcan en el pedido o en asignaciones
  const O = useMemo(() => {
    const codes = mks.map((m) => m.code);
    for (const e of Object.keys(pedido)) for (const m of Object.keys(pedido[e] || {})) if (+pedido[e][m] > 0 && !codes.includes(m)) codes.push(m);
    for (const a of asig) if (+a.cajas > 0 && !codes.includes(a.mercado)) codes.push(a.mercado);
    const cuad: Record<string, Cuadratura> = {}; const asg: Record<string, number> = {};
    for (const m of codes) {
      const r: Record<string, number> = {};
      for (const e of Object.keys(pedido)) if (+pedido[e]?.[m] > 0) r[e] = +pedido[e][m];
      cuad[m] = cuadrar(r); asg[m] = 0;
    }
    for (const a of asig) if (+a.cajas > 0) asg[a.mercado] = (asg[a.mercado] || 0) + +a.cajas;
    return { codes, cuad, asg };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pedido, asig, cat]);

  // ── Escrituras ──────────────────────────────────────────────────────────────
  const saveCat = useCallback((mut: (c: Catalogo) => void) => enqueue(async () => {
    const next = clonar(catRef.current || VACIO); mut(next);
    catRef.current = next; setCat(next);
    await store.setCatalogo(next);
  }), [enqueue, store]);

  const setPed = (esp: string, mk: string, v: number) => { const s = semana; enqueue(() => store.setPedidoCelda(s, esp, mk, v)); };
  const setCeldaAsig = (esp: string, prov: string, mk: string, v: number) => {
    const s = semana;
    const ex = asigRef.current.filter((a) => a.especie === esp && a.proveedor === prov && a.mercado === mk);
    enqueue(async () => {
      if (v > 0) {
        await store.setAsig(s, ex[0]?.id || asigId(esp, prov, mk), { proveedor: prov, especie: esp, mercado: mk, cajas: v, creado: Date.now() });
        if (ex.length > 1) await store.borrarAsig(s, ex.slice(1).map((a) => a.id));
      } else if (ex.length) await store.borrarAsig(s, ex.map((a) => a.id));
      await store.tocarSemana(s);
    });
  };
  const aprenderProv = (nombre: string, esp: string) => {
    const p = provBy(nombre);
    if (!p) saveCat((c) => { c.proveedores.push({ nombre, especies: [esp] }); c.proveedores.sort((a, b) => a.nombre.localeCompare(b.nombre)); });
    else if (!(p.especies || []).includes(esp)) saveCat((c) => {
      const x = c.proveedores.find((y) => y.nombre === nombre); if (!x) return;
      x.especies = [...(x.especies || []), esp].sort((a, b) => c.especies.indexOf(a) - c.especies.indexOf(b));
    });
  };
  const quitarProv = (esp: string, prov: string) => {
    const s = semana;
    setExtraProv((x) => { const w = { ...(x[s] || {}) }; w[esp] = (w[esp] || []).filter((n) => n !== prov); return { ...x, [s]: w }; });
    const ids = asigRef.current.filter((a) => a.especie === esp && a.proveedor === prov).map((a) => a.id);
    if (ids.length) enqueue(() => store.borrarAsig(s, ids));
  };
  const borrarEspecie = (esp: string) => {
    const s = semana;
    setExtraEsp((x) => ({ ...x, [s]: (x[s] || []).filter((e) => e !== esp) }));
    setExtraProv((x) => { const w = { ...(x[s] || {}) }; delete w[esp]; return { ...x, [s]: w }; });
    const ids = asigRef.current.filter((a) => a.especie === esp).map((a) => a.id);
    enqueue(async () => { if (ids.length) await store.borrarAsig(s, ids); await store.borrarEspeciePedido(s, esp); })
      .then(() => notify("success", `${esp} borrado de ${weekShort(s)}`, 2500));
  };
  const borrarSemana = () => {
    const s = semana; const ids = asigRef.current.map((a) => a.id);
    setExtraEsp((x) => ({ ...x, [s]: [] })); setExtraProv((x) => ({ ...x, [s]: {} }));
    enqueue(async () => { if (ids.length) await store.borrarAsig(s, ids); await store.borrarSemana(s); })
      .then(() => notify("success", `Semana ${weekShort(s)} borrada`, 2500));
  };

  // ── Copiar especies y proveedores (sin cantidades) de la última semana con datos ──
  const semanaOrigen = useMemo(() => {
    const prev = Object.keys(semanas).filter((w) => w < semana).sort().reverse();
    const conPedido = (w: string) => Object.values(semanas[w]?.pedido || {}).some((r) => Object.values(r || {}).some((v) => +v > 0));
    return prev.find(conPedido) || prev[0] || "";
  }, [semanas, semana]);
  const copiarEstructura = async () => {
    const src = semanaOrigen, dst = semana; if (!src || copiando) return;
    setCopiando(true);
    try {
      const lista = (await store.listAsig(src)).filter((a) => +a.cajas > 0);
      const ped = semanasRef.current[src]?.pedido || {};
      const esps = [...new Set([...Object.keys(ped).filter((e) => Object.values(ped[e] || {}).some((v) => +v > 0)), ...lista.map((a) => a.especie)])];
      if (!esps.length) { notify("info", `${weekShort(src)} no tiene especies ni proveedores para copiar.`, 4000); return; }
      const provs: Record<string, string[]> = {};
      for (const a of lista) { const l = (provs[a.especie] ||= []); if (!l.includes(a.proveedor)) l.push(a.proveedor); }
      setExtraEsp((x) => ({ ...x, [dst]: [...new Set([...(x[dst] || []), ...esps])] }));
      setExtraProv((x) => { const w = { ...(x[dst] || {}) }; for (const [e, l] of Object.entries(provs)) w[e] = [...new Set([...(w[e] || []), ...l])]; return { ...x, [dst]: w }; });
      setAbiertas((a) => ({ ...a, ...Object.fromEntries(esps.map((e) => [e, true])) }));
      notify("success", `Se copiaron ${esps.length} especies y sus proveedores de ${weekShort(src)}, sin cantidades.`, 5000);
    } catch (e) { console.error(e); notify("error", motivo(e, `No se pudo leer ${weekShort(src)}.`), 6000); }
    finally { setCopiando(false); }
  };

  // ── Renombrar / unir (se aplica en todas las semanas guardadas) ─────────────
  const cargarUso = useCallback(async () => {
    const u: Uso = { esp: {}, prov: {}, mk: {} };
    const add = (k: Kind, n: string, id: string) => (u[k][n] ||= new Set()).add(id);
    try {
      for (const [id, w] of Object.entries(semanasRef.current)) {
        for (const [e, r] of Object.entries(w.pedido || {})) for (const [m, v] of Object.entries(r || {})) if (+v > 0) { add("esp", e, id); add("mk", m, id); }
        for (const a of await store.listAsig(id)) if (+a.cajas > 0) { add("esp", a.especie, id); add("prov", a.proveedor, id); add("mk", a.mercado, id); }
      }
      setUso(u);
    } catch (e) { console.error(e); }
  }, [store]);
  useEffect(() => { if (vista === "ajustes" && semOk) cargarUso(); }, [vista, semOk, cargarUso]);

  const migrar = async (kind: Kind, from: string, to: string) => {
    const f = kind === "esp" ? "especie" : kind === "prov" ? "proveedor" : "mercado";
    let n = 0;
    for (const id of new Set([...Object.keys(semanasRef.current), semana])) {
      let ch = false;
      if (kind !== "prov") {
        const d = await store.getSemana(id);
        if (d?.pedido) {
          const P = clonar(d.pedido); let pc = false;
          if (kind === "esp" && P[from]) { const r = P[to] || {}; for (const [m, v] of Object.entries(P[from])) r[m] = (+r[m] || 0) + (+v || 0); P[to] = r; delete P[from]; pc = true; }
          if (kind === "mk") for (const e of Object.keys(P)) { const r = P[e]; if (r && from in r) { r[to] = (+r[to] || 0) + (+r[from] || 0); delete r[from]; pc = true; } }
          if (pc) { await store.setPedido(id, P); ch = true; }
        }
      }
      const lista = await store.listAsig(id);
      const afectadas = lista.filter((a) => (a as any)[f] === from);
      for (const a of afectadas) await store.patchAsig(id, a.id, { [f]: to } as any);
      if (afectadas.length) {
        // si quedaron dos filas iguales (misma especie, proveedor y mercado), se suman
        const vistas: Record<string, { id: string; cajas: number }> = {}; const sobran: string[] = [];
        for (const a0 of await store.listAsig(id)) {
          const k = `${a0.especie}|${a0.proveedor}|${a0.mercado}`;
          if (vistas[k]) { vistas[k].cajas += +a0.cajas || 0; await store.patchAsig(id, vistas[k].id, { cajas: vistas[k].cajas }); sobran.push(a0.id); }
          else vistas[k] = { id: a0.id, cajas: +a0.cajas || 0 };
        }
        if (sobran.length) await store.borrarAsig(id, sobran);
        ch = true;
      }
      if (ch) n++;
    }
    return n;
  };
  const renombrar = (kind: Kind, from: string, to: string) => {
    setAviso("Actualizando semanas guardadas…");
    if (kind === "esp") setExtraEsp((x) => Object.fromEntries(Object.entries(x).map(([w, l]) => [w, l.map((e) => (e === from ? to : e))])));
    enqueue(async () => { const n = await migrar(kind, from, to); setAviso(""); notify("success", `“${from}” ahora es “${to}”${n ? ` (${n} semanas actualizadas)` : ""}`); await cargarUso(); });
  };
  const unir = (kind: "esp" | "prov", keep: string, otros0: string[]) => {
    const otros = otros0.filter((o) => o !== keep); if (!otros.length) return;
    saveCat((c) => {
      if (kind === "esp") {
        c.especies = c.especies.filter((e) => !otros.includes(e)); if (!c.especies.includes(keep)) c.especies.push(keep);
        for (const p of c.proveedores) p.especies = [...new Set((p.especies || []).map((e) => (otros.includes(e) ? keep : e)))];
      } else {
        const k = c.proveedores.find((p) => p.nombre === keep);
        for (const o of otros) { const x = c.proveedores.find((p) => p.nombre === o); if (x && k) k.especies = [...new Set([...(k.especies || []), ...(x.especies || [])])]; }
        c.proveedores = c.proveedores.filter((p) => !otros.includes(p.nombre));
      }
    });
    setExtraProv((x) => { const y = clonar(x); for (const w of Object.values(y)) for (const e of Object.keys(w)) w[e] = [...new Set(w[e].map((n) => (otros.includes(n) ? keep : n)))]; return y; });
    setExtraEsp((x) => Object.fromEntries(Object.entries(x).map(([w, l]) => [w, [...new Set(l.map((e) => (otros.includes(e) ? keep : e)))]])));
    setAviso("Uniendo duplicados en las semanas guardadas…");
    enqueue(async () => { for (const o of otros) await migrar(kind, o, keep); setAviso(""); notify("success", `${otros.map((o) => `“${o}”`).join(", ")} quedó unido con “${keep}”`); await cargarUso(); });
  };

  // ── Navegación con teclado en la grilla ─────────────────────────────────────
  const ultimaCol = useRef("");   // columna (oficina) de la última celda visitada, para ↑ ↓
  const irA = (el: Element | null | undefined) => {
    if (!(el instanceof HTMLInputElement)) return;
    el.focus(); el.scrollIntoView({ block: "nearest" });
    if (el.dataset.nav === "cell") { el.select(); ultimaCol.current = el.dataset.mk || ultimaCol.current; }
  };
  const onGridKey = (e: React.KeyboardEvent<HTMLDivElement>) => {
    const t = e.target as HTMLInputElement;
    if (e.defaultPrevented || !t.dataset || !t.dataset.nav || e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
    const esCelda = t.dataset.nav === "cell"; const tr = t.closest("tr"); if (!tr || !gridRef.current) return;
    const todos = [...gridRef.current.querySelectorAll<HTMLInputElement>("input[data-nav]")];
    if (e.key === "Enter" && esCelda) { e.preventDefault(); irA(todos[todos.indexOf(t) + 1]); return; }
    if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
      if (!esCelda) return;
      if (t.dataset.edit) {   // editando dentro del número (segundo clic o F2): solo sale por los bordes
        const a = t.selectionStart ?? 0, z = t.selectionEnd ?? 0, L = t.value.length;
        if ((e.key === "ArrowLeft" && a > 0) || (e.key === "ArrowRight" && z < L)) return;
      }
      const fila = [...tr.querySelectorAll<HTMLInputElement>("input[data-nav='cell']")];
      e.preventDefault(); irA(fila[fila.indexOf(t) + (e.key === "ArrowRight" ? 1 : -1)]); return;
    }
    if (e.key === "ArrowUp" || e.key === "ArrowDown") {
      if (esCelda) ultimaCol.current = t.dataset.mk || "";
      const filas = [...gridRef.current.querySelectorAll("table[data-hoja] tbody tr")].filter((r) => r.querySelector("input[data-nav]"));
      const nr = filas[filas.indexOf(tr) + (e.key === "ArrowDown" ? 1 : -1)];
      e.preventDefault(); if (!nr) return;
      irA(nr.querySelector(`input[data-nav='cell'][data-mk="${CSS.escape(ultimaCol.current)}"]`) || nr.querySelector("input[data-nav]"));
    }
  };

  // ── Ordenar bloques: arrastrar desde ⠿ (o ↑ ↓ con el ⠿ enfocado) ───────────
  // El orden es el de la lista de especies del catálogo, así queda igual en todas
  // las semanas y en Proveedores. Al arrastrar, las especies se pliegan a su fila.
  const [arrastre, setArrastre] = useState<{ esp: string; orden: string[]; y: number } | null>(null);
  const arrRef = useRef(arrastre); arrRef.current = arrastre;
  const guardarOrden = useCallback((orden: string[]) => saveCat((c) => {
    for (const e of orden) if (!c.especies.includes(e)) c.especies.push(e);
    const lugares = orden.map((e) => c.especies.indexOf(e)).sort((a, b) => a - b);
    lugares.forEach((pos, i) => { c.especies[pos] = orden[i]; });
  }), [saveCat]);
  const moverBloque = (esp: string, d: number) => {
    const o = M.bloques.map((b) => b.esp); const i = o.indexOf(esp); const j = i + d;
    if (i < 0 || j < 0 || j >= o.length) return;
    [o[i], o[j]] = [o[j], o[i]]; guardarOrden(o);
    setTimeout(() => gridRef.current?.querySelector<HTMLElement>(`button[data-grip="${CSS.escape(esp)}"]`)?.focus(), 30);
  };
  const iniciarArrastre = (e: React.PointerEvent, esp: string) => {
    if (e.button !== 0 || M.bloques.length < 2) return;
    e.preventDefault();
    (document.activeElement as HTMLElement | null)?.blur?.();
    setArrastre({ esp, orden: M.bloques.map((b) => b.esp), y: e.clientY });
  };
  // Al empezar: dejar el título arrastrado bajo el puntero (los bloques se pliegan)
  useLayoutEffect(() => {
    const a = arrRef.current; if (!a) return;
    const el = gridRef.current?.querySelector<HTMLElement>(`tbody[data-esp="${CSS.escape(a.esp)}"]`);
    if (el) { const r = el.getBoundingClientRect(); window.scrollBy(0, r.top + r.height / 2 - a.y); }
  }, [arrastre?.esp]);
  useEffect(() => {
    if (!arrastre) return;
    const inicial = M.bloques.map((b) => b.esp).join("\u0000");
    let y = arrastre.y;
    const recolocar = () => {
      const a = arrRef.current; if (!a) return;
      const otros = [...(gridRef.current?.querySelectorAll<HTMLElement>("tbody[data-esp]") || [])].filter((s) => s.dataset.esp !== a.esp);
      let idx = otros.findIndex((s) => { const r = s.getBoundingClientRect(); return y < r.top + r.height / 2; });
      if (idx < 0) idx = otros.length;
      const resto = otros.map((s) => s.dataset.esp as string);
      const nuevo = [...resto.slice(0, idx), a.esp, ...resto.slice(idx)];
      if (nuevo.join("\u0000") !== a.orden.join("\u0000")) setArrastre({ ...a, orden: nuevo, y });
    };
    const mover = (ev: PointerEvent) => { y = ev.clientY; recolocar(); };
    const auto = setInterval(() => {
      const d = y < 90 ? -14 : y > window.innerHeight - 70 ? 14 : 0;
      if (d) { window.scrollBy(0, d); recolocar(); }
    }, 16);
    const soltar = () => {
      const a = arrRef.current; setArrastre(null);
      if (a && a.orden.join("\u0000") !== inicial) guardarOrden(a.orden);
      setTimeout(() => a && gridRef.current?.querySelector<HTMLElement>(`tbody[data-esp="${CSS.escape(a.esp)}"] tr`)?.scrollIntoView({ block: "nearest" }), 30);
    };
    const cancelar = () => setArrastre(null);
    const tecla = (ev: KeyboardEvent) => { if (ev.key === "Escape") cancelar(); };
    const prevSel = document.body.style.userSelect; document.body.style.userSelect = "none";
    window.addEventListener("pointermove", mover); window.addEventListener("pointerup", soltar);
    window.addEventListener("pointercancel", cancelar); window.addEventListener("keydown", tecla);
    return () => {
      clearInterval(auto); document.body.style.userSelect = prevSel;
      window.removeEventListener("pointermove", mover); window.removeEventListener("pointerup", soltar);
      window.removeEventListener("pointercancel", cancelar); window.removeEventListener("keydown", tecla);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [arrastre?.esp]);
  const bloquesVista = arrastre ? arrastre.orden.map((e) => M.bloques.find((b) => b.esp === e)).filter((b): b is Bloque => !!b) : M.bloques;

  // ── Opciones de los buscadores ──────────────────────────────────────────────
  const usadas = new Set(M.bloques.map((b) => b.esp));
  const opcionesEspecie = (raw: string): Opt[] => {
    const q = clave(raw); const all = C.especies;
    const hits = q ? all.filter((e) => clave(e).includes(q)).sort((a, b) => clave(a).indexOf(q) - clave(b).indexOf(q)) : all.filter((e) => !usadas.has(e));
    const sim = raw.trim() ? parecidos(all, raw).filter((e) => !hits.includes(e)).slice(0, 3) : [];
    const list: Opt[] = [...hits.filter((e) => !usadas.has(e)), ...hits.filter((e) => usadas.has(e))].slice(0, 10)
      .map((v) => ({ v, ya: usadas.has(v), hint: usadas.has(v) ? "ya está en la semana" : undefined }));
    for (const v of sim) list.push({ v, ya: usadas.has(v), hint: "parecido" });
    if (raw.trim() && !mismoQue(all, raw)) list.push({ v: mayusInicial(raw.trim()), nuevo: true });
    return list;
  };
  const elegirEspecie = (o: Opt) => {
    let v = o.v;
    if (o.nuevo) { const ex = mismoQue(C.especies, v); if (ex) v = ex; else saveCat((c) => { if (!c.especies.includes(v)) c.especies.push(v); }); }
    if (!usadas.has(v)) { setExtraEsp((x) => ({ ...x, [semana]: [...(x[semana] || []), v] })); setAbiertas((a) => ({ ...a, [v]: true })); }
    setTimeout(() => irA(gridRef.current?.querySelector(`tbody[data-esp="${CSS.escape(v)}"] input[data-nav='cell']`)), 50);
  };
  const opcionesProveedor = (esp: string, enBloque: string[]) => (raw: string): Opt[] => {
    const all = C.proveedores.map((p) => p.nombre); const q = clave(raw);
    const trabaja = (n: string) => (provBy(n)?.especies || []).includes(esp);
    const libres = all.filter((n) => !enBloque.includes(n));
    const hits = (q ? libres.filter((n) => clave(n).includes(q)) : libres.filter(trabaja))
      .sort((a, b) => (+trabaja(b) - +trabaja(a)) || (q ? clave(a).indexOf(q) - clave(b).indexOf(q) : 0) || a.localeCompare(b)).slice(0, 8);
    const sim = raw.trim() ? parecidos(libres, raw).filter((n) => !hits.includes(n)).slice(0, 3) : [];
    const list: Opt[] = [...hits.map((v) => ({ v, hint: trabaja(v) ? `trabaja ${esp}` : undefined })), ...sim.map((v) => ({ v, hint: "parecido" }))];
    const ex = mismoQue(all, raw);
    if (ex && enBloque.includes(ex)) list.unshift({ v: ex, ya: true, hint: "ya está en esta especie" });
    else if (raw.trim() && !ex) list.push({ v: raw.trim(), nuevo: true });
    return list;
  };
  const elegirProveedor = (esp: string) => (o: Opt) => {
    const ex = mismoQue(C.proveedores.map((p) => p.nombre), o.v); const n = ex || o.v.trim(); if (!n) return;
    aprenderProv(n, esp);
    setExtraProv((x) => { const w = { ...(x[semana] || {}) }; w[esp] = [...new Set([...(w[esp] || []), n])]; return { ...x, [semana]: w }; });
    if (o.nuevo && ex && ex !== o.v.trim()) notify("info", `Se usó “${ex}”, que ya existía.`, 3000);
    setTimeout(() => irA(gridRef.current?.querySelector(`tbody[data-esp="${CSS.escape(esp)}"] tr[data-prov="${CSS.escape(n)}"] input[data-nav='cell']`)), 50);
  };

  // ── Barra de la sección ─────────────────────────────────────────────────────
  const { y } = parseWeek(semana); const anioActual = parseWeek(currentWeek()).y;
  const ancho = full ? "max-w-none" : "max-w-6xl";
  const estadoGuardado = pendientes > 0
    ? <span className="inline-flex items-center gap-1 text-muted"><Spinner className="h-3 w-3" />Guardando…</span>
    : errorGuardar
      ? <span className="inline-flex items-center gap-1 text-danger"><Icon name="warning" className="h-3.5 w-3.5" />Sin guardar</span>
      : <span className="inline-flex items-center gap-1 text-success" title="Cada cambio se guarda al momento"><Icon name="check" className="h-3.5 w-3.5" />Guardado</span>;
  const tot = (n: number, label: string, color = "") => (
    <span className={`whitespace-nowrap ${color}`}><b className={`tabular-nums ${color || "text-fg"}`}>{fmt0(n)}</b> {label}</span>
  );
  const punto = <span aria-hidden="true" className="text-subtle">·</span>;
  const pestana = (v: Vista, label: string) => (
    <button type="button" onClick={() => setVista(v)} aria-pressed={vista === v} aria-label={label}
      className={`border-b-2 px-1 py-1.5 text-sm font-semibold transition-colors ${vista === v ? "border-primary text-primary" : "border-transparent text-muted hover:text-fg"}`}>
      {label}
    </button>
  );
  const barra = (
    <div ref={barraRef} className={`z-20 border-b border-line bg-bg/95 backdrop-blur md:sticky ${full ? "md:top-0" : "md:top-14"}`}>
      <div className={`mx-auto flex flex-wrap items-center gap-x-4 gap-y-2 px-3 py-2.5 sm:px-6 ${ancho}`}>
        <div className="min-w-0">
          <h2 className="text-base font-semibold text-fg">Pedidos semanales</h2>
          <p className="flex items-center gap-1.5 text-xs text-muted">{weekRange(semana)} {punto} {estadoGuardado}</p>
        </div>
        <div className="flex items-center gap-1.5">
          <Button size="sm" variant="secondary" icon="chevronLeft" aria-label="Semana anterior" title="Semana anterior" onClick={() => setSemana(shiftWeek(semana, -1))} />
          <select value={semana} onChange={(e) => setSemana(e.target.value)} aria-label="Semana"
            className="h-8 rounded-lg border border-line-strong bg-surface px-2 text-sm font-semibold text-fg">
            {Array.from({ length: 52 }, (_, i) => weekId(y, i + 1)).map((id) => (
              <option key={id} value={id}>{weekShort(id)}{y !== anioActual ? ` · ${y}` : ""}</option>
            ))}
          </select>
          <Button size="sm" variant="secondary" icon="chevronRight" aria-label="Semana siguiente" title="Semana siguiente" onClick={() => setSemana(shiftWeek(semana, 1))} />
        </div>
        <p data-totales className="mr-auto flex flex-wrap items-center gap-x-2 text-sm text-muted">
          {tot(M.pedTot, "pedido")} {punto} {tot(M.asgTot, "asignado")} {punto} {tot(M.falta, "falta", M.falta > 0 ? "font-semibold text-warning" : "")}
          {M.demas > 0 && <>{punto} {tot(M.demas, "de más", "font-semibold text-danger")}</>}
        </p>
        <div className="flex items-center gap-3">
          {pestana("pedido", "Pedido")}
          {pestana("proveedores", "Proveedores")}
          <div className="flex items-center">
            <IconButton label="Ajustes" icon="cog" size="sm" tone={vista === "ajustes" ? "primary" : "default"} aria-pressed={vista === "ajustes"} onClick={() => setVista("ajustes")} />
            <IconButton label={full ? "Salir de pantalla completa" : "Pantalla completa"} icon={full ? "collapse" : "expand"} size="sm" aria-pressed={full} onClick={alternarFull} />
          </div>
        </div>
      </div>
    </div>
  );

  const cargando = !catOk || !semOk || (vista !== "ajustes" && !asigOk);
  const nombreOfi = (c: string) => mks.find((m) => m.code === c)?.nombre || c;
  // Encabezado de oficinas fijo al bajar (solo si la tabla cabe sin deslizar hacia el lado)
  const cabFija = O.codes.length <= 9 || full;
  const thFija = cabFija ? "z-[11] bg-thead lg:sticky lg:top-[var(--semanal-top)]" : "bg-thead";
  const sinPermiso = aviso === SIN_PERMISO;
  const sinCatalogo = catOk && !cat && !sinPermiso;
  const todasAbiertas = M.bloques.length > 0 && M.bloques.every((b) => abiertas[b.esp]);
  // Contenedores por oficina: completos + pallets sueltos (los que no alcanzan a llenar otro)
  const contDe = (m: string) => { const q = O.cuad[m]; return q.completos ? q.necesarios : q.necesarios - 1; };
  const sueltosDe = (m: string) => { const q = O.cuad[m]; return q.completos ? 0 : q.sobran; };
  const contTot = O.codes.reduce((t, m) => t + contDe(m), 0);
  const filasPie = TOTALES.filter((t) => totales[t.id]);
  const alternarTodas = () => setAbiertas(todasAbiertas ? {} : Object.fromEntries(M.bloques.map((b) => [b.esp, true])));
  // Estado de la asignación de una especie: en texto, sin pintar las celdas del pedido
  const estadoAsig = (b: Bloque) => {
    if (b.pt === 0 && b.at === 0) return null;
    if (b.at === 0) return <span className="text-subtle">sin asignar</span>;
    if (!b.diff) return <Icon name="check" className="h-4 w-4 text-success" strokeWidth={2.5} aria-label="Asignación completa" />;
    return <span className="font-semibold text-warning">{[b.fp > 0 ? `faltan ${fmt0(b.fp)}` : "", b.fn > 0 ? `${fmt0(b.fn)} de más` : ""].filter(Boolean).join(" · ")}</span>;
  };

  return (
    <div className="flex min-w-0 flex-1 flex-col">
      {barra}
      <main ref={gridRef} onKeyDown={onGridKey} className={`mx-auto w-full flex-1 space-y-3 px-3 pt-4 pb-10 sm:px-6 ${ancho}`}>
        {aviso && <div role="status" className="rounded-xl border border-warning-line bg-warning-soft px-4 py-3 text-sm font-semibold text-warning">{aviso}</div>}
        {cargando ? (
          <div className="space-y-3">{[0, 1].map((i) => <div key={i} className="h-40 animate-pulse rounded-2xl border border-line bg-surface" />)}</div>
        ) : sinPermiso ? null : sinCatalogo ? (
          <section className="rounded-2xl border border-dashed border-line-strong bg-surface px-6 py-12 text-center">
            <p className="text-base font-semibold text-fg">Aún no hay catálogo de especies, mercados y proveedores.</p>
            <p className="mt-1 text-sm text-muted">Se puede cargar el de la planilla Pedidos_Semanales 2026 y luego editarlo en Ajustes.</p>
            <Button className="mt-4" variant="primary" icon="plus" onClick={() => enqueue(() => store.setCatalogo(CATALOGO_INICIAL))}>
              Cargar catálogo inicial ({CATALOGO_INICIAL.especies.length} especies · {CATALOGO_INICIAL.mercados.length} mercados · {CATALOGO_INICIAL.proveedores.length} proveedores)
            </Button>
          </section>
        ) : vista === "proveedores" ? (
          <Proveedores asig={asig} pedido={pedido} mks={mks} ordenar={ordenar} />
        ) : vista === "ajustes" ? (
          <Ajustes cat={C} uso={uso} saveCat={saveCat} renombrar={renombrar} unir={unir} notify={notify} />
        ) : (
          <>
            <div className={`rounded-2xl border border-line bg-surface shadow-xs ${cabFija ? "overflow-x-auto lg:overflow-visible" : "overflow-x-auto"}`}
              style={{ "--semanal-top": `${(full ? 0 : 56) + barH}px` } as React.CSSProperties}>
              <table data-hoja className="w-full table-fixed text-sm" style={{ minWidth: 380 + O.codes.length * 68 }}>
                <colgroup>
                  <col className="w-44 sm:w-60" />
                  {O.codes.map((m) => <col key={m} />)}
                  <col className="w-16" /><col className="w-28" /><col className="w-10" />
                </colgroup>
                <thead>
                  <tr className="text-center text-[11px] font-semibold tracking-wide text-thead-fg uppercase">
                    <th scope="col" className={`z-[12] bg-thead py-2 pr-2 pl-3 text-left align-middle sm:sticky sm:left-0 sm:pl-12 ${cabFija ? "lg:top-[var(--semanal-top)] lg:rounded-tl-2xl" : ""}`}>
                      <span className="flex flex-wrap items-center gap-x-3 gap-y-1">
                        Especie
                        {M.bloques.length > 0 && (
                          <button type="button" onClick={alternarTodas} aria-pressed={todasAbiertas}
                            title={todasAbiertas ? "Ocultar los proveedores de todas las especies" : "Mostrar los proveedores de todas las especies"}
                            className="inline-flex h-6 items-center gap-1 rounded-md border border-white/40 bg-white/10 px-2 text-[11px] font-semibold tracking-normal text-thead-fg normal-case transition-colors hover:bg-white/20">
                            <Icon name={todasAbiertas ? "chevronUp" : "chevronDown"} className="h-3 w-3" strokeWidth={2.5} />
                            {todasAbiertas ? "Cerrar todo" : "Abrir todo"}
                          </button>
                        )}
                      </span>
                    </th>
                    {O.codes.map((m) => <th key={m} scope="col" data-ofi={m} title={nombreOfi(m)} className={`${thFija} px-1 py-2 align-middle`}>{m}</th>)}
                    <th scope="col" className={`${thFija} px-1 py-2 align-middle`}>Total</th>
                    <th scope="col" className={`${thFija} px-2 py-2 text-left align-middle`}>Asignación</th>
                    <th scope="col" aria-label="Acciones" className={`${thFija} ${cabFija ? "lg:rounded-tr-2xl" : ""}`} />
                  </tr>
                </thead>
                {M.bloques.length === 0 && (
                  <tbody>
                    <tr><td colSpan={O.codes.length + 4} className="px-6 py-10 text-center">
                      <p className="text-base font-semibold text-fg">Sin pedido para {weekShort(semana)}.</p>
                      {semanaOrigen ? <>
                        <p className="mt-1 text-sm text-muted">Parte con las mismas especies y proveedores de {weekShort(semanaOrigen)} (sin cantidades) o agrega la primera especie con “+ Agregar especie”.</p>
                        <Button className="mt-4" variant="primary" icon="duplicate" disabled={copiando} onClick={copiarEstructura}>
                          Copiar especies y proveedores de {weekShort(semanaOrigen)}
                        </Button>
                      </> : <p className="mt-1 text-sm text-muted">Agrega la primera especie con “+ Agregar especie”.</p>}
                    </td></tr>
                  </tbody>
                )}
                {bloquesVista.map((b) => {
                  const abierta = !arrastre && !!abiertas[b.esp];
                  const movida = arrastre?.esp === b.esp;
                  const bg = movida ? "bg-primary-soft" : abierta ? "bg-surface-2" : "bg-surface";
                  const bgTd = movida || abierta ? bg : "bg-surface group-hover:bg-surface-2";
                  return (
                    <tbody key={b.esp} data-esp={b.esp}>
                      <tr className={`group border-t border-line ${bg} ${movida || abierta ? "" : "hover:bg-surface-2"}`}>
                        <td className={`z-[1] py-0.5 pr-1 pl-1 sm:sticky sm:left-0 ${bgTd}`}>
                          <div className="flex min-w-0 items-center">
                            {M.bloques.length > 1 ? (
                              <button type="button" data-grip={b.esp} aria-label={`Mover ${b.esp} (flechas arriba y abajo)`} title="Arrastra para ordenar · o usa ↑ ↓"
                                onPointerDown={(e) => iniciarArrastre(e, b.esp)}
                                onKeyDown={(e) => { if (e.key === "ArrowUp" || e.key === "ArrowDown") { e.preventDefault(); moverBloque(b.esp, e.key === "ArrowUp" ? -1 : 1); } }}
                                className={`flex h-7 w-5 shrink-0 touch-none items-center justify-center rounded-md text-subtle hover:bg-surface-3 hover:text-fg focus-visible:text-primary ${movida ? "cursor-grabbing text-primary" : "cursor-grab"}`}>
                                <Icon name="grip" className="h-4.5 w-4.5" />
                              </button>
                            ) : <span className="w-5 shrink-0" />}
                            <button type="button" data-abrir={b.esp} onClick={() => setAbiertas((a) => { const n = { ...a }; if (n[b.esp]) delete n[b.esp]; else n[b.esp] = true; return n; })}
                              aria-expanded={abierta} title={abierta ? "Ocultar proveedores" : `Ver y editar proveedores (${b.provs.length})`}
                              className="flex min-w-0 flex-1 items-center gap-1.5 rounded-md py-1 pr-1 pl-0.5 text-left hover:text-primary">
                              <Icon name={abierta ? "chevronDown" : "chevronRight"} className="h-3.5 w-3.5 shrink-0 text-subtle" strokeWidth={2.5} />
                              <span className="truncate font-semibold text-fg" title={b.esp}>{b.esp}</span>
                            </button>
                          </div>
                        </td>
                        {O.codes.map((m) => (
                          <td key={m} className="px-1 py-0.5">
                            <CeldaInput value={+b.ped[m] || 0} mk={m} tono="pedido" label={`Pedido ${b.esp} ${nombreOfi(m)}`} onCommit={(v) => setPed(b.esp, m, v)} />
                          </td>
                        ))}
                        <td className="px-1 text-center font-bold text-fg tabular-nums">{fmt0(b.pt)}</td>
                        <td data-asig className="px-2 text-left text-xs whitespace-nowrap">{estadoAsig(b)}</td>
                        <td className="text-center">
                          <IconButton label={`Borrar ${b.esp} de esta semana`} icon="trash" size="xs" tone="danger" className={AL_PASAR} onClick={() => borrarEspecie(b.esp)} />
                        </td>
                      </tr>
                      {abierta && b.provs.map((n) => {
                        const f = b.filas[n]; let t = 0;
                        return (
                          <tr key={n} data-prov={n} className="group border-t border-line bg-surface hover:bg-surface-2">
                            <td className="z-[1] bg-surface py-0.5 pr-2 pl-[3.25rem] group-hover:bg-surface-2 sm:sticky sm:left-0"><span className="block truncate text-[13px] text-muted" title={n}>{n}</span></td>
                            {O.codes.map((m) => { const v = f[m]?.cajas || 0; t += v; return (
                              <td key={m} className="px-1 py-0.5">
                                <CeldaInput value={v} mk={m} tono="prov" label={`${n} ${b.esp} ${nombreOfi(m)}`} onCommit={(x) => setCeldaAsig(b.esp, n, m, x)} />
                              </td>
                            ); })}
                            <td className="px-1 text-center text-[13px] text-muted tabular-nums">{t ? fmt0(t) : ""}</td>
                            <td />
                            <td className="text-center"><IconButton label={`Quitar ${n} de ${b.esp}`} icon="x" size="xs" tone="danger" className={AL_PASAR} onClick={() => quitarProv(b.esp, n)} /></td>
                          </tr>
                        );
                      })}
                      {abierta && b.diff && b.at > 0 && (
                        <tr data-falta className="border-t border-line bg-surface text-xs">
                          <td className="z-[1] bg-surface py-1 pr-2 pl-[3.25rem] font-medium text-muted sm:sticky sm:left-0">Falta</td>
                          {O.codes.map((m) => { const f = b.falta[m] || 0; return (
                            <td key={m} title={f < 0 ? `${fmt0(-f)} asignados de más` : undefined} className="px-1 text-center font-semibold text-warning tabular-nums">
                              {f > 0 ? fmt0(f) : f < 0 ? `+${fmt0(-f)}` : ""}
                            </td>
                          ); })}
                          <td className="px-1 text-center font-semibold text-warning tabular-nums">{b.fp > 0 ? fmt0(b.fp) : ""}</td>
                          <td colSpan={2} />
                        </tr>
                      )}
                      {abierta && (
                        <tr className="border-t border-line bg-surface">
                          <td className="z-[1] bg-surface py-1.5 pr-2 pl-11 sm:sticky sm:left-0">
                            <Combo placeholder="+ Agregar proveedor" ariaLabel={`Agregar proveedor a ${b.esp}`} nav="add" inputClassName={COMBO_AGREGAR}
                              getOptions={opcionesProveedor(b.esp, b.provs)} onPick={elegirProveedor(b.esp)} />
                          </td>
                          <td colSpan={O.codes.length + 3} />
                        </tr>
                      )}
                    </tbody>
                  );
                })}
                <tbody>
                  <tr className="border-t border-line bg-surface">
                    <td className={`z-[1] bg-surface py-2 pr-2 pl-6 sm:sticky sm:left-0 ${M.bloques.length && filasPie.length ? "" : "lg:rounded-bl-2xl"}`}>
                      <Combo placeholder="+ Agregar especie" ariaLabel="Agregar especie" nav="add" inputClassName={COMBO_AGREGAR}
                        getOptions={opcionesEspecie} onPick={elegirEspecie} />
                    </td>
                    <td colSpan={O.codes.length + 3} />
                  </tr>
                </tbody>
                {M.bloques.length > 0 && filasPie.length > 0 && (
                  <tfoot>
                    {filasPie.map((f, k) => (
                      <tr key={f.id} data-total={f.id} className={`text-muted ${f.id === filasPie[0].id || f.id === "contenedores" ? "border-t border-line-strong" : "border-t border-line"}`}>
                        <td title={f.desc} className={`z-[1] bg-surface py-1.5 pr-2 pl-3 text-left font-medium sm:sticky sm:left-0 sm:pl-12 ${cabFija && k === filasPie.length - 1 ? "lg:rounded-bl-2xl" : ""}`}>{f.label}</td>
                        {O.codes.map((m) => {
                          const q = O.cuad[m];
                          if (f.id === "pedido") return <td key={m} className="px-1 text-center tabular-nums">{q.pallets ? fmt0(q.pallets) : ""}</td>;
                          if (f.id === "asignado") return <td key={m} className="px-1 text-center tabular-nums">{O.asg[m] ? fmt0(O.asg[m]) : ""}</td>;
                          if (f.id === "contenedores") return <td key={m} className="px-1 text-center font-semibold text-fg tabular-nums">{q.pallets ? contDe(m) : ""}</td>;
                          const su = sueltosDe(m);
                          return <td key={m} title={su ? `Faltan ${q.faltan} pallets para completar otro contenedor (de ${q.cap} pallets)` : undefined}
                            className={`px-1 text-center tabular-nums ${su ? "font-semibold text-warning" : "text-subtle"}`}>{q.pallets ? su : ""}</td>;
                        })}
                        <td className={`px-1 text-center font-bold tabular-nums ${f.id === "contenedores" ? "text-fg" : ""}`}>
                          {f.id === "pedido" ? fmt0(M.pedTot) : f.id === "asignado" ? fmt0(M.asgTot) : f.id === "contenedores" ? contTot : ""}
                        </td>
                        <td colSpan={2} />
                      </tr>
                    ))}
                  </tfoot>
                )}
              </table>
            </div>
            {M.bloques.length > 0 && (
              <div className="flex flex-wrap items-start justify-end gap-x-3 gap-y-2 px-1">
                <p className="min-w-0 basis-full text-xs text-subtle sm:basis-0 sm:flex-1 sm:pt-1.5">La fila de cada especie es el pedido por oficina; ábrela para asignar proveedores. Pallets sueltos = los que no alcanzan a llenar un contenedor (21 de manzana, 20 de otras especies).</p>
                <MenuTotales valor={totales} onChange={setTotales} />
                <Button size="sm" variant="danger-ghost" icon="trash" onClick={borrarSemana}>Borrar semana</Button>
              </div>
            )}
          </>
        )}
      </main>
    </div>
  );
}

// ── Ajustes: catálogo como etiquetas agrupadas (crear, editar, eliminar, unir) ─
function Ajustes({ cat, uso, saveCat, renombrar, unir, notify }: {
  cat: Catalogo; uso: Uso | null; saveCat: (m: (c: Catalogo) => void) => unknown;
  renombrar: (k: Kind, from: string, to: string) => void; unir: (k: "esp" | "prov", keep: string, otros: string[]) => void; notify: Notify;
}) {
  const [edit, setEdit] = useState<{ kind: Kind; i: number } | null>(null);
  const [agregando, setAgregando] = useState<Kind | null>(null);
  const [f, setF] = useState<Record<string, string>>({});
  const provNames = cat.proveedores.map((p) => p.nombre);
  const usoTxt = (k: Kind, n: string) => { if (!uso) return "…"; const s = uso[k][n]; return s && s.size ? `usado en ${s.size} sem.` : "sin uso"; };

  const abrir = (kind: Kind, i: number) => {
    setAgregando(null); setEdit({ kind, i });
    if (kind === "mk") { const m = cat.mercados[i]; setF({ code: m.code, nombre: m.nombre, empresa: m.empresa || "Frutam" }); }
    else setF({ nombre: kind === "esp" ? cat.especies[i] : cat.proveedores[i].nombre });
  };
  const guardar = () => {
    if (!edit) return; const { kind, i } = edit;
    if (kind === "mk") {
      const from = cat.mercados[i].code; const code = (f.code || "").trim().toUpperCase();
      if (!code) return notify("error", "Escribe el código.");
      if (cat.mercados.some((m, j) => j !== i && m.code === code)) return notify("error", `Ya existe el mercado ${code}.`);
      setEdit(null);
      saveCat((c) => { const x = c.mercados.find((m) => m.code === from); if (!x) return; x.code = code; x.nombre = (f.nombre || code).trim(); x.empresa = f.empresa || "Frutam"; c.mercados.sort((a, b) => +(a.empresa === "Polar") - +(b.empresa === "Polar")); });
      if (code !== from) renombrar("mk", from, code); else notify("success", "Mercado actualizado", 2000);
      return;
    }
    const lista = kind === "esp" ? cat.especies : provNames; const from = lista[i]; const to = (f.nombre || "").trim();
    if (!to) return notify("error", "Escribe el nombre.");
    const otro = lista.find((x, j) => j !== i && clave(x) === clave(to));
    setEdit(null);
    if (otro) return unir(kind, otro, [from]);           // mismo nombre normalizado: se unen
    if (to === from) return;
    if (kind === "esp") saveCat((c) => { const k = c.especies.indexOf(from); if (k >= 0) c.especies[k] = to; for (const p of c.proveedores) p.especies = (p.especies || []).map((e) => (e === from ? to : e)); });
    else saveCat((c) => { const x = c.proveedores.find((p) => p.nombre === from); if (x) x.nombre = to; c.proveedores.sort((a, b) => a.nombre.localeCompare(b.nombre)); });
    renombrar(kind, from, to);
  };
  const eliminar = () => {
    if (!edit) return; const { kind, i } = edit; setEdit(null);
    const name = kind === "mk" ? cat.mercados[i].code : kind === "esp" ? cat.especies[i] : cat.proveedores[i].nombre;
    saveCat((c) => {
      if (kind === "mk") c.mercados = c.mercados.filter((m) => m.code !== name);
      else if (kind === "esp") { c.especies = c.especies.filter((e) => e !== name); for (const p of c.proveedores) p.especies = (p.especies || []).filter((e) => e !== name); }
      else c.proveedores = c.proveedores.filter((p) => p.nombre !== name);
    });
    notify("success", `${name} eliminado de la lista`, 2500);
  };
  const agregar = (kind: Kind) => {
    if (kind === "mk") {
      const code = (f.code || "").trim().toUpperCase(); if (!code) return;
      if (cat.mercados.some((m) => m.code === code)) return notify("error", `Ya existe el mercado ${code}.`);
      saveCat((c) => { c.mercados.push({ code, nombre: (f.nombre || code).trim(), empresa: f.empresa || "Frutam" }); c.mercados.sort((a, b) => +(a.empresa === "Polar") - +(b.empresa === "Polar")); });
    } else {
      const lista = kind === "esp" ? cat.especies : provNames; const v = (f.nombre || "").trim(); if (!v) return;
      const ex = mismoQue(lista, v); if (ex) return notify("info", `Ya existe como “${ex}”.`);
      const sim = parecidos(lista, v);
      if (kind === "esp") saveCat((c) => { c.especies.push(mayusInicial(v)); });
      else saveCat((c) => { c.proveedores.push({ nombre: v, especies: [] }); c.proveedores.sort((a, b) => a.nombre.localeCompare(b.nombre)); });
      notify("success", sim.length ? `${v} agregado. Ojo: se parece a “${sim[0]}”.` : `${v} agregado`, 3000);
    }
    setAgregando(null); setF({});
  };
  const onKey = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") { e.preventDefault(); if (edit) guardar(); else if (agregando) agregar(agregando); }
    if (e.key === "Escape") { setEdit(null); setAgregando(null); }
  };
  const inp = (k: string, ph: string, w: string, extra = "") => (
    <input autoFocus={k === "nombre" || k === "code"} value={f[k] || ""} placeholder={ph} aria-label={ph} onChange={(e) => setF({ ...f, [k]: e.target.value })} onKeyDown={onKey}
      className={`field h-8 py-0 ${w} ${extra}`} />
  );
  const selEmpresa = (
    <select value={f.empresa || "Frutam"} aria-label="Empresa" onChange={(e) => setF({ ...f, empresa: e.target.value })} className="field h-8 w-auto py-0">
      <option>Frutam</option><option>Polar</option>
    </select>
  );
  const editor = (kind: Kind, nombre: string) => (
    <span className="inline-flex flex-wrap items-center gap-1.5 rounded-xl border border-primary-line bg-primary-soft p-2">
      {kind === "mk" ? <>{inp("code", "Código", "w-20", "font-bold")}{inp("nombre", "País", "w-36")}{selEmpresa}</> : inp("nombre", kind === "esp" ? "Especie" : "Proveedor", "w-56")}
      <Button size="sm" variant="primary" onClick={guardar}>Guardar</Button>
      <Button size="sm" variant="secondary" onClick={() => setEdit(null)}>Cancelar</Button>
      <Button size="sm" variant="danger-ghost" icon="trash" onClick={eliminar}>Eliminar</Button>
      <span className="px-1 text-xs text-muted">{usoTxt(kind, nombre)}</span>
    </span>
  );
  const chip = (kind: Kind, i: number, label: string, sub?: string) =>
    edit && edit.kind === kind && edit.i === i ? <React.Fragment key={kind + i}>{editor(kind, label)}</React.Fragment> : (
      <button key={kind + i} type="button" onClick={() => abrir(kind, i)} title={`Editar ${label}`}
        className="inline-flex h-7 items-center gap-1.5 rounded-full bg-primary-soft px-3 text-xs font-semibold text-primary ring-primary-line hover:ring-1">
        {label}{sub && <span className="font-normal text-muted">{sub}</span>}
      </button>
    );
  const grupo = (titulo: string, chips: React.ReactNode) => (
    <div key={titulo} className="grid gap-2 border-t border-line py-3 sm:grid-cols-[130px_1fr]">
      <div className="pt-1 text-xs font-semibold tracking-wide text-muted uppercase">{titulo}</div>
      <div className="flex flex-wrap gap-1.5">{chips}</div>
    </div>
  );
  const seccion = (kind: Kind, titulo: string, n: number, cuerpo: React.ReactNode) => (
    <section className="rounded-2xl border border-line bg-surface p-4 shadow-xs sm:px-5">
      <div className="flex flex-wrap items-center gap-2 pb-3">
        <h3 className="mr-auto text-sm font-semibold text-fg">{titulo} · {n}</h3>
        {agregando === kind ? (
          <span className="inline-flex flex-wrap items-center gap-1.5 rounded-xl border border-primary-line bg-primary-soft p-2">
            {kind === "mk" ? <>{inp("code", "Código", "w-24")}{inp("nombre", "País", "w-36")}{selEmpresa}</> : inp("nombre", kind === "esp" ? "Nueva especie" : "Nuevo proveedor", "w-56")}
            <Button size="sm" variant="primary" onClick={() => agregar(kind)}>Agregar</Button>
            <Button size="sm" variant="secondary" onClick={() => setAgregando(null)}>Cancelar</Button>
          </span>
        ) : (
          <Button size="sm" variant="secondary" icon="plus" onClick={() => { setEdit(null); setF({}); setAgregando(kind); }}>Agregar</Button>
        )}
      </div>
      {cuerpo}
    </section>
  );
  const agrupar = <T,>(items: T[], key: (x: T) => string, orden: string[]) => {
    const g: Record<string, T[]> = {}; for (const it of items) (g[key(it)] ||= []).push(it);
    return Object.keys(g).sort((a, b) => { const ia = orden.indexOf(a), ib = orden.indexOf(b); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b); }).map((k) => [k, g[k]] as [string, T[]]);
  };
  const dups = [...gruposDuplicados(cat.especies).map((names) => ({ kind: "esp" as const, names })), ...gruposDuplicados(provNames).map((names) => ({ kind: "prov" as const, names }))];

  return (
    <>
      {dups.length > 0 && (
        <section className="rounded-2xl border border-warning-line bg-warning-soft p-4 sm:px-5">
          <h3 className="text-sm font-semibold text-warning">Posibles duplicados · {dups.length}</h3>
          <p className="mt-1 text-sm text-fg-2">Elige con qué nombre quedarte; los datos de todas las semanas se juntan en ese nombre.</p>
          {dups.map((g, gi) => (
            <div key={gi} className="mt-2 flex flex-wrap items-center gap-2 border-t border-warning-line pt-2">
              <span className="mr-auto text-sm text-fg">{g.kind === "esp" ? "Especie" : "Proveedor"}: <b>{g.names.join(" · ")}</b></span>
              {g.names.map((n) => <Button key={n} size="sm" variant="secondary" onClick={() => unir(g.kind, n, g.names)}>Dejar “{n}”</Button>)}
            </div>
          ))}
        </section>
      )}
      <p className="text-sm text-muted">Toca una etiqueta para editarla o eliminarla. Los cambios de nombre se aplican también en las semanas guardadas.</p>
      {seccion("mk", "Mercados", cat.mercados.length,
        agrupar(cat.mercados.map((m, i) => ({ m, i })), (x) => x.m.empresa || "Sin empresa", ["Frutam", "Polar"])
          .map(([g, xs]) => grupo(g, xs.map(({ m, i }) => chip("mk", i, m.code, m.nombre)))))}
      {seccion("esp", "Especies", cat.especies.length,
        agrupar(cat.especies.map((e, i) => ({ e, i })), (x) => familia(x.e), [...FAMILIAS.map((x) => x[0]), "Otras"])
          .map(([g, xs]) => grupo(g, xs.map(({ e, i }) => chip("esp", i, e)))))}
      {seccion("prov", "Proveedores", cat.proveedores.length,
        grupo("A–Z", cat.proveedores.map((p, i) => chip("prov", i, p.nombre))))}
    </>
  );
}

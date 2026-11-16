// ── Pedidos semanales: tipos, utilidades y contrato de almacenamiento (v56.0) ──
// La sección no importa Firebase directamente: recibe un SemanalStore. Así la
// lógica se prueba con un almacenamiento en memoria y en la app usa Firestore
// (ver ./firestore.ts).

export type Mercado = { code: string; nombre: string; empresa: string };
export type Proveedor = { nombre: string; especies: string[] };
export type Catalogo = { especies: string[]; mercados: Mercado[]; proveedores: Proveedor[] };
/** pedido[especie][mercado] = pallets (el campo de las asignaciones se llama "cajas" por historia) */
export type Pedido = Record<string, Record<string, number>>;
export type Semana = { semana: string; pedido?: Pedido; actualizado?: number };
export type Asig = { id: string; proveedor: string; especie: string; mercado: string; cajas: number; creado?: number };
export type AsigData = Omit<Asig, "id">;
type Unsub = () => void;

export interface SemanalStore {
  subCatalogo(cb: (c: Catalogo | null) => void, err: (e: unknown) => void): Unsub;
  subSemanas(cb: (s: Record<string, Semana>) => void, err: (e: unknown) => void): Unsub;
  subAsig(semana: string, cb: (a: Asig[]) => void, err: (e: unknown) => void): Unsub;
  setCatalogo(c: Catalogo): Promise<void>;
  /** Escribe una celda del pedido (crea la semana si no existe). */
  setPedidoCelda(semana: string, especie: string, mercado: string, cajas: number): Promise<void>;
  /** Reemplaza el pedido completo de la semana. */
  setPedido(semana: string, pedido: Pedido): Promise<void>;
  /** Quita una especie del pedido de la semana. */
  borrarEspeciePedido(semana: string, especie: string): Promise<void>;
  /** Marca la semana como usada (para que aparezca en la lista de semanas). */
  tocarSemana(semana: string): Promise<void>;
  borrarSemana(semana: string): Promise<void>;
  getSemana(semana: string): Promise<Semana | null>;
  listAsig(semana: string): Promise<Asig[]>;
  setAsig(semana: string, id: string, data: AsigData): Promise<void>;
  patchAsig(semana: string, id: string, patch: Partial<AsigData>): Promise<void>;
  /** Borra varias asignaciones en un solo lote. */
  borrarAsig(semana: string, ids: string[]): Promise<void>;
}

// ── Semanas ISO (W01–W52) ────────────────────────────────────────────────────
export function isoWeek(d: Date) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const y = t.getUTCFullYear();
  const w = Math.ceil(((+t - Date.UTC(y, 0, 1)) / 864e5 + 1) / 7);
  return { y, w };
}
export const weekId = (y: number, w: number) => `${y}-W${String(w).padStart(2, "0")}`;
export const parseWeek = (id: string) => { const [y, w] = id.split("-W"); return { y: +y, w: +w }; };
export const weekShort = (id: string) => `W${String(parseWeek(id).w).padStart(2, "0")}`;
/** El año se maneja con 52 semanas (la W53 ISO no se usa). */
export function shiftWeek(id: string, n: number) {
  let { y, w } = parseWeek(id);
  w += n;
  while (w > 52) { w -= 52; y++; }
  while (w < 1) { w += 52; y--; }
  return weekId(y, w);
}
export function currentWeek() { const { y, w } = isoWeek(new Date()); return weekId(y, Math.min(w, 52)); }
const MES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
export function weekRange(id: string) {
  const { y, w } = parseWeek(id);
  const j4 = new Date(Date.UTC(y, 0, 4)); const d = j4.getUTCDay() || 7;
  const a = new Date(j4); a.setUTCDate(j4.getUTCDate() - d + 1 + (w - 1) * 7);
  const b = new Date(a); b.setUTCDate(a.getUTCDate() + 6);
  return a.getUTCMonth() === b.getUTCMonth()
    ? `${a.getUTCDate()}–${b.getUTCDate()} ${MES[b.getUTCMonth()]} ${b.getUTCFullYear()}`
    : `${a.getUTCDate()} ${MES[a.getUTCMonth()]} – ${b.getUTCDate()} ${MES[b.getUTCMonth()]} ${b.getUTCFullYear()}`;
}

// ── Números (formato chileno) ────────────────────────────────────────────────
const nf = new Intl.NumberFormat("es-CL", { maximumFractionDigits: 1 });
export const fmt = (n: number) => (n ? nf.format(n) : "");
export const fmt0 = (n: number) => nf.format(n || 0);
export const parseCajas = (v: string) => { const n = parseFloat(String(v).replace(/\./g, "").replace(",", ".")); return isFinite(n) && n > 0 ? n : 0; };

// ── Nombres: comparación sin mayúsculas, tildes, espacios ni signos ──────────
export const norm = (x: string) => String(x).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
export const clave = (x: string) => norm(x).replace(/[^a-z0-9]/g, "");
export function lev(a: string, b: string) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (Math.abs(m - n) > 2) return 9;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[n];
}
/** Nombre existente equivalente (Garcés = garces = GaRcEs). */
export const mismoQue = (lista: string[], raw: string) => { const k = clave(raw); return k ? lista.find((x) => clave(x) === k) : undefined; };
/** Nombres parecidos (posible error de tipeo). */
export const parecidos = (lista: string[], raw: string) => {
  const k = clave(raw); if (k.length < 4) return [];
  const tol = k.length >= 7 ? 2 : 1;
  return lista.filter((x) => { const kx = clave(x); return kx !== k && lev(kx, k) <= tol; });
};
/** Grupos de posibles duplicados ya existentes en el catálogo. */
export function gruposDuplicados(lista: string[]) {
  const g: Record<string, string[]> = {};
  for (const x of lista) (g[clave(x)] ||= []).push(x);
  const ks = Object.keys(g);
  const out = ks.filter((k) => g[k].length > 1).map((k) => g[k]);
  for (let i = 0; i < ks.length; i++) for (let j = i + 1; j < ks.length; j++) {
    const a = ks[i], b = ks[j];
    if (Math.min(a.length, b.length) >= 9 && lev(a, b) <= 1) out.push([...g[a], ...g[b]]);
  }
  return out;
}

// ── Familias de especies (agrupación en Ajustes) ─────────────────────────────
export const FAMILIAS: [string, string[]][] = [
  ["Carozos", ["Cerezas", "Nectarin", "Durazno", "Ciruela", "Damasco"]],
  ["Manzanas", ["Manzana"]],
  ["Peras", ["Pera"]],
  ["Uvas", ["Uva", "Red Globe"]],
  ["Kiwis", ["Kiwi"]],
  ["Cítricos", ["Naranja", "Limón", "Limon", "Clementina", "Mandarina", "Pomelo"]],
];
export const familia = (e: string) => { for (const [f, pre] of FAMILIAS) if (pre.some((x) => e.startsWith(x))) return f; return "Otras"; };

/** Id determinístico de una asignación (especie + proveedor + mercado). */
export function asigId(especie: string, proveedor: string, mercado: string) {
  let x = 5381;
  for (const ch of `${especie}|${proveedor}|${mercado}`) x = ((x * 33) ^ (ch.codePointAt(0) || 0)) >>> 0;
  return "x" + x.toString(36);
}

// ── Contenedores ─────────────────────────────────────────────────────────────
// Un contenedor lleva 21 pallets de manzana o 20 de otra especie y se pueden mezclar
// especies, así que cada oficina suma pallets/capacidad y se mira si queda en
// contenedores enteros. Se calcula sobre el PEDIDO de cada oficina.
export const capacidad = (especie: string) => (familia(especie) === "Manzanas" ? 21 : 20);
export const fmt1 = (n: number) => n.toLocaleString("es-CL", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
export type Cuadratura = { pallets: number; cont: number; necesarios: number; completos: boolean; faltan: number; sobran: number; cap: number };

/** Cuadratura de una oficina a partir de sus pallets por especie. */
export function cuadrar(porEspecie: Record<string, number>): Cuadratura {
  let pallets = 0, cont = 0, manz = 0;
  for (const [e, v] of Object.entries(porEspecie)) {
    const n = +v || 0; if (n <= 0) continue;
    pallets += n; cont += n / capacidad(e); if (capacidad(e) === 21) manz += n;
  }
  const cap = manz * 2 >= pallets ? 21 : 20;            // capacidad de referencia: la que más pesa
  const resto = cont - Math.floor(cont + 1e-9);
  const faltan = resto > 1e-9 ? Math.round((1 - resto) * cap) : 0;
  const sobran = resto > 1e-9 ? Math.round(resto * cap) : 0;
  const completos = faltan === 0 || sobran === 0;       // queda a menos de medio pallet de un entero
  return { pallets, cont, necesarios: completos ? Math.round(cont) : Math.ceil(cont - 1e-9), completos, faltan: completos ? 0 : faltan, sobran: completos ? 0 : sobran, cap };
}

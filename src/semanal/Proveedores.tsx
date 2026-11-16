// ── Pedidos semanales · vista Proveedores (v58.0) ────────────────────────────
// Consulta de lo que se compra a cada proveedor, por oficina o por especie. Reúne lo
// que antes estaba en "Proveedor × Oficina", "Proveedor × Especie" y "Por proveedor".
import React, { useMemo, useState } from "react";
import { Icon } from "../ui";
import { type Asig, type Mercado, type Pedido, fmt0 } from "./datos";

type Modo = "ofi" | "esp";
const sumar = (xs: number[]) => xs.reduce((s, x) => s + x, 0);
const MODOS: { id: Modo; label: string }[] = [{ id: "ofi", label: "Por oficina" }, { id: "esp", label: "Por especie" }];

export default function Proveedores({ asig, pedido, mks, ordenar }: {
  asig: Asig[]; pedido: Pedido; mks: Mercado[]; ordenar: (a: string[]) => string[];
}) {
  const [modo, setModo] = useState<Modo>("ofi");
  const [abiertos, setAbiertos] = useState<Record<string, boolean>>({});

  const D = useMemo(() => {
    const A = asig.filter((a) => +a.cajas > 0);
    const val = (f: (a: Asig) => boolean) => sumar(A.filter(f).map((a) => +a.cajas));
    const pedDe = (m: string) => sumar(Object.keys(pedido).map((e) => +pedido[e]?.[m] || 0));
    const orden = mks.map((m) => m.code);
    const ofis = [...new Set([...orden, ...A.map((a) => a.mercado)])].filter((m) => pedDe(m) > 0 || val((a) => a.mercado === m) > 0);
    const provs = [...new Set(A.map((a) => a.proveedor))]
      .map((n) => ({ n, t: val((a) => a.proveedor === n) })).sort((a, b) => b.t - a.t || a.n.localeCompare(b.n));
    return { A, val, pedDe, ofis, provs, asgTot: sumar(A.map((a) => +a.cajas)) };
  }, [asig, pedido, mks]);
  const { A, val, pedDe, ofis, provs, asgTot } = D;
  const pedTot = sumar(Object.values(pedido).flatMap((r) => Object.values(r || {}).map((v) => +v || 0)));
  const nombreOfi = (c: string) => mks.find((m) => m.code === c)?.nombre || c;

  if (!provs.length) return (
    <section className="rounded-2xl border border-dashed border-line-strong bg-surface px-6 py-12 text-center">
      <p className="text-base font-semibold text-fg">Sin proveedores asignados esta semana.</p>
      <p className="mt-1 text-sm text-muted">Asígnalos en Pedido: abre una especie y usa “+ Agregar proveedor”.</p>
    </section>
  );

  const cols = modo === "ofi" ? ofis.map((m) => ({ k: m, label: m, title: nombreOfi(m) })) : ordenar(A.map((a) => a.especie)).map((e) => ({ k: e, label: e, title: e }));
  const de = (a: Asig) => (modo === "ofi" ? a.mercado : a.especie);
  const td = "px-2 py-1.5 text-center tabular-nums";
  const pie = (label: string, vals: number[], total: number, ref?: number[], primera = false) => (
    <tr className={`text-muted ${primera ? "border-t border-line-strong" : "border-t border-line"}`}>
      <td className="sticky left-0 z-[1] bg-surface px-4 py-1.5 text-left font-medium">{label}</td>
      {vals.map((v, j) => <td key={j} className={`${td} ${ref && v < ref[j] ? "font-bold text-warning" : ref && v > ref[j] ? "font-bold text-danger" : ""}`}>{v ? fmt0(v) : ""}</td>)}
      <td className={`${td} font-bold ${ref && total < sumar(ref) ? "text-warning" : "text-fg"}`}>{fmt0(total)}</td><td />
    </tr>
  );
  const pedCols = cols.map((c) => (modo === "ofi" ? pedDe(c.k) : sumar(Object.values(pedido[c.k] || {}).map((v) => +v || 0))));

  return (
    <>
      <div className="flex flex-wrap items-center gap-3">
        <div className="inline-flex overflow-hidden rounded-lg border border-line-strong bg-surface" role="group" aria-label="Ver proveedores">
          {MODOS.map((m) => (
            <button key={m.id} type="button" onClick={() => setModo(m.id)} aria-pressed={modo === m.id}
              className={`h-8 px-3 text-xs font-semibold transition-colors ${modo === m.id ? "bg-primary text-primary-fg" : "text-fg-2 hover:bg-surface-2"}`}>{m.label}</button>
          ))}
        </div>
        <span className="text-xs text-muted">
          {provs.length} {provs.length === 1 ? "proveedor" : "proveedores"} con pedido · solo consulta{modo === "ofi" ? " · clic en un proveedor para ver sus especies" : ""}
        </span>
      </div>
      <div className="overflow-x-auto rounded-2xl border border-line bg-surface shadow-xs">
        <table className="w-full text-sm" style={{ minWidth: 320 + cols.length * (modo === "ofi" ? 64 : 92) }}>
          <thead>
            <tr className="bg-thead text-[11px] font-semibold tracking-wide text-thead-fg uppercase">
              <th scope="col" className="sticky left-0 z-10 w-52 bg-thead px-4 py-2 text-left">Proveedor</th>
              {cols.map((c) => <th key={c.k} scope="col" title={c.title} className="px-2 py-2 text-center">{c.label}</th>)}
              <th scope="col" className="w-20 px-2 py-2 text-center">Total</th>
              <th scope="col" className="w-16 px-2 py-2 pr-4 text-right">%</th>
            </tr>
          </thead>
          <tbody>
            {provs.map((p) => {
              const abierto = modo === "ofi" && !!abiertos[p.n];
              const bg = abierto ? "bg-surface-2" : "bg-surface";
              return (
                <React.Fragment key={p.n}>
                  <tr className={`border-t border-line ${bg}`}>
                    <td className={`sticky left-0 z-[1] px-2 py-1.5 text-left ${bg}`}>
                      {modo === "ofi" ? (
                        <button type="button" onClick={() => setAbiertos((a) => ({ ...a, [p.n]: !a[p.n] }))} aria-expanded={abierto}
                          className="inline-flex max-w-full items-center gap-1.5 rounded-md px-1 py-0.5 font-semibold text-fg hover:text-primary">
                          <Icon name={abierto ? "chevronDown" : "chevronRight"} className="h-3.5 w-3.5 shrink-0 text-subtle" strokeWidth={2.5} />
                          <span className="truncate">{p.n}</span>
                        </button>
                      ) : <span className="block truncate px-2 font-semibold text-fg">{p.n}</span>}
                    </td>
                    {cols.map((c) => { const v = val((a) => a.proveedor === p.n && de(a) === c.k); return <td key={c.k} className={`${td} text-fg-2`}>{v ? fmt0(v) : ""}</td>; })}
                    <td className={`${td} font-bold text-fg`}>{fmt0(p.t)}</td>
                    <td className="px-2 py-1.5 pr-4 text-right text-xs text-muted tabular-nums">{asgTot ? Math.round((p.t / asgTot) * 100) : 0}%</td>
                  </tr>
                  {abierto && ordenar(A.filter((a) => a.proveedor === p.n).map((a) => a.especie)).map((e) => (
                    <tr key={e} className="border-t border-line bg-surface text-[13px] text-muted">
                      <td className="sticky left-0 z-[1] bg-surface py-1 pr-2 pl-9 text-left"><span className="block truncate">{e}</span></td>
                      {cols.map((c) => { const v = val((a) => a.proveedor === p.n && a.especie === e && a.mercado === c.k); return <td key={c.k} className="px-2 py-1 text-center tabular-nums">{v ? fmt0(v) : ""}</td>; })}
                      <td className="px-2 py-1 text-center tabular-nums">{fmt0(val((a) => a.proveedor === p.n && a.especie === e))}</td><td />
                    </tr>
                  ))}
                </React.Fragment>
              );
            })}
          </tbody>
          <tfoot>
            {pie("Pedido", pedCols, pedTot, undefined, true)}
            {pie("Asignado", cols.map((c) => val((a) => de(a) === c.k)), asgTot, pedCols)}
          </tfoot>
        </table>
      </div>
    </>
  );
}

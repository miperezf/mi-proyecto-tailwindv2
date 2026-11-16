// ── Menú lateral de secciones (v56.2) ────────────────────────────────────────
// Botón ☰ en la barra superior → menú lateral que se despliega desde la izquierda,
// bajo la barra (la barra queda visible y el botón pasa a ✕). Se cierra al elegir
// una sección, con ✕, con Esc o tocando fuera. Igual en escritorio y móvil.
// Para sumar una sección: agregarla a SECCIONES y renderizarla en App.tsx.
import { useEffect, useRef, useState } from "react";
import ReactDOM from "react-dom";
import { Icon } from "./ui";

export type SeccionId = "pedidos" | "semanal";
export type Seccion = { id: SeccionId; label: string; icon: string; desc: string };

export const SECCIONES: Seccion[] = [
  { id: "pedidos", label: "Pedidos por Mail ID", icon: "newMail", desc: "Armar y enviar pedidos por correo" },
  { id: "semanal", label: "Pedidos semanales", icon: "table", desc: "Pedido por mercado y asignación a proveedores" },
];

/** Sección según la dirección (#/semanal). Sin hash = Pedidos por Mail ID. */
export const seccionDesdeHash = (): SeccionId => (window.location.hash.replace(/^#\/?/, "") === "semanal" ? "semanal" : "pedidos");
export const hashDeSeccion = (id: SeccionId) => (id === "pedidos" ? "#/" : `#/${id}`);

export function MenuSecciones({ open, onClose, activa, onChange, anclaRef }: {
  open: boolean; onClose: () => void; activa: SeccionId; onChange: (id: SeccionId) => void;
  /** Barra superior: el menú se abre justo debajo de ella. */
  anclaRef: React.RefObject<HTMLElement | null>;
}) {
  const [top, setTop] = useState(56);
  const navRef = useRef<HTMLElement>(null);
  const onCloseRef = useRef(onClose); onCloseRef.current = onClose;

  useEffect(() => {
    if (!open) return;
    const ubicar = () => setTop(Math.max(0, Math.round(anclaRef.current?.getBoundingClientRect().bottom ?? 56)));
    ubicar();
    const previo = document.activeElement as HTMLElement | null;
    const t = setTimeout(() => navRef.current?.querySelector<HTMLElement>("a[aria-current='page'], a")?.focus(), 0);
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") { e.stopPropagation(); onCloseRef.current(); } };
    window.addEventListener("scroll", ubicar, true); window.addEventListener("resize", ubicar);
    document.addEventListener("keydown", onKey, true);
    return () => {
      clearTimeout(t);
      window.removeEventListener("scroll", ubicar, true); window.removeEventListener("resize", ubicar);
      document.removeEventListener("keydown", onKey, true);
      previo?.focus?.();
    };
  }, [open, anclaRef]);

  if (!open) return null;
  return ReactDOM.createPortal(
    <>
      <div className="fixed inset-x-0 bottom-0 z-40 bg-overlay/40" style={{ top }} onClick={onClose} aria-hidden="true" />
      <nav ref={navRef} aria-label="Secciones"
        className="fixed bottom-0 left-0 z-40 flex w-72 max-w-[85vw] flex-col gap-1 overflow-y-auto border-r border-line bg-surface p-2 shadow-2xl"
        style={{ top }}>
        {SECCIONES.map((s) => (
          <a
            key={s.id}
            href={hashDeSeccion(s.id)}
            onClick={(e) => { e.preventDefault(); onChange(s.id); onClose(); }}
            aria-current={activa === s.id ? "page" : undefined}
            className={`flex items-center gap-3 rounded-lg px-3 py-2.5 transition-colors ${activa === s.id ? "bg-primary-soft text-primary" : "text-fg-2 hover:bg-surface-3"}`}
          >
            <Icon name={s.icon} className="h-5 w-5 shrink-0" />
            <span className="min-w-0">
              <span className="block text-sm font-semibold">{s.label}</span>
              <span className="block text-xs text-muted">{s.desc}</span>
            </span>
          </a>
        ))}
      </nav>
    </>,
    document.body
  );
}

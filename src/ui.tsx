// ── Componentes de interfaz reutilizables (v52) ──────────────────────────────
// Sin dependencias externas. Colores vía tokens de src/index.css.
import React, { useCallback, useEffect, useRef, useState } from "react";
import ReactDOM from "react-dom";

// ── Íconos (Heroicons outline 24px, MIT) ─────────────────────────────────────
const ICONS: Record<string, string[]> = {
  plus: ["M12 4.5v15m7.5-7.5h-15"],
  trash: ["m14.74 9-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 0 1-2.244 2.077H8.084a2.25 2.25 0 0 1-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 0 0-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 0 1 3.478-.397m7.5 0v-.916c0-1.18-.91-2.164-2.09-2.201a51.964 51.964 0 0 0-3.32 0c-1.18.037-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 0 0-7.5 0"],
  note: ["M7.5 8.25h9m-9 3H12m-9.75 1.51c0 1.6 1.123 2.994 2.707 3.227 1.129.166 2.27.293 3.423.379.35.026.67.21.865.501L12 21l2.755-4.133a1.14 1.14 0 0 1 .865-.501 48.172 48.172 0 0 0 3.423-.379c1.584-.233 2.707-1.626 2.707-3.228V6.741c0-1.602-1.123-2.995-2.707-3.228A48.394 48.394 0 0 0 12 3c-2.392 0-4.744.175-7.043.513C3.373 3.746 2.25 5.14 2.25 6.741v6.018Z"],
  duplicate: ["M15.75 17.25v3.375c0 .621-.504 1.125-1.125 1.125h-9.75a1.125 1.125 0 0 1-1.125-1.125V7.875c0-.621.504-1.125 1.125-1.125H6.75a9.06 9.06 0 0 1 1.5.124m7.5 10.376h3.375c.621 0 1.125-.504 1.125-1.125V11.25c0-4.46-3.243-8.161-7.5-8.876a9.06 9.06 0 0 0-1.5-.124H9.375c-.621 0-1.125.504-1.125 1.125v3.5m7.5 10.375H9.375a1.125 1.125 0 0 1-1.125-1.125v-9.25m12 6.625v-1.875a3.375 3.375 0 0 0-3.375-3.375h-1.5a1.125 1.125 0 0 1-1.125-1.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H9.75"],
  x: ["M6 18 18 6M6 6l12 12"],
  ban: ["M18.364 18.364A9 9 0 0 0 5.636 5.636m12.728 12.728A9 9 0 0 1 5.636 5.636m12.728 12.728L5.636 5.636"],
  undo: ["M9 15 3 9m0 0 6-6M3 9h12a6 6 0 0 1 0 12h-3"],
  clock: ["M12 6v6h4.5m4.5 0a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z"],
  cog: [
    "M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z",
    "M15 12a3 3 0 11-6 0 3 3 0 016 0z",
  ],
  newMail: ["M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m3.75 9v6m3-3H9m1.5-12H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"],
  search: ["m21 21-5.197-5.197m0 0A7.5 7.5 0 1 0 5.196 5.196a7.5 7.5 0 0 0 10.607 10.607Z"],
  eye: [
    "M2.036 12.322a1.012 1.012 0 0 1 0-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178Z",
    "M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z",
  ],
  send: ["M6 12 3.269 3.125A59.769 59.769 0 0 1 21.485 12 59.768 59.768 0 0 1 3.27 20.875L5.999 12Zm0 0h7.5"],
  copy: ["M15.666 3.888A2.25 2.25 0 0 0 13.5 2.25h-3c-1.03 0-1.9.693-2.166 1.638m7.332 0c.055.194.084.4.084.612v0a.75.75 0 0 1-.75.75H9a.75.75 0 0 1-.75-.75v0c0-.212.03-.418.084-.612m7.332 0c.646.049 1.288.11 1.927.184 1.1.128 1.907 1.077 1.907 2.185V19.5a2.25 2.25 0 0 1-2.25 2.25H6.75A2.25 2.25 0 0 1 4.5 19.5V6.257c0-1.108.806-2.057 1.907-2.185a48.208 48.208 0 0 1 1.927-.184"],
  check: ["m4.5 12.75 6 6 9-13.5"],
  anchor: ["M12 3v18m0 0c-4.5 0-7.5-3-7.5-6.75M12 21c4.5 0 7.5-3 7.5-6.75M8.25 7.5h7.5M12 6a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3Z"],
  checkCircle: ["M9 12.75 11.25 15 15 9.75M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z"],
  warning: ["M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z"],
  info: ["m11.25 11.25.041-.02a.75.75 0 0 1 1.063.852l-.708 2.836a.75.75 0 0 0 1.063.853l.041-.021M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9-3.75h.008v.008H12V8.25Z"],
  chevronUp: ["m4.5 15.75 7.5-7.5 7.5 7.5"],
  chevronDown: ["m19.5 8.25-7.5 7.5-7.5-7.5"],
  chevronLeft: ["M15.75 19.5 8.25 12l7.5-7.5"],
  chevronRight: ["m8.25 4.5 7.5 7.5-7.5 7.5"],
  user: ["M17.982 18.725A7.488 7.488 0 0 0 12 15.75a7.488 7.488 0 0 0-5.982 2.975m11.963 0a9 9 0 1 0-11.963 0m11.963 0A8.966 8.966 0 0 1 12 21a8.966 8.966 0 0 1-5.982-2.275M15 9.75a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z"],
  refresh: ["M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0 3.181 3.183a8.25 8.25 0 0 0 13.803-3.7M4.031 9.865a8.25 8.25 0 0 1 13.803-3.7l3.181 3.182m0-4.991v4.99"],
  sun: ["M12 3v2.25m6.364.386-1.591 1.591M21 12h-2.25m-.386 6.364-1.591-1.591M12 18.75V21m-4.773-4.227-1.591 1.591M5.25 12H3m4.227-4.773L5.636 5.636M15.75 12a3.75 3.75 0 1 1-7.5 0 3.75 3.75 0 0 1 7.5 0Z"],
  moon: ["M21.752 15.002A9.72 9.72 0 0 1 18 15.75c-5.385 0-9.75-4.365-9.75-9.75 0-1.33.266-2.597.748-3.752A9.753 9.753 0 0 0 3 11.25C3 16.635 7.365 21 12.75 21a9.753 9.753 0 0 0 9.002-5.998Z"],
  users: ["M15 19.128a9.38 9.38 0 0 0 2.625.372 9.337 9.337 0 0 0 4.121-.952 4.125 4.125 0 0 0-7.533-2.493M15 19.128v-.003c0-1.113-.285-2.16-.786-3.07M15 19.128v.106A12.318 12.318 0 0 1 8.624 21c-2.331 0-4.512-.645-6.374-1.766l-.001-.109a6.375 6.375 0 0 1 11.964-3.07M12 6.375a3.375 3.375 0 1 1-6.75 0 3.375 3.375 0 0 1 6.75 0Zm8.25 2.25a2.625 2.625 0 1 1-5.25 0 2.625 2.625 0 0 1 5.25 0Z"],
  bulb: ["M12 18v-5.25m0 0a6.01 6.01 0 0 0 1.5-.189m-1.5.189a6.01 6.01 0 0 1-1.5-.189m3.75 7.478a12.06 12.06 0 0 1-4.5 0m3.75 2.383a14.406 14.406 0 0 1-3 0M14.25 18v-.192c0-.983.658-1.823 1.508-2.316a7.5 7.5 0 1 0-7.517 0c.85.493 1.509 1.333 1.509 2.316V18"],
  sparkles: ["M9.813 15.904 9 18.75l-.813-2.846a4.5 4.5 0 0 0-3.09-3.09L2.25 12l2.846-.813a4.5 4.5 0 0 0 3.09-3.09L9 5.25l.813 2.846a4.5 4.5 0 0 0 3.09 3.09L15.75 12l-2.846.813a4.5 4.5 0 0 0-3.09 3.09Z"],
  lock: ["M16.5 10.5V6.75a4.5 4.5 0 1 0-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 0 0 2.25-2.25v-6.75a2.25 2.25 0 0 0-2.25-2.25H6.75a2.25 2.25 0 0 0-2.25 2.25v6.75a2.25 2.25 0 0 0 2.25 2.25Z"],
  key: ["M15.75 5.25a3 3 0 0 1 3 3m3 0a6 6 0 0 1-7.029 5.912c-.563-.097-1.159.026-1.563.43L10.5 17.25H8.25v2.25H6v2.25H2.25v-2.818c0-.597.237-1.17.659-1.591l6.499-6.499c.404-.404.527-1 .43-1.563A6 6 0 1 1 21.75 8.25Z"],
  eyeOff: ["M3.98 8.223A10.477 10.477 0 0 0 1.934 12C3.226 16.338 7.244 19.5 12 19.5c.993 0 1.953-.138 2.863-.395M6.228 6.228A10.451 10.451 0 0 1 12 4.5c4.756 0 8.773 3.162 10.065 7.498a10.522 10.522 0 0 1-4.293 5.774M6.228 6.228 3 3m3.228 3.228 3.65 3.65m7.894 7.894L21 21m-3.228-3.228-3.65-3.65m0 0a3 3 0 1 0-4.243-4.243m4.242 4.242L9.88 9.88"],
  logout: ["M15.75 9V5.25A2.25 2.25 0 0 0 13.5 3h-6a2.25 2.25 0 0 0-2.25 2.25v13.5A2.25 2.25 0 0 0 7.5 21h6a2.25 2.25 0 0 0 2.25-2.25V15m3 0 3-3m0 0-3-3m3 3H9"],
  building: ["M3.75 21h16.5M4.5 3h15M5.25 3v18m13.5-18v18M9 6.75h1.5m-1.5 3h1.5m-1.5 3h1.5m3-6H15m-1.5 3H15m-1.5 3H15M9 21v-3.375c0-.621.504-1.125 1.125-1.125h3.75c.621 0 1.125.504 1.125 1.125V21"],
  grip: ["M8.25 6.75h.008v.008H8.25V6.75Zm0 5.25h.008v.008H8.25V12Zm0 5.25h.008v.008H8.25v-.008Zm7.5-10.5h.008v.008h-.008V6.75Zm0 5.25h.008v.008h-.008V12Zm0 5.25h.008v.008h-.008v-.008Z"],
};

export function Icon({ name, className = "h-5 w-5", strokeWidth = 1.75 }: { name: string; className?: string; strokeWidth?: number }) {
  const paths = ICONS[name] || [];
  return (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"
      strokeWidth={name === "grip" ? 3 : strokeWidth} className={className} aria-hidden="true">
      {paths.map((d, i) => <path key={i} strokeLinecap="round" strokeLinejoin="round" d={d} />)}
    </svg>
  );
}

export function Spinner({ className = "h-4 w-4" }: { className?: string }) {
  return (
    <svg className={`animate-spin ${className}`} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z" />
    </svg>
  );
}

// ── Botones ──────────────────────────────────────────────────────────────────
const VARIANTS: Record<string, string> = {
  primary: "bg-primary text-primary-fg hover:bg-primary-hover shadow-xs",
  secondary: "bg-surface text-fg-2 border border-line-strong hover:bg-surface-2 shadow-xs",
  ghost: "text-fg-2 hover:bg-surface-3",
  danger: "bg-danger text-white hover:bg-danger-hover shadow-xs dark:text-bg",
  "danger-ghost": "text-danger hover:bg-danger-soft",
};
const SIZES: Record<string, string> = {
  sm: "h-8 px-3 text-xs gap-1.5",
  md: "h-10 px-4 text-sm gap-2",
};

type ButtonProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: string; size?: string; icon?: string; loading?: boolean;
};
export const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(function Button(
  { variant = "secondary", size = "md", icon, loading = false, className = "", children, disabled, type = "button", ...rest }, ref
) {
  return (
    <button ref={ref} type={type} disabled={disabled || loading} aria-busy={loading || undefined}
      className={`inline-flex shrink-0 items-center justify-center rounded-lg font-semibold whitespace-nowrap transition-colors disabled:cursor-not-allowed disabled:opacity-50 ${VARIANTS[variant]} ${SIZES[size]} ${className}`}
      {...rest}>
      {loading ? <Spinner /> : icon ? <Icon name={icon} className={size === "sm" ? "h-4 w-4" : "h-[18px] w-[18px]"} /> : null}
      {children}
    </button>
  );
});

type IconButtonProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  label: string; icon: string; tone?: string; badge?: number | string | null; size?: string;
};
export function IconButton({ label, icon, tone = "default", badge = null, size = "md", className = "", type = "button", ...rest }: IconButtonProps) {
  const tones: Record<string, string> = {
    default: "text-muted hover:text-fg hover:bg-surface-3",
    primary: "text-primary hover:bg-primary-soft",
    danger: "text-muted hover:text-danger hover:bg-danger-soft",
    // v52.1: colores semánticos fijos para las acciones de fila
    success: "text-success hover:bg-success-soft",
    warning: "text-warning hover:bg-warning-soft",
    dangerSolid: "text-danger hover:bg-danger-soft",
  };
  const box = size === "xs" ? "h-7 w-7" : size === "sm" ? "h-8 w-8" : "h-9 w-9";
  return (
    <button type={type} aria-label={label} title={label}
      className={`relative inline-flex ${box} shrink-0 items-center justify-center rounded-lg transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${tones[tone]} ${className}`}
      {...rest}>
      <Icon name={icon} className={size === "xs" ? "h-4 w-4" : size === "sm" ? "h-[18px] w-[18px]" : "h-5 w-5"} strokeWidth={icon === "plus" ? 2.5 : 1.75} />
      {badge !== null && badge !== undefined && badge !== 0 && (
        <span className="absolute -top-0.5 -right-0.5 flex h-4 min-w-4 items-center justify-center rounded-full bg-primary px-1 text-[10px] font-bold text-primary-fg">
          {badge}
        </span>
      )}
    </button>
  );
}

export function Badge({ tone = "neutral", children, className = "" }: { tone?: string; children: React.ReactNode; className?: string }) {
  const tones: Record<string, string> = {
    neutral: "bg-surface-3 text-fg-2",
    primary: "bg-primary-soft text-primary",
    success: "bg-success-soft text-success",
    warning: "bg-warning-soft text-warning",
    danger: "bg-danger-soft text-danger",
  };
  return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-semibold ${tones[tone]} ${className}`}>{children}</span>;
}

// ── Modal / panel lateral accesible ──────────────────────────────────────────
// Esc cierra, foco al primer elemento, foco atrapado dentro, vuelve al disparador.
const FOCUSABLE = 'a[href],button:not([disabled]),textarea:not([disabled]),input:not([disabled]),select:not([disabled]),iframe,[tabindex]:not([tabindex="-1"])';

type ModalProps = {
  open: boolean; onClose: () => void; title: React.ReactNode; children: React.ReactNode;
  footer?: React.ReactNode; size?: string; drawer?: boolean; dismissable?: boolean; icon?: string; initialFocus?: React.RefObject<HTMLElement | null>;
};
export function Modal({ open, onClose, title, children, footer, size = "md", drawer = false, dismissable = true, icon, initialFocus }: ModalProps) {
  const panelRef = useRef<HTMLDivElement>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  const titleId = useRef(`dlg-${Math.random().toString(36).slice(2, 8)}`).current;

  useEffect(() => {
    if (!open) return;
    const previous = document.activeElement as HTMLElement | null;
    const t = setTimeout(() => {
      const target = initialFocus?.current || panelRef.current?.querySelector<HTMLElement>(FOCUSABLE);
      target?.focus();
    }, 0);
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && dismissable) { e.stopPropagation(); onCloseRef.current(); }
      if (e.key === "Tab" && panelRef.current) {
        const items = Array.from(panelRef.current.querySelectorAll<HTMLElement>(FOCUSABLE));
        if (items.length === 0) return;
        const first = items[0], last = items[items.length - 1];
        if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
        else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      }
    };
    document.addEventListener("keydown", onKey, true);
    return () => { clearTimeout(t); document.removeEventListener("keydown", onKey, true); previous?.focus?.(); };
  }, [open, dismissable, initialFocus]);

  if (!open) return null;
  const widths: Record<string, string> = { sm: "max-w-sm", md: "max-w-lg", lg: "max-w-3xl", xl: "max-w-5xl" };
  const panelClass = drawer
    ? "fixed inset-y-0 right-0 flex w-full max-w-sm flex-col border-l border-line bg-surface shadow-2xl"
    : `relative flex max-h-[92vh] w-full ${widths[size]} flex-col rounded-2xl border border-line bg-surface shadow-2xl`;

  return ReactDOM.createPortal(
    <div className={`fixed inset-0 z-50 ${drawer ? "" : "flex items-center justify-center p-3 sm:p-6"}`}>
      <div className="absolute inset-0 bg-overlay" onClick={() => dismissable && onClose()} aria-hidden="true" />
      <div ref={panelRef} role="dialog" aria-modal="true" aria-labelledby={titleId} className={panelClass}>
        <div className="flex items-center gap-3 border-b border-line px-5 py-4">
          {icon && <span className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary-soft text-primary"><Icon name={icon} className="h-[18px] w-[18px]" /></span>}
          <h2 id={titleId} className="min-w-0 flex-1 truncate text-base font-semibold text-fg">{title}</h2>
          {dismissable && <IconButton label="Cerrar" icon="x" size="sm" onClick={onClose} />}
        </div>
        <div className="min-h-0 flex-1 overflow-y-auto px-5 py-4">{children}</div>
        {footer && <div className="flex flex-wrap items-center justify-end gap-2 border-t border-line bg-surface-2 px-5 py-3 rounded-b-2xl">{footer}</div>}
      </div>
    </div>,
    document.body
  );
}

// ── Confirmación (reemplaza window.confirm) ──────────────────────────────────
type ConfirmOptions = { title: string; message?: React.ReactNode; confirmLabel?: string; cancelLabel?: string; danger?: boolean };
export function useConfirm(): [React.ReactNode, (o: ConfirmOptions) => Promise<boolean>] {
  const [state, setState] = useState<(ConfirmOptions & { resolve: (v: boolean) => void }) | null>(null);
  const cancelRef = useRef<HTMLButtonElement>(null);
  const ask = useCallback((o: ConfirmOptions) => new Promise<boolean>((resolve) => setState({ ...o, resolve })), []);
  const close = (v: boolean) => { state?.resolve(v); setState(null); };
  const el = (
    <Modal open={!!state} onClose={() => close(false)} title={state?.title} size="sm"
      icon={state?.danger ? "warning" : "info"} initialFocus={cancelRef}
      footer={<>
        <Button ref={cancelRef} variant="secondary" onClick={() => close(false)}>{state?.cancelLabel || "Cancelar"}</Button>
        <Button variant={state?.danger ? "danger" : "primary"} onClick={() => close(true)}>{state?.confirmLabel || "Confirmar"}</Button>
      </>}>
      <div className="text-sm leading-relaxed text-fg-2">{state?.message}</div>
    </Modal>
  );
  return [el, ask];
}

// ── Avisos (toasts) ──────────────────────────────────────────────────────────
type Toast = { id: number; tone: string; message: React.ReactNode };
export function useToasts(): [React.ReactNode, (tone: string, message: React.ReactNode, ms?: number) => void] {
  const [toasts, setToasts] = useState<Toast[]>([]);
  const idRef = useRef(0);
  const notify = useCallback((tone: string, message: React.ReactNode, ms = 4500) => {
    const id = ++idRef.current;
    setToasts((t) => [...t.slice(-3), { id, tone, message }]);
    if (ms > 0) setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), ms);
  }, []);
  const icons: Record<string, string> = { success: "checkCircle", error: "warning", info: "info" };
  const tones: Record<string, string> = {
    success: "border-success/30 text-success",
    error: "border-danger/40 text-danger",
    info: "border-primary-line text-primary",
  };
  const el = ReactDOM.createPortal(
    <div className="pointer-events-none fixed inset-x-0 bottom-20 z-[60] flex flex-col items-center gap-2 px-4 sm:bottom-6 sm:items-end" aria-live="polite" role="status">
      {toasts.map((t) => (
        <div key={t.id} className={`pointer-events-auto flex w-full max-w-sm items-start gap-3 rounded-xl border bg-surface px-4 py-3 shadow-lg ${tones[t.tone] || tones.info}`}>
          <Icon name={icons[t.tone] || "info"} className="mt-0.5 h-5 w-5 shrink-0" />
          <div className="min-w-0 flex-1 text-sm text-fg">{t.message}</div>
          <button type="button" aria-label="Cerrar aviso" className="text-subtle hover:text-fg" onClick={() => setToasts((x) => x.filter((y) => y.id !== t.id))}>
            <Icon name="x" className="h-4 w-4" />
          </button>
        </div>
      ))}
    </div>,
    document.body
  );
  return [el, notify];
}

// ── Pedidos semanales: almacenamiento en Firestore (v56.0) ───────────────────
// Rutas (mismo prefijo que los pedidos):
//   artifacts/{appId}/public/data/config/catalogo                 → Catalogo
//   artifacts/{appId}/public/data/semanas/{AAAA-Www}              → { semana, pedido, aereo, actualizado }
//        (pedido = marítimo, aereo = aéreo; v59.0)
//   artifacts/{appId}/public/data/semanas/{AAAA-Www}/asignaciones/{id}
//        → { proveedor, especie, mercado, cajas, creado, via? }   (via: "aereo"; sin via = marítimo)
// A diferencia de los pedidos por Mail ID, aquí cada cambio se guarda al momento.
import {
  type Firestore, collection, doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc,
  onSnapshot, writeBatch, deleteField, FieldPath,
} from "firebase/firestore";
import { type Asig, type AsigData, type Catalogo, type Semana, type SemanalStore, campoPedido } from "./datos";

export function firestoreStore(db: Firestore, appId: string): SemanalStore {
  const base = `artifacts/${appId}/public/data`;
  const semanaRef = (s: string) => doc(db, `${base}/semanas`, s);
  const asigCol = (s: string) => collection(db, `${base}/semanas/${s}/asignaciones`);
  const catRef = doc(db, `${base}/config`, "catalogo");
  const toAsig = (d: { id: string; data: () => any }): Asig => ({ id: d.id, ...(d.data() as AsigData) });

  return {
    subCatalogo: (cb, err) => onSnapshot(catRef, (s) => cb(s.exists() ? (s.data() as Catalogo) : null), err),
    subSemanas: (cb, err) => onSnapshot(collection(db, `${base}/semanas`), (qs) => {
      const out: Record<string, Semana> = {};
      qs.forEach((d) => { out[d.id] = d.data() as Semana; });
      cb(out);
    }, err),
    subAsig: (s, cb, err) => onSnapshot(asigCol(s), (qs) => cb(qs.docs.map(toAsig)), err),
    setCatalogo: (c) => setDoc(catRef, c),
    setPedidoCelda: (s, especie, mercado, cajas, via = "maritimo") =>
      setDoc(semanaRef(s), { semana: s, [campoPedido(via)]: { [especie]: { [mercado]: cajas } }, actualizado: Date.now() }, { merge: true }),
    // reemplaza solo el campo de esa vía (la otra queda igual)
    setPedido: (s, pedido, via = "maritimo") => updateDoc(semanaRef(s), campoPedido(via), pedido, "actualizado", Date.now()),
    borrarEspeciePedido: async (s, especie, via = "maritimo") => {
      const snap = await getDoc(semanaRef(s));
      if (snap.exists()) await updateDoc(semanaRef(s), new FieldPath(campoPedido(via), especie), deleteField(), "actualizado", Date.now());
    },
    tocarSemana: (s) => setDoc(semanaRef(s), { semana: s, actualizado: Date.now() }, { merge: true }),
    borrarSemana: (s) => deleteDoc(semanaRef(s)),
    getSemana: async (s) => { const snap = await getDoc(semanaRef(s)); return snap.exists() ? (snap.data() as Semana) : null; },
    listAsig: async (s) => (await getDocs(asigCol(s))).docs.map(toAsig),
    setAsig: (s, id, data) => setDoc(doc(asigCol(s), id), data),
    patchAsig: (s, id, patch) => updateDoc(doc(asigCol(s), id), patch as Record<string, unknown>),
    borrarAsig: async (s, ids) => {
      for (let i = 0; i < ids.length; i += 400) {
        const b = writeBatch(db);
        ids.slice(i, i + 400).forEach((id) => b.delete(doc(asigCol(s), id)));
        await b.commit();
      }
    },
  };
}

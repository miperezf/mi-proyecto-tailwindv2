--- a/App.tsx	2026-10-19 19:42:06.548763317 +0000
+++ b/App.tsx	2026-10-19 19:42:06.614367597 +0000
@@ -162,6 +162,17 @@
   isCanceled: false,
 };
 
+// ── v51.1: campos de header que el usuario puede editar a mano.
+// Se usan para persistir ediciones en pedidos "sent" no activos al reenviar.
+// No incluye status/createdBy/createdAt/mailId/emailSubject a propósito.
+const EDITABLE_HEADER_FIELDS = [
+  "reDestinatarios",
+  "deNombrePais",
+  "nave",
+  "fechaCarga",
+  "exporta",
+];
+
 const App = () => {
   const appId =
     typeof __app_id !== "undefined" ? __app_id : firebaseConfig.projectId;
@@ -189,6 +200,10 @@
   const incotermWriteQueueRef = useRef(Promise.resolve());
   const incotermSeqRef = useRef(0);
 
+  // ── v51.1: copia de cada header tal como se cargó desde Firestore (por id).
+  // Permite detectar qué campos editó el usuario en pedidos "sent" no activos.
+  const loadedHeadersRef = useRef<Record<string, any>>({});
+
   // ─── LAST SEND RECOVERY ─────────────────────────────────────────────────────
 
   // ─── HISTORY PANEL ──────────────────────────────────────────────────────────
@@ -1096,6 +1111,7 @@
       });
 
       if (orders.length === 0) {
+        loadedHeadersRef.current = {};
         setDisplayedOrders([]);
         setActiveOrderId(null);
         setCurrentOrderIndex(0);
@@ -1103,6 +1119,10 @@
         // ── BUG FIX: keep original status from Firestore — never force "draft"
         // Mutating sent→draft here was causing flushAndSave to persist "draft"
         // back to Firestore and handleNewMail to soft-delete sent orders.
+        // ── v51.1: guardar baseline de headers para detectar ediciones al enviar
+        loadedHeadersRef.current = Object.fromEntries(
+          orders.map((o) => [o.id, { ...o.header }])
+        );
         setDisplayedOrders(orders);
         const last = orders[orders.length - 1];
         setActiveOrderId(last.id);
@@ -2206,7 +2226,19 @@
               incoterm:     latestHeader.incoterm,
             });
           } else {
-            console.log("[FRUTAM:SEND] ✅ Orden SENT (no activa) — persistiendo solo items");
+            // ── v51.1: pedido sent NO activo: además de items, persistir solo los
+            // campos de header que el usuario cambió respecto de lo cargado.
+            const baseline = loadedHeadersRef.current[order.id];
+            if (baseline) {
+              EDITABLE_HEADER_FIELDS.forEach((key) => {
+                const current = order.header?.[key] ?? "";
+                const original = baseline[key] ?? "";
+                if (current !== original) {
+                  updatePayload[`header.${key}`] = current;
+                }
+              });
+            }
+            console.log("[FRUTAM:SEND] ✅ Orden SENT (no activa) — persistiendo items + campos de header modificados");
           }
 
           await updateDoc(orderDocRef, updatePayload);

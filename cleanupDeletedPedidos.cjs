const admin = require("firebase-admin");
const fs = require("fs");

const serviceAccount = require("./pedidosfrutam-160cb-firebase-adminsdk-fbsvc-9ad08f548e.json");

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

async function exportRioBlanco() {

  const snapshot = await db
    .collection("artifacts")
    .doc("pedidosfrutam-160cb")
    .collection("public")
    .doc("data")
    .collection("pedidos")
    .where("header.reDestinatarios", "==", "RIO BLANCO")
    .get();

  console.log("Pedidos RIO BLANCO encontrados:", snapshot.size);

  const pedidos = snapshot.docs.map(doc => ({
    id: doc.id,
    ...doc.data()
  }));

  fs.writeFileSync(
    "pedidos_rio_blanco.json",
    JSON.stringify(pedidos, null, 2)
  );

  console.log("Archivo pedidos_rio_blanco.json generado");
}

exportRioBlanco();
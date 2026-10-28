// ── Oficinas para compras DIRECTAS (Consignee & Notify) ─────────────────────
// Para agregar o corregir una oficina, editar OFICINAS. Campos:
//   pais     → como se escribe en el campo País (sin importar tildes/mayúsculas)
//   alias    → otras formas de escribir el país (ej. "SALVADOR")
//   puerto   → solo si el país tiene más de una oficina (Costa Rica: una por puerto).
//              En ese caso, al marcar Directo se elige el puerto en un menú.
//   nombre   → texto que va en Exporta: "[PROVEEDOR] DIRECTO A [nombre] (PUERTO)"
//   direccion, nota, ids [[rótulo, número]], contacto, correo, celular, telefono
//            → datos de la ficha del correo, en ese orden (los vacíos se omiten)
//   telefonoRotulo → rótulo del teléfono si no es "Teléfono" (ej. "Teléfono / Fax")
//   notify   → solo si el Notify es OTRA empresa (misma estructura). Si no se
//              indica, la ficha se titula "CONSIGNEE & NOTIFY".
// Todo en MAYÚSCULAS salvo los correos.
//
// El estado "Directo" se guarda en el propio campo Exporta (no requiere campos
// nuevos en Firestore): si Exporta termina en " DIRECTO A <etiqueta>", el pedido
// es directo a esa oficina.

export type Ficha = {
  nombre: string;
  direccion?: string;
  nota?: string;
  ids?: [string, string][];
  contacto?: string;
  correo?: string;
  celular?: string;
  telefono?: string;
  telefonoRotulo?: string;
};

export type Oficina = Ficha & {
  id: string;
  pais: string;
  alias?: string[];
  puerto?: string;
  notify?: Ficha;
};

const FRESH_GARDEN: Ficha = {
  nombre: "FRESH GARDEN S.A.",
  direccion:
    "INSTALACIONES DE CENADA, BARREAL DE HEREDIA 625 METROS ESTE DEL CENTRO COMERCIAL REAL CARIARI, GALPON #3/LOCAL 3306 HEREDIA -COSTA RICA.",
  ids: [["Cédula jurídica", "3101555258"]],
  contacto: "SANDRA ALPIZAR",
  correo: "salpizar@vpcom.com",
  telefono: "506-22885689",
};

const ALPHA_CARGO: Ficha = {
  nombre: "ALPHA CARGO",
  direccion: "PANAMA, CIUDAD DE PANAMA, AV. SAMUEL LEWIS, PH SL55, TORRE OPTIMA, PISO 36, OFICINA 36-2",
  ids: [["ID", "32500365524"], ["RUC", "155716129-2-2021"], ["DV", "47"]],
  contacto: "JOSEPH HERNANDEZ",
  celular: "+507 6393-6491",
  telefono: "+507 381-1331",
};

export const OFICINAS: Oficina[] = [
  {
    id: "CO",
    pais: "COLOMBIA",
    nombre: "VPC COLOMBIA S.A.S",
    direccion: "AV. CRA. 80 NO. 2-51 SUR CORABASTOS INT 24 BOGOTA - COLOMBIA.",
    nota: "ENTREGA DIRECTA - CODIGO UTS B00692.",
    ids: [["NIT", "830.045.768-2"]],
    contacto: "ANDREA NEIVA",
    correo: "aneiva@vpcom.com",
    telefono: "5714021844",
  },
  { id: "CR-BALBOA", pais: "COSTA RICA", puerto: "BALBOA", ...FRESH_GARDEN, notify: ALPHA_CARGO },
  { id: "CR-CALDERA", pais: "COSTA RICA", puerto: "CALDERA", ...FRESH_GARDEN },
  { id: "CR-MOIN", pais: "COSTA RICA", puerto: "MOÍN", ...FRESH_GARDEN },
  {
    id: "EC",
    pais: "ECUADOR",
    nombre: "VPCEC S.A.",
    direccion: "KM 6 VIA DURAN BOLICHE DURAN 092411, GUAYAQUIL- ECUADOR.",
    ids: [["RUC", "0990699119001"]],
    contacto: "KENIA BALLADARES",
    correo: "kballadares@vpcom.com",
    telefono: "593-95-9556958",
  },
  {
    id: "SV",
    pais: "EL SALVADOR",
    alias: ["SALVADOR"],
    nombre: "FRUTERIA VIDAURRI S.A. DE C.V.",
    direccion:
      "POLIGONO 2, CANTON JOYA GALANA LOTIFICACION AGRICOLA, LOTE 1, HACIENDA EL ANGEL, APOPA, SAN SALVADOR, EL SALVADOR.",
    contacto: "WILLIAM ARTIGA",
    correo: "wartiga@vpcom.com",
    telefono: "503 7047 9859",
  },
  {
    id: "GT",
    pais: "GUATEMALA",
    nombre: "FRUVIGUA SOCIEDAD ANONIMA",
    direccion: "3A CALLE 9-09 ZONA 1. CENTRO AMÉRICA - GUATEMALA.",
    ids: [["NIT", "80371132"]],
    contacto: "MAYRA SANTOS",
    correo: "msantos@vpcom.com",
    telefono: "(502)22181000",
    telefonoRotulo: "Teléfono / Fax",
  },
  {
    id: "MX",
    pais: "MÉXICO",
    nombre: "VIDIMPORT S.A. DE C.V.",
    direccion:
      "ANDADOR K-L BODEGA K 111 ZONA V SECTOR 2 NAVE 4 CENTRAL DE ABASTO IZTAPALAPA C.P.09040 CIUDAD DE MEXICO - MEXICO",
    ids: [["RFC", "VID910607UM2"]],
    contacto: "ALMA COLIN",
    correo: "acolin@vpcom.com",
    telefono: "525540809715",
  },
  {
    id: "PA",
    pais: "PANAMÁ",
    nombre: "MAFALDA S.A.",
    direccion:
      "PROVINCIA DE PANAMÁ, DISTRITO DEL PANAMÁ, CORREGIMIENTO DE LAS MAÑANITAS, CALLE VIA LAS MAÑANITAS, EDIFICIO EMERGENT COLD - GALORES - PANAMÁ.",
    ids: [["RUC", "1140-361-117549"]],
    contacto: "LILIANA DELGADO",
    correo: "ldelgado@vpcom.com",
    telefono: "507-62040490",
  },
  {
    id: "VE",
    pais: "VENEZUELA",
    nombre: "VIDINTER, C.A.",
    direccion:
      "CARRETERA.LAS MAYAS VIA LA MARIPOSA 105 URBANIZACION TURMERITO, GALPON LAS MAYAS ROSARIGUA, CARACAS, VENEZUELA",
    ids: [["RIF", "J-00111368-1"]],
    contacto: "MARLENE SUBERO",
    correo: "msubero@frutmail.net",
    telefono: "0800-2255345",
  },
];

/** Normaliza para comparar países: sin tildes, mayúsculas, espacios simples. */
export const normalizarTexto = (s: unknown): string =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

/** País sin lo que va entre paréntesis: "COLOMBIA (BUENAVENTURA)" → "COLOMBIA". */
export const paisBase = (pais: unknown): string => normalizarTexto(String(pais ?? "").replace(/\([^)]*\)/g, " "));

/** Oficinas del país (por nombre o alias), ordenadas por puerto. */
export const oficinasDePais = (pais: unknown, lista: Oficina[] = OFICINAS): Oficina[] => {
  const p = paisBase(pais);
  if (!p) return [];
  return lista
    .filter((o) => [o.pais, ...(o.alias || [])].some((x) => normalizarTexto(x) === p))
    .sort((a, b) => normalizarTexto(a.puerto).localeCompare(normalizarTexto(b.puerto)));
};

/** Texto de la oficina en Exporta. Si el país tiene varias, se agrega el puerto. */
export const etiquetaOficina = (o: Oficina, lista: Oficina[] = OFICINAS): string => {
  const variasEnPais = oficinasDePais(o.pais, lista).length > 1;
  return variasEnPais && o.puerto ? `${o.nombre} (${o.puerto.toUpperCase()})` : o.nombre;
};

/** Datos de una ficha como pares [rótulo, valor], en orden y sin los vacíos. */
export const datosFicha = (f: Ficha): [string, string][] =>
  ([
    ["Dirección", f.direccion],
    ["Nota", f.nota],
    ...(f.ids || []),
    ["Contacto", f.contacto],
    ["Correo", f.correo],
    ["Celular", f.celular],
    [f.telefonoRotulo || "Teléfono", f.telefono],
  ] as [string, string | undefined][])
    .filter(([, v]) => String(v ?? "").trim() !== "")
    .map(([k, v]) => [k, String(v).trim()]);

/** Compatibilidad: datos del consignee de la oficina. */
export const datosOficina = (o: Oficina): [string, string][] => datosFicha(o);

/** Fichas que van en el correo: [título, ficha]. Una si Consignee = Notify, dos si no. */
export const fichasOficina = (o: Oficina): [string, Ficha][] =>
  o.notify ? [["CONSIGNEE", o], ["NOTIFY", o.notify]] : [["CONSIGNEE & NOTIFY", o]];

export const SEPARADOR_DIRECTO = " DIRECTO A ";

export const exportaDirecto = (proveedor: unknown, o: Oficina, lista: Oficina[] = OFICINAS): string =>
  `${String(proveedor ?? "").trim().toUpperCase()}${SEPARADOR_DIRECTO}${etiquetaOficina(o, lista)}`;

/** Devuelve la oficina si Exporta indica una compra directa; si no, null. */
export const detectarDirecto = (exporta: unknown, lista: Oficina[] = OFICINAS): Oficina | null => {
  const e = normalizarTexto(exporta);
  if (!e.includes(SEPARADOR_DIRECTO.trim())) return null;
  return lista.find((o) => e.endsWith(normalizarTexto(SEPARADOR_DIRECTO + etiquetaOficina(o, lista)))) || null;
};

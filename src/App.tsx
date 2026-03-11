import React, { useState, useEffect, useRef } from "react";
import ReactDOM from "react-dom";
import { initializeApp } from "firebase/app";
import {
  getAuth,
  signInAnonymously,
  signInWithCustomToken,
  onAuthStateChanged,
} from "firebase/auth";
import {
  getFirestore,
  collection,
  doc,
  getDocs,
  addDoc,
  setDoc,
  updateDoc,
  // deleteDoc, // Removed hard delete
  onSnapshot,
  query,
  where,
  writeBatch,
  deleteDoc,
} from "firebase/firestore";

const firebaseConfig =
  typeof __firebase_config !== "undefined"
    ? JSON.parse(__firebase_config)
    : {
        apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
        authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
        projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
        storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
        messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
        appId: import.meta.env.VITE_FIREBASE_APP_ID,
        measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID,
      };

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const InputField = React.forwardRef(
  (
    {
      label,
      name,
      value,
      onChange,
      onBlur,
      onFocus,
      placeholder = "",
      className = "",
      type = "text",
      list,
      readOnly = false,
      darkMode = false,
    },
    ref
  ) => {
    const safeValue =
      value === null || value === undefined ? "" : String(value);
    return (
      <div className="mb-2 w-full">
        <label
          htmlFor={name}
          className="block text-sm font-medium"
          style={{ color: darkMode ? "#ffffff" : "#374151" }}
        >
          {label}
        </label>
        <input
          type={type}
          id={name}
          name={name}
          value={safeValue}
          onChange={onChange}
          onFocus={onFocus}
          onBlur={onBlur}
          placeholder={placeholder}
          className={`mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 bg-gray-50 border ${className}`}
          ref={ref}
          list={list}
          readOnly={readOnly}
        />
      </div>
    );
  }
);

const TableInput = React.forwardRef(
  (
    {
      name,
      value,
      onChange,
      onBlur,
      onFocus,
      type = "text",
      placeholder = "",
      readOnly = false,
      isCanceledProp = false,
      list,
    },
    ref
  ) => {
    const safeValue =
      value === null || value === undefined ? "" : String(value);
    return (
      <input
        type={type}
        name={name}
        value={safeValue}
        onChange={onChange}
        onFocus={onFocus}
        onBlur={onBlur}
        placeholder={placeholder}
        className={`w-full h-full p-px border-none bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded-md ${
          isCanceledProp ? "line-through" : ""
        }`}
        ref={ref}
        readOnly={readOnly}
        list={list}
      />
    );
  }
);

const initialHeaderState = {
  reDestinatarios: "",
  deNombrePais: "",
  nave: "",
  fechaCarga: "",
  exporta: "",
  emailSubject: "",
  mailId: "",
  incoterm: "FOB",
  status: "draft",
  createdAt: null,
  createdBy: null,
  lastModifiedBy: null,
  updatedAt: null,
};

const initialItemState = {
  id: "",  // always overridden with crypto.randomUUID() at call site
  pallets: "",
  especie: "",
  variedad: "",
  formato: "",
  calibre: "",
  categoria: "",
  preciosFOB: "",
  estado: "",
  isCanceled: false,
};

const App = () => {
  const appId =
    typeof __app_id !== "undefined" ? __app_id : firebaseConfig.projectId;

  const [userId, setUserId] = useState(null);
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [allOrdersFromFirestore, setAllOrdersFromFirestore] = useState([]);
  const [displayedOrders, setDisplayedOrders] = useState([]);

  const [searchTerm, setSearchTerm] = useState("");
  const [committedSearchTerm, setCommittedSearchTerm] = useState("");

  const [activeOrderId, setActiveOrderId] = useState(null);

  // ─── PRESENCE SYSTEM ────────────────────────────────────────────────────────
  const [presenceMailId, setPresenceMailId] = useState(null);
  const [otherUsersPresent, setOtherUsersPresent] = useState([]);
  const [showPresenceModal, setShowPresenceModal] = useState(false);
  const [presenceConflictUsers, setPresenceConflictUsers] = useState([]);
  const [pendingSearchTerm, setPendingSearchTerm] = useState("");
  const presenceUnsubscribeRef = useRef(null);

  // ─── LAST SEND RECOVERY ─────────────────────────────────────────────────────

  // ─── HISTORY PANEL ──────────────────────────────────────────────────────────
  const [showHistoryPanel, setShowHistoryPanel] = useState(false);
  const [showWhatsNew, setShowWhatsNew] = useState(() => {
    return localStorage.getItem("frutam_whatsnew_v47") !== "seen";
  });
  const [historyResendingId, setHistoryResendingId] = useState(null);

  // ─── EMAIL CLIENT CONFIG ─────────────────────────────────────────────────────
  const EMAIL_CLIENT_KEY = "frutam_email_client";
  const [emailClientPref, setEmailClientPref] = useState(
    () => localStorage.getItem(EMAIL_CLIENT_KEY) || "auto"
  );
  const [showConfigPanel, setShowConfigPanel] = useState(false);

  // ─── AUTOCOMPLETADO CONTEXTUAL ──────────────────────────────────────────────
  const [proveedorSuggestion, setProveedorSuggestion] = useState(null);
  const suggestionTimerRef = useRef(null);
  const lookupDebounceRef = useRef(null);
  const suggestionAppliedRef = useRef(""); // stores proveedor name — empty means no suggestion applied yet
  const committedSearchTermRef = useRef("");

  // ─── DARK MODE ───────────────────────────────────────────────────────────────
  const DARK_MODE_KEY = "frutam_dark_mode";
  const [darkMode, setDarkMode] = useState(
    () => localStorage.getItem(DARK_MODE_KEY) === "true"
  );
  const toggleDarkMode = () => {
    setDarkMode((prev) => {
      localStorage.setItem(DARK_MODE_KEY, String(!prev));
      return !prev;
    });
  };

  // ─── PROFILE CRUD ────────────────────────────────────────────────────────────

  // ─── BUSCAR ÚLTIMO PEDIDO DE PROVEEDOR ───────────────────────────────────────
  const lookupLastOrderForProveedor = async (proveedorName) => {
    const TAG = "[PRELOAD]";
    const name = (proveedorName || "").trim();

    const nameLower = name.toLowerCase();

    // ── GUARDS ──────────────────────────────────────────────────────────────────
    if (!db || !name) {
      return;
    }
    if (committedSearchTermRef.current) {
      return;
    }
    if (suggestionAppliedRef.current === nameLower) {
      return;
    }
    const currentItems = orderItemsRef.current;
    const hasData = currentItems.some(it => it.especie || it.variedad || it.formato);
    if (hasData) {
      return;
    }

    // ── FIRESTORE QUERY ─────────────────────────────────────────────────────────
    try {
      const snap = await getDocs(query(
        collection(db, `artifacts/${appId}/public/data/pedidos`),
        where("header.status", "==", "sent")
      ));

      const allSent = parseFirestoreOrders(snap.docs);

      // ── FIND MATCHING ORDERS ─────────────────────────────────────────────────
      const matches = allSent
        .filter(o => (o.header?.reDestinatarios || "").toLowerCase() === nameLower)
        .sort((a, b) => (b.header?.updatedAt || 0) - (a.header?.updatedAt || 0));

      if (matches.length === 0) {
        return;
      }

      // ── FIND MOST RECENT ORDER WITH ACTUAL ITEM DATA ─────────────────────────
      let bestOrder = null;
      let bestItems = [];

      for (const order of matches) {
        const raw = Array.isArray(order.items) ? order.items : [];
        const candidates = raw.filter(it => it.especie || it.variedad);
        if (candidates.length > 0) {
          bestOrder = order;
          bestItems = candidates;
          break;
        }
      }

      if (!bestOrder || bestItems.length === 0) {
        return;
      }

      // ── BUILD SUGGESTION ─────────────────────────────────────────────────────
      // Include ALL item fields — pallets and prices included
      const suggestionItems = bestItems.map(it => ({
        especie:    it.especie    || "",
        variedad:   it.variedad   || "",
        formato:    it.formato    || "",
        calibre:    it.calibre    || "",
        categoria:  it.categoria  || "",
        pallets:    it.pallets    || "",
        preciosFOB: it.preciosFOB || "",
      }));

      const suggestion = {
        proveedor:      name,
        nave:           bestOrder.header?.nave          || "",
        deNombrePais:   bestOrder.header?.deNombrePais  || "",
        exporta:        bestOrder.header?.exporta       || "",
        incoterm:       bestOrder.header?.incoterm      || "FOB",
        items:          suggestionItems,
      };

      setProveedorSuggestion(suggestion);
      if (suggestionTimerRef.current) clearTimeout(suggestionTimerRef.current);
      suggestionTimerRef.current = setTimeout(() => {
        setProveedorSuggestion(null);
      }, 12000);

    } catch (e) {
      console.error(`${TAG} ERROR:`, e);
    }
  };

    const applyProveedorSuggestion = () => {
    if (!proveedorSuggestion) return;
    if (suggestionTimerRef.current) clearTimeout(suggestionTimerRef.current);

    const TAG = "[PRELOAD]";
    // Fill table with all fields — only reset estado and isCanceled
    const newItems = proveedorSuggestion.items.map(it => ({
      id: crypto.randomUUID(),
      especie:    it.especie    || "",
      variedad:   it.variedad   || "",
      formato:    it.formato    || "",
      calibre:    it.calibre    || "",
      categoria:  it.categoria  || "",
      pallets:    it.pallets    || "",
      preciosFOB: it.preciosFOB || "",
      estado:     "",
      isCanceled: false,
    }));
    setOrderItems(newItems);
    orderItemsRef.current = newItems;
    suggestionAppliedRef.current = proveedorSuggestion.proveedor.toLowerCase();
    setProveedorSuggestion(null);
  };

  const dismissProveedorSuggestion = () => {
    if (suggestionTimerRef.current) clearTimeout(suggestionTimerRef.current);
    if (proveedorSuggestion) {
      suggestionAppliedRef.current = proveedorSuggestion.proveedor.toLowerCase();
    }
    setProveedorSuggestion(null);
  };

  const saveEmailClientPref = (pref) => {
    localStorage.setItem(EMAIL_CLIENT_KEY, pref);
    setEmailClientPref(pref);
  };

  const openEmailClient = (subject) => {
    const outlookWebUrl = `https://outlook.office.com/mail/deeplink/compose?subject=${encodeURIComponent(subject)}`;
    const mailtoUrl = `mailto:?subject=${encodeURIComponent(subject)}&body=`;

    if (isMobileDevice()) {
      window.location.href = mailtoUrl;
      return;
    }

    const pref = emailClientPref;

    if (pref === "web") {
      window.open(outlookWebUrl, "_blank");
      return;
    }

    if (pref === "desktop") {
      window.location.href = mailtoUrl;
      return;
    }

    const fallbackTab = window.open("about:blank", "_blank");
    window.location.href = mailtoUrl;

    let desktopOpened = false;
    const onBlur = () => { desktopOpened = true; };
    window.addEventListener("blur", onBlur, { once: true });

    setTimeout(() => {
      window.removeEventListener("blur", onBlur);
      if (desktopOpened) {
        try { fallbackTab?.close(); } catch (_) {}
      } else {
        try {
          if (fallbackTab && !fallbackTab.closed) {
            fallbackTab.location.href = outlookWebUrl;
          }
        } catch (_) {}
      }
    }, 600);
  };

  // ─── DERIVED: last 5 sent Mail IDs for the current user ─────────────────────
  const sentHistory = React.useMemo(() => {
    const sentOrders = allOrdersFromFirestore.filter(
      (o) =>
        o.header?.status === "sent" &&
        o.header?.lastModifiedBy === userId
    );
    const grouped = {};
    sentOrders.forEach((o) => {
      const mid = o.header?.mailId;
      if (!mid) return;
      if (!grouped[mid]) {
        grouped[mid] = { mailId: mid, orders: [], latestUpdatedAt: 0 };
      }
      grouped[mid].orders.push(o);
      if ((o.header?.updatedAt || 0) > grouped[mid].latestUpdatedAt) {
        grouped[mid].latestUpdatedAt = o.header.updatedAt;
      }
    });
    return Object.values(grouped)
      .sort((a, b) => b.latestUpdatedAt - a.latestUpdatedAt)
      .slice(0, 5);
  }, [allOrdersFromFirestore, userId]);

  const resendFromHistory = async (mailId) => {
    setHistoryResendingId(mailId);
    try {
      const group = sentHistory.find((g) => g.mailId === mailId);
      if (!group) return;

      const ordersToProcess = [...group.orders].sort(
        (a, b) => (a.header?.createdAt || 0) - (b.header?.createdAt || 0)
      );

      let innerEmailContentHtml = "";
      const allProveedores = new Set();
      const allEspecies = new Set();

      // ── v26: compute the widest card so all cards render at the same width
      const maxWidth = Math.max(...ordersToProcess.map(o => measureOrderWidth(o.header, o.items)));

      ordersToProcess.forEach((order, index) => {
        innerEmailContentHtml += `
            <div style="margin-bottom:8px;margin-left:16px;">
              <h3 style="margin:0 0 6px 0;font-size:16px;color:#2563eb;font-family:Arial,sans-serif;">Pedido #${index + 1}</h3>
            </div>
            ${generateSingleOrderHtml(order.header, order.items, index + 1, maxWidth)}
        `;
        if (order.header.reDestinatarios) allProveedores.add(order.header.reDestinatarios);
        order.items.forEach((item) => { if (item.especie) allEspecies.add(item.especie); });
      });

      const consolidatedSubject = generateEmailSubjectValue(
        Array.from(allProveedores),
        Array.from(allEspecies),
        ""
      );

      const fullEmailBodyHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Detalle de Pedido</title>
</head>
<body style="margin:0;padding:16px 16px 16px 16px;background-color:#f8f8f8;font-family:Arial,sans-serif;">
  <div style="width:100%;text-align:right;margin-bottom:12px;font-family:Arial,sans-serif;font-weight:bold;font-size:14px;color:#ef4444;">Mail ID: ${mailId}</div>
  <div>
    ${innerEmailContentHtml}
  </div>
</body>
</html>`;

      await copyFormattedContentToClipboard(fullEmailBodyHtml);

      openEmailClient(consolidatedSubject);
    } catch (err) {
    } finally {
      setHistoryResendingId(null);
    }
  };

  const formatRelativeTime = (timestamp) => {
    if (!timestamp) return "";
    const diff = Date.now() - timestamp;
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return "justo ahora";
    if (mins < 60) return `hace ${mins} min`;
    const hrs = Math.floor(mins / 60);
    if (hrs < 24) return `hace ${hrs} h`;
    const days = Math.floor(hrs / 24);
    return `hace ${days} día${days > 1 ? "s" : ""}`;
  };

  // Firebase Initialization and Authentication Effect
  useEffect(() => {
    // ── v22 FIX (memory leak): Previously `return () => unsubscribe()` was
    // inside the async initAuth() function. useEffect ignores return values
    // from async functions (they return Promises, not cleanup functions).
    // The onAuthStateChanged listener was therefore NEVER unsubscribed.
    // Fix: capture the unsubscribe ref outside the async block and return
    // the cleanup directly to useEffect.
    let unsubscribeAuth = null;

    const initAuth = async () => {
      try {
        if (
          typeof __initial_auth_token !== "undefined" &&
          __initial_auth_token
        ) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          // Sign in anonymously only if not already authenticated
          if (!auth.currentUser) {
            await signInAnonymously(auth);
          }
        }
      } catch (error) {
        // Auth failure — will still attach the listener below
      }
    };

    unsubscribeAuth = onAuthStateChanged(auth, async (user) => {
      if (user) {
        setUserId(user.uid);
        setIsAuthReady(true);
      } else {
        try {
          if (
            typeof __initial_auth_token !== "undefined" &&
            __initial_auth_token
          ) {
            await signInWithCustomToken(auth, __initial_auth_token);
          } else {
            await signInAnonymously(auth);
          }
          setUserId(auth.currentUser?.uid ?? null);
        } catch (anonError) {
          setUserId(null);
        } finally {
          setIsAuthReady(true);
        }
      }
    });

    initAuth();

    // This cleanup NOW reaches useEffect correctly
    return () => {
      if (unsubscribeAuth) unsubscribeAuth();
    };
  }, []);

  const getCurrentWeekNumber = () => {
    const date = new Date();
    date.setHours(0, 0, 0, 0);
    date.setDate(date.getDate() + 3 - ((date.getDay() + 6) % 7));
    const week1 = new Date(date.getFullYear(), 0, 4);
    return (
      1 +
      Math.round(
        ((date.getTime() - week1.getTime()) / 86400000 -
          3 +
          ((week1.getDay() + 6) % 7)) /
          7
      )
    );
  };

  const SPECIES_NORMALIZATION_MAP = {
    MANZANA: "MANZANA",
    MANZANAS: "MANZANA",
    MZ: "MANZANA",
    APPLE: "MANZANA",
    APPLES: "MANZANA",
    PERA: "PERA",
    PERAS: "PERA",
    PR: "PERA",
    PEAR: "PERA",
    PEARS: "PERA",
    UVA: "UVA",
    UVAS: "UVA",
    GRAPE: "UVA",
    GRAPES: "UVA",
    UV: "UVA",
    KIWI: "KIWI",
    KIWIS: "KIWI",
    NARANJA: "NARANJA",
    NARANJAS: "NARANJA",
    LIMON: "LIMON",
    LIMONES: "LIMON",
    LM: "LIMON",
    CEREZA: "CEREZA",
    CEREZAS: "CEREZA",
    CLEMENTINA: "CLEMENTINA",
    CLEMENTINAS: "CLEMENTINA",
    CIRUELA: "CIRUELA",
    CIRUELAS: "CIRUELA",
    DURAZNO: "DURAZNO",
    DURAZNOS: "DURAZNO",
    NECTARIN: "NECTARIN",
    NECTARINES: "NECTARIN",
  };

  const normalizeSpeciesName = (name) => {
    if (typeof name !== "string" || name.trim() === "") {
      return "";
    }
    const upperName = name.toUpperCase().trim();
    return SPECIES_NORMALIZATION_MAP[upperName] || upperName;
  };

  const generateEmailSubjectValue = (
    proveedores,
    especies,
    mailGlobalId = ""
  ) => {
    const weekNumber = getCurrentWeekNumber();

    const formatPart = (arr, defaultValue) => {
      const safeArr = Array.isArray(arr) ? arr : [];
      const uniqueValues = Array.from(
        new Set(
          safeArr
            .filter((val) => typeof val === "string" && val.trim() !== "")
            .map((val) => normalizeSpeciesName(val))
        )
      );
      return uniqueValues.length > 0 ? uniqueValues.join("-") : defaultValue;
    };

    const uniqueProveedores = Array.from(
      new Set(Array.isArray(proveedores) ? proveedores : [])
    ).filter((val) => typeof val === "string" && val.trim() !== "");
    const formattedProveedor =
      uniqueProveedores.length > 0
        ? uniqueProveedores[0].toUpperCase().replace(/[^A-Z0-9]/g, "")
        : "PROVEEDOR";

    const formattedEspecie = formatPart(especies, "ESPECIE");

    const subject = `PED–W${weekNumber}–${formattedProveedor}–${formattedEspecie}`;
    return subject;
  };

  // initialHeaderState and initialItemState defined at module level (below)

  const [headerInfo, setHeaderInfo] = useState(() => ({
    ...initialHeaderState,
    emailSubject: generateEmailSubjectValue([], []),
  }));

  const [orderItems, setOrderItems] = useState([{ ...initialItemState, id: crypto.randomUUID() }]);
  const [currentOrderIndex, setCurrentOrderIndex] = useState(0);
  const [showOrderActionsModal, setShowOrderActionsModal] = useState(false);
  const [previewHtmlContent, setPreviewHtmlContent] = useState("");
  const [isShowingPreview, setIsShowingPreview] = useState(false);
  const [emailActionTriggered, setEmailActionTriggered] = useState(false);

  // ── v22 FIX (setState after unmount): guard all async operations
  const isMountedRef = useRef(true);
  useEffect(() => {
    isMountedRef.current = true;
    return () => { isMountedRef.current = false; };
  }, []);

  const [showObservationModal, setShowObservationModal] = useState(false);
  const [currentEditingItemData, setCurrentEditingItemData] = useState(null);
  const [modalObservationText, setModalObservationText] = useState("");
  const observationTextareaRef = useRef(null);

  const headerInputRefs = useRef({});
  const tableInputRefs = useRef({});
  const isUserEditing = useRef(false);
  const pendingSaveTimeout = useRef(null);
  const orderItemsRef = useRef([]);
  orderItemsRef.current = orderItems;
  // ── v20: headerInfo also needs a ref — it has the same stale-closure risk
  // as orderItems. Any action function that fires synchronously after a blur
  // must read from this ref instead of the closure value.
  const headerInfoRef = useRef(headerInfo);
  headerInfoRef.current = headerInfo;
  const activeOrderIdRef = useRef(null);
  activeOrderIdRef.current = activeOrderId;

  // ─── DRAG & DROP REFS ────────────────────────────────────────────────────────
  const dragItemIndex = useRef(null);
  const dragOverIndex = useRef(null);

  const headerInputOrder = [
    "reDestinatarios",
    "deNombrePais",
    "nave",
    "fechaCarga",
    "exporta",
    "emailSubject",
  ];
  const tableColumnOrder = [
    "pallets",
    "especie",
    "variedad",
    "formato",
    "calibre",
    "categoria",
    "preciosFOB",
  ];

  const saveOrderToFirestore = async (orderToSave) => {
    if (!db || !userId) {
      return;
    }
    try {
      const ordersCollectionRef = collection(
        db,
        `artifacts/${appId}/public/data/pedidos`
      );
      const { header, items } = orderToSave;
      const orderDocId = orderToSave.id || doc(ordersCollectionRef).id;
      const orderDocRef = doc(ordersCollectionRef, orderDocId);

      if (orderToSave.isNew) {
        await setDoc(orderDocRef, {
          header: {
            ...header,
            lastModifiedBy: userId,
            updatedAt: Date.now(),
          },
          items: JSON.stringify(items),
        });
      } else {
        const updatePayload = { items: JSON.stringify(items) };
        const skipFields = new Set(["createdBy", "createdAt"]);
        Object.entries(header).forEach(([key, value]) => {
          if (!skipFields.has(key)) {
            updatePayload[`header.${key}`] = value;
          }
        });
        updatePayload["header.lastModifiedBy"] = userId;
        updatePayload["header.updatedAt"] = Date.now();
        await updateDoc(orderDocRef, updatePayload);
      }
    } catch (error) {
      throw error;
    }
  };

  const handleSoftDeleteOrderInFirestore = async (orderDocIdToDelete) => {
    if (!db || !userId) {
      return;
    }
    try {
      const orderDocRef = doc(
        db,
        `artifacts/${appId}/public/data/pedidos`,
        orderDocIdToDelete
      );
      await updateDoc(orderDocRef, {
        "header.status": "deleted",
        "header.lastModifiedBy": userId,
        "header.updatedAt": Date.now(),
      });
    } catch (error) {
    }
  };

  // ─── PRESENCE HELPERS ───────────────────────────────────────────────────────

  const writePresence = async (mailId) => {
    if (!db || !userId || !mailId) return;
    try {
      const presenceRef = doc(
        db,
        `artifacts/${appId}/public/data/presence/${mailId}/users/${userId}`
      );
      await setDoc(presenceRef, {
        userId,
        openedAt: Date.now(),
        expiresAt: Date.now() + 5 * 60 * 1000,
      });
    } catch (error) {
    }
  };

  const clearPresence = async (mailId) => {
    if (!db || !userId || !mailId) return;
    try {
      const presenceRef = doc(
        db,
        `artifacts/${appId}/public/data/presence/${mailId}/users/${userId}`
      );
      await deleteDoc(presenceRef);
    } catch (error) {
    }
  };

  const subscribeToPresence = (mailId) => {
    if (presenceUnsubscribeRef.current) {
      presenceUnsubscribeRef.current();
      presenceUnsubscribeRef.current = null;
    }
    if (!db || !mailId) return;

    const presenceCollectionRef = collection(
      db,
      `artifacts/${appId}/public/data/presence/${mailId}/users`
    );

    const unsubscribe = onSnapshot(presenceCollectionRef, (snapshot) => {
      const now = Date.now();
      const others = snapshot.docs
        .map((d) => d.data())
        .filter(
          (p) =>
            p.userId !== userId &&
            p.expiresAt > now
        )
        .map((p) => p.userId.substring(0, 8).toUpperCase());

      setOtherUsersPresent(others);
    });

    presenceUnsubscribeRef.current = unsubscribe;
  };

  const checkPresenceConflict = async (mailId) => {
    if (!db || !mailId) return [];
    try {
      const presenceCollectionRef = collection(
        db,
        `artifacts/${appId}/public/data/presence/${mailId}/users`
      );
      const snapshot = await getDocs(presenceCollectionRef);
      const now = Date.now();
      const others = snapshot.docs
        .map((d) => d.data())
        .filter((p) => p.userId !== userId && p.expiresAt > now)
        .map((p) => p.userId.substring(0, 8).toUpperCase());
      return others;
    } catch (error) {
      return [];
    }
  };

  useEffect(() => {
    if (!presenceMailId) return;
    const interval = setInterval(() => {
      writePresence(presenceMailId);
    }, 2 * 60 * 1000);
    return () => clearInterval(interval);
  }, [presenceMailId]);

  useEffect(() => {
    const handleUnload = () => {
      if (presenceMailId && db && userId) {
        clearPresence(presenceMailId);
      }
    };
    window.addEventListener("beforeunload", handleUnload);
    return () => window.removeEventListener("beforeunload", handleUnload);
  }, [presenceMailId]);

  useEffect(() => {
    // ── v22 FIX (unnecessary renders / infinite loop risk): Previously
    // `headerInfo.emailSubject` was in the dependency array. The effect
    // reads it to compare AND writes it via setHeaderInfo — any mismatch
    // between the computed and stored value could cause a perpetual re-run.
    // Fix: remove emailSubject from deps. The effect only needs to react to
    // the inputs that feed the subject calculation (proveedor, especie, mailId).
    const currentProveedor = headerInfo.reDestinatarios;
    const currentEspecie = orderItems[0]?.especie || "";

    const newSubject = generateEmailSubjectValue(
      [currentProveedor],
      [currentEspecie],
      headerInfo.mailId
    );

    // Use functional setter to compare against the truly-current value,
    // avoiding a stale closure read of headerInfo.emailSubject
    setHeaderInfo((prevInfo) => {
      if (prevInfo.emailSubject === newSubject) return prevInfo; // no-op, no re-render
      return { ...prevInfo, emailSubject: newSubject };
    });
  }, [
    headerInfo.reDestinatarios,
    orderItems[0]?.especie,
    headerInfo.mailId,
    // emailSubject intentionally omitted — it is the OUTPUT, not an input
  ]);

  const parseFirestoreOrders = (docs) => {
    const orders = docs.map((d) => {
      const data = d.data();
      let parsedItems = [];
      if (data.items) {
        if (typeof data.items === "string") {
          try { parsedItems = JSON.parse(data.items); } catch { parsedItems = []; }
        } else if (Array.isArray(data.items)) {
          parsedItems = data.items;
        }
      }
      return {
        id: d.id,
        header: data.header,
        items: Array.isArray(parsedItems) ? parsedItems : [],
      };
    });
    return orders.sort((a, b) => (a.header?.createdAt || 0) - (b.header?.createdAt || 0));
  };

  const loadSentHistory = async (currentDrafts = []) => {
    if (!db || !userId) return;
    try {
      const ordersCollectionRef = collection(db, `artifacts/${appId}/public/data/pedidos`);
      // Single where clause only — avoids composite index requirement
      const q = query(
        ordersCollectionRef,
        where("header.status", "==", "sent")
      );
      const snapshot = await getDocs(q);
      // Filter by userId in memory
      const sentOrders = parseFirestoreOrders(snapshot.docs).filter(
        (o) => o.header?.lastModifiedBy === userId
      );
      setAllOrdersFromFirestore([...currentDrafts, ...sentOrders]);
    } catch (err) {
      console.error("loadSentHistory error:", err);
      setAllOrdersFromFirestore((prev) => {
        const existingSent = prev.filter((o) => o.header?.status === "sent");
        return [...currentDrafts, ...existingSent];
      });
    }
  };

  const purgeMyDrafts = async () => {
    // Soft-delete todos los drafts existentes del usuario antes de crear uno nuevo
    if (!db || !userId) return;
    const ordersCollectionRef = collection(db, `artifacts/${appId}/public/data/pedidos`);
    const q = query(
      ordersCollectionRef,
      where("header.createdBy", "==", userId),
      where("header.status", "==", "draft")
    );
    const snapshot = await getDocs(q);
    await Promise.all(
      snapshot.docs.map((d) => handleSoftDeleteOrderInFirestore(d.id))
    );
  };

  const loadMyDrafts = async ({ silent = false } = {}) => {
    if (!db || !userId) return;
    if (!silent) setIsLoading(true);
    try {
      const ordersCollectionRef = collection(db, `artifacts/${appId}/public/data/pedidos`);
      const q = query(
        ordersCollectionRef,
        where("header.createdBy", "==", userId),
        where("header.status", "==", "draft")
      );
      const snapshot = await getDocs(q);
      // ── v22 FIX: guard against setState after unmount
      if (!isMountedRef.current) return;
      // Si es llamada silenciosa (limpiar/nuevo mail) purgar y crear draft limpio
      // Si es llamada normal (post-send) usar drafts existentes
      let drafts = parseFirestoreOrders(snapshot.docs);

      if (silent) {
        // Purgar todos los drafts acumulados y empezar con uno nuevo
        await Promise.all(drafts.map((d) => handleSoftDeleteOrderInFirestore(d.id)));
        drafts = [];
      }

      if (drafts.length === 0) {
        const newRef = doc(ordersCollectionRef);
        const newHeader = {
          ...initialHeaderState,
          mailId: crypto.randomUUID().substring(0, 8).toUpperCase(),
          emailSubject: generateEmailSubjectValue([], []),
          status: "draft",
          createdAt: Date.now(),
          createdBy: userId,
          updatedAt: Date.now(),
        };
        const newItems = [{ ...initialItemState, id: crypto.randomUUID() }];
        await saveOrderToFirestore({ id: newRef.id, isNew: true, header: newHeader, items: newItems });
        if (!isMountedRef.current) return;
        const newDraft = { id: newRef.id, header: newHeader, items: newItems };
        setDisplayedOrders([newDraft]);
        setActiveOrderId(newRef.id);
        activeOrderIdRef.current = newRef.id;
        headerInfoRef.current = newHeader;
        orderItemsRef.current = newItems;
        setHeaderInfo(newHeader);
        setOrderItems(newItems);
        setCurrentOrderIndex(0);
        await loadSentHistory([newDraft]);
      } else {
        setDisplayedOrders(drafts);
        const firstId = drafts[0].id;
        setActiveOrderId(firstId);
        activeOrderIdRef.current = firstId;
        headerInfoRef.current = drafts[0].header;
        orderItemsRef.current = drafts[0].items;
        setHeaderInfo(drafts[0].header);
        setOrderItems(drafts[0].items);
        setCurrentOrderIndex(0);
        await loadSentHistory(drafts);
      }
    } catch (err) {
      console.error("loadMyDrafts error:", err);
    } finally {
      if (isMountedRef.current) setIsLoading(false);
    }
  };

  const searchByMailId = async (mailId) => {
    if (!db || !mailId) return;
    setIsLoading(true);
    try {
      const ordersCollectionRef = collection(db, `artifacts/${appId}/public/data/pedidos`);
      const q = query(
        ordersCollectionRef,
        where("header.mailId", "==", mailId)
      );
      const snapshot = await getDocs(q);
      // ── v22 FIX: guard against setState after unmount
      if (!isMountedRef.current) return;
      const orders = parseFirestoreOrders(snapshot.docs).filter(
        (o) => o.header?.status !== "deleted"
      );

      // Re-fetch sent history so it's always fresh, not stale from prev state
      let freshSent = [];
      try {
        const sentQ = query(
          ordersCollectionRef,
          where("header.status", "==", "sent")
        );
        const sentSnap = await getDocs(sentQ);
        freshSent = parseFirestoreOrders(sentSnap.docs).filter(
          (o) => o.header?.lastModifiedBy === userId
        );
      } catch (_) {
        freshSent = [];
      }
      setAllOrdersFromFirestore((prev) => {
        const existingSent = freshSent.length > 0
          ? freshSent
          : prev.filter((o) => o.header?.status === "sent");
        return [...orders, ...existingSent];
      });

      if (orders.length === 0) {
        setDisplayedOrders([]);
        setActiveOrderId(null);
        setCurrentOrderIndex(0);
      } else {
        // ── BUG FIX: keep original status from Firestore — never force "draft"
        // Mutating sent→draft here was causing flushAndSave to persist "draft"
        // back to Firestore and handleNewMail to soft-delete sent orders.
        setDisplayedOrders(orders);
        const last = orders[orders.length - 1];
        setActiveOrderId(last.id);
        activeOrderIdRef.current = last.id;
        headerInfoRef.current = last.header;
        orderItemsRef.current = last.items;
        setHeaderInfo(last.header);
        setOrderItems(last.items);
        setCurrentOrderIndex(orders.length - 1);
      }
    } catch (err) {
      console.error("searchByMailId error:", err);
    } finally {
      if (isMountedRef.current) setIsLoading(false);
    }
  };

  useEffect(() => {
    if (!isAuthReady || !userId || !db) return;
    // Al montar: crear un pedido en blanco nuevo, sin cargar drafts anteriores.
    // El usuario parte desde cero cada sesión — puede buscar un mailID existente
    // o simplemente empezar a armar un pedido nuevo.
    const initBlankOrder = async () => {
      setIsLoading(true);
      try {
        // Purgar drafts huérfanos antes de crear uno nuevo
        await purgeMyDrafts();
        const ordersCollectionRef = collection(db, `artifacts/${appId}/public/data/pedidos`);
        const newRef = doc(ordersCollectionRef);
        const newHeader = {
          ...initialHeaderState,
          mailId: crypto.randomUUID().substring(0, 8).toUpperCase(),
          emailSubject: generateEmailSubjectValue([], []),
          status: "draft",
          createdAt: Date.now(),
          createdBy: userId,
          updatedAt: Date.now(),
        };
        const newItems = [{ ...initialItemState, id: crypto.randomUUID() }];
        await saveOrderToFirestore({ id: newRef.id, isNew: true, header: newHeader, items: newItems });
        if (!isMountedRef.current) return;
        const newDraft = { id: newRef.id, header: newHeader, items: newItems };
        setDisplayedOrders([newDraft]);
        setActiveOrderId(newRef.id);
        activeOrderIdRef.current = newRef.id;
        setHeaderInfo(newHeader);
        headerInfoRef.current = newHeader;
        setOrderItems(newItems);
        orderItemsRef.current = newItems;
        setCurrentOrderIndex(0);
      } catch (err) {
        console.error("initBlankOrder error:", err);
      } finally {
        if (isMountedRef.current) setIsLoading(false);
      }
    };
    initBlankOrder();
  }, [isAuthReady, userId]);

  // ── v20 ARCHITECTURE: Single flush helper used by ALL action functions.
  // Reads from refs (always current) unless caller passes explicit overrides.
  // The debounce-autosave also calls this with no arguments — it reads from
  // refs too, so both paths are consistent.
  const saveCurrentFormDataToDisplayed = async (
    explicitHeader = null,
    explicitItems = null
  ) => {
    if (!db || !userId || !activeOrderIdRef.current) return;
    // Guard: don't save if user hasn't actually edited anything since the order was loaded.
    if (!isUserEditing.current) return;

    // Always read from refs — never from closures
    const snapshotHeader = explicitHeader ?? headerInfoRef.current;
    const snapshotItems  = explicitItems  ?? orderItemsRef.current;

    const currentOrderData = {
      id: activeOrderIdRef.current,
      header: { ...snapshotHeader },
      items: snapshotItems.map((item) => ({ ...item })),
    };

    // Always update local displayedOrders — preview and navigation rely on this
    setDisplayedOrders((prev) =>
      prev.map((o) => (o.id === activeOrderIdRef.current ? { ...o, ...currentOrderData } : o))
    );

    // ── SENT MODE GUARD: if we are viewing a searched mailID (committedSearchTermRef set)
    // AND the active order is sent, do NOT write to Firestore.
    // Changes are kept in local state only — Firestore is only updated by performSendEmail.
    const activeOrderStatus = displayedOrders.find(
      (o) => o.id === activeOrderIdRef.current
    )?.header?.status;
    const isViewingSearchedMail = !!committedSearchTermRef.current;
    if (isViewingSearchedMail && activeOrderStatus === "sent") return;

    await saveOrderToFirestore(currentOrderData);
  };

  // ── v20: Central flush used by all action functions (navigate/send/preview/
  // finalize/search). Cancels any pending debounce and immediately persists
  // the current ref state to both displayedOrders and Firestore.
  //
  // ── v21: Forces blur on the currently focused element before reading refs.
  // This ensures that if the user clicks an action button while a field still
  // has focus (no blur has fired yet), the onBlur formatting handler runs
  // and the formatted value is captured in the ref before we save.
  const flushAndSave = async () => {
    // Force blur so that any pending onBlur handlers (formatters) run
    if (document.activeElement && document.activeElement !== document.body) {
      document.activeElement.blur();
      // Yield one microtask tick so that React processes the blur setState
      // before we read from refs
      await new Promise((r) => setTimeout(r, 0));
    }

    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
      pendingSaveTimeout.current = null;
    }
    // ── IMPORTANT: set isUserEditing=true so saveCurrentFormDataToDisplayed
    // doesn't skip the save (guard requires isUserEditing to be true).
    // flushAndSave is ONLY called from explicit user actions (navigate/send/preview)
    // so it is always correct to force-save here regardless of editing state.
    isUserEditing.current = true;
    await saveCurrentFormDataToDisplayed();
    isUserEditing.current = false;
  };

  const handleHeaderChange = (e) => {
    const { name, value } = e.target;
    setHeaderInfo((prevInfo) => ({ ...prevInfo, [name]: value }));
  };

  const handleHeaderBlur = (e) => {
    const { name, value, type } = e.target;
    if (
      type !== "date" &&
      type !== "number" &&
      name !== "emailSubject" &&
      name !== "mailId"
    ) {
      setHeaderInfo((prevInfo) => ({
        ...prevInfo,
        [name]: value.toUpperCase(),
      }));
    }
  };

  const handleFocusField = () => {
    isUserEditing.current = true;
    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
      pendingSaveTimeout.current = null;
    }
  };

  // ── v20 ARCHITECTURE: The v19 fix introduced pendingItemsRef/pendingHeaderRef
  // as a two-step intermediate. This added complexity without full coverage
  // (headerInfo still had a stale closure risk in v19).
  //
  // The correct and complete solution is simpler:
  //   • orderItemsRef.current  — always the freshest orderItems (synced above)
  //   • headerInfoRef.current  — always the freshest headerInfo (synced above)
  //
  // All action functions (navigate, finalize, send, preview, search) read from
  // these refs directly. No intermediate pending refs needed.
  //
  // The debounce-based autosave is KEPT but its only job is Firestore
  // persistence during idle time. It is never relied upon for correctness —
  // all critical paths flush explicitly before acting.

  const handleBlurField = (originalBlurFn) => (e) => {
    if (originalBlurFn) originalBlurFn(e);

    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
    }
    // Debounce only for idle Firestore persistence. Action functions do NOT
    // wait for this — they read from refs directly (see flushAndSave below).
    pendingSaveTimeout.current = setTimeout(async () => {
      pendingSaveTimeout.current = null;
      await saveCurrentFormDataToDisplayed();
      isUserEditing.current = false;
    }, 300);
  };

  // ── v22 FIX (unnecessary renders): handleItemChange and handleItemBlur are
  // called on every input in the table. If they are inline arrow functions,
  // React creates a new function reference on every render of the parent,
  // which forces every input to re-render even when unrelated state changes
  // (e.g. headerInfo updates cause all table inputs to re-render).
  // Stabilizing with useCallback ensures stable references between renders.
  // Note: these still update correctly because they use functional setState.
  const handleItemChange = React.useCallback((itemId, e) => {
    const { name, value } = e.target;
    // ── v40: clamp pallets to 0 minimum on every change
    const sanitized = name === "pallets" && value !== ""
      ? String(Math.max(0, parseFloat(value) || 0))
      : value;
    setOrderItems((prevItems) =>
      prevItems.map((item) =>
        item.id === itemId ? { ...item, [name]: sanitized } : item
      )
    );
  }, []); // no deps — uses only stable setState

  const handleItemBlur = React.useCallback((itemId, e) => {
    const { name, value, type } = e.target;
    setOrderItems((prevItems) => {
      const updatedItems = prevItems.map((item) => {
        if (item.id === itemId) {
          let newValue = value;
          if (name === "preciosFOB") {
            const matches = value.match(/\d+([.,]\d+)?/g);
            const formattedPrices = [];
            if (matches && matches.length > 0) {
              for (const match of matches) {
                const numericValue = parseFloat(match.replace(",", "."));
                if (!isNaN(numericValue)) {
                  formattedPrices.push(`$ ${numericValue.toFixed(2).replace(".", ",")}`);
                }
              }
            }
            newValue = formattedPrices.length > 0 ? formattedPrices.join(" - ") : "";
          } else if (name === "calibre") {
            const parts = value.split(/[,;\s-]+/).filter((part) => part.trim() !== "").map((part) => part.trim().toUpperCase());
            newValue = parts.join(" - ");
          } else if (name === "categoria") {
            const matches = value.match(/[a-zA-Z0-9]+/g);
            newValue = matches && matches.length > 0 ? matches.join(" - ").toUpperCase() : "";
          } else if (name === "pallets") {
            // ── v40: enforce minimum 0 on blur (covers paste and autofill)
            const parsed = parseFloat(value);
            newValue = isNaN(parsed) ? "" : String(Math.max(0, parsed));
          } else if (type !== "number") {
            newValue = value.toUpperCase();
          }
          return { ...item, [name]: newValue };
        }
        return item;
      });
      return updatedItems;
    });
  }, []); // no deps — pure transformation using only event data

  const handleAddItem = (sourceItemId = null) => {
    // ── v22 FIX (duplicate writes + stale closure): Previously saveOrderToFirestore
    // was called INSIDE the setOrderItems setter. This caused two problems:
    //   1. `headerInfo` captured in the setter is the stale closure value, not the
    //      current header (same class of bug as the original stale-closure series).
    //   2. The debounce autosave would also fire ~300ms later — two writes in flight.
    // Fix: compute the next items first, then write once using refs.
    setOrderItems((prevItems) => {
      let updatedItems;
      if (sourceItemId) {
        const sourceItem = prevItems.find((item) => item.id === sourceItemId);
        if (sourceItem) {
          const newItem = { ...sourceItem, id: crypto.randomUUID(), isCanceled: false };
          const index = prevItems.findIndex((item) => item.id === sourceItemId);
          updatedItems = [
            ...prevItems.slice(0, index + 1),
            newItem,
            ...prevItems.slice(index + 1),
          ];
        } else {
          updatedItems = [
            ...prevItems,
            {
              id: crypto.randomUUID(), pallets: "", especie: "", variedad: "",
              formato: "", calibre: "", categoria: "", preciosFOB: "", estado: "",
              isCanceled: false,
            },
          ];
        }
      } else {
        updatedItems = [
          ...prevItems,
          {
            id: crypto.randomUUID(), pallets: "", especie: "", variedad: "",
            formato: "", calibre: "", categoria: "", preciosFOB: "", estado: "",
            isCanceled: false,
          },
        ];
      }
      // Schedule the Firestore write AFTER React commits the state update,
      // reading from refs (always current) to avoid stale closures.
      setTimeout(() => {
        saveOrderToFirestore({
          id: activeOrderIdRef.current,
          header: { ...headerInfoRef.current },
          items: updatedItems,
        });
      }, 0);
      return updatedItems;
    });
  };

  const handleDeleteItem = (idToDelete) => {
    setOrderItems((prevItems) => {
      if (prevItems.length <= 1) return prevItems;
      const updatedItems = prevItems.filter((item) => item.id !== idToDelete);
      // ── v22 FIX: write outside setter using refs (not stale closures)
      setTimeout(() => {
        saveOrderToFirestore({
          id: activeOrderIdRef.current,
          header: { ...headerInfoRef.current },
          items: updatedItems,
        });
      }, 0);
      return updatedItems;
    });
  };

  const handleDragReorder = (fromIndex, toIndex) => {
    if (fromIndex === toIndex) return;
    setOrderItems((prev) => {
      const updated = [...prev];
      const [moved] = updated.splice(fromIndex, 1);
      updated.splice(toIndex, 0, moved);
      // ── v22 FIX: write outside setter using refs
      setTimeout(() => {
        saveOrderToFirestore({
          id: activeOrderIdRef.current,
          header: { ...headerInfoRef.current },
          items: updated,
        });
      }, 0);
      return updated;
    });
  };

  const toggleItemCancellation = (itemId) => {
    setOrderItems((prevItems) => {
      const updatedItems = prevItems.map((item) => {
        if (item.id === itemId) {
          const newIsCanceled = !item.isCanceled;
          return { ...item, isCanceled: newIsCanceled, estado: newIsCanceled ? "CANCELADO" : "" };
        }
        return item;
      });
      // ── v22 FIX: write outside setter using refs
      setTimeout(() => {
        saveOrderToFirestore({
          id: activeOrderIdRef.current,
          header: { ...headerInfoRef.current },
          items: updatedItems,
        });
      }, 0);
      return updatedItems;
    });
  };

  const currentOrderTotalPallets = orderItems.reduce((sum, item) => {
    if (item.isCanceled) {
      return sum;
    }
    // ── v39: Math.max(0, ...) prevents negative pallet values from
    // pulling the total below zero
    const pallets = Math.max(0, parseFloat(item.pallets) || 0);
    return sum + pallets;
  }, 0);

  const formatDateToSpanish = (dateString) => {
    const months = [
      "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
      "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ];
    const daysOfWeek = [
      "Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado",
    ];
    try {
      const date = new Date(dateString + "T00:00:00");
      if (isNaN(date.getTime())) {
        return dateString;
      }
      const dayOfWeek = daysOfWeek[date.getDay()];
      const dayOfMonth = date.getDate();
      const month = months[date.getMonth()];
      const year = date.getFullYear();
      return `${dayOfWeek} ${dayOfMonth} de ${month} de ${year}`;
    } catch (e) {
      return dateString;
    }
  };

  const escapeHtml = (str) => {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  };

  // ── v26: calculates the approximate pixel width a single order's table
  // would need, based on character counts in each column. Used to find the
  // widest order so all cards can be rendered at the same width.
  const measureOrderWidth = (orderHeader, orderItemsData) => {
    // px per character (Arial 11px ≈ 7px/char) + column padding (28px = 14px each side)
    const PX_PER_CHAR = 7;
    const COL_PADDING = 28;

    // Column headers (minimum width per column)
    const headers = ["Pallets", "Especie", "Variedad", "Formato", "Calibre", "Categoría",
      `Precios ${orderHeader.incoterm || "FOB"}`];

    // For each column, find the max char length across header + all data rows
    const colWidths = headers.map((h, colIdx) => {
      const colKey = ["pallets","especie","variedad","formato","calibre","categoria","preciosFOB"][colIdx];
      const maxDataLen = orderItemsData.reduce((max, item) => {
        const val = String(item[colKey] || "");
        return Math.max(max, val.length);
      }, 0);
      return Math.max(h.length, maxDataLen) * PX_PER_CHAR + COL_PADDING;
    });

    // Total table width = sum of all columns
    // Add card padding (15px each side = 30px) + border (2px)
    return colWidths.reduce((a, b) => a + b, 0) + 32;
  };

  const generateSingleOrderHtml = (
    orderHeader,
    orderItemsData,
    orderNumber,
    fixedWidth = null   // ── v26: if provided, card renders at this exact px width
  ) => {
    const nonCancelledItems = orderItemsData.filter((item) => !item.isCanceled);
    const singleOrderTotalPallets = nonCancelledItems.reduce((sum, item) => {
      const pallets = Math.max(0, parseFloat(item.pallets) || 0);
      return sum + pallets;
    }, 0);

    const allObservations = orderItemsData
      .map((item) => item.estado)
      .filter(
        (obs) => obs && obs.trim() !== "" && obs.toUpperCase() !== "CANCELADO"
      );

    const consolidatedObservationsText =
      allObservations.length > 0 ? escapeHtml(allObservations.join(" – ")) : "";

    const formattedNave       = escapeHtml(orderHeader.nave         || "");
    const formattedPais       = escapeHtml(orderHeader.deNombrePais || "");
    const formattedFechaCarga = escapeHtml(
      orderHeader.fechaCarga ? formatDateToSpanish(orderHeader.fechaCarga) : ""
    );
    const formattedExporta    = escapeHtml(orderHeader.exporta  || "");
    const incotermLabel       = escapeHtml(orderHeader.incoterm || "FOB");

    const orderBlockExtra = orderHeader.status === "deleted"
      ? "opacity:0.6;text-decoration:line-through;"
      : "";

    // ── v32: header fields use <table> rows instead of <p> tags
    // so Outlook respects font-size and color when HTML is pasted

    // ── v18: padding duplicado para mayor legibilidad en el email ──────────────
    const thStyle =
      "font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#ffffff;mso-color-alt:#ffffff;" +
      "background-color:#2563eb;padding:4px 8px;" +
      "border-top:1px solid #1e40af;border-bottom:1px solid #1e40af;" +
      "border-left:1px solid #1e40af;border-right:1px solid #1e40af;" +
      "text-align:center;white-space:nowrap;vertical-align:middle;";

    const tdBase =
      "font-family:Arial,sans-serif;font-size:11px;color:#333333;mso-color-alt:#333333;" +
      "padding:4px 8px;text-align:center;white-space:nowrap;vertical-align:middle;" +
      "border-top:1px solid #dddddd;border-bottom:1px solid #dddddd;" +
      "border-left:1px solid #dddddd;border-right:1px solid #dddddd;";

    const tdTotalLabel =
      "font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#333333;mso-color-alt:#333333;" +
      "background-color:#f0f0f0;padding:6px 14px 6px 6px;text-align:right;" +
      "white-space:nowrap;vertical-align:middle;" +
      "border-top:1px solid #dddddd;border-bottom:1px solid #dddddd;" +
      "border-left:1px solid #dddddd;border-right:1px solid #dddddd;";

    const tdTotalValue =
      "font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#333333;mso-color-alt:#333333;" +
      "background-color:#f0f0f0;padding:6px 14px;text-align:center;" +
      "white-space:nowrap;vertical-align:middle;" +
      "border-top:1px solid #dddddd;border-bottom:1px solid #dddddd;" +
      "border-left:1px solid #dddddd;border-right:1px solid #dddddd;";

    const dataRowsHtml = orderItemsData
      .map((item, idx) => {
        const rowBg = idx % 2 === 0 ? "#f9f9f9" : "#ffffff";
        const cancelExtra = item.isCanceled
          ? "color:#ef4444;text-decoration:line-through;"
          : "";
        const td = tdBase + `background-color:${rowBg};` + cancelExtra;
        return (
          `<tr style="background-color:${rowBg};">` +
          `<td style="${td}">${escapeHtml(item.pallets    || "")}</td>` +
          `<td style="${td}">${escapeHtml(item.especie    || "")}</td>` +
          `<td style="${td}">${escapeHtml(item.variedad   || "")}</td>` +
          `<td style="${td}">${escapeHtml(item.formato    || "")}</td>` +
          `<td style="${td}">${escapeHtml(item.calibre    || "")}</td>` +
          `<td style="${td}">${escapeHtml(item.categoria  || "")}</td>` +
          `<td style="${td}">${escapeHtml(item.preciosFOB || "")}</td>` +
          `</tr>`
        );
      })
      .join("");

    const cardWidthStyle = fixedWidth ? `width:${fixedWidth}px;` : "";
    const tableWidthStyle = fixedWidth ? `width:${fixedWidth - 32}px;` : "width:auto;";

    return `
<div style="margin-bottom:48px;margin-left:16px;display:inline-block;${cardWidthStyle}${orderBlockExtra}">
<div style="font-family:Arial,sans-serif;font-size:11px;color:#333333;mso-color-alt:#333333;background-color:#f8f8f8;border:1px solid #dddddd;border-radius:8px;padding:15px;text-align:left;box-sizing:border-box;">

  <div style="margin-bottom:10px;">
    <p style="margin:0 0 3px 0;font-family:Arial,sans-serif;font-size:14px;line-height:1.4;"><span style="font-weight:bold;">País:</span> ${formattedPais}</p>
    <p style="margin:0 0 3px 0;font-family:Arial,sans-serif;font-size:14px;line-height:1.4;"><span style="font-weight:bold;">Nave:</span> ${formattedNave}</p>
    <p style="margin:0 0 3px 0;font-family:Arial,sans-serif;font-size:14px;line-height:1.4;"><span style="font-weight:bold;">Fecha de carga:</span> ${formattedFechaCarga}</p>
    <p style="margin:0;font-family:Arial,sans-serif;font-size:14px;line-height:1.4;"><span style="font-weight:bold;">Exporta:</span> ${formattedExporta}</p>
  </div>

  <table cellpadding="0" cellspacing="0" border="0"
    style="border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;${tableWidthStyle}table-layout:auto;margin-top:8px;">
    <thead>
      <tr style="background-color:#2563eb;">
        <th style="${thStyle}">Pallets</th>
        <th style="${thStyle}">Especie</th>
        <th style="${thStyle}">Variedad</th>
        <th style="${thStyle}">Formato</th>
        <th style="${thStyle}">Calibre</th>
        <th style="${thStyle}">Categoría</th>
        <th style="${thStyle}">Precios ${incotermLabel}</th>
      </tr>
    </thead>
    <tbody>
      ${dataRowsHtml}
      <tr style="background-color:#f0f0f0;">
        <td colspan="6" style="${tdTotalLabel}">Total de Pallets:</td>
        <td colspan="1" style="${tdTotalValue}">${singleOrderTotalPallets} Pallets</td>
      </tr>
    </tbody>
  </table>

  <p style="margin-top:10px;margin-bottom:0;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;">
    Observaciones: <span style="font-weight:normal;font-style:italic;font-size:11px;font-family:Arial,sans-serif;">${consolidatedObservationsText}</span>
  </p>

</div>
</div>`;
  };

  const copyFormattedContentToClipboard = async (content) => {
    if (navigator.clipboard && window.ClipboardItem) {
      try {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([content], { type: "text/html" }),
            "text/plain": new Blob([content], { type: "text/plain" }),
          }),
        ]);
        return;
      } catch (err) {
      }
    }

    try {
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = content;
      tempDiv.style.cssText = "position:fixed;top:0;left:-9999px;opacity:0;pointer-events:none;";
      document.body.appendChild(tempDiv);

      const range = document.createRange();
      range.selectNodeContents(tempDiv);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);

      const success = document.execCommand("copy");
      selection.removeAllRanges();
      document.body.removeChild(tempDiv);

      if (success) {
        return;
      }
    } catch (err) {
    }

    try {
      await navigator.clipboard.writeText(content);
    } catch (err) {
    }
  };

  const handleAddOrder = async () => {
    if (!db || !userId) return;

    if (activeOrderId) {
      await flushAndSave();
    }

    // Read from refs after flush
    const currentHeader = headerInfoRef.current;

    let mailIdToAssignForNewOrder = "";
    if (committedSearchTerm) {
      mailIdToAssignForNewOrder = committedSearchTerm;
    } else {
      if (currentHeader.mailId && currentHeader.status === "draft") {
        mailIdToAssignForNewOrder = currentHeader.mailId;
      } else {
        mailIdToAssignForNewOrder = crypto.randomUUID().substring(0, 8).toUpperCase();
      }
    }

    const ordersCollectionRef = collection(db, `artifacts/${appId}/public/data/pedidos`);
    const newOrderDocRef = doc(ordersCollectionRef);
    const newOrderId = newOrderDocRef.id;

    const newBlankHeader = {
      ...initialHeaderState,
      reDestinatarios: currentHeader.reDestinatarios,
      emailSubject: generateEmailSubjectValue(
        [currentHeader.reDestinatarios],
        [],
        mailIdToAssignForNewOrder
      ),
      mailId: mailIdToAssignForNewOrder,
      status: "draft",
      createdAt: Date.now(),
      createdBy: userId,
      updatedAt: Date.now(),
    };
    const newBlankItems = [{ ...initialItemState, id: crypto.randomUUID() }];

    await saveOrderToFirestore({
      id: newOrderId,
      isNew: true,
      header: newBlankHeader,
      items: newBlankItems,
    });

    const newOrder = { id: newOrderId, header: newBlankHeader, items: newBlankItems };
    setDisplayedOrders((prev) => [...prev, newOrder]);
    setCurrentOrderIndex(displayedOrders.length); // length before adding = new last index
    setActiveOrderId(newOrderId);
    setHeaderInfo(newBlankHeader);
    setOrderItems(newBlankItems);
    // Only allow suggestion again if new order has no proveedor (user will type a new one)
    if (!committedSearchTerm && !newBlankHeader.reDestinatarios) suggestionAppliedRef.current = "";
  };

  const handlePreviousOrder = async () => {
    if (currentOrderIndex === 0) return;
    await flushAndSave();
    const newIndex = currentOrderIndex - 1;
    const target = displayedOrders[newIndex];
    if (target) {
      setActiveOrderId(target.id);
      setHeaderInfo(target.header);
      setOrderItems(target.items);
      setCurrentOrderIndex(newIndex);
    }
  };

  const handleNextOrder = async () => {
    if (currentOrderIndex === displayedOrders.length - 1) return;
    await flushAndSave();
    const newIndex = currentOrderIndex + 1;
    const target = displayedOrders[newIndex];
    if (target) {
      setActiveOrderId(target.id);
      setHeaderInfo(target.header);
      setOrderItems(target.items);
      setCurrentOrderIndex(newIndex);
    }
  };

  const handleDeleteCurrentOrder = async () => {
    try {
      if (displayedOrders.length <= 1) return;

      const orderIdToSoftDelete = displayedOrders[currentOrderIndex].id;
      await handleSoftDeleteOrderInFirestore(orderIdToSoftDelete);

      const updated = displayedOrders.filter((o) => o.id !== orderIdToSoftDelete);
      setDisplayedOrders(updated);

      const newIndex = Math.min(currentOrderIndex, updated.length - 1);
      const target = updated[newIndex];
      if (target) {
        setActiveOrderId(target.id);
        setHeaderInfo(target.header);
        setOrderItems(target.items);
        setCurrentOrderIndex(newIndex);
      } else {
        // No quedan pedidos — dejar UI vacía, no recargar drafts automáticamente
        setDisplayedOrders([]);
        setActiveOrderId(null);
        activeOrderIdRef.current = null;
        setCurrentOrderIndex(0);
        setHeaderInfo({ ...initialHeaderState });
        headerInfoRef.current = { ...initialHeaderState };
        setOrderItems([{ ...initialItemState, id: crypto.randomUUID() }]);
        orderItemsRef.current = [{ ...initialItemState, id: crypto.randomUUID() }];
      }
    } catch (error) {
    }
  };

  const handleSearchClick = async () => {
    const term = searchTerm.toUpperCase().trim();
    if (!term) return;

    suggestionAppliedRef.current = "";

    if (activeOrderId) await flushAndSave();

    const conflicts = await checkPresenceConflict(term);
    if (conflicts.length > 0) {
      setPresenceConflictUsers(conflicts);
      setPendingSearchTerm(term);
      setShowPresenceModal(true);
    } else {
      await enterMailId(term);
    }
  };

  const enterMailId = async (term) => {
    if (presenceMailId && presenceMailId !== term) {
      await clearPresence(presenceMailId);
      if (presenceUnsubscribeRef.current) {
        presenceUnsubscribeRef.current();
        presenceUnsubscribeRef.current = null;
      }
    }
    setPresenceMailId(term);
    await writePresence(term);
    subscribeToPresence(term);
    setCommittedSearchTerm(term);
    committedSearchTermRef.current = term;

    await searchByMailId(term);
  };

  const handleClearSearch = async () => {
    // Use ref for activeOrderId to avoid stale closure — check status via displayedOrders
    const activeStatus = displayedOrders.find(
      (o) => o.id === activeOrderIdRef.current
    )?.header?.status;
    if (activeOrderIdRef.current && activeStatus !== "sent") {
      await flushAndSave();
    }

    if (presenceMailId) {
      await clearPresence(presenceMailId);
      if (presenceUnsubscribeRef.current) {
        presenceUnsubscribeRef.current();
        presenceUnsubscribeRef.current = null;
      }
      setPresenceMailId(null);
      setOtherUsersPresent([]);
    }

    // Reset UI immediately — don't wait for Firestore
    setSearchTerm("");
    setCommittedSearchTerm("");
    committedSearchTermRef.current = "";
    setDisplayedOrders([]);
    setActiveOrderId(null);
    activeOrderIdRef.current = null;
    setCurrentOrderIndex(0);
    setHeaderInfo({ ...initialHeaderState });
    headerInfoRef.current = { ...initialHeaderState };
    setOrderItems([{ ...initialItemState, id: crypto.randomUUID() }]);
    orderItemsRef.current = [{ ...initialItemState, id: crypto.randomUUID() }];
    setProveedorSuggestion(null);
    if (suggestionTimerRef.current) clearTimeout(suggestionTimerRef.current);
    suggestionAppliedRef.current = "";

    await loadMyDrafts({ silent: true });
  };

  const handleNewMail = async () => {
    if (!db || !userId) return;

    // Reset refs BEFORE soft-delete so no stale proveedor leaks into new draft
    headerInfoRef.current = { ...initialHeaderState };
    orderItemsRef.current = [{ ...initialItemState, id: crypto.randomUUID() }];
    activeOrderIdRef.current = null;

    // Soft-delete all current drafts
    const currentDrafts = displayedOrders.filter(o => o.header?.status === "draft");
    await Promise.all(currentDrafts.map(o => handleSoftDeleteOrderInFirestore(o.id)));

    // Reset all UI state
    setDisplayedOrders([]);
    setActiveOrderId(null);
    setCurrentOrderIndex(0);
    setHeaderInfo({ ...initialHeaderState });
    setOrderItems([{ ...initialItemState, id: crypto.randomUUID() }]);
    setProveedorSuggestion(null);
    if (suggestionTimerRef.current) clearTimeout(suggestionTimerRef.current);
    suggestionAppliedRef.current = "";
    committedSearchTermRef.current = "";
    setSearchTerm("");
    setCommittedSearchTerm("");

    // Load drafts — finds 0, creates fresh one with new mailId and empty proveedor
    await loadMyDrafts({ silent: true });
  };

  const isMobileDevice = () => {
    return window.innerWidth <= 767;
  };

  const performSendEmail = async () => {
    try {
      if (!db || !userId) {
        return;
      }

      // Flush any pending blur state then persist to Firestore before acting
      await flushAndSave();

      // Read the authoritative state from refs (always current after flush)
      const latestHeader = headerInfoRef.current;
      const latestItems  = orderItemsRef.current;

      const mailGlobalId = latestHeader.mailId;

      if (!mailGlobalId) {
        setShowOrderActionsModal(false);
        return;
      }

      const ordersCollectionRef = collection(
        db,
        `artifacts/${appId}/public/data/pedidos`
      );

      const currentProveedor = latestHeader.reDestinatarios;
      if (currentProveedor && mailGlobalId) {
        const ordersToGroupQuery = query(
          ordersCollectionRef,
          where("header.status", "==", "draft"),
          where("header.reDestinatarios", "==", currentProveedor)
        );
        const ordersToGroupSnapshot = await getDocs(ordersToGroupQuery);
        const batch = writeBatch(db);
        let mutationCount = 0;

        ordersToGroupSnapshot.docs.forEach((docSnapshot) => {
          const orderData = docSnapshot.data();
          const existingMailId = orderData.header?.mailId;
          const docCreatedBy = orderData.header?.createdBy;

          if (docCreatedBy !== userId) {
            return;
          }

          if (!existingMailId || existingMailId !== mailGlobalId) {
            const orderRef = doc(
              db,
              `artifacts/${appId}/public/data/pedidos`,
              docSnapshot.id
            );
            batch.update(orderRef, { "header.mailId": mailGlobalId });
            mutationCount++;
          }
        });

        if (mutationCount > 0) {
          await batch.commit();
        }
      }
      const q = query(
        ordersCollectionRef,
        where("header.mailId", "==", mailGlobalId)
      );
      const querySnapshot = await getDocs(q);

      let ordersToProcess = querySnapshot.docs.map((doc) => ({
        id: doc.id,
        header: doc.data().header,
        items: JSON.parse(doc.data().items || "[]"),
      }));

      ordersToProcess = ordersToProcess.filter(
        (order) => order.header?.status !== "deleted"
      );

      const activeOrderCurrentState = {
        id: activeOrderId,
        // ── v19 FIX (Bug 2): use the ref-flushed snapshot, not the stale closure
        header: { ...latestHeader },
        items: latestItems.map((item) => ({ ...item })),
      };

      const existingIndexInProcess = ordersToProcess.findIndex(
        (o) => o.id === activeOrderId
      );
      if (existingIndexInProcess !== -1) {
        if (activeOrderCurrentState.header?.status !== "deleted") {
          ordersToProcess[existingIndexInProcess] = activeOrderCurrentState;
        } else {
          ordersToProcess.splice(existingIndexInProcess, 1);
        }
      } else if (
        activeOrderCurrentState.id &&
        activeOrderCurrentState.header?.mailId === mailGlobalId &&
        activeOrderCurrentState.header?.status !== "deleted"
      ) {
        ordersToProcess.push(activeOrderCurrentState);
      }

      ordersToProcess.sort((a, b) => {
        const dateA = a.header?.createdAt || 0;
        const dateB = b.header?.createdAt || 0;
        return dateA - dateB;
      });
      for (const order of ordersToProcess) {
        if (order.header?.status === "draft") {
          // ── v21 SECURITY FIX: only flip status on documents owned by the
          // current user. The mailId query (above) returns ALL docs with that
          // mailId — including docs created by other users who share the same
          // mailID. Without this guard, User A's "Send" would mark User B's
          // draft as "sent" in Firestore without B's consent.
          if (order.header?.createdBy && order.header.createdBy !== userId) {
            continue;
          }

          const orderDocRef = doc(
            db,
            `artifacts/${appId}/public/data/pedidos`,
            order.id
          );
          const isActiveOrder = order.id === activeOrderId;
          const updatePayload = {
            "header.mailId": mailGlobalId,
            "header.status": "sent",
            "header.lastModifiedBy": userId,
            "header.updatedAt": Date.now(),
          };
          if (isActiveOrder) {
            updatePayload["items"] = JSON.stringify(latestItems);
          }
          await updateDoc(orderDocRef, updatePayload);
          // Avoid direct mutation of React state objects — update via index
          const idx = ordersToProcess.indexOf(order);
          if (idx !== -1) {
            ordersToProcess[idx] = {
              ...order,
              header: { ...order.header, status: "sent", mailId: mailGlobalId }
            };
          }
        }
      }

      if (ordersToProcess.length === 0) {
        setPreviewHtmlContent(
          '<p style="text-align: center; color: #888;">No hay pedidos activos para enviar con este ID de Mail. Todos los pedidos asociados están eliminados o no existen.</p>'
        );
        setIsShowingPreview(true);
        setEmailActionTriggered(false);
        return;
      }

      let innerEmailContentHtml = "";
      const allProveedores = new Set();
      const allEspecies = new Set();

      // ── v26: compute the widest card so all cards render at the same width
      const maxWidth = Math.max(...ordersToProcess.map(o => measureOrderWidth(o.header, o.items)));

      ordersToProcess.forEach((order, index) => {
        innerEmailContentHtml += `
            <div style="margin-bottom:8px;margin-left:16px;">
              <h3 style="margin:0 0 6px 0;font-size:16px;color:#2563eb;font-family:Arial,sans-serif;">Pedido #${index + 1}</h3>
            </div>
            ${generateSingleOrderHtml(order.header, order.items, index + 1, maxWidth)}
        `;

        if (order.header.reDestinatarios)
          allProveedores.add(order.header.reDestinatarios);
        order.items.forEach((item) => {
          if (item.especie) allEspecies.add(item.especie);
        });
      });

      const consolidatedSubject = generateEmailSubjectValue(
        Array.from(allProveedores),
        Array.from(allEspecies),
        ""
      );

      const fullEmailBodyHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Detalle de Pedido</title>
</head>
<body style="margin:0;padding:16px 16px 16px 16px;background-color:#f8f8f8;font-family:Arial,sans-serif;">
  <div style="width:100%;text-align:right;margin-bottom:12px;font-family:Arial,sans-serif;font-weight:bold;font-size:14px;color:#ef4444;">Mail ID: ${mailGlobalId}</div>
  <div>
    ${innerEmailContentHtml}
  </div>
</body>
</html>`;

      await copyFormattedContentToClipboard(fullEmailBodyHtml);

      openEmailClient(consolidatedSubject);

      setShowOrderActionsModal(false);
      setSearchTerm("");
      setCommittedSearchTerm("");
      setEmailActionTriggered(false);
      setIsShowingPreview(false);
      setPreviewHtmlContent("");

      // Optimistically update allOrdersFromFirestore with just-sent orders
      // so sentHistory badge/panel refreshes immediately
      setAllOrdersFromFirestore((prev) => {
        const notSentYet = prev.filter((o) => o.header?.status !== "sent");
        const nowSent = ordersToProcess.map((o) => ({
          ...o,
          header: { ...o.header, status: "sent", lastModifiedBy: userId },
        }));
        const alreadySent = prev.filter((o) => o.header?.status === "sent");
        return [...notSentYet, ...nowSent, ...alreadySent];
      });

      await loadMyDrafts({ silent: true });
    } catch (error) {
    }
  };

  const handlePreviewOrder = async () => {
    await flushAndSave();

    // Read from refs after flush
    const latestItems  = orderItemsRef.current;
    const latestHeader = headerInfoRef.current;

    const previewGlobalId = latestHeader.mailId;

    if (!previewGlobalId) {
      setPreviewHtmlContent(
        '<p style="text-align: center; color: #888;">No hay pedidos para previsualizar sin un ID de Mail asociado.</p>'
      );
      setIsShowingPreview(true);
      return;
    }

    const ordersCollectionRef = collection(
      db,
      `artifacts/${appId}/public/data/pedidos`
    );

    // ── v19 FIX: use latestHeader throughout (not stale headerInfo closure)
    const currentProveedor = latestHeader.reDestinatarios;
    if (currentProveedor && previewGlobalId) {
      const ordersToGroupQuery = query(
        ordersCollectionRef,
        where("header.status", "==", "draft"),
        where("header.reDestinatarios", "==", currentProveedor)
      );
      const ordersToGroupSnapshot = await getDocs(ordersToGroupQuery);
      const batch = writeBatch(db);
      let mutationCount = 0;

      ordersToGroupSnapshot.docs.forEach((docSnapshot) => {
        const orderData = docSnapshot.data();
        const existingMailId = orderData.header?.mailId;
        const docCreatedBy = orderData.header?.createdBy;

        if (docCreatedBy !== userId) return;

        if (!existingMailId || existingMailId !== previewGlobalId) {
          const orderRef = doc(
            db,
            `artifacts/${appId}/public/data/pedidos`,
            docSnapshot.id
          );
          batch.update(orderRef, { "header.mailId": previewGlobalId });
          mutationCount++;
        }
      });

      if (mutationCount > 0) {
        await batch.commit();
      }
    }

    const q = query(
      ordersCollectionRef,
      where("header.mailId", "==", previewGlobalId)
    );
    const querySnapshot = await getDocs(q);

    let ordersForPreview = querySnapshot.docs.map((doc) => ({
      id: doc.id,
      header: doc.data().header,
      items: JSON.parse(doc.data().items || "[]"),
    }));

    ordersForPreview = ordersForPreview.filter(
      (order) => order.header?.status !== "deleted"
    );

    const activeOrderCurrentState = {
      id: activeOrderId,
      header: { ...latestHeader },
      items: latestItems.map((item) => ({ ...item })),
    };

    const existingIndex = ordersForPreview.findIndex(
      (o) => o.id === activeOrderId
    );
    if (existingIndex !== -1) {
      if (activeOrderCurrentState.header?.status !== "deleted") {
        ordersForPreview[existingIndex] = activeOrderCurrentState;
      } else {
        ordersForPreview.splice(existingIndex, 1);
      }
    } else if (
      activeOrderCurrentState.id &&
      activeOrderCurrentState.header?.mailId === previewGlobalId &&
      activeOrderCurrentState.header?.status !== "deleted"
    ) {
      ordersForPreview.push(activeOrderCurrentState);
    }

    ordersForPreview.sort((a, b) => {
      const dateA = a.header?.createdAt || 0;
      const dateB = b.header?.createdAt || 0;
      return dateA - dateB;
    });

    if (ordersForPreview.length === 0) {
      setPreviewHtmlContent(
        '<p style="text-align: center; color: #888;">No hay pedidos activos para previsualizar con este ID de Mail. Todos los pedidos asociados están eliminados o no existen.</p>'
      );
    } else {
      let innerPreviewHtml = "";

      // ── v26: compute the widest card so all cards render at the same width
      const maxWidth = Math.max(...ordersForPreview.map(o => measureOrderWidth(o.header, o.items)));

      ordersForPreview.forEach((order, index) => {
        innerPreviewHtml += `
            <div style="margin-bottom:8px;margin-left:16px;">
              <h3 style="margin:0 0 6px 0;font-size:16px;color:#2563eb;font-family:Arial,sans-serif;">Pedido #${index + 1}</h3>
            </div>
            ${generateSingleOrderHtml(order.header, order.items, index + 1, maxWidth)}
        `;
      });

      const finalPreviewHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Previsualización de Pedido</title>
</head>
<body style="margin:0;padding:16px 16px 16px 16px;background-color:#f8f8f8;font-family:Arial,sans-serif;">
  <div style="width:100%;text-align:right;margin-bottom:12px;font-family:Arial,sans-serif;font-weight:bold;font-size:14px;color:#ef4444;">Mail ID: ${previewGlobalId}</div>
  <div>
    ${innerPreviewHtml}
  </div>
</body>
</html>`;
      setPreviewHtmlContent(finalPreviewHtml);
    }
    setIsShowingPreview(true);
  };

  const handleFinalizeOrder = async () => {
    // Force blur so any pending onBlur formatters run before reading refs
    if (document.activeElement && document.activeElement !== document.body) {
      document.activeElement.blur();
      await new Promise((r) => setTimeout(r, 0));
    }
    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
      pendingSaveTimeout.current = null;
    }
    // Update local state only — do NOT write to Firestore here.
    // Firestore is only updated by performSendEmail.
    const snapshotHeader = headerInfoRef.current;
    const snapshotItems  = orderItemsRef.current;
    if (activeOrderIdRef.current) {
      setDisplayedOrders((prev) =>
        prev.map((o) => o.id === activeOrderIdRef.current
          ? { ...o, header: { ...snapshotHeader }, items: snapshotItems.map(i => ({ ...i })) }
          : o
        )
      );
    }

    // Read mailId from ref (not stale closure)
    let mailIdToAssign = headerInfoRef.current.mailId;
    if (!mailIdToAssign) {
      mailIdToAssign = crypto.randomUUID().substring(0, 8).toUpperCase();
      setHeaderInfo((prev) => ({ ...prev, mailId: mailIdToAssign }));
    }

    setShowOrderActionsModal(true);
    setEmailActionTriggered(false);
    setIsShowingPreview(false);
    setPreviewHtmlContent("");
  };

  const handleOpenObservationModal = (itemId) => {
    const itemToEdit = orderItems.find((item) => item.id === itemId);
    if (itemToEdit) {
      setCurrentEditingItemData(itemToEdit);
      setModalObservationText(itemToEdit.estado);
      setShowObservationModal(true);
    }
  };

  useEffect(() => {
    if (showObservationModal && observationTextareaRef.current) {
      observationTextareaRef.current.focus();
      if (observationTextareaRef.current.value) {
        observationTextareaRef.current.select();
      }
    }
  }, [showObservationModal]);

  const handleSaveObservation = () => {
    if (currentEditingItemData) {
      const formattedObservation = modalObservationText
        ? modalObservationText.charAt(0).toUpperCase() +
          modalObservationText.slice(1).toLowerCase()
        : "";

      const updatedItems = orderItems.map((item) =>
        item.id === currentEditingItemData.id
          ? { ...item, estado: formattedObservation }
          : item
      );
      setOrderItems(updatedItems);
      // ── v22 FIX: use refs instead of stale headerInfo closure
      saveOrderToFirestore({
        id: activeOrderIdRef.current,
        header: { ...headerInfoRef.current },
        items: updatedItems,
      });
    }
    setShowObservationModal(false);
    setCurrentEditingItemData(null);
    setModalObservationText("");
  };

  const handleCloseObservationModal = () => {
    setShowObservationModal(false);
    setCurrentEditingItemData(null);
    setModalObservationText("");
  };

  useEffect(() => {
    const handleKeyDown = (e) => {
      const activeElement = document.activeElement;
      let isHandled = false;

      const isOurInput = (element) => {
        return (
          element.tagName === "INPUT" &&
          element.name &&
          (headerInputOrder.includes(element.name) ||
            tableColumnOrder.includes(element.name))
        );
      };

      if (!isOurInput(activeElement)) {
        return;
      }

      const headerIndex = headerInputOrder.indexOf(activeElement.name);
      let tablePosition = null;
      if (headerIndex === -1) {
        for (const rowId in tableInputRefs.current) {
          for (const colName in tableInputRefs.current[rowId]) {
            if (tableInputRefs.current[rowId][colName] === activeElement) {
              tablePosition = { rowId, colName };
              break;
            }
          }
          if (tablePosition) break;
        }
      }

      if (
        e.key === "ArrowRight" ||
        e.key === "ArrowLeft" ||
        e.key === "ArrowUp" ||
        e.key === "ArrowDown"
      ) {
        e.preventDefault();

        if (headerIndex !== -1) {
          let nextIndex = headerIndex;
          if (e.key === "ArrowRight") {
            nextIndex = (headerIndex + 1) % headerInputOrder.length;
          } else if (e.key === "ArrowLeft") {
            nextIndex =
              (headerIndex - 1 + headerInputOrder.length) %
              headerInputOrder.length;
          } else if (e.key === "ArrowDown") {
            if (orderItemsRef.current.length > 0) {
              const firstRowId = orderItemsRef.current[0].id;
              const firstColName = tableColumnOrder[0];
              if (
                tableInputRefs.current[firstRowId] &&
                tableInputRefs.current[firstRowId][firstColName]
              ) {
                tableInputRefs.current[firstRowId][firstColName].focus();
                isHandled = true;
              }
            }
          } else if (e.key === "ArrowUp") {
            nextIndex =
              (headerIndex - 1 + headerInputOrder.length) %
              headerInputOrder.length;
            if (headerInputRefs.current[headerInputOrder[nextIndex]]) {
              headerInputRefs.current[headerInputOrder[nextIndex]].focus();
              isHandled = true;
            }
          }
          if (
            !isHandled &&
            headerInputRefs.current[headerInputOrder[nextIndex]]
          ) {
            headerInputRefs.current[headerInputOrder[nextIndex]].focus();
            isHandled = true;
          }
        } else if (tablePosition) {
          const { rowId, colName } = tablePosition;
          const currentRowIndex = orderItemsRef.current.findIndex(
            (item) => item.id === rowId
          );
          const currentColIndex = tableColumnOrder.indexOf(colName);

          let targetRowIndex = currentRowIndex;
          let targetColIndex = currentColIndex;

          if (e.key === "ArrowRight") {
            targetColIndex++;
            if (targetColIndex >= tableColumnOrder.length) {
              targetColIndex = 0;
              targetRowIndex++;
            }
          } else if (e.key === "ArrowLeft") {
            targetColIndex--;
            if (targetColIndex < 0) {
              targetColIndex = tableColumnOrder.length - 1;
              targetRowIndex--;
            }
          }
          else if (e.key === "ArrowDown") {
            targetRowIndex++;
          } else if (e.key === "ArrowUp") {
            targetRowIndex--;
          }

          let nextElementToFocus = null;

          if (targetRowIndex >= 0 && targetRowIndex < orderItemsRef.current.length) {
            const targetRowId = orderItemsRef.current[targetRowIndex]?.id;
            const targetColName = tableColumnOrder[targetColIndex];
            if (
              tableInputRefs.current[targetRowId] &&
              tableInputRefs.current[targetRowId][targetColName]
            ) {
              nextElementToFocus =
                tableInputRefs.current[targetRowId][targetColName];
            }
          } else if (targetRowIndex < 0) {
            const lastHeaderInputName =
              headerInputOrder[headerInputOrder.length - 1];
            nextElementToFocus = headerInputRefs.current[lastHeaderInputName];
          } else if (targetRowIndex >= orderItemsRef.current.length) {
            const lastValidRowIndex = orderItemsRef.current.length - 1;
            const lastValidRowId = orderItemsRef.current[lastValidRowIndex]?.id;
            const targetColName = tableColumnOrder[currentColIndex];
            if (
              lastValidRowId &&
              tableInputRefs.current[lastValidRowId] &&
              tableInputRefs.current[lastValidRowId][targetColName]
            ) {
              nextElementToFocus =
                tableInputRefs.current[lastValidRowId][targetColName];
            }
          }

          if (nextElementToFocus) {
            nextElementToFocus.focus();
            isHandled = true;
          }
        }
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, []);

  const currentOrderIsDeleted = headerInfo.status === "deleted";

  // ── v38 CLS FIX: Instead of swapping the entire DOM (which caused a 0.25 CLS
  // score), we always render the same layout shell and replace only the inner
  // content area with a skeleton while loading. This keeps the page geometry
  // stable and eliminates the largest layout shift.
  if (isLoading) {
    return (
      <div className="min-h-screen bg-gray-100 p-4 sm:p-6 lg:p-8 font-inter">
        <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-lg p-4 sm:p-6 space-y-6 relative">
          {/* Same header shell as real UI */}
          <div className="absolute top-4 right-4 z-10 flex items-center gap-2">
            <div className="w-20 sm:w-24 md:w-28 h-9 bg-gray-200 rounded-md animate-pulse" />
          </div>
          <div className="">
            <div className="h-7 w-64 bg-gray-200 rounded animate-pulse mx-auto mb-4" />
            <div className="flex items-end gap-2 mb-4">
              <div className="flex-grow h-9 bg-gray-100 rounded-md border border-gray-200 animate-pulse" />
              <div className="w-20 h-9 bg-blue-200 rounded-md animate-pulse" />
              <div className="w-28 h-9 bg-gray-200 rounded-md animate-pulse" />
            </div>
            {/* Skeleton form card */}
            <div className="border border-gray-200 rounded-xl p-4 space-y-3 animate-pulse">
              <div className="grid grid-cols-2 gap-3">
                {[...Array(6)].map((_, i) => (
                  <div key={i} className="space-y-1">
                    <div className="h-3 w-24 bg-gray-200 rounded" />
                    <div className="h-8 bg-gray-100 rounded border border-gray-200" />
                  </div>
                ))}
              </div>
              <div className="h-px bg-gray-100 my-2" />
              {/* Skeleton table */}
              <div className="space-y-2">
                <div className="h-8 bg-blue-100 rounded" />
                {[...Array(3)].map((_, i) => (
                  <div key={i} className="h-7 bg-gray-50 rounded border border-gray-100" />
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div
      className="min-h-screen p-4 sm:p-6 lg:p-8 font-inter transition-colors duration-200"
      style={{ backgroundColor: darkMode ? "#111827" : "#f3f4f6" }}
    >
      <div
        className="max-w-4xl mx-auto rounded-xl shadow-lg p-4 sm:p-6 space-y-6 relative transition-colors duration-200"
        style={{
          backgroundColor: darkMode ? "#1f2937" : "#ffffff",
          borderColor: darkMode ? "#374151" : "transparent",
          border: darkMode ? "1px solid #374151" : undefined,
        }}
      >

        <div className="absolute top-4 right-4 z-10 flex items-center gap-2">

          {!committedSearchTerm && (
            <button
              onClick={() => {
                if (window.confirm("¿Iniciar un nuevo mail? Los pedidos actuales serán eliminados.")) {
                  handleNewMail();
                }
              }}
              className="p-1.5 rounded-full border shadow-sm transition-colors"
              style={{ backgroundColor: darkMode ? "#374151" : "#ffffff", borderColor: darkMode ? "#4b5563" : "#e5e7eb" }}
              title="Nuevo mail — descarta drafts actuales"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
            </button>
          )}
          <button
            onClick={() => setShowConfigPanel(true)}
            className="p-1.5 rounded-full border shadow-sm transition-colors" style={{ backgroundColor: darkMode ? "#374151" : "#ffffff", borderColor: darkMode ? "#4b5563" : "#e5e7eb" }}
            title="Configuración"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
              <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
          </button>

          {sentHistory.length > 0 && (
            <button
              onClick={() => setShowHistoryPanel(true)}
              className="relative p-1.5 rounded-full border shadow-sm transition-colors" style={{ backgroundColor: darkMode ? "#374151" : "#ffffff", borderColor: darkMode ? "#4b5563" : "#e5e7eb" }}
              title="Ver historial de envíos"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-gray-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <span className="absolute -top-1 -right-1 bg-blue-600 text-white text-xs rounded-full w-4 h-4 flex items-center justify-center font-bold">
                {sentHistory.length}
              </span>
            </button>
          )}

          <img
            src="https://www.vpcom.com/images/logo-vpc.png"
            onError={(e) => {
              e.target.onerror = null;
              e.target.src =
                "https://placehold.co/100x40/FFFFFF/000000?text=LogoVPC";
            }}
            alt="Logo VPC"
            width="112"
            height="45"
            className="w-20 sm:w-24 md:w-28 h-auto object-contain rounded-md"
          />
        </div>

        {userId && (
          <div className="absolute top-4 left-4 text-xs text-gray-500 flex items-center gap-2">
            <span title={`ID: ${userId}`}>
              🙍 {userId.substring(0, 8).toUpperCase()}
            </span>
            <button
              onClick={() => {
                if (window.confirm("¿Resetear tu identidad de sesión? Esto creará un nuevo ID y perderás acceso a tus borradores actuales.")) {
                  window.location.reload();
                }
              }}
              className="text-gray-400 hover:text-red-500 transition-colors"
              title="Resetear identidad de sesión"
            >
              ↺
            </button>
          </div>
        )}

        <div className="">
          <h1 className="text-xl sm:text-2xl font-bold text-center mb-4" style={{ color: darkMode ? "#f9fafb" : "#1f2937" }}>
            Pedidos Comercial Frutam
          </h1>
          <div className="flex items-end gap-2 mb-4">
            <div className="flex-grow">
              <label
                htmlFor="searchTerm"
                className="block text-sm font-medium" style={{ color: darkMode ? "#d1d5db" : "#374151" }}
              >
                Buscar por Mail ID:
              </label>
              <input
                type="text"
                id="searchTerm"
                name="searchTerm"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Ingrese el Mail ID (ej. 5B7D7692)"
                className="mt-1 block w-full rounded-md shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 border transition-colors" style={{ backgroundColor: darkMode ? "#374151" : "#f9fafb", borderColor: darkMode ? "#4b5563" : "#d1d5db", color: darkMode ? "#f9fafb" : "#111827" }}
              />
            </div>
            <button
              onClick={handleSearchClick}
              className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-md shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition duration-150 ease-in-out mt-auto"
            >
              Buscar
            </button>
            <button
              onClick={handleClearSearch}
              className="px-4 py-2 bg-gray-300 text-gray-800 font-semibold rounded-md shadow-md hover:bg-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 transition duration-150 ease-in-out mt-auto"
            >
              Limpiar Búsqueda
            </button>

          </div>

          {otherUsersPresent.length > 0 && committedSearchTerm && (
            <div className="flex items-center gap-2 px-4 py-2 mb-2 bg-yellow-50 border border-yellow-300 rounded-lg text-yellow-800 text-sm">
              <span className="text-lg">⚠️</span>
              <span>
                <strong>
                  {otherUsersPresent.length === 1
                    ? otherUsersPresent[0]
                    : otherUsersPresent.join(", ")}
                </strong>{" "}
                {otherUsersPresent.length === 1 ? "también está" : "también están"} editando este pedido ahora mismo. Los cambios de ambos podrían solaparse.
              </span>
            </div>
          )}

          {/* ── SUGERENCIA DE AUTOCOMPLETADO ── */}
          {proveedorSuggestion && (
            <div className="flex items-center gap-2 px-4 py-2 mb-2 rounded-lg text-sm"
              style={{
                background: darkMode ? "#1e3a5f" : "#eff6ff",
                border: `1px solid ${darkMode ? "#3b82f6" : "#93c5fd"}`,
                color: darkMode ? "#bfdbfe" : "#1e40af",
              }}
            >
              <span className="text-lg">💡</span>
              <span style={{ flex: 1 }}>
                <strong>Último pedido a {proveedorSuggestion.proveedor}:</strong>
                {" "}{[...new Set(proveedorSuggestion.items.map(i => i.especie).filter(Boolean))].join(", ")}
                {proveedorSuggestion.nave ? ` · ${proveedorSuggestion.nave}` : ""}
                {proveedorSuggestion.deNombrePais ? ` · ${proveedorSuggestion.deNombrePais}` : ""}
              </span>
              <button
                onClick={applyProveedorSuggestion}
                className="px-3 py-1 rounded-md text-xs font-bold text-white"
                style={{ background: "#3b82f6", border: "none", cursor: "pointer", whiteSpace: "nowrap" }}
              >
                Precargar
              </button>
              <button
                onClick={dismissProveedorSuggestion}
                className="px-2 py-1 rounded-md text-xs font-semibold"
                style={{ background: "transparent", border: `1px solid ${darkMode ? "#3b82f6" : "#93c5fd"}`, color: darkMode ? "#93c5fd" : "#1d4ed8", cursor: "pointer" }}
              >
                ✕
              </button>
            </div>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mt-4">
            <div>
              <div className="flex items-center justify-between mb-1">
                <label className="block text-sm font-medium" style={{ color: darkMode ? "#d1d5db" : "#374151" }}>
                  Nombre de Proveedor:
                </label>
                <div className={`flex rounded border shadow-sm overflow-hidden
                  ${currentOrderIsDeleted ? "opacity-50 pointer-events-none" : ""}
                  border-gray-300`}
                >
                  {["FOB", "FOT"].map((term, idx) => (
                    <button
                      key={term}
                      type="button"
                      onClick={() => {
                        if (!currentOrderIsDeleted) {
                          setHeaderInfo((prev) => ({ ...prev, incoterm: term }));
                        }
                      }}
                      className={`px-2 py-0.5 text-xs font-medium transition-colors duration-100 focus:outline-none
                        ${idx === 0 ? "" : "border-l border-gray-300"}
                        ${
                          headerInfo.incoterm === term
                            ? "bg-blue-50 text-blue-700 font-semibold"
                            : "bg-gray-50 text-gray-400 hover:bg-gray-100 hover:text-gray-600"
                        }
                      `}
                      disabled={currentOrderIsDeleted}
                    >
                      {term}
                    </button>
                  ))}
                </div>
              </div>
              <input
                type="text"
                name="reDestinatarios"
                value={headerInfo.reDestinatarios === null || headerInfo.reDestinatarios === undefined ? "" : String(headerInfo.reDestinatarios)}
                onFocus={handleFocusField}
                onBlur={handleBlurField(handleHeaderBlur)}
                onChange={(e) => {
                  handleHeaderChange(e);
                  const val = e.target.value;
                  if (lookupDebounceRef.current) clearTimeout(lookupDebounceRef.current);
                  lookupDebounceRef.current = setTimeout(() => {
                    lookupLastOrderForProveedor(val);
                  }, 600);
                }}
                placeholder="Ingrese nombre de proveedor"
                className={`mt-0 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 border transition-colors ${
                  currentOrderIsDeleted || (currentOrderIndex > 0 && !!headerInfo.reDestinatarios)
                    ? "bg-gray-100 text-gray-500 cursor-not-allowed"
                    : "bg-gray-50"
                }`}
                ref={(el) => (headerInputRefs.current.reDestinatarios = el)}
                readOnly={currentOrderIsDeleted || (currentOrderIndex > 0 && !!headerInfo.reDestinatarios)}
              />
            </div>

            <div>
              <InputField
                label="País:"
                  darkMode={darkMode}
                name="deNombrePais"
                value={headerInfo.deNombrePais}
                onChange={handleHeaderChange}
                onFocus={handleFocusField}
                onBlur={handleBlurField(handleHeaderBlur)}
                placeholder="País de destino"
                ref={(el) => (headerInputRefs.current.deDestinatarios = el)}
                readOnly={currentOrderIsDeleted}
              />
            </div>
            <div>
              <InputField
                label="Nave:"
                  darkMode={darkMode}
                name="nave"
                value={headerInfo.nave}
                onChange={handleHeaderChange}
                onFocus={handleFocusField}
                onBlur={handleBlurField(handleHeaderBlur)}
                placeholder="Nombre de Nave"
                ref={(el) => (headerInputRefs.current.nave = el)}
                readOnly={currentOrderIsDeleted}
              />
            </div>
            <div>
              <InputField
                label="Fecha de carga:"
                  darkMode={darkMode}
                name="fechaCarga"
                value={headerInfo.fechaCarga}
                onChange={handleHeaderChange}
                onFocus={handleFocusField}
                onBlur={handleBlurField(handleHeaderBlur)}
                placeholder="FECHA DE CARGA"
                type="date"
                ref={(el) => (headerInputRefs.current.fechaCarga = el)}
                readOnly={currentOrderIsDeleted}
              />
            </div>
            <div>
              <InputField
                label="Exporta:"
                  darkMode={darkMode}
                name="exporta"
                value={headerInfo.exporta}
                onChange={handleHeaderChange}
                onFocus={handleFocusField}
                onBlur={handleBlurField(handleHeaderBlur)}
                placeholder="Exportadora"
                ref={(el) => (headerInputRefs.current.exporta = el)}
                readOnly={currentOrderIsDeleted}
              />
            </div>
            <div>
              <InputField
                label="Asunto del Email:"
                  darkMode={darkMode}
                name="emailSubject"
                value={headerInfo.emailSubject}
                onChange={handleHeaderChange}
                onBlur={handleHeaderBlur}
                placeholder="Asunto del Correo (Se auto-completa)"
                ref={(el) => (headerInputRefs.current.emailSubject = el)}
                readOnly={true}
              />
            </div>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row justify-center items-center gap-4 mt-4 mb-6">
          <div className="flex items-center justify-center w-full sm:w-auto">
            <button
              onClick={handlePreviousOrder}
              disabled={currentOrderIndex === 0}
              className="flex items-center justify-center px-3 py-2 bg-gray-300 text-gray-800 font-semibold rounded-lg shadow-md hover:bg-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 transition duration-150 ease-in-out disabled:opacity-50 disabled:cursor-not-allowed"
              title="Ir al pedido anterior (más antiguo)"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M9.707 4.293a1 1 0 010 1.414L5.414 10l4.293 4.293a1 1 0 01-1.414 1.414l-5-5a1 1 0 010-1.414l5-5a1 1 0 011.414 0z" clipRule="evenodd" />
              </svg>
              <span className="hidden sm:inline ml-1">Anterior</span>
            </button>

            <span className="text-center font-semibold text-lg mx-2 sm:mx-4 min-w-[150px] sm:min-w-0" style={{ color: darkMode ? "#d1d5db" : "#374151" }}>
              {`Pedido ${currentOrderIndex + 1} de ${displayedOrders.length}`}
              {currentOrderIsDeleted && (
                <span className="text-red-500 ml-2">(Eliminado)</span>
              )}
            </span>

            <button
              onClick={handleNextOrder}
              disabled={currentOrderIndex === displayedOrders.length - 1}
              className="flex items-center justify-center px-3 py-2 bg-gray-300 text-gray-800 font-semibold rounded-lg shadow-md hover:bg-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 transition duration-150 ease-in-out disabled:opacity-50 disabled:cursor-not-allowed"
              title="Ir al siguiente pedido (más reciente)"
            >
              <span className="hidden sm:inline mr-1">Siguiente</span>
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M10.293 15.707a1 1 0 010-1.414L14.586 10l-4.293-4.293a1 1 0 111.414-1.414l5 5a1 1 0 010 1.414l-5 5a1 1 0 01-1.414 0z" clipRule="evenodd" />
              </svg>
            </button>
          </div>
        </div>

        {displayedOrders.length === 0 && committedSearchTerm ? (
          <div className="text-center text-gray-600 text-lg my-8">
            No se encontraron pedidos con el ID:{" "}
            <span className="font-semibold">{committedSearchTerm}</span>.
          </div>
        ) : (
          <>
            <div className="overflow-x-auto rounded-lg shadow-md">
              <table className="min-w-full divide-y divide-gray-200 hidden md:table">
                <thead className="bg-blue-600 text-white">
                  <tr style={{ backgroundColor: "#2563eb", color: "#ffffff" }}>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider rounded-tl-lg whitespace-nowrap" title="Arrastrar para reordenar">⠿</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Pallets</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Especie</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Variedad</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Formato</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Calibre</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider rounded-tr-lg whitespace-nowrap">Categoría</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Precios {headerInfo.incoterm || "FOB"}</th>
                    <th scope="col" className="px-1 py-px text-xs font-medium uppercase tracking-wider whitespace-nowrap">Acciones</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {orderItems.map((item, index) => (
                    <tr
                      key={item.id}
                      draggable={!currentOrderIsDeleted}
                      onDragStart={() => { dragItemIndex.current = index; }}
                      onDragEnter={() => { dragOverIndex.current = index; }}
                      onDragOver={(e) => e.preventDefault()}
                      onDragEnd={() => {
                        if (dragItemIndex.current !== null && dragOverIndex.current !== null) {
                          handleDragReorder(dragItemIndex.current, dragOverIndex.current);
                        }
                        dragItemIndex.current = null;
                        dragOverIndex.current = null;
                      }}
                      className={`hover:bg-gray-50 ${
                        index % 2 === 0 ? "bg-gray-50" : "bg-white"
                      } ${
                        item.isCanceled || currentOrderIsDeleted
                          ? "text-red-500"
                          : ""
                      } ${!currentOrderIsDeleted ? "cursor-grab active:cursor-grabbing" : ""}`}
                      style={{ userSelect: "none" }}
                    >
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap text-center text-gray-300 select-none" title="Arrastrar para reordenar" style={{ cursor: "grab", fontSize: "16px", width: "20px" }}>
                        ⠿
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          type="number"
                          name="pallets"
                          min="0"
                          value={item.pallets}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="21"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].pallets = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          name="especie"
                          value={item.especie}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="Manzana"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          {...(item.especie && item.especie.toUpperCase() === "MANZANAS" ? { list: "apple-varieties" } : {})}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].especie = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          name="variedad"
                          value={item.variedad}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="Galas"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          {...(item.especie && item.especie.toUpperCase() === "MANZANAS" ? { list: "apple-varieties" } : {})}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].variedad = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          name="formato"
                          value={item.formato}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="20 Kg"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].formato = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          name="calibre"
                          value={item.calibre}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="100;113"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].calibre = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          name="categoria"
                          value={item.categoria}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="PRE:XFY"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].categoria = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center", borderColor: darkMode ? "#374151" : "#e5e7eb" }}>
                        <TableInput
                          name="preciosFOB"
                          value={item.preciosFOB}
                          onChange={(e) => handleItemChange(item.id, e)}
                          onFocus={handleFocusField}
                          onBlur={handleBlurField((e) => handleItemBlur(item.id, e))}
                          placeholder="$14"
                          readOnly={item.isCanceled || currentOrderIsDeleted}
                          isCanceledProp={item.isCanceled || currentOrderIsDeleted}
                          ref={(el) => {
                            if (!tableInputRefs.current[item.id])
                              tableInputRefs.current[item.id] = {};
                            tableInputRefs.current[item.id].preciosFOB = el;
                          }}
                        />
                      </td>
                      <td className="px-1 py-px text-right text-xs font-medium flex items-center justify-center h-full">
                        <button
                          onClick={() => handleOpenObservationModal(item.id)}
                          className="text-blue-600 hover:text-blue-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 p-1 rounded-md"
                          title="Editar Observación"
                          disabled={currentOrderIsDeleted}
                        >
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
                            <path strokeLinecap="round" strokeLinejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                          </svg>
                        </button>
                        <button
                          onClick={() => handleAddItem(item.id)}
                          className="text-green-600 hover:text-green-900 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 p-1 rounded-md"
                          title="Duplicar fila"
                          disabled={currentOrderIsDeleted}
                        >
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                            <path fillRule="evenodd" d="M10 5a1 1 0 011 1v3h3a1 1 0 110 2h-3v3a1 1 0 11-2 0v-3H6a1 1 0 110-2h3V6a1 1 0 011-1z" clipRule="evenodd" />
                          </svg>
                        </button>
                        <button
                          onClick={() => toggleItemCancellation(item.id)}
                          className={`ml-1 ${item.isCanceled ? "text-gray-600 hover:text-gray-900" : "text-red-600 hover:text-red-900"} focus:outline-none focus:ring-2 focus:ring-offset-2 p-1 rounded-md`}
                          title={item.isCanceled ? "Revertir cancelación" : "Cancelar fila"}
                          disabled={currentOrderIsDeleted}
                        >
                          {item.isCanceled ? (
                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm.707-10.293a1 1 0 00-1.414-1.414L7.586 9H6a1 1 0 000 2h1.586l1.707 1.707a1 1 0 001.414-1.414L10.414 10H12a1 1 0 000-2h-1.586z" clipRule="evenodd" />
                            </svg>
                          ) : (
                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                              <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
                            </svg>
                          )}
                        </button>
                        <button
                          onClick={() => handleDeleteItem(item.id)}
                          className={`ml-1 text-red-600 hover:text-red-900 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-offset-2 p-1 rounded-md ${orderItems.length <= 1 || currentOrderIsDeleted ? "opacity-50 cursor-not-allowed" : ""}`}
                          title={orderItems.length <= 1 ? "No se puede eliminar la última fila" : "Eliminar fila"}
                          disabled={orderItems.length <= 1 || currentOrderIsDeleted}
                        >
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                            <path fillRule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 011-1h4a1 1 0 110 2H8a1 1 0 01-1-1zm2 3a1 1 0 011-1h4a1 1 0 110 2H8a1 1 0 01-1-1zm-1 4a1 1 0 002 0v-4a1 1 0 00-2 0v4z" />
                          </svg>
                        </button>
                      </td>
                    </tr>
                  ))}
                  <tr style={{ backgroundColor: "#e0e0e0" }}>
                    <td colSpan="8" style={{ padding: "6px 15px 6px 6px", textAlign: "right", fontWeight: "bold", border: "1px solid #ccc", borderBottomLeftRadius: "8px", marginTop: "15px" }}>
                      Total de Pallets:
                    </td>
                    <td colSpan="1" style={{ padding: "6px", fontWeight: "bold", border: "1px solid #ccc", borderBottomRightRadius: "8px", textAlign: "center" }}>
                      {currentOrderTotalPallets} Pallets
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Mobile View - Cards */}
            <div className="md:hidden space-y-4 p-2">
              {orderItems.map((item, index) => (
                <div
                  key={item.id}
                  className={`bg-white border border-gray-200 rounded-lg p-4 shadow-sm space-y-2 ${
                    item.isCanceled || currentOrderIsDeleted
                      ? "line-through text-red-500 opacity-70"
                      : ""
                  }`}
                >
                  <div className="flex justify-between items-center pb-2 border-b border-gray-100">
                    <div className="flex items-center gap-2">
                      <span
                        className="text-gray-300 text-lg select-none"
                        title="Arrastrar para reordenar"
                        style={{ cursor: "grab", lineHeight: 1 }}
                        draggable={!currentOrderIsDeleted}
                        onDragStart={() => { dragItemIndex.current = index; }}
                        onDragEnter={() => { dragOverIndex.current = index; }}
                        onDragOver={(e) => e.preventDefault()}
                        onDragEnd={() => {
                          if (dragItemIndex.current !== null && dragOverIndex.current !== null) {
                            handleDragReorder(dragItemIndex.current, dragOverIndex.current);
                          }
                          dragItemIndex.current = null;
                          dragOverIndex.current = null;
                        }}
                      >⠿</span>
                      <span className="text-xs font-semibold text-blue-600">
                        Artículo #{index + 1}
                      </span>
                    </div>
                    <div className="flex space-x-2">
                      <button onClick={() => handleOpenObservationModal(item.id)} className="text-blue-600 hover:text-blue-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 p-1 rounded-md" title="Editar Observación" disabled={currentOrderIsDeleted}>
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" className="h-5 w-5">
                          <path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z" />
                        </svg>
                      </button>
                      <button onClick={() => handleAddItem(item.id)} className="text-green-600 hover:text-green-900 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 p-1 rounded-md" title="Duplicar fila" disabled={currentOrderIsDeleted}>
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                          <path fillRule="evenodd" d="M10 5a1 1 0 011 1v3h3a1 1 0 110 2h-3v3a1 1 0 11-2 0v-3H6a1 1 0 110-2h3V6a1 1 0 011-1z" />
                        </svg>
                      </button>
                      <button onClick={() => toggleItemCancellation(item.id)} className={`ml-1 ${item.isCanceled ? "text-gray-600 hover:text-gray-900" : "text-red-600 hover:text-red-900"} focus:outline-none focus:ring-2 focus:ring-offset-2 p-1 rounded-md`} title={item.isCanceled ? "Revertir cancelación" : "Cancelar fila"} disabled={currentOrderIsDeleted}>
                        {item.isCanceled ? (
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm.707-10.293a1 1 0 00-1.414-1.414L7.586 9H6a1 1 0 000 2h1.586l1.707 1.707a1 1 0 001.414-1.414L10.414 10H12a1 1 0 000-2h-1.586z" clipRule="evenodd" /></svg>
                        ) : (
                          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" /></svg>
                        )}
                      </button>
                      <button onClick={() => handleDeleteItem(item.id)} className={`ml-1 text-red-600 hover:text-red-900 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-offset-2 p-1 rounded-md ${orderItems.length <= 1 || currentOrderIsDeleted ? "opacity-50 cursor-not-allowed" : ""}`} title={orderItems.length <= 1 ? "No se puede eliminar la última fila" : "Eliminar fila"} disabled={orderItems.length <= 1 || currentOrderIsDeleted}>
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 011-1h4a1 1 0 110 2H8a1 1 0 01-1-1zm2 3a1 1 0 011-1h4a1 1 0 110 2H8a1 1 0 01-1-1zm-1 4a1 1 0 002 0v-4a1 1 0 00-2 0v4z" /></svg>
                      </button>
                    </div>
                  </div>

                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Pallets:</span>
                    <TableInput name="pallets" value={item.pallets} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="21" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} />
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Especie:</span>
                    <TableInput name="especie" value={item.especie} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="Manzana" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} {...(item.especie && item.especie.toUpperCase() === "MANZANAS" ? { list: "apple-varieties" } : {})} />
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Variedad:</span>
                    <TableInput name="variedad" value={item.variedad} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="Galas" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} {...(item.especie && item.especie.toUpperCase() === "MANZANAS" ? { list: "apple-varieties" } : {})} />
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Formato:</span>
                    <TableInput name="formato" value={item.formato} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="20 Kg" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} />
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Calibre:</span>
                    <TableInput name="calibre" value={item.calibre} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="100;113" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} />
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Categoría:</span>
                    <TableInput name="categoria" value={item.categoria} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="PRE:XFY" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} />
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="font-semibold text-gray-500 w-1/2">Precios {headerInfo.incoterm || "FOB"}:</span>
                    <TableInput name="preciosFOB" value={item.preciosFOB} onChange={(e) => handleItemChange(item.id, e)} onFocus={handleFocusField} onBlur={handleBlurField((e) => handleItemBlur(item.id, e))} placeholder="$14" readOnly={item.isCanceled || currentOrderIsDeleted} isCanceledProp={item.isCanceled || currentOrderIsDeleted} />
                  </div>
                </div>
              ))}
              <div className="bg-blue-100 border border-blue-200 rounded-lg py-2 px-3 shadow-sm text-center font-bold text-base text-blue-800 mt-4">
                Total de Pallets: {currentOrderTotalPallets} Pallets
              </div>
            </div>
          </>
        )}

        {/* Action Buttons */}
        <div className="flex flex-col sm:flex-row justify-center gap-4 mt-6">
          <button
            onClick={handleAddOrder}
            className="flex items-center justify-center px-6 py-3 bg-indigo-600 text-white font-semibold rounded-lg shadow-md hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto"
            title="Crear un nuevo pedido en blanco"
            disabled={currentOrderIsDeleted}
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm1-11a1 1 0 10-2 0v2H7a1 1 0 100 2h2v2a1 1 0 102 0v-2h2a1 1 0 100-2h-2V7z" clipRule="evenodd" />
            </svg>
            Agregar Pedido
          </button>

          <button
            onClick={handleDeleteCurrentOrder}
            disabled={displayedOrders.length <= 1 || currentOrderIsDeleted}
            className={`flex items-center justify-center px-6 py-3 bg-red-600 text-white font-semibold rounded-lg shadow-md hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto ${displayedOrders.length <= 1 || currentOrderIsDeleted ? "opacity-50 cursor-not-allowed" : ""}`}
            title="Eliminar el pedido actual"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 011-1h4a1 1 0 110 2H8a1 1 0 01-1-1zm2 3a1 1 0 011-1h4a1 1 0 110 2H8a1 1 0 01-1-1zm-1 4a1 1 0 002 0v-4a1 1 0 00-2 0v4z" />
            </svg>
            Eliminar Pedido
          </button>

          <button
            onClick={handleFinalizeOrder}
            className="flex items-center justify-center px-6 py-3 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto"
            title="Finalizar el pedido y ver opciones de envío"
            disabled={currentOrderIsDeleted}
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
            </svg>
            Finalizar Pedido
          </button>
        </div>

        {/* Unified Order Actions Modal */}
        {showOrderActionsModal && (
          <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center p-4 z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-screen-lg mx-auto my-8 relative flex flex-col max-h-[90vh]">
              <h2 className="text-xl font-bold mb-4 text-gray-800 text-center">
                Opciones de Pedido Finalizado
              </h2>

              {!isShowingPreview && (
                <h3 className="text-sm text-gray-600 text-center mb-4 leading-relaxed">
                  <strong>Instrucción importante:</strong> Después de enviar el
                  email, abre tu aplicación de correo y pega (Ctrl+V o Cmd+V) el
                  contenido manualmente en el cuerpo del mensaje.
                </h3>
              )}

              {!isShowingPreview ? (
                <div className="flex justify-center gap-4 mb-4 flex-wrap sm:flex-nowrap">
                  <button
                    onClick={handlePreviewOrder}
                    className="flex items-center justify-center px-4 py-2 text-sm bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700 focus:outline-none focus:ring-2 focus:ring-gray-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto"
                    title="Previsualizar el pedido completo"
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
                      <path fillRule="evenodd" d="M8 4a4 4 0 100 8c1.65 0 3-1.35 3-3V7a1 1 0 112 0v1a5 5 0 01-5 5H4a5 5 0 01-5-5v-1c0-1.65 1.35-3 3-3h1V4a1 1 0 11-2 0V3h-1a1 1 0 110-2h1a1 1 0 011 1v1h1a1 1 0 011 1V4zm0 2a2 2 0 100 4 2 2 0 000-4z" clipRule="evenodd" />
                    </svg>
                    Previsualizar
                  </button>
                  <button
                    onClick={performSendEmail}
                    className="flex items-center justify-center px-4 py-2 text-sm bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto"
                    title="Copiar contenido y abrir cliente de correo"
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
                      <path d="M2.003 5.884L10 9.882l7.997-3.998A2 2 0 0016 4H4a2 2 0 00-1.997 1.884z" />
                      <path d="M18 8.118l-8 4-8-4V14a2 2 0 002 2h12a2 2 0 002-2V8.118z" />
                    </svg>
                    Enviar Email
                  </button>
                </div>
              ) : (
                <>
                  <h3 className="text-lg font-semibold mb-3 text-gray-800 text-center">
                    Previsualización del Pedido
                  </h3>
                  <div className="bg-gray-50 rounded-md border border-gray-200 flex-grow overflow-hidden" style={{minHeight: "300px"}}>
                    <iframe
                      srcDoc={previewHtmlContent}
                      style={{width:"100%", height:"100%", minHeight:"360px", border:"none", display:"block"}}
                      title="Previsualización del pedido"
                    />
                  </div>
                  <div className="flex justify-center mt-3 gap-2 flex-wrap sm:flex-nowrap">
                    <button
                      onClick={() => setIsShowingPreview(false)}
                      className="flex items-center justify-center px-4 py-2 text-sm bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700 focus:outline-none focus:ring-2 focus:ring-gray-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto"
                    >
                      Volver a Opciones
                    </button>
                    <button
                      onClick={performSendEmail}
                      className="flex items-center justify-center px-4 py-2 text-sm bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition duration-150 ease-in-out transform hover:scale-105 w-full sm:w-auto"
                      title="Copiar contenido y abrir cliente de correo"
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-1" viewBox="0 0 20 20" fill="currentColor">
                        <path d="M2.003 5.884L10 9.882l7.997-3.998A2 2 0 0016 4H4a2 2 0 00-1.997 1.884z" />
                        <path d="M18 8.118l-8 4-8-4V14a2 2 0 002 2h12a2 2 0 002-2V8.118z" />
                      </svg>
                      Enviar Email
                    </button>
                  </div>
                </>
              )}

              <div className="flex justify-center mt-4">
                <button
                  onClick={() => setShowOrderActionsModal(false)}
                  className="px-6 py-3 bg-red-500 text-white font-semibold rounded-lg shadow-md hover:bg-red-600 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-offset-2 transition duration-150 ease-in-out"
                >
                  Cerrar
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Observation Modal */}
        {showObservationModal && currentEditingItemData && (
          <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center p-4 z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md mx-auto relative">
              <h2 className="text-xl font-bold mb-4 text-gray-800 text-center">
                Editar Observación para Línea
              </h2>
              <div className="mb-4">
                <label htmlFor="modalObservation" className="block text-sm font-medium text-gray-700 mb-1">
                  Observación:
                </label>
                <textarea
                  id="modalObservation"
                  className="mt-1 block w-full rounded-md shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 border transition-colors" style={{ backgroundColor: darkMode ? "#374151" : "#f9fafb", borderColor: darkMode ? "#4b5563" : "#d1d5db", color: darkMode ? "#f9fafb" : "#111827" }}
                  rows="4"
                  value={modalObservationText}
                  onChange={(e) => setModalObservationText(e.target.value)}
                  placeholder="Ingrese la observación aquí..."
                  ref={observationTextareaRef}
                ></textarea>
              </div>
              <div className="flex justify-end gap-3">
                <button onClick={handleCloseObservationModal} className="px-4 py-2 bg-gray-300 text-gray-800 font-semibold rounded-lg shadow-md hover:bg-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 transition duration-150 ease-in-out">
                  Cancelar
                </button>
                <button onClick={handleSaveObservation} className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition duration-150 ease-in-out">
                  Guardar
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Recovery Modal */}

        {/* Presence Conflict Modal */}
        {showPresenceModal && (
          <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center p-4 z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md mx-auto relative">
              <div className="flex items-center gap-3 mb-4">
                <span className="text-3xl">⚠️</span>
                <h2 className="text-xl font-bold text-gray-800">Pedido en uso</h2>
              </div>
              <p className="text-gray-600 mb-2">
                El pedido <span className="font-bold text-blue-600">{pendingSearchTerm}</span> está siendo editado por:
              </p>
              <ul className="mb-4 space-y-1">
                {presenceConflictUsers.map((u) => (
                  <li key={u} className="flex items-center gap-2 text-gray-700 font-medium">
                    <span className="inline-block w-2 h-2 rounded-full bg-yellow-400"></span>
                    {u}
                  </li>
                ))}
              </ul>
              <p className="text-sm text-gray-500 mb-6">
                Si lo abres igual, los cambios de ambos usuarios podrían solaparse y perderse información.
              </p>
              <div className="flex justify-end gap-3">
                <button onClick={() => { setShowPresenceModal(false); setPendingSearchTerm(""); setPresenceConflictUsers([]); }} className="px-4 py-2 bg-gray-200 text-gray-800 font-semibold rounded-lg hover:bg-gray-300 transition">
                  Cancelar
                </button>
                <button
                  onClick={async () => {
                    setShowPresenceModal(false);
                    setPresenceConflictUsers([]);
                    await enterMailId(pendingSearchTerm);
                    setPendingSearchTerm("");
                  }}
                  className="px-4 py-2 bg-yellow-500 text-white font-semibold rounded-lg hover:bg-yellow-600 transition"
                >
                  Abrir de todas formas
                </button>
              </div>
            </div>
          </div>
        )}

        <datalist id="apple-varieties">
          <option value="GALA" />
          <option value="GRANNY" />
          <option value="FUJI" />
          <option value="PINK LADY" />
          <option value="ROJA" />
          <option value="CRIPPS PINK" />
        </datalist>
      </div>

      {/* ── HISTORY PANEL (Portal) ────────────────────────────────────────────── */}
      {showHistoryPanel && ReactDOM.createPortal(
        <>
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.2)", zIndex: 9998 }} onClick={() => setShowHistoryPanel(false)} />
          <div style={{ position: "fixed", top: 0, right: 0, height: "100vh", width: "320px", background: "#fff", boxShadow: "-4px 0 24px rgba(0,0,0,0.12)", zIndex: 9999, display: "flex", flexDirection: "column" }}>
            <div className="flex items-center justify-between px-5 py-4 border-b border-gray-100">
              <div className="flex items-center gap-2">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-blue-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                <h2 className="text-base font-bold text-gray-800">Historial de envíos</h2>
              </div>
              <button onClick={() => setShowHistoryPanel(false)} className="text-gray-400 hover:text-gray-600 transition-colors p-1 rounded">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
                </svg>
              </button>
            </div>
            <div className="flex-1 overflow-y-auto px-4 py-3 space-y-3">
              {sentHistory.length === 0 ? (
                <p className="text-sm text-gray-400 text-center mt-8">No hay envíos recientes.</p>
              ) : (
                sentHistory.map((entry) => {
                  const proveedores = [...new Set(entry.orders.map(o => o.header?.reDestinatarios).filter(Boolean))];
                  const especies = [...new Set(entry.orders.flatMap(o => o.items.map(i => i.especie)).filter(Boolean))];
                  const subject = generateEmailSubjectValue(proveedores, especies, "");
                  return (
                    <div key={entry.mailId} className="border border-gray-100 rounded-lg p-4 bg-gray-50 space-y-2">
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <p className="text-xs text-gray-400 mb-0.5">Mail ID</p>
                          <p className="font-bold text-blue-600 text-sm tracking-wide">{entry.mailId}</p>
                        </div>
                        <span className="text-xs text-gray-400 whitespace-nowrap mt-1">{formatRelativeTime(entry.latestUpdatedAt)}</span>
                      </div>
                      <div>
                        <p className="text-xs text-gray-400 mb-0.5">Asunto</p>
                        <p className="text-xs text-gray-700 font-medium break-all leading-snug">{subject}</p>
                      </div>
                      <div className="text-xs text-gray-400">
                        {entry.orders.length} pedido{entry.orders.length > 1 ? "s" : ""}
                        {entry.orders[0]?.header?.reDestinatarios ? ` · ${entry.orders[0].header.reDestinatarios}` : ""}
                      </div>
                      <div className="flex gap-2 pt-1">
                        <button onClick={async () => { setShowHistoryPanel(false); setSearchTerm(entry.mailId); setCommittedSearchTerm(entry.mailId); await enterMailId(entry.mailId); }} className="flex-1 px-3 py-1.5 text-xs font-semibold bg-white border border-gray-200 text-gray-700 rounded-md hover:bg-gray-100 transition">
                          Ver pedido
                        </button>
                        <button onClick={() => resendFromHistory(entry.mailId)} disabled={historyResendingId === entry.mailId} className="flex-1 px-3 py-1.5 text-xs font-semibold bg-blue-600 text-white rounded-md hover:bg-blue-700 transition disabled:opacity-50 flex items-center justify-center gap-1">
                          {historyResendingId === entry.mailId ? (
                            <><svg className="animate-spin h-3 w-3" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/></svg> Copiando...</>
                          ) : (
                            <><svg xmlns="http://www.w3.org/2000/svg" className="h-3 w-3" viewBox="0 0 20 20" fill="currentColor"><path d="M8 3a1 1 0 011-1h2a1 1 0 110 2H9a1 1 0 01-1-1z"/><path d="M6 3a2 2 0 00-2 2v11a2 2 0 002 2h8a2 2 0 002-2V5a2 2 0 00-2-2 3 3 0 01-3 3H9a3 3 0 01-3-3z"/></svg> Re-enviar</>
                          )}
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
            <div className="px-4 py-3 border-t border-gray-100">
              <p className="text-xs text-gray-400 text-center">
                Últimos {sentHistory.length} envío{sentHistory.length !== 1 ? "s" : ""} · Solo los tuyos
              </p>
            </div>
          </div>
        </>,
        document.body
      )}

      {/* ── WHAT'S NEW MODAL (una sola vez por versión) ────────────────────────── */}
      {showWhatsNew && ReactDOM.createPortal(
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.5)" }}>
          <div className="rounded-2xl shadow-2xl w-full max-w-sm p-6" style={{ background: "#ffffff" }}>
            <div className="flex items-center gap-3 mb-4">
              <span style={{ fontSize: "28px" }}>🚀</span>
              <div>
                <h2 className="text-base font-bold text-gray-800">Novedades de esta versión</h2>
                <p className="text-xs text-gray-400">Frutam · v47</p>
              </div>
            </div>
            <ul className="space-y-3 mb-6">
              {[
                { icon: "💡", title: "Precarga de datos", desc: "Al escribir un proveedor la app sugiere automáticamente los datos del último pedido enviado, incluyendo tabla completa, pallets y precios." },
                { icon: "🌙", title: "Modo oscuro", desc: "Activable desde el panel de configuración." },
                { icon: "📄", title: "Nuevo Mail", desc: "Ícono en la esquina superior derecha para descartar los drafts actuales y comenzar un mail nuevo." },
                { icon: "🔢", title: "Pallets sin negativos", desc: "El campo de pallets no acepta valores negativos." },
                { icon: "✉️", title: "Mejoras en el HTML del mail", desc: "Render del correo mejorado visualmente para una presentación más prolija." },
              ].map(({ icon, title, desc }) => (
                <li key={title} className="flex gap-3 items-start">
                  <span className="text-lg mt-0.5">{icon}</span>
                  <div>
                    <p className="text-sm font-semibold text-gray-700">{title}</p>
                    <p className="text-xs text-gray-500 leading-snug">{desc}</p>
                  </div>
                </li>
              ))}
            </ul>
            <button
              onClick={() => {
                localStorage.setItem("frutam_whatsnew_v47", "seen");
                setShowWhatsNew(false);
              }}
              className="w-full py-2.5 rounded-xl text-sm font-bold text-white transition-colors hover:opacity-90"
              style={{ background: "#2563eb" }}
            >
              Entendido
            </button>
          </div>
        </div>,
        document.body
      )}

      {/* ── CONFIG PANEL (Portal) ─────────────────────────────────────────────── */}
      {showConfigPanel && ReactDOM.createPortal(
        <>
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.2)", zIndex: 9998 }} onClick={() => setShowConfigPanel(false)} />
          <div style={{ position: "fixed", top: 0, right: 0, height: "100vh", width: "300px", background: darkMode ? "#1f2937" : "#fff", boxShadow: "-4px 0 24px rgba(0,0,0,0.12)", zIndex: 9999, display: "flex", flexDirection: "column" }}>
            <div className="flex items-center justify-between px-5 py-4" style={{ borderBottom: `1px solid ${darkMode ? "#374151" : "#f3f4f6"}` }}>
              <div className="flex items-center gap-2">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" style={{ color: darkMode ? "#9ca3af" : "#6b7280" }} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
                  <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                </svg>
                <h2 className="text-base font-bold" style={{ color: darkMode ? "#f9fafb" : "#1f2937" }}>Configuración</h2>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="p-1 rounded transition-colors" style={{ color: darkMode ? "#6b7280" : "#9ca3af" }}>
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
                </svg>
              </button>
            </div>
            <div className="flex-1 px-5 py-5 space-y-6" style={{ overflowY: "auto" }}>

              {/* ── DARK MODE TOGGLE ── */}
              <div>
                <p className="text-sm font-semibold mb-1" style={{ color: darkMode ? "#d1d5db" : "#374151" }}>Apariencia</p>
                <p className="text-xs mb-3" style={{ color: darkMode ? "#6b7280" : "#9ca3af" }}>Cambia entre modo claro y oscuro</p>
                <button
                  onClick={toggleDarkMode}
                  className="w-full flex items-center justify-between px-4 py-3 rounded-lg border transition-all"
                  style={{
                    backgroundColor: darkMode ? "#374151" : "#f9fafb",
                    borderColor: darkMode ? "#4b5563" : "#e5e7eb",
                  }}
                >
                  <div className="flex items-center gap-3">
                    <span className="text-lg">{darkMode ? "🌙" : "☀️"}</span>
                    <span className="text-sm font-semibold" style={{ color: darkMode ? "#f9fafb" : "#374151" }}>
                      {darkMode ? "Modo oscuro" : "Modo claro"}
                    </span>
                  </div>
                  {/* Toggle pill */}
                  <div
                    className="relative w-11 h-6 rounded-full transition-colors duration-200 flex-shrink-0"
                    style={{ backgroundColor: darkMode ? "#3b82f6" : "#d1d5db" }}
                  >
                    <div
                      className="absolute top-0.5 w-5 h-5 bg-white rounded-full shadow transition-transform duration-200"
                      style={{ transform: darkMode ? "translateX(20px)" : "translateX(2px)" }}
                    />
                  </div>
                </button>
              </div>

              <div style={{ borderTop: `1px solid ${darkMode ? "#374151" : "#f3f4f6"}`, paddingTop: "1.5rem" }}>
                <p className="text-sm font-semibold mb-1" style={{ color: darkMode ? "#d1d5db" : "#374151" }}>Cliente de correo</p>
                <p className="text-xs mb-3" style={{ color: darkMode ? "#6b7280" : "#9ca3af" }}>¿Cómo abrir el correo al hacer "Enviar Email"?</p>
                <div className="space-y-2">
                  {[
                    { value: "auto", label: "Auto-detectar", desc: "Intenta Outlook Desktop primero, abre Outlook Web si no hay app instalada", icon: "🔍" },
                    { value: "desktop", label: "Outlook Desktop", desc: "Siempre abre la app de Outlook instalada en Windows", icon: "🖥️" },
                    { value: "web", label: "Outlook Web", desc: "Siempre abre outlook.office.com en el navegador", icon: "🌐" }
                  ].map((opt) => (
                    <button
                      key={opt.value}
                      onClick={() => saveEmailClientPref(opt.value)}
                      className="w-full text-left px-4 py-3 rounded-lg border transition-all"
                      style={{
                        borderColor: emailClientPref === opt.value ? "#3b82f6" : darkMode ? "#4b5563" : "#e5e7eb",
                        backgroundColor: emailClientPref === opt.value
                          ? darkMode ? "#1e3a5f" : "#eff6ff"
                          : darkMode ? "#374151" : "#ffffff",
                      }}
                    >
                      <div className="flex items-center gap-2">
                        <span className="text-base">{opt.icon}</span>
                        <span className="text-sm font-semibold" style={{ color: emailClientPref === opt.value ? "#3b82f6" : darkMode ? "#d1d5db" : "#374151" }}>{opt.label}</span>
                        {emailClientPref === opt.value && <span className="ml-auto text-blue-500 text-xs font-bold">✓ Activo</span>}
                      </div>
                      <p className="text-xs mt-1 ml-6" style={{ color: darkMode ? "#6b7280" : "#9ca3af" }}>{opt.desc}</p>
                    </button>
                  ))}
                </div>
              </div>
            </div>
            <div className="px-5 py-3" style={{ borderTop: `1px solid ${darkMode ? "#374151" : "#f3f4f6"}` }}>
              <p className="text-xs text-center" style={{ color: darkMode ? "#6b7280" : "#9ca3af" }}>Preferencia guardada en este dispositivo</p>
            </div>
          </div>
        </>,
        document.body
      )}
      {/* ── VERSION FOOTER ───────────────────────────────────────────────────── */}
      <div style={{
        position: "fixed",
        bottom: "8px",
        left: "50%",
        transform: "translateX(-50%)",
        fontSize: "11px",
        color: darkMode ? "#4b5563" : "#9ca3af",
        pointerEvents: "none",
        userSelect: "none",
        whiteSpace: "nowrap",
        zIndex: 10,
      }}>
        v48.1 · 11 Mar 2026
      </div>
    </div>
  );
};

export default App;

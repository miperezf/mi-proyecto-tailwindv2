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

// Configuración de Firebase usando variables de entorno de Vite
// Es crucial que las variables de entorno en Vite comiencen con `VITE_`
// Note: In the Canvas environment, __firebase_config and __initial_auth_token are provided globally.
// This setup assumes a Vite environment for local development, and will fall back to global vars if available.
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

// Initialize Firebase outside the component to prevent re-initializations
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);


// Component for rendering a single input field with styling
const InputField = React.forwardRef(
  (
    {
      label,
      name,
      value,
      onChange,
      onBlur,
      onFocus, // Added onFocus prop
      placeholder = "",
      className = "",
      type = "text",
      list, // Add list prop for datalist
      readOnly = false, // Add readOnly prop
    },
    ref
  ) => {
    // Ensure value is explicitly a string, falling back to an empty string if null/undefined
    const safeValue =
      value === null || value === undefined ? "" : String(value);
    return (
      <div className="mb-2 w-full">
        <label
          htmlFor={name}
          className="block text-sm font-medium text-gray-700"
        >
          {label}
        </label>
        <input
          type={type}
          id={name} // Keep ID for accessibility in header fields
          name={name}
          value={safeValue} // Use the safeValue
          onChange={onChange}
          onFocus={onFocus} // Added onFocus handler
          onBlur={onBlur} // Add onBlur event for capitalization
          placeholder={placeholder}
          className={`mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 bg-gray-50 border ${className}`}
          ref={ref} // Pass the ref here
          list={list} // Pass the list prop here
          readOnly={readOnly} // Apply readOnly property
        />
      </div>
    );
  }
);

// Component for rendering a table cell input
const TableInput = React.forwardRef(
  (
    {
      name,
      value,
      onChange,
      onBlur,
      onFocus, // Added onFocus prop
      type = "text",
      placeholder = "",
      readOnly = false,
      isCanceledProp = false,
      list,
    },
    ref
  ) => {
    // Added 'list' prop
    // Ensure value is explicitly a string, falling back to an an empty string if null/undefined
    const safeValue =
      value === null || value === undefined ? "" : String(value);
    return (
      <input
        type={type}
        // Removed dynamic 'id' prop from TableInput to prevent re-rendering/selection issues.
        name={name}
        value={safeValue} // Use the safeValue
        onChange={onChange}
        onFocus={onFocus} // Added onFocus handler
        onBlur={onBlur} // Add onBlur event for capitalization
        placeholder={placeholder}
        // Apply line-through directly to the input if canceled
        className={`w-full h-full p-px border-none bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded-md ${
          isCanceledProp ? "line-through" : ""
        }`}
        ref={ref} // Pass the ref here
        readOnly={readOnly} // Apply readOnly property
        list={list} // Pass the list prop here
      />
    );
  }
);

// Main App component
const App = () => {
  // Using projectId as appId for the Firestore collection path, prioritizing __app_id
  const appId =
    typeof __app_id !== "undefined" ? __app_id : firebaseConfig.projectId;

  const [userId, setUserId] = useState(null);
  const [isAuthReady, setIsAuthReady] = useState(false); // To know when authentication is ready
  const [isLoading, setIsLoading] = useState(true); // Loading state for data fetching
  const [allOrdersFromFirestore, setAllOrdersFromFirestore] = useState([]); // Stores raw data from Firestore
  const [displayedOrders, setDisplayedOrders] = useState([]); // Orders currently displayed (after search filter)

  // State for search functionality
  const [searchTerm, setSearchTerm] = useState("");
  const [committedSearchTerm, setCommittedSearchTerm] = useState(""); // New state for search triggered by button

  // State to explicitly track the ID of the currently active order being edited
  const [activeOrderId, setActiveOrderId] = useState(null);

  // ─── PRESENCE SYSTEM ────────────────────────────────────────────────────────
  const [presenceMailId, setPresenceMailId] = useState(null);
  const [otherUsersPresent, setOtherUsersPresent] = useState([]);
  const [showPresenceModal, setShowPresenceModal] = useState(false);
  const [presenceConflictUsers, setPresenceConflictUsers] = useState([]);
  const [pendingSearchTerm, setPendingSearchTerm] = useState("");
  const presenceUnsubscribeRef = useRef(null);

  // ─── LAST SEND RECOVERY ─────────────────────────────────────────────────────
  const LAST_SEND_KEY = "frutam_last_send";

  const [lastSendData, setLastSendData] = useState(() => {
    try {
      const raw = localStorage.getItem(LAST_SEND_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (parsed && !parsed.confirmed && Date.now() - parsed.timestamp < 24 * 60 * 60 * 1000) {
        return parsed;
      }
      return null;
    } catch {
      return null;
    }
  });

  const [showRecoveryModal, setShowRecoveryModal] = useState(false);

  // ─── HISTORY PANEL ──────────────────────────────────────────────────────────
  const [showHistoryPanel, setShowHistoryPanel] = useState(false);
  const [historyResendingId, setHistoryResendingId] = useState(null);

  // ─── EMAIL CLIENT CONFIG ─────────────────────────────────────────────────────
  const EMAIL_CLIENT_KEY = "frutam_email_client";
  const [emailClientPref, setEmailClientPref] = useState(
    () => localStorage.getItem(EMAIL_CLIENT_KEY) || "auto"
  );
  const [showConfigPanel, setShowConfigPanel] = useState(false);

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

  const saveLastSend = (mailId, subject, html) => {
    const record = { mailId, subject, html, timestamp: Date.now(), confirmed: false };
    localStorage.setItem(LAST_SEND_KEY, JSON.stringify(record));
    setLastSendData(record);
  };

  const confirmLastSend = () => {
    try {
      const raw = localStorage.getItem(LAST_SEND_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        parsed.confirmed = true;
        localStorage.setItem(LAST_SEND_KEY, JSON.stringify(parsed));
      }
    } catch { /* ignore */ }
    setLastSendData(null);
  };

  const recoverLastSend = async () => {
    if (!lastSendData) return;
    await copyFormattedContentToClipboard(lastSendData.html);
    openEmailClient(lastSendData.subject);
    setShowRecoveryModal(false);
  };

  // ─── DERIVED: last 3 sent Mail IDs for the current user ─────────────────────
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
      .slice(0, 3);
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

      ordersToProcess.forEach((order, index) => {
        innerEmailContentHtml += `
          <h3 style="font-size:18px;color:#2563eb;margin-top:40px;margin-bottom:15px;text-align:left;">
            Pedido #${index + 1}
          </h3>
          ${generateSingleOrderHtml(order.header, order.items, index + 1)}
        `;
        if (order.header.reDestinatarios) allProveedores.add(order.header.reDestinatarios);
        order.items.forEach((item) => { if (item.especie) allEspecies.add(item.especie); });
      });

      const consolidatedSubject = generateEmailSubjectValue(
        Array.from(allProveedores),
        Array.from(allEspecies),
        ""
      );

      const fullEmailBodyHtml = `
        <!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Detalle de Pedido</title>
</head>
<body style="margin:0;padding:16px;background-color:#f8f8f8;font-family:Arial,sans-serif;">
  <div style="text-align:right;margin-bottom:12px;font-family:Arial,sans-serif;font-weight:bold;font-size:14px;color:#ef4444;">
    Mail ID: ${mailId}
  </div>
  ${innerEmailContentHtml}
</body>
</html>`;

      await copyFormattedContentToClipboard(fullEmailBodyHtml);
      saveLastSend(mailId, consolidatedSubject, fullEmailBodyHtml);

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
    const initAuth = async () => {
      try {
        const unsubscribe = onAuthStateChanged(auth, async (user) => {
          if (user) {
            // Always use the real Firebase Auth UID — never localStorage.
            // localStorage caused desync: the stored UID could differ from
            // request.auth.uid in Security Rules, causing permission errors.
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
              // auth.currentUser is set synchronously after signIn
              setUserId(auth.currentUser?.uid ?? null);
            } catch (anonError) {
              setUserId(null);
            } finally {
              setIsAuthReady(true);
            }
          }
        });

        return () => unsubscribe();
      } catch (error) {
        setIsAuthReady(true);
      }
    };

    initAuth();
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
    id: crypto.randomUUID(),
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

  const [headerInfo, setHeaderInfo] = useState(() => ({
    ...initialHeaderState,
    emailSubject: generateEmailSubjectValue([], []),
  }));

  const [orderItems, setOrderItems] = useState([initialItemState]);
  const [currentOrderIndex, setCurrentOrderIndex] = useState(0);
  const [showOrderActionsModal, setShowOrderActionsModal] = useState(false);
  const [previewHtmlContent, setPreviewHtmlContent] = useState("");
  const [isShowingPreview, setIsShowingPreview] = useState(false);
  const [emailActionTriggered, setEmailActionTriggered] = useState(false);

  const [showObservationModal, setShowObservationModal] = useState(false);
  const [currentEditingItemData, setCurrentEditingItemData] = useState(null);
  const [modalObservationText, setModalObservationText] = useState("");
  const observationTextareaRef = useRef(null);

  const headerInputRefs = useRef({});
  const tableInputRefs = useRef({});
  const isUserEditing = useRef(false);
  const pendingSaveTimeout = useRef(null);
  // Stable ref so the keyboard handler always sees current orderItems
  // without needing to be re-registered on every state change.
  const orderItemsRef = useRef([]);
  // Mirror into ref every render so the keyboard handler always has fresh data
  orderItemsRef.current = orderItems;
  // Stable ref so the Firestore listener snapshot callback can read the current
  // activeOrderId without needing it as a dependency (which would cause the
  // listener to re-subscribe on every order change, creating an infinite loop).
  const activeOrderIdRef = useRef(null);
  activeOrderIdRef.current = activeOrderId;

  // ─── DRAG & DROP REFS ────────────────────────────────────────────────────────
  const dragItemIndex = useRef(null);   // index being dragged
  const dragOverIndex = useRef(null);  // index currently hovered

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

      // Distinguish create vs update so Security Rules evaluate correctly.
      // - create: send full document including createdBy + createdAt (required by rule)
      // - update: use updateDoc so we never overwrite createdBy/createdAt,
      //   which avoids the createdAtUnchanged() rule failure on legacy docs
      //   that were saved without a numeric createdAt.
      if (orderToSave.isNew) {
        // Brand-new document — setDoc triggers the "create" rule.
        // Must include createdBy + createdAt (required by rule).
        await setDoc(orderDocRef, {
          header: {
            ...header,
            lastModifiedBy: userId,
            updatedAt: Date.now(),
          },
          items: JSON.stringify(items),
        });
      } else {
        // Existing document — use dot-notation so Firestore merges individual
        // fields instead of replacing the entire "header" object.
        // This means createdBy and createdAt in Firestore are NEVER touched,
        // so createdByUnchanged() in Security Rules always passes regardless
        // of what the local header state contains.
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

  // ────────────────────────────────────────────────────────────────────────────

  useEffect(() => {
    const currentProveedor = headerInfo.reDestinatarios;
    const currentEspecie = orderItems[0]?.especie || "";

    const newSubjectForCurrentOrder = generateEmailSubjectValue(
      [currentProveedor],
      [currentEspecie],
      headerInfo.mailId
    );

    if (newSubjectForCurrentOrder !== headerInfo.emailSubject) {
      setHeaderInfo((prevInfo) => ({
        ...prevInfo,
        emailSubject: newSubjectForCurrentOrder,
      }));
    }
  }, [
    headerInfo.reDestinatarios,
    orderItems[0]?.especie,
    headerInfo.emailSubject,
    headerInfo.mailId,
  ]);

  useEffect(() => {
    if (!db || !userId || !isAuthReady) {
      return;
    }

    setIsLoading(true);
    const ordersCollectionRef = collection(
      db,
      `artifacts/${appId}/public/data/pedidos`
    );

    // Normal mode: filter server-side to only fetch the current user's orders.
    // Requires the composite index on header.createdBy + header.createdAt
    // declared in firestore.indexes.json.
    // Search mode: fetch the full collection so cross-user mailId lookups work.
    const q = committedSearchTerm
      ? query(ordersCollectionRef)
      : query(ordersCollectionRef, where("header.createdBy", "==", userId));


    const unsubscribe = onSnapshot(
      q,
      async (snapshot) => {
        let fetchedOrders = snapshot.docs.map((doc) => {
          const data = doc.data();
          let parsedItems = [];

          if (data.items) {
            if (typeof data.items === "string") {
              try {
                parsedItems = JSON.parse(data.items);
              } catch (e) {
                parsedItems = [];
              }
            } else if (Array.isArray(data.items)) {
              parsedItems = data.items;
            }
          }

          return {
            id: doc.id,
            header: data.header,
            items: parsedItems,
          };
        });

        fetchedOrders = fetchedOrders.map((order) => ({
          ...order,
          items: Array.isArray(order.items)
            ? order.items
            : [order.items].filter(Boolean),
        }));

        fetchedOrders.sort((a, b) => {
          const dateA = a.header?.createdAt || 0;
          const dateB = b.header?.createdAt || 0;
          return dateA - dateB;
        });

        setAllOrdersFromFirestore(fetchedOrders);
        setIsLoading(false);

        const draftOrders = fetchedOrders.filter(
          (order) =>
            order.header?.status === "draft" &&
            order.header?.createdBy === userId
        );

        if (
          draftOrders.length === 0 &&
          !committedSearchTerm &&
          activeOrderIdRef.current === null
        ) {
          const newOrderDocRef = doc(ordersCollectionRef);
          const newOrderId = newOrderDocRef.id;
          const newBlankHeader = {
            ...initialHeaderState,
            mailId: crypto.randomUUID().substring(0, 8).toUpperCase(),
            emailSubject: generateEmailSubjectValue([], []),
            status: "draft",
            createdAt: Date.now(),
            createdBy: userId,
            updatedAt: Date.now(),
          };
          const newBlankItems = [
            { ...initialItemState, id: crypto.randomUUID() },
          ];
          try {
            await saveOrderToFirestore({
              id: newOrderId,
              isNew: true,
              header: newBlankHeader,
              items: newBlankItems,
            });
            setActiveOrderId(newOrderId);
          } catch (saveError) {
            setIsLoading(false);
          }
        } else if (
          draftOrders.length > 0 &&
          activeOrderIdRef.current === null &&
          !committedSearchTerm
        ) {
          setActiveOrderId(draftOrders[0].id);
        }
      },
      (error) => {
        setIsLoading(false);
      }
    );

    return () => unsubscribe();
  }, [db, userId, isAuthReady, appId, committedSearchTerm]);

  useEffect(() => {
    let filtered = [];

    if (committedSearchTerm) {
      filtered = allOrdersFromFirestore.filter((order) => {
        const mailIdMatch = (order.header?.mailId || "")
          .toLowerCase()
          .includes(committedSearchTerm.toLowerCase());
        const isNotDeleted = order.header?.status !== "deleted";
        return mailIdMatch && isNotDeleted;
      });
    } else {
      filtered = allOrdersFromFirestore.filter((order) => {
        const isDraft = order.header?.status === "draft";
        const isOwner = order.header?.createdBy === userId;
        return isDraft && isOwner;
      });
    }

    filtered.sort((a, b) => {
      const dateA = a.header?.createdAt || 0;
      const dateB = b.header?.createdAt || 0;
      return dateA - dateB;
    });

    setDisplayedOrders(filtered);

    let newActiveOrderIdCandidate = null;

    if (committedSearchTerm) {
      const currentActiveInFiltered = filtered.find(
        (order) => order.id === activeOrderId
      );
      if (currentActiveInFiltered) {
        newActiveOrderIdCandidate = activeOrderId;
      } else if (filtered.length > 0) {
        newActiveOrderIdCandidate = filtered[filtered.length - 1].id;
      } else {
        newActiveOrderIdCandidate = null;
      }
    } else {
      const currentActiveInDrafts = filtered.find(
        (order) => order.id === activeOrderId
      );
      if (currentActiveInDrafts) {
        newActiveOrderIdCandidate = activeOrderId;
      } else if (filtered.length > 0) {
        newActiveOrderIdCandidate = filtered[filtered.length - 1].id;
      } else {
        newActiveOrderIdCandidate = null;
      }
    }

    if (newActiveOrderIdCandidate !== activeOrderId) {
      setActiveOrderId(newActiveOrderIdCandidate);
    } else if (
      newActiveOrderIdCandidate &&
      newActiveOrderIdCandidate === activeOrderId
    ) {
      const newIndex = filtered.findIndex(
        (order) => order.id === activeOrderId
      );
      if (newIndex !== -1 && newIndex !== currentOrderIndex) {
        setCurrentOrderIndex(newIndex);
      }
    }
  }, [allOrdersFromFirestore, committedSearchTerm, activeOrderId]);

  useEffect(() => {

    if (activeOrderId && displayedOrders.length > 0) {
      const orderToDisplay = displayedOrders.find(
        (order) => order.id === activeOrderId
      );
      if (orderToDisplay) {
        if (!isUserEditing.current) {
          setHeaderInfo(orderToDisplay.header);
          setOrderItems(orderToDisplay.items);
        } else {
        }
        const foundIndex = displayedOrders.findIndex(
          (order) => order.id === activeOrderId
        );
        if (foundIndex !== -1 && foundIndex !== currentOrderIndex) {
          setCurrentOrderIndex(foundIndex);
        }
      } else {
      }
    }
    else if (!activeOrderId && displayedOrders.length > 0) {
      const firstDraftOrder = displayedOrders[0];
      if (firstDraftOrder) {
        setActiveOrderId(firstDraftOrder.id);
      } else {
        setHeaderInfo({
          ...initialHeaderState,
          emailSubject: generateEmailSubjectValue([], []),
        });
        setOrderItems([initialItemState]);
        setCurrentOrderIndex(0);
      }
    }
    else if (!activeOrderId && displayedOrders.length === 0) {
      setHeaderInfo({
        ...initialHeaderState,
        emailSubject: generateEmailSubjectValue([], []),
      });
      setOrderItems([initialItemState]);
      setCurrentOrderIndex(0);
    }
    else if (activeOrderId && displayedOrders.length === 0) {
      setActiveOrderId(null);
    }
  }, [activeOrderId, displayedOrders]);

  const saveCurrentFormDataToDisplayed = async () => {
    if (!db || !userId) {
      return;
    }
    if (!activeOrderId) {
      return;
    }

    const currentOrderData = {
      id: activeOrderId,
      header: { ...headerInfo },
      items: orderItems.map((item) => ({ ...item })),
    };
    await saveOrderToFirestore(currentOrderData);
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
    // Cancel any pending save-and-release so focus restores protection
    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
      pendingSaveTimeout.current = null;
    }
  };

  const handleBlurField = (originalBlurFn) => (e) => {
    // Run the original blur handler first (e.g. capitalize)
    if (originalBlurFn) originalBlurFn(e);

    // Schedule a save-and-release after a short delay.
    // The delay lets React batch the state update from the blur handler
    // before we snapshot headerInfo/orderItems for saving.
    // If focus moves to another field, handleFocusField will cancel
    // this timeout, keeping the guard active the whole time the user
    // is inside the form.
    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
    }
    pendingSaveTimeout.current = setTimeout(async () => {
      pendingSaveTimeout.current = null;
      // Save to Firestore while isUserEditing is still true so no
      // concurrent snapshot can overwrite while we are mid-flight.
      await saveCurrentFormDataToDisplayed();
      // Only release the guard AFTER the save is committed.
      isUserEditing.current = false;
    }, 300);
  };

  const handleItemChange = (itemId, e) => {
    const { name, value } = e.target;
    setOrderItems((prevItems) => {
      const updatedItems = prevItems.map((item) =>
        item.id === itemId
          ? {
              ...item,
              [name]: value,
            }
          : item
      );
      return updatedItems;
    });
  };

  const handleItemBlur = (itemId, e) => {
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
                  formattedPrices.push(
                    `$ ${numericValue.toFixed(2).replace(".", ",")}`
                  );
                }
              }
            }

            if (formattedPrices.length > 0) {
              newValue = formattedPrices.join(" - ");
            } else {
              newValue = "";
            }
          } else if (name === "calibre") {
            const parts = value
              .split(/[,;\s-]+/)
              .filter((part) => part.trim() !== "")
              .map((part) => part.trim().toUpperCase());
            newValue = parts.join(" - ");
          } else if (name === "categoria") {
            const matches = value.match(/[a-zA-Z0-9]+/g);
            if (matches && matches.length > 0) {
              newValue = matches.join(" - ");
            } else {
              newValue = "";
            }
            newValue = newValue.toUpperCase();
          }
          else if (type !== "number") {
            newValue = value.toUpperCase();
          }
          return { ...item, [name]: newValue };
        }
        return item;
      });
      return updatedItems;
    });
  };

  const handleAddItem = (sourceItemId = null) => {
    setOrderItems((prevItems) => {
      let updatedItems;
      if (sourceItemId) {
        const sourceItem = prevItems.find((item) => item.id === sourceItemId);
        if (sourceItem) {
          const newItem = {
            ...sourceItem,
            id: crypto.randomUUID(),
            isCanceled: false,
          };
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
              id: crypto.randomUUID(),
              pallets: "",
              especie: "",
              variedad: "",
              formato: "",
              calibre: "",
              categoria: "",
              preciosFOB: "",
              estado: "",
              isCanceled: false,
            },
          ];
        }
      } else {
        updatedItems = [
          ...prevItems,
          {
            id: crypto.randomUUID(),
            pallets: "",
            especie: "",
            variedad: "",
            formato: "",
            calibre: "",
            categoria: "",
            preciosFOB: "",
            estado: "",
            isCanceled: false,
          },
        ];
      }
      saveOrderToFirestore({
        id: activeOrderId,
        header: { ...headerInfo },
        items: updatedItems,
      });
      return updatedItems;
    });
  };

  const handleDeleteItem = (idToDelete) => {
    setOrderItems((prevItems) => {
      if (prevItems.length <= 1) {
        return prevItems;
      }
      const updatedItems = prevItems.filter((item) => item.id !== idToDelete);
      saveOrderToFirestore({
        id: activeOrderId,
        header: { ...headerInfo },
        items: updatedItems,
      });
      return updatedItems;
    });
  };

  const handleDragReorder = (fromIndex, toIndex) => {
    if (fromIndex === toIndex) return;
    setOrderItems((prev) => {
      const updated = [...prev];
      const [moved] = updated.splice(fromIndex, 1);
      updated.splice(toIndex, 0, moved);
      // Save new order to Firestore immediately
      saveOrderToFirestore({
        id: activeOrderId,
        header: { ...headerInfo },
        items: updated,
      });
      return updated;
    });
  };

  const toggleItemCancellation = (itemId) => {
    setOrderItems((prevItems) => {
      const updatedItems = prevItems.map((item) => {
        if (item.id === itemId) {
          const newIsCanceled = !item.isCanceled;
          return {
            ...item,
            isCanceled: newIsCanceled,
            estado: newIsCanceled ? "CANCELADO" : "",
          };
        }
        return item;
      });
      saveOrderToFirestore({
        id: activeOrderId,
        header: { ...headerInfo },
        items: updatedItems,
      });
      return updatedItems;
    });
  };

  const currentOrderTotalPallets = orderItems.reduce((sum, item) => {
    if (item.isCanceled) {
      return sum;
    }
    const pallets = parseFloat(item.pallets) || 0;
    return sum + pallets;
  }, 0);

  const formatDateToSpanish = (dateString) => {
    const months = [
      "Enero",
      "Febrero",
      "Marzo",
      "Abril",
      "Mayo",
      "Junio",
      "Julio",
      "Agosto",
      "Septiembre",
      "Octubre",
      "Noviembre",
      "Diciembre",
    ];
    const daysOfWeek = [
      "Domingo",
      "Lunes",
      "Martes",
      "Miércoles",
      "Jueves",
      "Viernes",
      "Sábado",
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

  // ─── EMAIL HTML GENERATOR ────────────────────────────────────────────────────
  // Generates a single order block as a self-contained HTML string.
  //
  // DESIGN PRINCIPLES for cross-client compatibility:
  //  1. Zero class references — every style is 100% inline on the element itself.
  //     Outlook Desktop strips <style> blocks when pasting; inline styles survive.
  //  2. Explicit font-family on every text element — Outlook resets inherited fonts.
  //  3. No shorthand borders on <table> — use border-collapse + per-cell borders.
  //  4. No CSS variables, no calc(), no flexbox inside the table — not supported
  //     in Outlook's Word rendering engine (used by Outlook Desktop on Windows).
  //  5. All colours are hex — Outlook 2016 ignores rgba/hsl.
  //  6. mso-padding-alt / mso-line-height-rule omitted intentionally (not needed
  //     here) but the structure is already compatible with the Word engine.
  // ─────────────────────────────────────────────────────────────────────────────
  // ─── EMAIL HTML GENERATOR ────────────────────────────────────────────────────
  // Visual design: matches exactly the published version (screenshot reference)
  //   • Outer card: white bg, 1px #dddddd border, 8px border-radius, 15px padding
  //   • Header section: País/Nave/Fecha/Exporta in 13px Arial, separated by a
  //     1px #eeeeee bottom border
  //   • Table: blue (#2563eb) header row, alternating #f9f9f9/#ffffff body rows,
  //     light gray (#f0f0f0) total row — all with 1px #dddddd cell borders
  //   • Observaciones: 13px bold label + italic value
  //
  // Cross-client compatibility (Outlook Desktop / Outlook Web / Gmail / Apple Mail):
  //   • Every style is 100% inline — no <style> block (Outlook Desktop strips them)
  //   • Explicit font-family on every element (Outlook resets inherited fonts)
  //   • table width:100% + table-layout:fixed → fills container on all clients
  //   • mso-table-lspace/rspace:0pt → removes Outlook's default table spacing
  //   • Conditional comments <!--[if mso]> → tells Outlook's Word engine exactly
  //     how to render the outer wrapper (which ignores border-radius anyway)
  //   • No overflow:hidden wrappers (Outlook Desktop clips content inside them)
  //   • All colors are 6-digit hex (Outlook 2016 ignores rgba/hsl)
  // ─────────────────────────────────────────────────────────────────────────────
  // Escape user-supplied strings before inserting them into HTML templates.
  // This prevents stored XSS if a user types HTML/script tags in a field.
  const escapeHtml = (str) => {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  };

  const generateSingleOrderHtml = (
    orderHeader,
    orderItemsData,
    orderNumber
  ) => {
    const nonCancelledItems = orderItemsData.filter((item) => !item.isCanceled);
    const singleOrderTotalPallets = nonCancelledItems.reduce((sum, item) => {
      const pallets = parseFloat(item.pallets) || 0;
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

    // ── Styles ───────────────────────────────────────────────────────────────
    // <p> inside the header section
    const pStyle     = "margin:0;margin-bottom:3px;font-family:Arial,sans-serif;font-size:13px;color:#333333;line-height:1.5;";
    const pLastStyle = "margin:0;font-family:Arial,sans-serif;font-size:13px;color:#333333;line-height:1.5;";

    // Table <th> — blue header matching the published version
    const thStyle =
      "font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#ffffff;" +
      "background-color:#2563eb;padding:4px 6px;" +
      "border-top:1px solid #1e40af;border-bottom:1px solid #1e40af;" +
      "border-left:1px solid #1e40af;border-right:1px solid #1e40af;" +
      "text-align:center;white-space:nowrap;vertical-align:middle;";

    // Table <td> base — border matches the card border (#dddddd) so there's no
    // visual seam between the last column and the card edge
    const tdBase =
      "font-family:Arial,sans-serif;font-size:11px;color:#333333;" +
      "padding:4px 6px;text-align:center;white-space:nowrap;vertical-align:middle;" +
      "border-top:1px solid #dddddd;border-bottom:1px solid #dddddd;" +
      "border-left:1px solid #dddddd;border-right:1px solid #dddddd;";

    // Total row — light gray matching the screenshot
    const tdTotalLabel =
      "font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#333333;" +
      "background-color:#f0f0f0;padding:5px 12px 5px 6px;text-align:right;" +
      "white-space:nowrap;vertical-align:middle;" +
      "border-top:1px solid #dddddd;border-bottom:1px solid #dddddd;" +
      "border-left:1px solid #dddddd;border-right:1px solid #dddddd;";

    const tdTotalValue =
      "font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#333333;" +
      "background-color:#f0f0f0;padding:5px 6px;text-align:center;" +
      "white-space:nowrap;vertical-align:middle;" +
      "border-top:1px solid #dddddd;border-bottom:1px solid #dddddd;" +
      "border-left:1px solid #dddddd;border-right:1px solid #dddddd;";

    // ── Data rows ─────────────────────────────────────────────────────────────
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

    // ── HTML output ───────────────────────────────────────────────────────────
    // The outer <!--[if mso]> wrapper forces Outlook Desktop to allocate 100% width
    // for the card div (Outlook ignores max-width on divs without this hint).
    // Modern clients (Gmail, Apple Mail, Outlook Web) use the div directly.
    return `
<!--[if mso]><table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td style="padding:0;"><![endif]-->
<div style="font-family:Arial,sans-serif;font-size:14px;color:#333333;margin-bottom:20px;background-color:#ffffff;border:1px solid #dddddd;border-radius:8px;padding:15px;width:100%;max-width:900px;text-align:left;box-sizing:border-box;${orderBlockExtra}">

  <div style="padding-bottom:10px;border-bottom:1px solid #eeeeee;margin-bottom:12px;">
    <p style="${pStyle}"><strong style="font-weight:bold;">País:</strong> ${formattedPais}</p>
    <p style="${pStyle}"><strong style="font-weight:bold;">Nave:</strong> ${formattedNave}</p>
    <p style="${pStyle}"><strong style="font-weight:bold;">Fecha de carga:</strong> ${formattedFechaCarga}</p>
    <p style="${pLastStyle}"><strong style="font-weight:bold;">Exporta:</strong> ${formattedExporta}</p>
  </div>

  <table cellpadding="0" cellspacing="0" border="0"
    style="border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;width:100%;table-layout:fixed;margin-top:8px;">
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

  <p style="margin-top:10px;margin-bottom:0;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#333333;">
    Observaciones: <span style="font-weight:normal;font-style:italic;color:#555555;">${consolidatedObservationsText}</span>
  </p>

</div>
<!--[if mso]></td></tr></table><![endif]-->`;
  };

  // ─── CLIPBOARD HELPER ────────────────────────────────────────────────────────
  // Uses the modern Clipboard API (navigator.clipboard.write) which writes the
  // HTML string directly as a text/html Blob — bypassing the DOM entirely.
  // This guarantees the exact same HTML lands in Outlook Desktop, Outlook Web,
  // Mail on Mac, and every other client, regardless of OS or browser.
  //
  // Fallback chain (for older browsers / non-HTTPS / Safari quirks):
  //   1. navigator.clipboard.write  → modern, direct, no DOM rendering needed
  //   2. navigator.clipboard.writeText → copies plain HTML as text (last resort)
  //   3. execCommand('copy')         → legacy DOM selection method (original code)
  //
  // The plain-text fallback intentionally copies the raw HTML string so the user
  // at least has the content — they can paste it in a text editor to inspect it.
  // ─────────────────────────────────────────────────────────────────────────────
  const copyFormattedContentToClipboard = async (content) => {
    // Strategy 1: Modern Clipboard API — writes HTML Blob directly, no DOM needed.
    // Supported: Chrome 76+, Edge 79+, Firefox 87+ (requires HTTPS or localhost).
    if (navigator.clipboard && window.ClipboardItem) {
      try {
        await navigator.clipboard.write([
          new ClipboardItem({
            // The HTML MIME type tells Outlook/Mail to paste as rich text, not plain text.
            "text/html": new Blob([content], { type: "text/html" }),
            // Always include a plain-text fallback so paste still works in plain-text fields.
            "text/plain": new Blob([content], { type: "text/plain" }),
          }),
        ]);
        return;
      } catch (err) {
      }
    }

    // Strategy 2: execCommand('copy') — legacy DOM selection method.
    // Less reliable across OS/browser combinations but still works in most cases.
    try {
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = content;
      // Position off-screen but still in the document so it can be selected.
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

    // Strategy 3: Plain-text fallback — at minimum the user gets the raw HTML.
    try {
      await navigator.clipboard.writeText(content);
    } catch (err) {
    }
  };

  const handleAddOrder = async () => {
    if (!db || !userId) {
      return;
    }

    if (activeOrderId) {
      await saveCurrentFormDataToDisplayed();
    }

    let mailIdToAssignForNewOrder = "";
    if (committedSearchTerm) {
      mailIdToAssignForNewOrder = committedSearchTerm;
    } else {
      if (headerInfo.mailId && headerInfo.status === "draft") {
        mailIdToAssignForNewOrder = headerInfo.mailId;
      } else {
        mailIdToAssignForNewOrder = crypto
          .randomUUID()
          .substring(0, 8)
          .toUpperCase();
      }
    }

    const ordersCollectionRef = collection(
      db,
      `artifacts/${appId}/public/data/pedidos`
    );
    const newOrderDocRef = doc(ordersCollectionRef);
    const newOrderId = newOrderDocRef.id;

    const newBlankHeader = {
      ...initialHeaderState,
      reDestinatarios: headerInfo.reDestinatarios,
      emailSubject: generateEmailSubjectValue(
        [headerInfo.reDestinatarios],
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

    setActiveOrderId(newOrderId);
  };

  const handlePreviousOrder = () => {

    if (currentOrderIndex === 0) {
      return;
    }

    // Cancel any pending blur-triggered save; navigation save covers it
    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
      pendingSaveTimeout.current = null;
    }
    isUserEditing.current = false;

    saveCurrentFormDataToDisplayed();

    const newIndex = currentOrderIndex - 1;
    if (displayedOrders[newIndex]) {
      const nextActiveId = displayedOrders[newIndex].id;
      setActiveOrderId(nextActiveId);
    } else {
    }
  };

  const handleNextOrder = () => {

    if (currentOrderIndex === displayedOrders.length - 1) {
      return;
    }

    // Cancel any pending blur-triggered save; navigation save covers it
    if (pendingSaveTimeout.current) {
      clearTimeout(pendingSaveTimeout.current);
      pendingSaveTimeout.current = null;
    }
    isUserEditing.current = false;

    saveCurrentFormDataToDisplayed();

    const newIndex = currentOrderIndex + 1;
    if (displayedOrders[newIndex]) {
      const nextActiveId = displayedOrders[newIndex].id;
      setActiveOrderId(nextActiveId);
    } else {
    }
  };

  const handleDeleteCurrentOrder = async () => {
    try {
      await saveCurrentFormDataToDisplayed();

      if (displayedOrders.length <= 1) {
        return;
      }


      const orderIdToSoftDelete = displayedOrders[currentOrderIndex].id;

      await handleSoftDeleteOrderInFirestore(orderIdToSoftDelete);
    } catch (error) {
    }
  };

  const handleSearchClick = async () => {
    const term = searchTerm.toUpperCase();
    if (!term) return;

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
  };

  const handleClearSearch = async () => {

    if (activeOrderId) {
      await saveCurrentFormDataToDisplayed();
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

    setSearchTerm("");
    setCommittedSearchTerm("");
  };

  const isMobileDevice = () => {
    return window.innerWidth <= 767;
  };

  const performSendEmail = async () => {
    try {
      if (!db || !userId) {
        return;
      }

      await saveCurrentFormDataToDisplayed();

      const mailGlobalId = headerInfo.mailId;

      if (!mailGlobalId) {
        setShowOrderActionsModal(false);
        return;
      }

      const ordersCollectionRef = collection(
        db,
        `artifacts/${appId}/public/data/pedidos`
      );

      const currentProveedor = headerInfo.reDestinatarios;
      if (currentProveedor && mailGlobalId) {
        const ordersToGroupQuery = query(
          ordersCollectionRef,
          where("header.status", "==", "draft"),
          where("header.reDestinatarios", "==", currentProveedor)
        );
        const ordersToGroupSnapshot = await getDocs(ordersToGroupQuery);
        const batch = writeBatch(db);

        ordersToGroupSnapshot.docs.forEach((docSnapshot) => {
          const orderData = docSnapshot.data();
          const existingMailId = orderData.header?.mailId;
          if (!existingMailId || existingMailId !== mailGlobalId) {
            const orderRef = doc(
              db,
              `artifacts/${appId}/public/data/pedidos`,
              docSnapshot.id
            );
            batch.update(orderRef, { "header.mailId": mailGlobalId });
          }
        });

        if (batch._mutations.length > 0) {
          await batch.commit();
        } else {
        }
      } else {
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
        header: { ...headerInfo },
        items: orderItems.map((item) => ({ ...item })),
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
          const orderDocRef = doc(
            db,
            `artifacts/${appId}/public/data/pedidos`,
            order.id
          );
          await updateDoc(orderDocRef, {
            "header.mailId": mailGlobalId,
            "header.status": "sent",
            "header.lastModifiedBy": userId,
            "header.updatedAt": Date.now(),
          });
          order.header.status = "sent";
          order.header.mailId = mailGlobalId;
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

      ordersToProcess.forEach((order, index) => {
        innerEmailContentHtml += `
            <h3 style="font-size: 18px; color: #2563eb; margin-top: 40px; margin-bottom: 15px; text-align: left;">Pedido #${
              index + 1
            }</h3>
            ${generateSingleOrderHtml(order.header, order.items, index + 1)}
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

      // All styles are 100% inline inside generateSingleOrderHtml.
      // No <style> block — Outlook Desktop strips stylesheet blocks on paste.
      const fullEmailBodyHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Detalle de Pedido</title>
</head>
<body style="margin:0;padding:16px;background-color:#f8f8f8;font-family:Arial,sans-serif;">
  <div style="text-align:right;margin-bottom:12px;font-family:Arial,sans-serif;font-weight:bold;font-size:14px;color:#ef4444;">
    Mail ID: ${mailGlobalId}
  </div>
  ${innerEmailContentHtml}
</body>
</html>`;

      await copyFormattedContentToClipboard(fullEmailBodyHtml);
      saveLastSend(mailGlobalId, consolidatedSubject, fullEmailBodyHtml);

      openEmailClient(consolidatedSubject);

      setShowOrderActionsModal(false);
      setSearchTerm("");
      setCommittedSearchTerm("");
      setEmailActionTriggered(false);
      setIsShowingPreview(false);
      setPreviewHtmlContent("");
    } catch (error) {
    }
  };

  const handlePreviewOrder = async () => {
    await saveCurrentFormDataToDisplayed();

    const previewGlobalId = headerInfo.mailId;

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

    const currentProveedor = headerInfo.reDestinatarios;
    if (currentProveedor && previewGlobalId) {
      const ordersToGroupQuery = query(
        ordersCollectionRef,
        where("header.status", "==", "draft"),
        where("header.reDestinatarios", "==", currentProveedor)
      );
      const ordersToGroupSnapshot = await getDocs(ordersToGroupQuery);
      const batch = writeBatch(db);

      ordersToGroupSnapshot.docs.forEach((docSnapshot) => {
        const orderData = docSnapshot.data();
        const existingMailId = orderData.header?.mailId;
        if (!existingMailId || existingMailId !== previewGlobalId) {
          const orderRef = doc(
            db,
            `artifacts/${appId}/public/data/pedidos`,
            docSnapshot.id
          );
          batch.update(orderRef, { "header.mailId": previewGlobalId });
        }
      });

      if (batch._mutations.length > 0) {
        await batch.commit();
      } else {
      }
    } else {
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
      header: { ...headerInfo },
      items: orderItems.map((item) => ({ ...item })),
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
      ordersForPreview.forEach((order, index) => {
        innerPreviewHtml += `
          <h3 style="font-size: 18px; color: #2563eb; margin-top: 20px; margin-bottom: 10px; text-align: left;">Pedido #${
            index + 1
          }</h3>
          ${generateSingleOrderHtml(order.header, order.items, index + 1)}
        `;
      });

      // All styles are 100% inline — no <style> block needed.
      const finalPreviewHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Previsualización de Pedido</title>
</head>
<body style="margin:0;padding:16px;background-color:#f8f8f8;font-family:Arial,sans-serif;">
  <div style="text-align:right;margin-bottom:12px;font-family:Arial,sans-serif;font-weight:bold;font-size:14px;color:#ef4444;">
    Mail ID: ${previewGlobalId}
  </div>
  ${innerPreviewHtml}
</body>
</html>`;
      setPreviewHtmlContent(finalPreviewHtml);
    }
    setIsShowingPreview(true);
  };

  const handleFinalizeOrder = async () => {

    let mailIdToAssign = headerInfo.mailId;
    if (!headerInfo.mailId) {
      mailIdToAssign = crypto.randomUUID().substring(0, 8).toUpperCase();
      setHeaderInfo((prev) => ({ ...prev, mailId: mailIdToAssign }));
    } else {
    }

    await saveCurrentFormDataToDisplayed();

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
      saveOrderToFirestore({
        id: activeOrderId,
        header: { ...headerInfo },
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
  }, []); // stable: reads orderItems via orderItemsRef, uses only refs for DOM

  if (isLoading) {
    return (
      <div className="min-h-screen bg-gray-100 flex items-center justify-center">
        <div className="text-blue-600 text-lg font-semibold">
          Cargando pedidos...
        </div>
      </div>
    );
  }

  const currentOrderIsDeleted = headerInfo.status === "deleted";


  return (
    <div className="min-h-screen bg-gray-100 p-4 sm:p-6 lg:p-8 font-inter">
      <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-lg p-4 sm:p-6 space-y-6 relative">

        {/* ── TOP-RIGHT TOOLBAR ──────────────────────────────────────────────────
             All three elements (config gear, history clock, logo) live in a
             single flex container anchored to the top-right corner.
             This prevents the overlapping caused by independent absolute offsets.
        ──────────────────────────────────────────────────────────────────────── */}
        <div className="absolute top-4 right-4 z-10 flex items-center gap-2">

          {/* Config gear button — email client preference */}
          <button
            onClick={() => setShowConfigPanel(true)}
            className="p-1.5 rounded-full bg-white border border-gray-200 shadow-sm hover:bg-gray-50 hover:border-gray-400 transition-colors"
            title="Configuración"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
              <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
          </button>

          {/* History clock button — only shown when there are sent orders */}
          {sentHistory.length > 0 && (
            <button
              onClick={() => setShowHistoryPanel(true)}
              className="relative p-1.5 rounded-full bg-white border border-gray-200 shadow-sm hover:bg-blue-50 hover:border-blue-300 transition-colors"
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

          {/* Logo VPC */}
          <img
            src="https://www.vpcom.com/images/logo-vpc.png"
            onError={(e) => {
              e.target.onerror = null;
              e.target.src =
                "https://placehold.co/100x40/FFFFFF/000000?text=LogoVPC";
            }}
            alt="Logo VPC"
            className="w-20 sm:w-24 md:w-28 h-auto object-contain rounded-md"
          />
        </div>

        {/* Display persistent user identity */}
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

        <div className="border-b pb-4 mb-4 pt-16 sm:pt-0">
          <h1 className="text-xl sm:text-2xl font-bold text-center text-gray-800 mb-4">
            Pedidos Comercial Frutam
          </h1>
          {/* Search Input and Buttons */}
          <div className="flex items-end gap-2 mb-4">
            <div className="flex-grow">
              <label
                htmlFor="searchTerm"
                className="block text-sm font-medium text-gray-700"
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
                className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 bg-gray-50 border"
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

          {/* ── LAST SEND RECOVERY BANNER ─────────────────────────────────────── */}
          {lastSendData && !lastSendData.confirmed && (
            <div className="flex items-center justify-between gap-2 px-4 py-2 mb-2 bg-blue-50 border border-blue-200 rounded-lg text-blue-800 text-sm">
              <div className="flex items-center gap-2">
                <span className="text-lg">📋</span>
                <span>
                  Tienes un envío listo sin confirmar —{" "}
                  <strong>Mail ID: {lastSendData.mailId}</strong>
                  <span className="text-blue-500 ml-2 text-xs">
                    ({Math.round((Date.now() - lastSendData.timestamp) / 60000)} min atrás)
                  </span>
                </span>
              </div>
              <div className="flex gap-2 shrink-0">
                <button
                  onClick={() => setShowRecoveryModal(true)}
                  className="px-3 py-1 bg-blue-600 text-white text-xs font-semibold rounded-md hover:bg-blue-700 transition"
                >
                  Recuperar
                </button>
                <button
                  onClick={confirmLastSend}
                  className="px-3 py-1 bg-gray-200 text-gray-700 text-xs font-semibold rounded-md hover:bg-gray-300 transition"
                  title="Ya pegué el contenido correctamente"
                >
                  ✓ Ya envié
                </button>
              </div>
            </div>
          )}

          {/* Presence warning banner */}
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

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mt-4">
            <div>
              {/* Incoterm selector sits inline to the right of the supplier label */}
              <div className="flex items-center justify-between mb-1">
                <label className="block text-sm font-medium text-gray-700">
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
                onChange={handleHeaderChange}
                onFocus={handleFocusField}
                onBlur={handleBlurField(handleHeaderBlur)}
                placeholder="Ingrese nombre de proveedor"
                className="mt-0 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 bg-gray-50 border"
                ref={(el) => (headerInputRefs.current.reDestinatarios = el)}
                readOnly={currentOrderIsDeleted}
              />
            </div>

            <div>
              <InputField
                label="País:"
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

        {/* Navigation Buttons and Order Indicator */}
        <div className="flex flex-col sm:flex-row justify-center items-center gap-4 mt-4 mb-6">
          <div className="flex items-center justify-center w-full sm:w-auto">
            <button
              onClick={handlePreviousOrder}
              disabled={currentOrderIndex === 0}
              className="flex items-center justify-center px-3 py-2 bg-gray-300 text-gray-800 font-semibold rounded-lg shadow-md hover:bg-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 transition duration-150 ease-in-out disabled:opacity-50 disabled:cursor-not-allowed"
              title="Ir al pedido anterior (más antiguo)"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                className="h-5 w-5"
                viewBox="0 0 20 20"
                fill="currentColor"
              >
                <path
                  fillRule="evenodd"
                  d="M9.707 4.293a1 1 0 010 1.414L5.414 10l4.293 4.293a1 1 0 01-1.414 1.414l-5-5a1 1 0 010-1.414l5-5a1 1 0 011.414 0z"
                  clipRule="evenodd"
                />
              </svg>
              <span className="hidden sm:inline ml-1">Anterior</span>
            </button>

            <span className="text-center text-gray-700 font-semibold text-lg mx-2 sm:mx-4 min-w-[150px] sm:min-w-0">
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
              <svg
                xmlns="http://www.w3.org/2000/svg"
                className="h-5 w-5"
                viewBox="0 0 20 20"
                fill="currentColor"
              >
                <path
                  fillRule="evenodd"
                  d="M10.293 15.707a1 1 0 010-1.414L14.586 10l-4.293-4.293a1 1 0 111.414-1.414l5 5a1 1 0 010 1.414l-5 5a1 1 0 01-1.414 0z"
                  clipRule="evenodd"
                />
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
              {/* Table for desktop view */}
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
                        <TableInput
                          type="number"
                          name="pallets"
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
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
                      <td className="px-1 py-px text-xs border-r whitespace-nowrap" style={{ textAlign: "center" }}>
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
                    <td colSpan="6" style={{ padding: "6px 15px 6px 6px", textAlign: "right", fontWeight: "bold", border: "1px solid #ccc", borderBottomLeftRadius: "8px", marginTop: "15px" }}>
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
                  contenido manualmente en el cuerpo del mensaje. Una vez pegado
                  correctamente, haz click en{" "}
                  <strong>"✓ Ya envié"</strong> en el banner azul para confirmar.
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
                  <div className="bg-gray-50 p-4 rounded-md border border-gray-200 text-left flex-grow overflow-y-auto">
                    <div dangerouslySetInnerHTML={{ __html: previewHtmlContent }} />
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
                  className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 bg-gray-50 border"
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
        {showRecoveryModal && lastSendData && (
          <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center p-4 z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md mx-auto">
              <div className="flex items-center gap-3 mb-4">
                <span className="text-3xl">📋</span>
                <h2 className="text-xl font-bold text-gray-800">Recuperar envío</h2>
              </div>
              <p className="text-sm text-gray-600 mb-1">Mail ID:</p>
              <p className="font-bold text-blue-600 text-lg mb-3">{lastSendData.mailId}</p>
              <p className="text-sm text-gray-600 mb-1">Asunto del email:</p>
              <div className="flex items-center gap-2 mb-4">
                <code className="flex-1 bg-gray-50 border border-gray-200 rounded px-3 py-2 text-sm text-gray-800 break-all">
                  {lastSendData.subject}
                </code>
                <button onClick={() => { navigator.clipboard.writeText(lastSendData.subject); }} className="shrink-0 px-2 py-2 bg-gray-100 hover:bg-gray-200 rounded text-xs text-gray-600 transition" title="Copiar solo el asunto">
                  Copiar
                </button>
              </div>
              <p className="text-sm text-gray-500 mb-5">
                El contenido del pedido se volverá a copiar al portapapeles y se abrirá tu cliente de correo.
              </p>
              <div className="flex flex-col gap-2">
                <button onClick={recoverLastSend} className="w-full px-4 py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 transition flex items-center justify-center gap-2">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                    <path d="M8 3a1 1 0 011-1h2a1 1 0 110 2H9a1 1 0 01-1-1z" />
                    <path d="M6 3a2 2 0 00-2 2v11a2 2 0 002 2h8a2 2 0 002-2V5a2 2 0 00-2-2 3 3 0 01-3 3H9a3 3 0 01-3-3z" />
                  </svg>
                  Volver a copiar y abrir correo
                </button>
                <div className="flex gap-2">
                  <button onClick={() => { confirmLastSend(); setShowRecoveryModal(false); }} className="flex-1 px-4 py-2 bg-green-50 text-green-700 font-semibold rounded-lg border border-green-200 hover:bg-green-100 transition text-sm">
                    ✓ Ya lo pegué correctamente
                  </button>
                  <button onClick={() => setShowRecoveryModal(false)} className="flex-1 px-4 py-2 bg-gray-100 text-gray-700 font-semibold rounded-lg hover:bg-gray-200 transition text-sm">
                    Cerrar
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

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

        {/* Datalist for apple varieties */}
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
                        <button onClick={() => { setShowHistoryPanel(false); setSearchTerm(entry.mailId); setCommittedSearchTerm(entry.mailId); }} className="flex-1 px-3 py-1.5 text-xs font-semibold bg-white border border-gray-200 text-gray-700 rounded-md hover:bg-gray-100 transition">
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

      {/* ── CONFIG PANEL (Portal) ─────────────────────────────────────────────── */}
      {showConfigPanel && ReactDOM.createPortal(
        <>
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.2)", zIndex: 9998 }} onClick={() => setShowConfigPanel(false)} />
          <div style={{ position: "fixed", top: 0, right: 0, height: "100vh", width: "300px", background: "#fff", boxShadow: "-4px 0 24px rgba(0,0,0,0.12)", zIndex: 9999, display: "flex", flexDirection: "column" }}>
            <div className="flex items-center justify-between px-5 py-4 border-b border-gray-100">
              <div className="flex items-center gap-2">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-gray-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
                  <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                </svg>
                <h2 className="text-base font-bold text-gray-800">Configuración</h2>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="text-gray-400 hover:text-gray-600 p-1 rounded transition-colors">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
                </svg>
              </button>
            </div>
            <div className="flex-1 px-5 py-5 space-y-6">
              <div>
                <p className="text-sm font-semibold text-gray-700 mb-1">Cliente de correo</p>
                <p className="text-xs text-gray-400 mb-3">¿Cómo abrir el correo al hacer "Enviar Email"?</p>
                <div className="space-y-2">
                  {[
                    { value: "auto", label: "Auto-detectar", desc: "Intenta Outlook Desktop primero, abre Outlook Web si no hay app instalada", icon: "🔍" },
                    { value: "desktop", label: "Outlook Desktop", desc: "Siempre abre la app de Outlook instalada en Windows", icon: "🖥️" },
                    { value: "web", label: "Outlook Web", desc: "Siempre abre outlook.office.com en el navegador", icon: "🌐" }
                  ].map((opt) => (
                    <button
                      key={opt.value}
                      onClick={() => saveEmailClientPref(opt.value)}
                      className={`w-full text-left px-4 py-3 rounded-lg border transition-all ${emailClientPref === opt.value ? "border-blue-500 bg-blue-50" : "border-gray-200 bg-white hover:bg-gray-50"}`}
                    >
                      <div className="flex items-center gap-2">
                        <span className="text-base">{opt.icon}</span>
                        <span className={`text-sm font-semibold ${emailClientPref === opt.value ? "text-blue-700" : "text-gray-700"}`}>{opt.label}</span>
                        {emailClientPref === opt.value && <span className="ml-auto text-blue-600 text-xs font-bold">✓ Activo</span>}
                      </div>
                      <p className="text-xs text-gray-400 mt-1 ml-6">{opt.desc}</p>
                    </button>
                  ))}
                </div>
              </div>
            </div>
            <div className="px-5 py-3 border-t border-gray-100">
              <p className="text-xs text-gray-400 text-center">Preferencia guardada en este dispositivo</p>
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
        color: "#9ca3af",
        pointerEvents: "none",
        userSelect: "none",
        whiteSpace: "nowrap",
        zIndex: 10,
      }}>
        v13 · 03 Mar 2026
      </div>
    </div>
  );
};

export default App;

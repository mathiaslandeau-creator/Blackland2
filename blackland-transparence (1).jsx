import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Building2, Plus, Trash2, ChevronDown, ChevronLeft, ChevronRight, Sparkles, Zap, Droplets, Wifi,
  Receipt, Wallet, LogIn, ArrowUpRight, ArrowDownRight, X, AlertCircle, RefreshCw, Share2,
  LogOut, Users, Copy, Check, Lock
} from "lucide-react";

const C = {
  ink: "#0A0A0A",
  inkPanel: "#141412",
  inkPanel2: "#1C1B18",
  ivory: "#F3EEE3",
  ivoryDim: "#C9C3B4",
  ivoryFaint: "#8C8577",
  brass: "#B08D57",
  brassLight: "#C9A876",
  moss: "#8DA07A",
  wine: "#B06A5C",
  line: "rgba(243,238,227,0.09)",
  lineStrong: "rgba(243,238,227,0.16)",
};

const EXPENSE_TYPES = [
  { id: "electricite", label: "Électricité", icon: Zap },
  { id: "eau", label: "Eau", icon: Droplets },
  { id: "wifi", label: "Wifi", icon: Wifi },
  { id: "autre", label: "Autre charge", icon: Receipt },
];

const SEED_PROPERTIES = [
  { id: "p1", name: "2+1", address: "Marseille", ownerName: "Propriétaire", commissionPct: 18, icalUrl: "https://www.airbnb.fr/calendar/ical/1711686575949791907.ics?t=244970255d964961ba655936c63108e0" },
  { id: "p2", name: "Penthouse", address: "Marseille", ownerName: "Propriétaire", commissionPct: 18, icalUrl: "https://www.airbnb.fr/calendar/ical/1710223286544021736.ics?t=2a3b56e888194457877b71de058b407b" },
];

const eur = (n) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(Number(n) || 0);
const eur2 = (n) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR" }).format(Number(n) || 0);
const fmtDate = (s) => s ? new Date(s + "T00:00:00").toLocaleDateString("fr-FR", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const uid = () => Math.random().toString(36).slice(2, 10);
function parseNum(v) {
  if (v === null || v === undefined) return NaN;
  const cleaned = String(v).trim().replace(/\s/g, "").replace(",", ".");
  if (cleaned === "") return NaN;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

function parseICS(text) {
  const events = [];
  const blocks = text.split("BEGIN:VEVENT").slice(1);
  for (const block of blocks) {
    const body = block.split("END:VEVENT")[0];
    const dtstart = body.match(/DTSTART[^:]*:(\d{8})/);
    const dtend = body.match(/DTEND[^:]*:(\d{8})/);
    const uid = body.match(/UID:([^\r\n]+)/);
    if (dtstart) {
      const s = dtstart[1];
      const e = dtend ? dtend[1] : s;
      events.push({
        start: `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`,
        end: `${e.slice(0, 4)}-${e.slice(4, 6)}-${e.slice(6, 8)}`,
        uid: uid ? uid[1].trim() : Math.random().toString(36).slice(2),
      });
    }
  }
  return events;
}

async function fetchICalBlocks(url) {
  const attempts = [
    {
      name: "allorigins",
      url: "https://api.allorigins.win/get?url=" + encodeURIComponent(url),
      parse: async (res) => {
        const j = await res.json();
        if (!j || typeof j.contents !== "string") throw new Error("format de réponse inattendu");
        return j.contents;
      },
    },
    {
      name: "corsproxy.io",
      url: "https://corsproxy.io/?url=" + encodeURIComponent(url),
      parse: async (res) => res.text(),
    },
    {
      name: "codetabs",
      url: "https://api.codetabs.com/v1/proxy?quest=" + encodeURIComponent(url),
      parse: async (res) => res.text(),
    },
  ];
  const errors = [];
  for (const attempt of attempts) {
    try {
      const res = await fetch(attempt.url);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const text = await attempt.parse(res);
      if (!text || !text.includes("BEGIN:VCALENDAR")) throw new Error("contenu invalide (pas un calendrier iCal)");
      return parseICS(text);
    } catch (e) {
      errors.push(attempt.name + " : " + e.message);
    }
  }
  throw new Error(errors.join(" | "));
}

function hashPassword(pw) {
  // Simple, synchronous, dependency-free obfuscation (not cryptographic security — this app has no
  // real backend, so no client-side scheme here would be a true security boundary anyway; see the
  // caveat given to the user). Synchronous and free of any browser API that could be unavailable.
  const s = String(pw == null ? "" : pw);
  try {
    return btoa(unescape(encodeURIComponent(s)));
  } catch (e) {
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
    return "fb_" + Math.abs(h).toString(16);
  }
}

function genPassword() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789";
  let out = "";
  for (let i = 0; i < 10; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

function genUsername(name) {
  const base = (name || "user").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
  return (base.slice(0, 10) || "user") + Math.floor(100 + Math.random() * 900);
}

function useCountUp(target, duration = 700) {
  const [val, setVal] = useState(0);
  const raf = useRef(null);
  useEffect(() => {
    const start = performance.now();
    cancelAnimationFrame(raf.current);
    function tick(now) {
      const t = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - t, 3);
      setVal(target * eased);
      if (t < 1) raf.current = requestAnimationFrame(tick);
    }
    raf.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf.current);
  }, [target, duration]);
  return val;
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12, color: C.ivoryFaint, flex: 1, minWidth: 120 }}>
      {label}
      {children}
    </label>
  );
}

const inputStyle = {
  background: C.inkPanel2,
  border: `1px solid ${C.line}`,
  borderRadius: 4,
  color: C.ivory,
  padding: "8px 10px",
  fontSize: 14,
  fontFamily: "'Inter', sans-serif",
  outline: "none",
};

function Button({ children, onClick, variant = "ghost", type = "button", style }) {
  const base = {
    fontFamily: "'Inter', sans-serif",
    fontSize: 13,
    padding: "9px 16px",
    borderRadius: 4,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    border: `1px solid ${variant === "brass" ? C.brass : C.line}`,
    background: variant === "brass" ? C.brass : "transparent",
    color: variant === "brass" ? C.ink : C.ivory,
    transition: "opacity 0.15s",
  };
  return (
    <button type={type} onClick={onClick} style={{ ...base, ...style }}
      onMouseEnter={(e) => (e.currentTarget.style.opacity = "0.8")}
      onMouseLeave={(e) => (e.currentTarget.style.opacity = "1")}>
      {children}
    </button>
  );
}

function LoginScreen({ onLogin, usersCount }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [trace, setTrace] = useState([]);

  function log(msg) {
    setTrace((t) => [...t.slice(-6), msg]);
  }

  async function submit(e) {
    if (e && e.preventDefault) e.preventDefault();
    log("clic reçu");
    setLoading(true);
    setError("");
    try {
      log("vérification en cours…");
      const ok = await onLogin(username.trim(), password);
      log("résultat : " + (ok ? "succès" : "échec (identifiants non reconnus)"));
      if (!ok) setError("Identifiant ou mot de passe incorrect.");
    } catch (err) {
      log("erreur : " + (err && err.message ? err.message : String(err)));
      setError("Une erreur est survenue pendant la connexion. Réessayez.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ background: C.ink, minHeight: 500, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", border: `1px solid ${C.lineStrong}`, padding: 24 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .bl-serif { font-family: 'Cormorant Garamond', serif; }
      `}</style>
      <div onKeyDown={(e) => { if (e.key === "Enter") submit(e); }} className="bl-fade" style={{ width: "100%", maxWidth: 320, display: "flex", flexDirection: "column", alignItems: "center", gap: 18, fontFamily: "'Inter', sans-serif" }}>
        <div style={{ textAlign: "center" }}>
          <div className="bl-serif" style={{ fontSize: 22, color: C.ivory, fontWeight: 600 }}>Blackland</div>
          <div style={{ fontSize: 11, color: C.ivoryFaint, letterSpacing: 1, textTransform: "uppercase", marginTop: 2 }}>Registre de transparence</div>
        </div>
        <div style={{ width: "100%", display: "flex", flexDirection: "column", gap: 12 }}>
          <Field label="Identifiant">
            <input style={inputStyle} value={username} onChange={(e) => setUsername(e.target.value)} autoFocus />
          </Field>
          <Field label="Mot de passe">
            <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} />
          </Field>
        </div>
        {error && <ErrorLine text={error} />}
        <Button type="button" variant="brass" onClick={submit} style={{ width: "100%", justifyContent: "center" }}>
          <Lock size={13} /> {loading ? "Connexion…" : "Se connecter"}
        </Button>
        <div style={{ width: "100%", background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 4, padding: 10, fontSize: 10.5, color: C.ivoryFaint, fontFamily: "'IBM Plex Mono', monospace" }}>
          <div>diagnostic — utilisateurs chargés : {usersCount}</div>
          {trace.length === 0 ? <div>(aucune action détectée pour l'instant)</div> : trace.map((t, i) => <div key={i}>{t}</div>)}
        </div>
      </div>
    </div>
  );
}

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("Blackland app error:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ background: C.ink, color: C.ivory, padding: 24, borderRadius: 10, border: `1px solid ${C.lineStrong}`, fontFamily: "'Inter', sans-serif" }}>
          <div style={{ color: C.wine, fontWeight: 600, marginBottom: 8 }}>Une erreur est survenue dans l'application.</div>
          <div style={{ fontSize: 12, color: C.ivoryFaint, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "pre-wrap" }}>
            {String((this.state.error && this.state.error.message) || this.state.error)}
          </div>
          <div style={{ fontSize: 12, color: C.ivoryFaint, marginTop: 10 }}>
            Copiez ce message et envoyez-le pour qu'on corrige le problème.
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

function BlacklandAppInner() {
  const [properties, setProperties] = useState([]);
  const [reservations, setReservations] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [ready, setReady] = useState(false);
  const [view, setView] = useState("admin");
  const [adminPropId, setAdminPropId] = useState(null);
  const [ownerPropId, setOwnerPropId] = useState(null);
  const [adminTab, setAdminTab] = useState("reservations");
  const [showResForm, setShowResForm] = useState(false);
  const [showExpForm, setShowExpForm] = useState(false);
  const [icalBlocks, setIcalBlocks] = useState({});
  const [syncStatus, setSyncStatus] = useState({});
  const [qrUrl, setQrUrlState] = useState("");
  const [users, setUsers] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);

  useEffect(() => {
    (async () => {
      let props = SEED_PROPERTIES;
      try {
        const p = await window.storage.get("bl_properties", true);
        if (p && p.value) props = JSON.parse(p.value);
        else await window.storage.set("bl_properties", JSON.stringify(SEED_PROPERTIES), true);
      } catch (e) {
        try { await window.storage.set("bl_properties", JSON.stringify(SEED_PROPERTIES), true); } catch (e2) {}
      }
      setProperties(props);
      setAdminPropId(props[0]?.id ?? null);
      setOwnerPropId(props[0]?.id ?? null);

      try {
        const r = await window.storage.get("bl_reservations", true);
        setReservations(r && r.value ? JSON.parse(r.value) : []);
      } catch (e) { setReservations([]); }

      try {
        const ex = await window.storage.get("bl_expenses", true);
        setExpenses(ex && ex.value ? JSON.parse(ex.value) : []);
      } catch (e) { setExpenses([]); }

      try {
        const ic = await window.storage.get("bl_ical_blocks", true);
        setIcalBlocks(ic && ic.value ? JSON.parse(ic.value) : {});
      } catch (e) { setIcalBlocks({}); }

      try {
        const su = await window.storage.get("bl_share_url", true);
        if (su && su.value) setQrUrlState(su.value);
        else if (typeof window !== "undefined" && window.location && window.location.href) {
          setQrUrlState(window.location.href);
        }
      } catch (e) {
        try { if (window.location && window.location.href) setQrUrlState(window.location.href); } catch (e2) {}
      }

      let loadedUsers = [];
      try {
        const us = await window.storage.get("bl_users_v2", true);
        loadedUsers = us && us.value ? JSON.parse(us.value) : [];
      } catch (e) { loadedUsers = []; }

      if (loadedUsers.length === 0) {
        const adminHash = await hashPassword("blackland2026");
        const ownerHash1 = await hashPassword("bien2plus1");
        const ownerHash2 = await hashPassword("penthouse26");
        loadedUsers = [
          { id: "u_admin", username: "admin", passwordHash: adminHash, name: "Administrateur Blackland", role: "admin", propertyIds: [] },
          { id: "u_owner1", username: "proprio1", passwordHash: ownerHash1, name: "Propriétaire — 2+1", role: "owner", propertyIds: ["p1"] },
          { id: "u_owner2", username: "proprio2", passwordHash: ownerHash2, name: "Propriétaire — Penthouse", role: "owner", propertyIds: ["p2"] },
        ];
        try { await window.storage.set("bl_users_v2", JSON.stringify(loadedUsers), true); } catch (e) {}
      }
      setUsers(loadedUsers);

      try {
        const sess = await window.storage.get("bl_session", false);
        if (sess && sess.value) {
          const { userId } = JSON.parse(sess.value);
          const found = loadedUsers.find((u) => u.id === userId);
          if (found) {
            setCurrentUser(found);
            setOwnerPropId(found.propertyIds?.[0] ?? props[0]?.id ?? null);
          }
        }
      } catch (e) {}

      setReady(true);
    })();
  }, []);

  async function persist(key, value, setter) {
    setter(value);
    try { await window.storage.set(key, JSON.stringify(value), true); } catch (e) { console.error("storage error", e); }
  }

  function addReservation(data) {
    const next = [...reservations, { ...data, id: uid() }];
    persist("bl_reservations", next, setReservations);
  }
  function deleteReservation(id) {
    persist("bl_reservations", reservations.filter((r) => r.id !== id), setReservations);
  }
  function addExpense(data) {
    const next = [...expenses, { ...data, id: uid() }];
    persist("bl_expenses", next, setExpenses);
  }
  function deleteExpense(id) {
    persist("bl_expenses", expenses.filter((e) => e.id !== id), setExpenses);
  }
  function updateCommission(propId, pct) {
    const next = properties.map((p) => p.id === propId ? { ...p, commissionPct: pct } : p);
    persist("bl_properties", next, setProperties);
  }

  function updateIcalUrl(propId, url) {
    const next = properties.map((p) => p.id === propId ? { ...p, icalUrl: url } : p);
    persist("bl_properties", next, setProperties);
  }

  async function syncIcal(propId) {
    const prop = properties.find((p) => p.id === propId);
    if (!prop?.icalUrl) return;
    setSyncStatus((s) => ({ ...s, [propId]: { loading: true, error: null } }));
    try {
      const blocks = await fetchICalBlocks(prop.icalUrl);
      const prevDismissed = icalBlocks[propId]?.dismissed || [];
      const next = { ...icalBlocks, [propId]: { blocks, lastSync: new Date().toISOString(), dismissed: prevDismissed } };
      await persist("bl_ical_blocks", next, setIcalBlocks);
      setSyncStatus((s) => ({ ...s, [propId]: { loading: false, error: null } }));
    } catch (e) {
      setSyncStatus((s) => ({ ...s, [propId]: { loading: false, error: "Synchro automatique impossible (" + e.message + "). Utilisez le collage manuel ci-dessous." } }));
    }
  }

  function importIcalText(propId, text) {
    try {
      const blocks = parseICS(text);
      const prevDismissed = icalBlocks[propId]?.dismissed || [];
      const next = { ...icalBlocks, [propId]: { blocks, lastSync: new Date().toISOString(), manual: true, dismissed: prevDismissed } };
      persist("bl_ical_blocks", next, setIcalBlocks);
      setSyncStatus((s) => ({ ...s, [propId]: { loading: false, error: null } }));
      return true;
    } catch (e) {
      setSyncStatus((s) => ({ ...s, [propId]: { loading: false, error: "Le contenu collé n'a pas pu être lu." } }));
      return false;
    }
  }

  function dismissBlock(propId, uid) {
    const current = icalBlocks[propId] || { blocks: [], dismissed: [] };
    const nextDismissed = [...(current.dismissed || []), uid];
    const next = { ...icalBlocks, [propId]: { ...current, dismissed: nextDismissed } };
    persist("bl_ical_blocks", next, setIcalBlocks);
  }

  function setQrUrl(url) {
    persist("bl_share_url", url, setQrUrlState);
  }

  async function login(username, password) {
    try {
      const hash = await hashPassword(password);
      const match = users.find((u) => u.username.toLowerCase() === (username || "").toLowerCase() && u.passwordHash === hash);
      if (!match) return false;
      setCurrentUser(match);
      setOwnerPropId(match.propertyIds?.[0] ?? properties[0]?.id ?? null);
      try { await window.storage.set("bl_session", JSON.stringify({ userId: match.id }), false); } catch (e) {}
      return true;
    } catch (e) {
      console.error("login error", e);
      return false;
    }
  }

  async function logout() {
    setCurrentUser(null);
    try { await window.storage.delete("bl_session", false); } catch (e) {}
  }

  async function addUser({ name, role, propertyIds }) {
    const username = genUsername(name);
    const password = genPassword();
    const passwordHash = await hashPassword(password);
    const newUser = { id: uid(), username, passwordHash, name, role, propertyIds: propertyIds || [] };
    const next = [...users, newUser];
    await persist("bl_users_v2", next, setUsers);
    return { username, password };
  }

  function deleteUser(id) {
    persist("bl_users_v2", users.filter((u) => u.id !== id), setUsers);
  }

  function updatePropertyAccess(userId, propertyIds) {
    const next = users.map((u) => u.id === userId ? { ...u, propertyIds } : u);
    persist("bl_users_v2", next, setUsers);
  }

  function addProperty({ name, address }) {
    const newProp = { id: uid(), name: name || "Nouveau bien", address: address || "", ownerName: "", commissionPct: 15, icalUrl: "" };
    const next = [...properties, newProp];
    persist("bl_properties", next, setProperties);
    setAdminPropId(newProp.id);
  }

  function updatePropertyMeta(propId, patch) {
    const next = properties.map((p) => p.id === propId ? { ...p, ...patch } : p);
    persist("bl_properties", next, setProperties);
  }

  function statsFor(propId) {
    const res = reservations.filter((r) => r.propertyId === propId);
    const exp = expenses.filter((e) => e.propertyId === propId);
    const revenue = res.reduce((s, r) => s + Number(r.amount || 0), 0);
    const cleaning = res.reduce((s, r) => s + Number(r.cleaningFee || 0), 0);
    const other = exp.reduce((s, e) => s + Number(e.amount || 0), 0);
    const prop = properties.find((p) => p.id === propId);
    const pct = prop?.commissionPct ?? 0;
    const commission = revenue * (pct / 100);
    const net = revenue - cleaning - other - commission;
    return { revenue, cleaning, other, commission, net, pct, count: res.length, res, exp };
  }

  if (!ready) {
    return (
      <div style={{ background: C.ink, minHeight: 400, display: "flex", alignItems: "center", justifyContent: "center", borderRadius: 8 }}>
        <span style={{ color: C.ivoryFaint, fontFamily: "'Inter', sans-serif", fontSize: 13 }}>Chargement du registre…</span>
      </div>
    );
  }

  if (!currentUser) {
    return <LoginScreen onLogin={login} usersCount={users.length} />;
  }

  const ownerVisibleProperties = currentUser.role === "admin" ? properties : properties.filter((p) => (currentUser.propertyIds || []).includes(p.id));

  return (
    <div style={{
      background: C.ink, minHeight: 500, borderRadius: 10, overflow: "hidden",
      fontFamily: "'Inter', sans-serif", color: C.ivory, border: `1px solid ${C.lineStrong}`,
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .bl-mono { font-family: 'IBM Plex Mono', monospace; }
        .bl-serif { font-family: 'Cormorant Garamond', serif; }
        .bl-fade { animation: blFadeIn 0.4s ease both; }
        @keyframes blFadeIn { from { opacity: 0; transform: translateY(4px);} to { opacity:1; transform:translateY(0);} }
        .bl-row:hover { background: rgba(243,238,227,0.03); }
        select.bl-select { appearance: none; -webkit-appearance:none; }
      `}</style>

      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "16px 24px", borderBottom: `1px solid ${C.line}`, flexWrap: "wrap", gap: 12,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div>
            <div className="bl-serif" style={{ fontSize: 20, fontWeight: 600, letterSpacing: 0.5 }}>Blackland</div>
            <div style={{ fontSize: 11, color: C.ivoryFaint, letterSpacing: 1, textTransform: "uppercase" }}>Registre de transparence</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {currentUser.role === "admin" && (
            <div style={{ display: "flex", gap: 4, background: C.inkPanel, borderRadius: 6, padding: 4, border: `1px solid ${C.line}` }}>
              {[{ k: "admin", label: "Espace admin" }, { k: "owner", label: "Aperçu propriétaire" }].map((t) => (
                <button key={t.k} onClick={() => setView(t.k)} style={{
                  padding: "7px 14px", fontSize: 12.5, borderRadius: 4, border: "none", cursor: "pointer",
                  background: view === t.k ? C.brass : "transparent",
                  color: view === t.k ? C.ink : C.ivoryDim,
                  fontWeight: view === t.k ? 600 : 400,
                }}>{t.label}</button>
              ))}
            </div>
          )}
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: C.ivoryFaint }}>
            <span>{currentUser.name}</span>
            <Button onClick={logout}><LogOut size={13} /> Déconnexion</Button>
          </div>
        </div>
      </div>

      {(currentUser.role === "admin" && view === "admin") ? (
        <AdminView
          properties={properties}
          adminPropId={adminPropId}
          setAdminPropId={setAdminPropId}
          adminTab={adminTab}
          setAdminTab={setAdminTab}
          stats={adminPropId ? statsFor(adminPropId) : null}
          showResForm={showResForm}
          setShowResForm={setShowResForm}
          showExpForm={showExpForm}
          setShowExpForm={setShowExpForm}
          addReservation={addReservation}
          deleteReservation={deleteReservation}
          addExpense={addExpense}
          deleteExpense={deleteExpense}
          updateCommission={updateCommission}
          updateIcalUrl={updateIcalUrl}
          syncIcal={syncIcal}
          importIcalText={importIcalText}
          syncStatus={adminPropId ? syncStatus[adminPropId] : null}
          icalData={adminPropId ? icalBlocks[adminPropId] : null}
          dismissBlock={dismissBlock}
          qrUrl={qrUrl}
          setQrUrl={setQrUrl}
          users={users}
          addUser={addUser}
          deleteUser={deleteUser}
          updatePropertyAccess={updatePropertyAccess}
          addProperty={addProperty}
          updatePropertyMeta={updatePropertyMeta}
        />
      ) : (
        <OwnerView
          properties={ownerVisibleProperties}
          ownerPropId={ownerPropId}
          setOwnerPropId={setOwnerPropId}
          stats={ownerPropId ? statsFor(ownerPropId) : null}
          icalData={ownerPropId ? icalBlocks[ownerPropId] : null}
        />
      )}
    </div>
  );
}

function PropertyPicker({ properties, value, onChange }) {
  return (
    <div style={{ position: "relative", display: "inline-block" }}>
      <select className="bl-select" value={value || ""} onChange={(e) => onChange(e.target.value)} style={{
        ...inputStyle, paddingRight: 32, cursor: "pointer", fontSize: 13.5,
      }}>
        {properties.map((p) => <option key={p.id} value={p.id} style={{ background: C.inkPanel }}>{p.name}</option>)}
      </select>
      <ChevronDown size={14} style={{ position: "absolute", right: 10, top: 10, color: C.ivoryFaint, pointerEvents: "none" }} />
    </div>
  );
}

function StatCard({ label, value, tone }) {
  const color = tone === "up" ? C.moss : tone === "down" ? C.wine : C.ivory;
  return (
    <div style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: "14px 16px", flex: 1, minWidth: 140 }}>
      <div style={{ fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 6 }}>{label}</div>
      <div className="bl-mono" style={{ fontSize: 19, fontWeight: 500, color }}>{value}</div>
    </div>
  );
}

function AdminView(props) {
  const {
    properties, adminPropId, setAdminPropId, adminTab, setAdminTab, stats,
    showResForm, setShowResForm, showExpForm, setShowExpForm,
    addReservation, deleteReservation, addExpense, deleteExpense, updateCommission,
    updateIcalUrl, syncIcal, importIcalText, syncStatus, icalData, dismissBlock,
    qrUrl, setQrUrl,
    users, addUser, deleteUser, updatePropertyAccess, addProperty, updatePropertyMeta,
  } = props;
  const prop = properties.find((p) => p.id === adminPropId);
  const [showShare, setShowShare] = useState(false);
  const [showUsers, setShowUsers] = useState(false);
  const [showNewProp, setShowNewProp] = useState(false);
  const [resPrefill, setResPrefill] = useState(null);

  function handleCompleteBlock(block) {
    setResPrefill({ checkIn: block.start, checkOut: block.end });
    setAdminTab("reservations");
    setShowResForm(true);
  }

  return (
    <div className="bl-fade" style={{ padding: 24 }}>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <Button onClick={() => setShowUsers((s) => !s)}>
          <Users size={14} />
          {showUsers ? "Masquer les utilisateurs" : "Gérer les utilisateurs"}
        </Button>
        <Button onClick={() => setShowShare((s) => !s)}>
          <Share2 size={14} />
          {showShare ? "Masquer le QR investisseurs" : "QR investisseurs"}
        </Button>
      </div>
      {showUsers && (
        <UsersPanel users={users} properties={properties} addUser={addUser} deleteUser={deleteUser} updatePropertyAccess={updatePropertyAccess} />
      )}
      {showShare && <SharePanel qrUrl={qrUrl} setQrUrl={setQrUrl} />}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <Building2 size={16} color={C.brassLight} />
          <PropertyPicker properties={properties} value={adminPropId} onChange={setAdminPropId} />
          <Button onClick={() => setShowNewProp((s) => !s)}>
            {showNewProp ? <X size={13} /> : <Plus size={13} />}
            {showNewProp ? "Annuler" : "Ajouter un bien"}
          </Button>
        </div>
        {prop && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: C.ivoryFaint }}>
            Commission de gestion
            <input type="number" min="0" max="100" value={prop.commissionPct}
              onChange={(e) => updateCommission(prop.id, Number(e.target.value))}
              style={{ ...inputStyle, width: 56, padding: "5px 8px", textAlign: "center" }} />
            %
          </div>
        )}
      </div>

      {showNewProp && <NewPropertyForm onSubmit={(d) => { addProperty(d); setShowNewProp(false); }} />}

      {prop && (
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
          <Field label="Nom du bien">
            <input style={inputStyle} value={prop.name} onChange={(e) => updatePropertyMeta(prop.id, { name: e.target.value })} />
          </Field>
          <Field label="Localisation">
            <input style={inputStyle} value={prop.address} onChange={(e) => updatePropertyMeta(prop.id, { address: e.target.value })} />
          </Field>
        </div>
      )}

      {prop && (
        <div style={{ fontSize: 12, color: C.ivoryFaint, marginBottom: 18 }}>
          Propriétaire(s) : {users.filter((u) => u.role === "owner" && u.propertyIds.includes(prop.id)).map((u) => u.name).join(", ") || "aucun utilisateur assigné — gérez les accès dans « Gérer les utilisateurs »"}
        </div>
      )}

      {stats && (
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 22 }}>
          <StatCard label="Revenu réservations" value={eur(stats.revenue)} />
          <StatCard label="Ménage" value={"− " + eur(stats.cleaning)} tone="down" />
          <StatCard label="Charges (élec/eau/wifi)" value={"− " + eur(stats.other)} tone="down" />
          <StatCard label={`Commission Blackland (${stats.pct}%)`} value={"− " + eur(stats.commission)} tone="down" />
          <StatCard label="Net propriétaire" value={eur(stats.net)} tone="up" />
        </div>
      )}

      <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${C.line}`, marginBottom: 16 }}>
        {[{ k: "reservations", label: "Réservations" }, { k: "expenses", label: "Charges" }, { k: "calendar", label: "Calendrier" }].map((t) => (
          <button key={t.k} onClick={() => setAdminTab(t.k)} style={{
            background: "none", border: "none", cursor: "pointer",
            padding: "8px 4px", marginRight: 20, fontSize: 13.5,
            color: adminTab === t.k ? C.ivory : C.ivoryFaint,
            borderBottom: adminTab === t.k ? `2px solid ${C.brass}` : "2px solid transparent",
          }}>{t.label}</button>
        ))}
      </div>

      {adminTab === "reservations" ? (
        <div>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 10 }}>
            <Button variant="brass" onClick={() => setShowResForm((s) => !s)}>
              {showResForm ? <X size={14} /> : <Plus size={14} />}
              {showResForm ? "Fermer" : "Ajouter une réservation"}
            </Button>
          </div>
          {showResForm && <ReservationForm propertyId={adminPropId} initial={resPrefill} onSubmit={(d) => { addReservation(d); setShowResForm(false); setResPrefill(null); }} />}
          <ReservationTable list={stats?.res || []} onDelete={deleteReservation} />
        </div>
      ) : adminTab === "expenses" ? (
        <div>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 10 }}>
            <Button variant="brass" onClick={() => setShowExpForm((s) => !s)}>
              {showExpForm ? <X size={14} /> : <Plus size={14} />}
              {showExpForm ? "Fermer" : "Ajouter une charge"}
            </Button>
          </div>
          {showExpForm && <ExpenseForm propertyId={adminPropId} onSubmit={(d) => { addExpense(d); setShowExpForm(false); }} />}
          <ExpenseTable list={stats?.exp || []} onDelete={deleteExpense} />
        </div>
      ) : (
        <AdminCalendarPanel
          prop={properties.find((p) => p.id === adminPropId)}
          stats={stats}
          updateIcalUrl={updateIcalUrl}
          syncIcal={syncIcal}
          importIcalText={importIcalText}
          syncStatus={syncStatus}
          icalData={icalData}
          onCompleteBlock={handleCompleteBlock}
          onDismissBlock={dismissBlock}
        />
      )}
    </div>
  );
}

function NewPropertyForm({ onSubmit }) {
  const [name, setName] = useState("");
  const [address, setAddress] = useState("");
  const [error, setError] = useState("");

  function submit(e) {
    e.preventDefault();
    if (!name.trim()) { setError("Merci d'indiquer un nom pour ce bien."); return; }
    setError("");
    onSubmit({ name: name.trim(), address: address.trim() });
    setName("");
    setAddress("");
  }

  return (
    <div onKeyDown={(e) => { if (e.key === "Enter") submit(e); }} className="bl-fade" style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 16, marginBottom: 16, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
      <Field label="Nom du bien"><input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="Ex. Studio La Joliette" /></Field>
      <Field label="Localisation"><input style={inputStyle} value={address} onChange={(e) => setAddress(e.target.value)} placeholder="Ex. Marseille" /></Field>
      <Button type="button" variant="brass" onClick={submit}>Créer le bien</Button>
      <ErrorLine text={error} />
    </div>
  );
}

function UsersPanel({ users, properties, addUser, deleteUser, updatePropertyAccess }) {
  const [showForm, setShowForm] = useState(false);
  const [name, setName] = useState("");
  const [role, setRole] = useState("owner");
  const [selectedProps, setSelectedProps] = useState([]);
  const [generated, setGenerated] = useState(null);
  const [copied, setCopied] = useState(false);

  async function submit(e) {
    e.preventDefault();
    if (!name.trim()) return;
    const creds = await addUser({ name: name.trim(), role, propertyIds: role === "owner" ? selectedProps : [] });
    setGenerated(creds);
    setName("");
    setRole("owner");
    setSelectedProps([]);
    setShowForm(false);
  }

  function toggleProp(id) {
    setSelectedProps((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  }

  function copyCreds() {
    if (!generated) return;
    const text = `Identifiant : ${generated.username}\nMot de passe : ${generated.password}`;
    if (navigator.clipboard?.writeText) navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  return (
    <div className="bl-fade" style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 20, marginBottom: 22 }}>
      <div className="bl-serif" style={{ fontSize: 17, marginBottom: 12 }}>Utilisateurs</div>

      <div style={{ border: `1px solid ${C.line}`, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1.2fr 1fr 1.6fr 32px", padding: "10px 14px", fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 0.5, borderBottom: `1px solid ${C.line}` }}>
          <div>Nom</div><div>Rôle</div><div>Biens accessibles</div><div />
        </div>
        {users.map((u) => (
          <div key={u.id} className="bl-row" style={{ display: "grid", gridTemplateColumns: "1.2fr 1fr 1.6fr 32px", padding: "10px 14px", fontSize: 13, borderBottom: `1px solid ${C.line}`, alignItems: "center" }}>
            <div>{u.name} <span style={{ color: C.ivoryFaint, fontSize: 11.5 }}>@{u.username}</span></div>
            <div style={{ color: C.ivoryFaint }}>{u.role === "admin" ? "Admin" : "Propriétaire"}</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {u.role === "admin" ? (
                <span style={{ fontSize: 11.5, color: C.ivoryFaint }}>Tous les biens</span>
              ) : (
                properties.map((p) => {
                  const active = u.propertyIds.includes(p.id);
                  return (
                    <button key={p.id} onClick={() => updatePropertyAccess(u.id, active ? u.propertyIds.filter((x) => x !== p.id) : [...u.propertyIds, p.id])} style={{
                      fontSize: 11, padding: "3px 8px", borderRadius: 12, cursor: "pointer",
                      border: `1px solid ${active ? C.brass : C.line}`,
                      background: active ? C.brass : "transparent",
                      color: active ? C.ink : C.ivoryFaint,
                    }}>{p.name}</button>
                  );
                })
              )}
            </div>
            <button onClick={() => deleteUser(u.id)} style={{ background: "none", border: "none", cursor: "pointer", color: C.ivoryFaint }}><Trash2 size={14} /></button>
          </div>
        ))}
      </div>

      {generated && (
        <div className="bl-fade" style={{ background: C.inkPanel2, border: `1px solid ${C.brass}`, borderRadius: 6, padding: 14, marginBottom: 16 }}>
          <div style={{ fontSize: 12.5, color: C.ivoryDim, marginBottom: 8 }}>
            Identifiants générés — notez-les, ils ne seront plus affichés en clair ensuite :
          </div>
          <div className="bl-mono" style={{ fontSize: 14, color: C.ivory, marginBottom: 10 }}>
            {generated.username} / {generated.password}
          </div>
          <Button variant="brass" onClick={copyCreds}>
            {copied ? <Check size={13} /> : <Copy size={13} />} {copied ? "Copié" : "Copier"}
          </Button>
        </div>
      )}

      <button onClick={() => setShowForm((s) => !s)} style={{ background: "none", border: "none", color: C.brassLight, fontSize: 12.5, cursor: "pointer", padding: 0, display: "flex", alignItems: "center", gap: 6 }}>
        <Plus size={13} /> {showForm ? "Annuler" : "Nouvel utilisateur"}
      </button>

      {showForm && (
        <div onKeyDown={(e) => { if (e.key === "Enter") submit(e); }} className="bl-fade" style={{ marginTop: 12, display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            <Field label="Nom"><input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="Ex. M. Fabre" /></Field>
            <Field label="Rôle">
              <select className="bl-select" style={inputStyle} value={role} onChange={(e) => setRole(e.target.value)}>
                <option value="owner" style={{ background: C.inkPanel }}>Propriétaire</option>
                <option value="admin" style={{ background: C.inkPanel }}>Admin</option>
              </select>
            </Field>
          </div>
          {role === "owner" && (
            <div>
              <div style={{ fontSize: 12, color: C.ivoryFaint, marginBottom: 6 }}>Biens accessibles</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {properties.map((p) => {
                  const active = selectedProps.includes(p.id);
                  return (
                    <button type="button" key={p.id} onClick={() => toggleProp(p.id)} style={{
                      fontSize: 11.5, padding: "4px 10px", borderRadius: 12, cursor: "pointer",
                      border: `1px solid ${active ? C.brass : C.line}`,
                      background: active ? C.brass : "transparent",
                      color: active ? C.ink : C.ivoryFaint,
                    }}>{p.name}</button>
                  );
                })}
              </div>
            </div>
          )}
          <div><Button type="button" variant="brass" onClick={submit}>Générer les identifiants</Button></div>
        </div>
      )}
    </div>
  );
}

function SharePanel({ qrUrl, setQrUrl }) {
  const [draft, setDraft] = useState(qrUrl || "");
  useEffect(() => { setDraft(qrUrl || ""); }, [qrUrl]);
  const qrSrc = draft ? `https://api.qrserver.com/v1/create-qr-code/?size=400x400&margin=12&data=${encodeURIComponent(draft)}` : null;

  return (
    <div className="bl-fade" style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 20, marginBottom: 22, display: "flex", gap: 24, flexWrap: "wrap", alignItems: "center" }}>
      <div style={{ flex: 1, minWidth: 240 }}>
        <div className="bl-serif" style={{ fontSize: 17, marginBottom: 8 }}>Partager avec vos investisseurs</div>
        <div style={{ fontSize: 12.5, color: C.ivoryFaint, marginBottom: 14, lineHeight: 1.5 }}>
          Publiez d'abord cet artefact via le bouton "Share" en haut de la fenêtre Claude pour obtenir un lien public, puis collez-le ci-dessous. Le QR code se génère automatiquement. Une fois le lien ouvert sur leur téléphone, vos investisseurs peuvent utiliser "Ajouter à l'écran d'accueil" pour l'installer comme une application.
        </div>
        <Field label="Lien public de l'application">
          <input style={inputStyle} value={draft} onChange={(e) => setDraft(e.target.value)} placeholder="https://claude.ai/public/artifacts/..." />
        </Field>
        <div style={{ marginTop: 10 }}>
          <Button variant="brass" onClick={() => setQrUrl(draft)}>Enregistrer ce lien</Button>
        </div>
      </div>
      <div style={{ background: "#FFFFFF", borderRadius: 8, padding: 16, display: "flex", alignItems: "center", justifyContent: "center", minWidth: 180 }}>
        {qrSrc ? (
          <img src={qrSrc} alt="QR code" style={{ width: 180, height: 180, display: "block" }} />
        ) : (
          <div style={{ width: 180, height: 180, display: "flex", alignItems: "center", justifyContent: "center", color: "#999", fontSize: 12, textAlign: "center", padding: 12 }}>
            Collez un lien pour générer le QR code
          </div>
        )}
      </div>
    </div>
  );
}

function AdminCalendarPanel({ prop, stats, updateIcalUrl, syncIcal, importIcalText, syncStatus, icalData, onCompleteBlock, onDismissBlock }) {
  const [pasteText, setPasteText] = useState("");
  const [copied, setCopied] = useState(false);
  if (!prop) return null;
  const loading = syncStatus?.loading;
  const error = syncStatus?.error;

  const existingRanges = (stats?.res || []).map((r) => r.checkIn);
  const dismissed = icalData?.dismissed || [];
  const blocksToComplete = (icalData?.blocks || []).filter((b) => !existingRanges.includes(b.start) && !dismissed.includes(b.uid));

  function copyUrl() {
    if (!prop.icalUrl) return;
    if (navigator.clipboard?.writeText) navigator.clipboard.writeText(prop.icalUrl);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  return (
    <div>
      <div style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 16, marginBottom: 16 }}>
        <Field label="Lien iCal Airbnb de ce bien">
          <input style={inputStyle} value={prop.icalUrl || ""} onChange={(e) => updateIcalUrl(prop.id, e.target.value)} placeholder="https://www.airbnb.fr/calendar/ical/..." />
        </Field>

        <div style={{ marginTop: 14, padding: 12, background: C.inkPanel2, borderRadius: 6, border: `1px solid ${C.line}` }}>
          <div style={{ fontSize: 12.5, color: C.ivoryDim, marginBottom: 10, lineHeight: 1.5 }}>
            La synchronisation automatique est bloquée ici par une règle de sécurité de la plateforme (elle empêche l'application de contacter un site externe) — ce n'est pas un bug côté Blackland et ça ne se réglera pas avec une nouvelle tentative. La méthode ci-dessous fonctionne à 100 % et prend 10 secondes :
          </div>
          <ol style={{ fontSize: 12.5, color: C.ivoryDim, margin: "0 0 12px 18px", padding: 0, lineHeight: 1.7 }}>
            <li>Copiez le lien ci-dessus <Button onClick={copyUrl} style={{ padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{copied ? <Check size={11} /> : <Copy size={11} />} {copied ? "Copié" : "copier"}</Button></li>
            <li>Ouvrez-le dans un nouvel onglet de votre navigateur</li>
            <li>Sélectionnez tout le texte affiché (Ctrl+A ou Cmd+A) puis copiez-le (Ctrl+C ou Cmd+C)</li>
            <li>Collez-le dans le champ ci-dessous et cliquez sur "Importer"</li>
          </ol>
          <textarea value={pasteText} onChange={(e) => setPasteText(e.target.value)} rows={5}
            style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11 }} placeholder="BEGIN:VCALENDAR..." />
          <div style={{ marginTop: 10 }}>
            <Button variant="brass" onClick={() => { const ok = importIcalText(prop.id, pasteText); if (ok) setPasteText(""); }}>Importer ce contenu</Button>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
          <Button onClick={() => syncIcal(prop.id)} style={{ opacity: loading ? 0.6 : 1 }}>
            <RefreshCw size={14} />
            {loading ? "Tentative…" : "Essayer la synchro automatique quand même"}
          </Button>
          {icalData?.lastSync && (
            <span style={{ fontSize: 11.5, color: C.ivoryFaint }}>
              Dernière mise à jour : {new Date(icalData.lastSync).toLocaleString("fr-FR")}{icalData.manual ? " (import manuel)" : ""}
            </span>
          )}
        </div>
        {error && <div style={{ marginTop: 10 }}><ErrorLine text={error} /></div>}
      </div>

      {blocksToComplete.length > 0 && (
        <div style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 16, marginBottom: 16 }}>
          <div style={{ fontSize: 12.5, color: C.ivoryDim, marginBottom: 10 }}>
            Séjours détectés sur Airbnb sans montant enregistré ({blocksToComplete.length}) :
          </div>
          {blocksToComplete.map((b) => (
            <div key={b.uid} className="bl-row" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 4px", borderBottom: `1px solid ${C.line}`, fontSize: 13, gap: 8 }}>
              <span>Du {fmtDate(b.start)} au {fmtDate(b.end)}</span>
              <div style={{ display: "flex", gap: 6 }}>
                <Button variant="brass" onClick={() => onCompleteBlock(b)} style={{ padding: "5px 12px", fontSize: 12 }}>Compléter cette réservation</Button>
                <Button onClick={() => onDismissBlock(prop.id, b.uid)} style={{ padding: "5px 8px", fontSize: 12 }} title="Retirer de cette liste">
                  <X size={13} />
                </Button>
              </div>
            </div>
          ))}
        </div>
      )}

      <ReservationCalendar reservations={stats?.res || []} icalBlocks={icalData?.blocks || []} />
    </div>
  );
}

function ErrorLine({ text }) {
  if (!text) return null;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, color: C.wine, fontSize: 12.5, width: "100%" }}>
      <AlertCircle size={13} /> {text}
    </div>
  );
}

function ReservationForm({ propertyId, onSubmit, initial }) {
  const [guestName, setGuestName] = useState("");
  const [checkIn, setCheckIn] = useState(initial?.checkIn || "");
  const [checkOut, setCheckOut] = useState(initial?.checkOut || "");
  const [amount, setAmount] = useState("");
  const [cleaningFee, setCleaningFee] = useState("");
  const [platform, setPlatform] = useState("Airbnb");
  const [error, setError] = useState("");

  function submit(e) {
    e.preventDefault();
    const amt = parseNum(amount);
    const clean = cleaningFee === "" ? 0 : parseNum(cleaningFee);
    if (!checkIn) { setError("Merci d'indiquer la date d'arrivée."); return; }
    if (!Number.isFinite(amt) || amt <= 0) { setError("Merci d'indiquer un montant de réservation valide (ex. 350 ou 350,50)."); return; }
    if (!Number.isFinite(clean)) { setError("Le coût de ménage n'est pas un nombre valide."); return; }
    setError("");
    onSubmit({ propertyId, guestName: guestName || "Voyageur", checkIn, checkOut, amount: amt, cleaningFee: clean, platform });
  }

  return (
    <div onKeyDown={(e) => { if (e.key === "Enter") submit(e); }} className="bl-fade" style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 16, marginBottom: 16, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
      <Field label="Voyageur"><input style={inputStyle} value={guestName} onChange={(e) => setGuestName(e.target.value)} placeholder="Nom du voyageur" /></Field>
      <Field label="Arrivée"><input type="date" style={inputStyle} value={checkIn} onChange={(e) => setCheckIn(e.target.value)} /></Field>
      <Field label="Départ"><input type="date" style={inputStyle} value={checkOut} onChange={(e) => setCheckOut(e.target.value)} /></Field>
      <Field label="Montant réservation (€)"><input type="text" inputMode="decimal" style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="350" /></Field>
      <Field label="Coût ménage (€)"><input type="text" inputMode="decimal" style={inputStyle} value={cleaningFee} onChange={(e) => setCleaningFee(e.target.value)} placeholder="45" /></Field>
      <Field label="Plateforme">
        <select className="bl-select" style={inputStyle} value={platform} onChange={(e) => setPlatform(e.target.value)}>
          {["Airbnb", "Booking", "Direct", "Autre"].map((p) => <option key={p} style={{ background: C.inkPanel }}>{p}</option>)}
        </select>
      </Field>
      <Button type="button" variant="brass" onClick={submit}>Enregistrer</Button>
      <ErrorLine text={error} />
    </div>
  );
}

function ExpenseForm({ propertyId, onSubmit }) {
  const [type, setType] = useState("electricite");
  const [amount, setAmount] = useState("");
  const [date, setDate] = useState("");
  const [label, setLabel] = useState("");
  const [error, setError] = useState("");

  function submit(e) {
    e.preventDefault();
    const amt = parseNum(amount);
    if (!date) { setError("Merci d'indiquer une date."); return; }
    if (!Number.isFinite(amt) || amt <= 0) { setError("Merci d'indiquer un montant valide (ex. 60 ou 60,50)."); return; }
    setError("");
    onSubmit({ propertyId, type, amount: amt, date, label });
  }

  return (
    <div onKeyDown={(e) => { if (e.key === "Enter") submit(e); }} className="bl-fade" style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 16, marginBottom: 16, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
      <Field label="Type de charge">
        <select className="bl-select" style={inputStyle} value={type} onChange={(e) => setType(e.target.value)}>
          {EXPENSE_TYPES.map((t) => <option key={t.id} value={t.id} style={{ background: C.inkPanel }}>{t.label}</option>)}
        </select>
      </Field>
      <Field label="Date"><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      <Field label="Montant (€)"><input type="text" inputMode="decimal" style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="60" /></Field>
      <Field label="Note (facultatif)"><input style={inputStyle} value={label} onChange={(e) => setLabel(e.target.value)} placeholder="Ex. facture EDF juin" /></Field>
      <Button type="button" variant="brass" onClick={submit}>Enregistrer</Button>
      <ErrorLine text={error} />
    </div>
  );
}

function ReservationTable({ list, onDelete }) {
  const sorted = [...list].sort((a, b) => (b.checkIn || "").localeCompare(a.checkIn || ""));
  if (sorted.length === 0) return <EmptyState text="Aucune réservation enregistrée pour ce bien." />;
  return (
    <div style={{ border: `1px solid ${C.line}`, borderRadius: 6, overflow: "hidden" }}>
      <div style={{ display: "grid", gridTemplateColumns: "1.2fr 1fr 1fr 0.9fr 0.9fr 0.8fr 32px", padding: "10px 14px", fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 0.5, borderBottom: `1px solid ${C.line}` }}>
        <div>Voyageur</div><div>Arrivée</div><div>Départ</div><div>Montant</div><div>Ménage</div><div>Plateforme</div><div />
      </div>
      {sorted.map((r) => (
        <div key={r.id} className="bl-row" style={{ display: "grid", gridTemplateColumns: "1.2fr 1fr 1fr 0.9fr 0.9fr 0.8fr 32px", padding: "10px 14px", fontSize: 13, borderBottom: `1px solid ${C.line}`, alignItems: "center" }}>
          <div>{r.guestName}</div>
          <div style={{ color: C.ivoryDim }}>{fmtDate(r.checkIn)}</div>
          <div style={{ color: C.ivoryDim }}>{fmtDate(r.checkOut)}</div>
          <div className="bl-mono" style={{ color: C.moss }}>{eur2(r.amount)}</div>
          <div className="bl-mono" style={{ color: C.wine }}>{eur2(r.cleaningFee)}</div>
          <div style={{ color: C.ivoryFaint, fontSize: 12 }}>{r.platform}</div>
          <button onClick={() => onDelete(r.id)} style={{ background: "none", border: "none", cursor: "pointer", color: C.ivoryFaint }}><Trash2 size={14} /></button>
        </div>
      ))}
    </div>
  );
}

function ExpenseTable({ list, onDelete }) {
  const sorted = [...list].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  if (sorted.length === 0) return <EmptyState text="Aucune charge enregistrée pour ce bien." />;
  return (
    <div style={{ border: `1px solid ${C.line}`, borderRadius: 6, overflow: "hidden" }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1.4fr 32px", padding: "10px 14px", fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 0.5, borderBottom: `1px solid ${C.line}` }}>
        <div>Type</div><div>Date</div><div>Montant</div><div>Note</div><div />
      </div>
      {sorted.map((e) => {
        const meta = EXPENSE_TYPES.find((t) => t.id === e.type) || EXPENSE_TYPES[3];
        const Icon = meta.icon;
        return (
          <div key={e.id} className="bl-row" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1.4fr 32px", padding: "10px 14px", fontSize: 13, borderBottom: `1px solid ${C.line}`, alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}><Icon size={13} color={C.brassLight} />{meta.label}</div>
            <div style={{ color: C.ivoryDim }}>{fmtDate(e.date)}</div>
            <div className="bl-mono" style={{ color: C.wine }}>{eur2(e.amount)}</div>
            <div style={{ color: C.ivoryFaint, fontSize: 12 }}>{e.label || "—"}</div>
            <button onClick={() => onDelete(e.id)} style={{ background: "none", border: "none", cursor: "pointer", color: C.ivoryFaint }}><Trash2 size={14} /></button>
          </div>
        );
      })}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div style={{ border: `1px dashed ${C.line}`, borderRadius: 6, padding: "28px 16px", textAlign: "center", color: C.ivoryFaint, fontSize: 13 }}>
      {text}
    </div>
  );
}

function localDateKey(d) {
  // Local calendar date as YYYY-MM-DD, deliberately NOT going through toISOString()/UTC —
  // doing so shifts dates by a day depending on the browser's timezone (the original bug).
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function buildOccupancy(reservations) {
  const map = {};
  reservations.forEach((r) => {
    if (!r.checkIn) return;
    const start = new Date(r.checkIn + "T00:00:00");
    const end = r.checkOut ? new Date(r.checkOut + "T00:00:00") : new Date(start.getTime() + 86400000);
    for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {
      map[localDateKey(d)] = r.guestName || "Réservé";
    }
  });
  return map;
}

function ReservationCalendar({ reservations, icalBlocks = [] }) {
  const today = new Date();
  const [cursor, setCursor] = useState(new Date(today.getFullYear(), today.getMonth(), 1));
  const occ = buildOccupancy(reservations);
  const extOcc = {};
  icalBlocks.forEach((b) => {
    const start = new Date(b.start + "T00:00:00");
    const end = new Date(b.end + "T00:00:00");
    for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {
      extOcc[localDateKey(d)] = true;
    }
  });

  const year = cursor.getFullYear();
  const monthIdx = cursor.getMonth();
  const first = new Date(year, monthIdx, 1);
  const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
  let startWeekday = first.getDay();
  startWeekday = startWeekday === 0 ? 6 : startWeekday - 1;

  const cells = [];
  for (let i = 0; i < startWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  const monthLabel = cursor.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  const dayLabels = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"];

  function dateKey(d) {
    return `${year}-${String(monthIdx + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  return (
    <div style={{ background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6, padding: 16 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <button onClick={() => setCursor(new Date(year, monthIdx - 1, 1))} style={{ background: "none", border: "none", cursor: "pointer", color: C.ivoryFaint }}>
          <ChevronLeft size={16} />
        </button>
        <div className="bl-serif" style={{ fontSize: 16, textTransform: "capitalize" }}>{monthLabel}</div>
        <button onClick={() => setCursor(new Date(year, monthIdx + 1, 1))} style={{ background: "none", border: "none", cursor: "pointer", color: C.ivoryFaint }}>
          <ChevronRight size={16} />
        </button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4, marginBottom: 6 }}>
        {dayLabels.map((d) => (
          <div key={d} style={{ textAlign: "center", fontSize: 10.5, color: C.ivoryFaint, textTransform: "uppercase" }}>{d}</div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4 }}>
        {cells.map((d, idx) => {
          if (d === null) return <div key={"e" + idx} />;
          const key = dateKey(d);
          const confirmed = occ[key];
          const external = !confirmed && extOcc[key];
          const label = confirmed ? confirmed : external ? "Bloqué sur Airbnb — à compléter" : "Libre";
          return (
            <div key={key} title={label} style={{
              aspectRatio: "1", display: "flex", alignItems: "center", justifyContent: "center",
              borderRadius: 4, fontSize: 12.5,
              background: confirmed ? C.brass : "transparent",
              color: confirmed ? C.ink : external ? C.brassLight : C.ivoryDim,
              border: confirmed ? "none" : external ? `1.5px dashed ${C.brassLight}` : `1px solid ${C.line}`,
              fontWeight: confirmed ? 600 : 400,
            }}>
              {d}
            </div>
          );
        })}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 16, marginTop: 12, fontSize: 11.5, color: C.ivoryFaint, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ width: 10, height: 10, borderRadius: 3, background: C.brass, display: "inline-block" }} /> Réservation enregistrée
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ width: 10, height: 10, borderRadius: 3, border: `1.5px dashed ${C.brassLight}`, display: "inline-block" }} /> Bloqué sur Airbnb (à compléter)
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ width: 10, height: 10, borderRadius: 3, border: `1px solid ${C.line}`, display: "inline-block" }} /> Libre
        </div>
      </div>
    </div>
  );
}

function monthLabelFr(monthStr) {
  const [y, m] = monthStr.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
}

function MonthlyStatementDownload({ prop, reservations, expenses }) {
  const months = Array.from(new Set([
    ...reservations.filter((r) => r.checkIn).map((r) => r.checkIn.slice(0, 7)),
    ...expenses.filter((e) => e.date).map((e) => e.date.slice(0, 7)),
  ])).sort().reverse();

  const currentMonth = new Date().toISOString().slice(0, 7);
  const defaultMonth = months.includes(currentMonth) ? currentMonth : (months[0] || currentMonth);
  const [month, setMonth] = useState(defaultMonth);
  const options = months.includes(currentMonth) ? months : [currentMonth, ...months];

  function download() {
    const monthRes = reservations.filter((r) => r.checkIn && r.checkIn.slice(0, 7) === month);
    const monthExp = expenses.filter((e) => e.date && e.date.slice(0, 7) === month);
    const revenue = monthRes.reduce((s, r) => s + Number(r.amount || 0), 0);
    const cleaning = monthRes.reduce((s, r) => s + Number(r.cleaningFee || 0), 0);
    const other = monthExp.reduce((s, e) => s + Number(e.amount || 0), 0);
    const pct = prop.commissionPct || 0;
    const commission = revenue * (pct / 100);
    const net = revenue - cleaning - other - commission;

    const rows = [
      [`Relevé mensuel — ${prop.name}`],
      [`Mois : ${monthLabelFr(month)}`],
      [`Bien : ${prop.address || ""}`],
      [],
      ["Date", "Type", "Description", "Entrée (€)", "Sortie (€)"],
    ];
    monthRes.forEach((r) => {
      rows.push([r.checkIn, "Réservation", `${r.guestName} (${r.platform || ""})`, Number(r.amount || 0), ""]);
      if (r.cleaningFee > 0) rows.push([r.checkIn, "Ménage", r.guestName, "", Number(r.cleaningFee)]);
    });
    monthExp.forEach((e) => {
      const label = (EXPENSE_TYPES.find((t) => t.id === e.type) || {}).label || "Charge";
      rows.push([e.date, label, e.label || "", "", Number(e.amount || 0)]);
    });
    rows.push([]);
    rows.push(["", "", "Revenu brut", revenue, ""]);
    rows.push(["", "", "Ménage", "", cleaning]);
    rows.push(["", "", "Charges (élec/eau/wifi)", "", other]);
    rows.push(["", "", `Commission Blackland (${pct}%)`, "", commission]);
    rows.push(["", "", "NET À PERCEVOIR", net, ""]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 28 }, { wch: 14 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relevé");
    const safeName = (prop.name || "bien").replace(/[^a-zA-Z0-9]+/g, "_");
    XLSX.writeFile(wb, `Blackland_${safeName}_${month}.xlsx`);
  }

  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 18, padding: "12px 14px", background: C.inkPanel, border: `1px solid ${C.line}`, borderRadius: 6 }}>
      <span style={{ fontSize: 12.5, color: C.ivoryFaint }}>Bilan comptable mensuel</span>
      <div style={{ position: "relative", display: "inline-block" }}>
        <select className="bl-select" value={month} onChange={(e) => setMonth(e.target.value)} style={{ ...inputStyle, paddingRight: 32, cursor: "pointer", fontSize: 13 }}>
          {options.map((m) => <option key={m} value={m} style={{ background: C.inkPanel, textTransform: "capitalize" }}>{monthLabelFr(m)}</option>)}
        </select>
        <ChevronDown size={14} style={{ position: "absolute", right: 10, top: 10, color: C.ivoryFaint, pointerEvents: "none" }} />
      </div>
      <Button variant="brass" onClick={download}>Télécharger (Excel)</Button>
    </div>
  );
}

function OwnerView({ properties, ownerPropId, setOwnerPropId, stats, icalData }) {
  const prop = properties.find((p) => p.id === ownerPropId);
  const netAnim = useCountUp(stats ? stats.net : 0);

  const items = [
    ...(stats?.res || []).map((r) => ({
      key: "r" + r.id, date: r.checkIn, label: `Réservation — ${r.guestName}`, sub: r.platform, amount: r.amount, sign: 1,
    })),
    ...(stats?.res || []).filter((r) => r.cleaningFee > 0).map((r) => ({
      key: "c" + r.id, date: r.checkIn, label: "Ménage", sub: r.guestName, amount: r.cleaningFee, sign: -1,
    })),
    ...(stats?.exp || []).map((e) => {
      const meta = EXPENSE_TYPES.find((t) => t.id === e.type) || EXPENSE_TYPES[3];
      return { key: "e" + e.id, date: e.date, label: meta.label, sub: e.label || "", amount: e.amount, sign: -1 };
    }),
  ].sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  return (
    <div className="bl-fade" style={{ padding: "28px 24px 32px", maxWidth: 640, margin: "0 auto" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, marginBottom: 4, color: C.ivoryFaint, fontSize: 12.5 }}>
        <LogIn size={13} />
        <span>Connecté en tant que</span>
        <PropertyPicker properties={properties} value={ownerPropId} onChange={setOwnerPropId} />
      </div>

      {!prop && (
        <EmptyState text="Aucun bien ne vous a été assigné pour l'instant. Contactez votre gestionnaire Blackland pour qu'il vous associe à un bien depuis l'espace admin (« Gérer les utilisateurs »)." />
      )}

      {prop && stats && (
        <>
          <div style={{ textAlign: "center", margin: "26px 0 30px" }}>
            <div style={{ fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 1.2, marginBottom: 6 }}>Relevé — {prop.name}</div>
            <div className="bl-serif" style={{ fontSize: 15, color: C.ivoryDim, marginBottom: 14 }}>{prop.address}</div>
            <div className="bl-mono" style={{ fontSize: 42, fontWeight: 500, color: C.ivory }}>{eur2(netAnim)}</div>
            <div style={{ fontSize: 12, color: C.ivoryFaint, marginTop: 4 }}>net à percevoir · {stats.count} réservation{stats.count > 1 ? "s" : ""}</div>
          </div>

          <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
            <StatCard label="Revenu brut" value={eur(stats.revenue)} tone="up" />
            <StatCard label="Ménage" value={"− " + eur(stats.cleaning)} tone="down" />
            <StatCard label="Charges (élec/eau/wifi)" value={"− " + eur(stats.other)} tone="down" />
            <StatCard label={`Commission (${stats.pct}%)`} value={"− " + eur(stats.commission)} tone="down" />
          </div>

          <MonthlyStatementDownload prop={prop} reservations={stats.res} expenses={stats.exp} />

          <div style={{ fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 8, borderTop: `1px solid ${C.line}`, paddingTop: 18 }}>
            Calendrier des séjours
          </div>
          <div style={{ marginBottom: 26 }}>
            <ReservationCalendar reservations={stats.res} icalBlocks={icalData?.blocks || []} />
            {icalData?.lastSync && (
              <div style={{ fontSize: 11, color: C.ivoryFaint, marginTop: 8, textAlign: "right" }}>
                Calendrier actualisé le {new Date(icalData.lastSync).toLocaleDateString("fr-FR")}
              </div>
            )}
          </div>

          <div style={{ fontSize: 11, color: C.ivoryFaint, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 8, borderTop: `1px solid ${C.line}`, paddingTop: 18 }}>
            Détail des opérations
          </div>

          {items.length === 0 ? (
            <EmptyState text="Aucune opération enregistrée pour l'instant." />
          ) : (
            <div>
              {items.map((it) => (
                <div key={it.key} className="bl-row" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 4px", borderBottom: `1px solid ${C.line}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    {it.sign > 0 ? <ArrowUpRight size={14} color={C.moss} /> : <ArrowDownRight size={14} color={C.wine} />}
                    <div>
                      <div style={{ fontSize: 13.5 }}>{it.label}</div>
                      <div style={{ fontSize: 11.5, color: C.ivoryFaint }}>{fmtDate(it.date)}{it.sub ? " · " + it.sub : ""}</div>
                    </div>
                  </div>
                  <div className="bl-mono" style={{ fontSize: 13.5, color: it.sign > 0 ? C.moss : C.wine }}>
                    {it.sign > 0 ? "+ " : "− "}{eur2(it.amount)}
                  </div>
                </div>
              ))}
            </div>
          )}

          <div style={{ textAlign: "center", marginTop: 28, fontSize: 11, color: C.ivoryFaint, letterSpacing: 0.4 }}>
            Blackland — gestion transparente de votre bien
          </div>
        </>
      )}
    </div>
  );
}

export default function BlacklandApp() {
  return (
    <ErrorBoundary>
      <BlacklandAppInner />
    </ErrorBoundary>
  );
}

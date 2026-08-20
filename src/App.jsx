import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Search, Check, X, ChevronLeft, UserCheck, Users, UtensilsCrossed, ClipboardList, LayoutGrid, Plus, Loader2, WifiOff, Upload, FileSpreadsheet } from "lucide-react";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://onrdmjrrvgztkzmwdzzq.supabase.co";
const SUPABASE_KEY = "sb_publishable_NbdrdOQHYblAME-Ur4kX1g_PhDUm51m";

const KNOWN_DIETS = ["Celíaco", "Vegetariano", "Vegano", "Diabético", "Alergia"];

function parseDietCell(raw) {
  if (!raw) return [];
  return String(raw)
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => KNOWN_DIETS.find((d) => d.toLowerCase() === s.toLowerCase()) || s);
}

function parseGuestsWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows
    .map((row) => {
      const name = String(row["Nombre"] ?? row["nombre"] ?? "").trim();
      const table = Number(row["Mesa"] ?? row["mesa"] ?? 0);
      const diet = parseDietCell(row["Dieta"] ?? row["dieta"]);
      const notes = String(row["Notas"] ?? row["notas"] ?? "").trim();
      return { name, table, diet, notes };
    })
    .filter((g) => g.name && g.table);
}

async function sb(path, options = {}, retry = true) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${currentAccessToken || SUPABASE_KEY}`,
      "Content-Type": "application/json",
      Prefer: options.prefer || "return=representation",
      ...(options.headers || {}),
    },
  });
  if (res.status === 401 && retry) {
    const refreshed = await tryRefreshSession();
    if (refreshed) return sb(path, options, false);
  }
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`${res.status} ${text}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

// --- Autenticación ---
let currentAccessToken = null;

function saveSession(data) {
  currentAccessToken = data.access_token;
  localStorage.setItem("ci_access_token", data.access_token);
  localStorage.setItem("ci_refresh_token", data.refresh_token);
}

function clearSession() {
  currentAccessToken = null;
  localStorage.removeItem("ci_access_token");
  localStorage.removeItem("ci_refresh_token");
}

async function authRequest(grantType, body) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=${grantType}`, {
    method: "POST",
    headers: { apikey: SUPABASE_KEY, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.msg || "No se pudo iniciar sesión");
  return data;
}

async function signIn(email, password) {
  const data = await authRequest("password", { email, password });
  saveSession(data);
  return data;
}

async function tryRefreshSession() {
  const refreshToken = localStorage.getItem("ci_refresh_token");
  if (!refreshToken) return false;
  try {
    const data = await authRequest("refresh_token", { refresh_token: refreshToken });
    saveSession(data);
    return true;
  } catch {
    clearSession();
    return false;
  }
}

function restoreSession() {
  const token = localStorage.getItem("ci_access_token");
  if (token) currentAccessToken = token;
  return !!token;
}

const DIET_OPTIONS = ["Celíaco", "Vegetariano", "Vegano", "Diabético", "Alergia"];

const DIET_STYLES = {
  "Celíaco": { bg: "#F3E7DE", text: "#8A4A2A" },
  "Vegetariano": { bg: "#E4EEE0", text: "#4C7A5E" },
  "Vegano": { bg: "#DCEAE0", text: "#3D6B4F" },
  "Diabético": { bg: "#EAE3F2", text: "#6B4A96" },
  "Alergia": { bg: "#F4DEDC", text: "#A8493D" },
};

function initials(name) {
  return name.split(" ").map((p) => p[0]).slice(0, 2).join("").toUpperCase();
}

function normalizeGuest(row) {
  return {
    id: row.id,
    name: row.name,
    table: row.table_number,
    status: row.status,
    diet: row.diet || [],
    notes: row.notes || "",
  };
}

export default function App() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [checkingSession, setCheckingSession] = useState(true);

  useEffect(() => {
    const hasToken = restoreSession();
    setLoggedIn(hasToken);
    setCheckingSession(false);
  }, []);

  if (checkingSession) {
    return (
      <div style={{ minHeight: "100vh", background: "#FAF7F2", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <Loader2 size={22} className="spin" color="#8A7F6B" />
      </div>
    );
  }

  if (!loggedIn) {
    return <LoginScreen onSuccess={() => setLoggedIn(true)} />;
  }

  return <MainApp onLogout={() => { clearSession(); setLoggedIn(false); }} />;
}

function LoginScreen({ onSuccess }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      await signIn(email.trim(), password);
      onSuccess();
    } catch (err) {
      setError("Email o contraseña incorrectos.");
    }
    setLoading(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#1B2430", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "'Public Sans', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Public+Sans:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        button { font-family: inherit; cursor: pointer; }
        input { font-family: inherit; }
      `}</style>
      <form
        onSubmit={handleSubmit}
        style={{ background: "#FAF7F2", borderRadius: 20, padding: "32px 26px", width: "100%", maxWidth: 380 }}
      >
        <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 24, color: "#1B2430", marginBottom: 4, textAlign: "center" }}>
          Control de Invitados
        </div>
        <div style={{ fontSize: 13, color: "#8A7F6B", textAlign: "center", marginBottom: 26 }}>
          Ingresá con tu usuario del salón
        </div>

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Email</label>
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          style={{ width: "100%", padding: "11px 13px", borderRadius: 10, border: "1.5px solid #E5DFD3", fontSize: 15, marginTop: 6, marginBottom: 16, outline: "none" }}
        />

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Contraseña</label>
        <input
          type="password"
          required
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={{ width: "100%", padding: "11px 13px", borderRadius: 10, border: "1.5px solid #E5DFD3", fontSize: 15, marginTop: 6, marginBottom: 20, outline: "none" }}
        />

        {error && (
          <div style={{ fontSize: 13, color: "#A8493D", background: "#F4DEDC", borderRadius: 10, padding: "10px 12px", marginBottom: 16 }}>
            {error}
          </div>
        )}

        <button
          type="submit"
          disabled={loading}
          style={{ width: "100%", padding: "13px", borderRadius: 12, border: "none", background: "#1B2430", color: "#FAF7F2", fontWeight: 700, fontSize: 14.5 }}
        >
          {loading ? "Ingresando..." : "Ingresar"}
        </button>
      </form>
    </div>
  );
}

function MainApp({ onLogout }) {
  const [guests, setGuests] = useState([]);
  const [eventId, setEventId] = useState(null);
  const [eventName, setEventName] = useState("");
  const [view, setView] = useState("reception");
  const [query, setQuery] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const pollRef = useRef(null);

  const fetchGuests = useCallback(async (evId) => {
    const rows = await sb(`guests?event_id=eq.${evId}&order=table_number.asc,name.asc`);
    setGuests(rows.map(normalizeGuest));
  }, []);

  // Cargar (o crear) el evento activo, y arrancar sincronización
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true);
        let events = await sb(`events?order=created_at.desc&limit=1`);
        let ev;
        if (!events.length) {
          const created = await sb(`events`, {
            method: "POST",
            body: JSON.stringify({ name: "Mi evento" }),
          });
          ev = created[0];
        } else {
          ev = events[0];
        }
        if (cancelled) return;
        setEventId(ev.id);
        setEventName(ev.name);
        await fetchGuests(ev.id);
        setError(null);
      } catch (e) {
        if (!cancelled) setError(`No se pudo conectar: ${e.message}`);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [fetchGuests]);

  // Sincronización cada 4s para que recepción y admin vean lo mismo
  useEffect(() => {
    if (!eventId) return;
    pollRef.current = setInterval(() => {
      fetchGuests(eventId).catch(() => setError("Se perdió la conexión, reintentando..."));
    }, 4000);
    return () => clearInterval(pollRef.current);
  }, [eventId, fetchGuests]);

  const toggleArrived = async (id) => {
    const g = guests.find((x) => x.id === id);
    const newStatus = g.status === "arrived" ? "pending" : "arrived";
    setGuests((gs) => gs.map((x) => (x.id === id ? { ...x, status: newStatus } : x)));
    try {
      await sb(`guests?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ status: newStatus }), prefer: "return=minimal" });
    } catch {
      setGuests((gs) => gs.map((x) => (x.id === id ? { ...x, status: g.status } : x)));
      setError("No se pudo guardar el cambio. Revisá tu conexión.");
    }
  };

  const updateGuest = async (id, patch) => {
    setGuests((gs) => gs.map((g) => (g.id === id ? { ...g, ...patch } : g)));
    try {
      const body = {};
      if (patch.table !== undefined) body.table_number = patch.table;
      if (patch.diet !== undefined) body.diet = patch.diet;
      if (patch.notes !== undefined) body.notes = patch.notes;
      await sb(`guests?id=eq.${id}`, { method: "PATCH", body: JSON.stringify(body), prefer: "return=minimal" });
    } catch {
      setError("No se pudo guardar el cambio. Revisá tu conexión.");
    }
  };

  const addGuest = async ({ name, table }) => {
    try {
      const created = await sb(`guests`, {
        method: "POST",
        body: JSON.stringify({ event_id: eventId, name, table_number: table, status: "pending", diet: [], notes: "" }),
      });
      setGuests((gs) => [...gs, normalizeGuest(created[0])]);
      setShowAdd(false);
    } catch (e) {
      setError(`No se pudo agregar: ${e.message}`);
    }
  };

  const importGuests = async (parsedRows) => {
    const payload = parsedRows.map((g) => ({
      event_id: eventId,
      name: g.name,
      table_number: g.table,
      status: "pending",
      diet: g.diet,
      notes: g.notes,
    }));
    const created = await sb(`guests`, {
      method: "POST",
      body: JSON.stringify(payload),
    });
    setGuests((gs) => [...gs, ...created.map(normalizeGuest)]);
    return created.length;
  };

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return guests;
    return guests.filter((g) => g.name.toLowerCase().includes(q) || String(g.table).includes(q));
  }, [guests, query]);

  const byTable = useMemo(() => {
    const map = {};
    for (const g of filtered) {
      if (!map[g.table]) map[g.table] = [];
      map[g.table].push(g);
    }
    return Object.entries(map).sort((a, b) => Number(a[0]) - Number(b[0]));
  }, [filtered]);

  const stats = useMemo(() => {
    const total = guests.length;
    const arrived = guests.filter((g) => g.status === "arrived").length;
    const withDiet = guests.filter((g) => g.diet.length > 0);
    return { total, arrived, pending: total - arrived, withDiet };
  }, [guests]);

  const tables = useMemo(() => {
    const map = {};
    for (const g of guests) {
      if (!map[g.table]) map[g.table] = { total: 0, arrived: 0 };
      map[g.table].total += 1;
      if (g.status === "arrived") map[g.table].arrived += 1;
    }
    return Object.entries(map).sort((a, b) => Number(a[0]) - Number(b[0]));
  }, [guests]);

  const editingGuest = guests.find((g) => g.id === editingId);

  return (
    <div style={{ minHeight: "100vh", background: "#FAF7F2", fontFamily: "'Public Sans', sans-serif", color: "#1B2430" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Public+Sans:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        button { font-family: inherit; cursor: pointer; }
        input { font-family: inherit; }
        ::selection { background: #B8935F44; }
      `}</style>

      {/* Top bar */}
      <div style={{ background: "#1B2430", padding: "18px 20px", position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ maxWidth: 560, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 20, color: "#FAF7F2", letterSpacing: 0.2 }}>
              {eventName || "Cargando..."}
            </div>
            <div style={{ fontSize: 12.5, color: loading ? "#8A8578" : "#B8935F", fontWeight: 500, marginTop: 1, display: "flex", alignItems: "center", gap: 5 }}>
              {loading ? (
                <>
                  <Loader2 size={12} className="spin" /> Conectando...
                </>
              ) : error ? (
                <>
                  <WifiOff size={12} /> {error}
                </>
              ) : (
                `${guests.length} invitados cargados`
              )}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ display: "flex", background: "#FAF7F215", borderRadius: 10, padding: 3 }}>
              <ViewToggleBtn active={view === "reception"} onClick={() => setView("reception")} icon={<UserCheck size={15} />} label="Recepción" />
              <ViewToggleBtn active={view === "admin"} onClick={() => setView("admin")} icon={<LayoutGrid size={15} />} label="Admin" />
            </div>
            <button
              onClick={onLogout}
              aria-label="Cerrar sesión"
              style={{ width: 34, height: 34, borderRadius: 9, border: "none", background: "#FAF7F215", color: "#CFC9BE", display: "flex", alignItems: "center", justifyContent: "center" }}
            >
              <X size={15} />
            </button>
          </div>
        </div>
      </div>

      <style>{`.spin { animation: spin 1s linear infinite; } @keyframes spin { to { transform: rotate(360deg); } }`}</style>

      <div style={{ maxWidth: 560, margin: "0 auto", padding: "18px 16px 40px" }}>
        {loading ? (
          <div style={{ textAlign: "center", padding: "60px 0", color: "#9A9284" }}>Cargando invitados...</div>
        ) : view === "reception" ? (
          <ReceptionView
            query={query}
            setQuery={setQuery}
            byTable={byTable}
            stats={stats}
            toggleArrived={toggleArrived}
            openEdit={setEditingId}
            onAdd={() => setShowAdd(true)}
          />
        ) : (
          <AdminView guests={guests} stats={stats} tables={tables} onImport={importGuests} />
        )}
      </div>

      {editingGuest && (
        <EditModal
          guest={editingGuest}
          onClose={() => setEditingId(null)}
          onSave={(patch) => {
            updateGuest(editingGuest.id, patch);
            setEditingId(null);
          }}
        />
      )}

      {showAdd && <AddGuestModal onClose={() => setShowAdd(false)} onSave={addGuest} />}
    </div>
  );
}

function ViewToggleBtn({ active, onClick, icon, label }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 6,
        padding: "7px 13px",
        borderRadius: 8,
        border: "none",
        background: active ? "#B8935F" : "transparent",
        color: active ? "#1B2430" : "#CFC9BE",
        fontWeight: 600,
        fontSize: 13,
        transition: "all 0.15s ease",
      }}
    >
      {icon}
      {label}
    </button>
  );
}

function ReceptionView({ query, setQuery, byTable, stats, toggleArrived, openEdit, onAdd }) {
  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
        <StatPill label="Llegaron" value={stats.arrived} tone="#4C7A5E" />
        <StatPill label="Faltan" value={stats.pending} tone="#A8493D" />
        <StatPill label="Total" value={stats.total} tone="#1B2430" />
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>
        <div style={{ position: "relative", flex: 1 }}>
          <Search size={17} style={{ position: "absolute", left: 13, top: 12.5, color: "#9A9284" }} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Buscar invitado o mesa"
            style={{
              width: "100%",
              padding: "11px 14px 11px 38px",
              borderRadius: 11,
              border: "1.5px solid #E5DFD3",
              background: "#fff",
              fontSize: 15,
              outline: "none",
            }}
          />
        </div>
        <button
          onClick={onAdd}
          style={{
            width: 44,
            borderRadius: 11,
            border: "none",
            background: "#1B2430",
            color: "#FAF7F2",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
          aria-label="Agregar invitado"
        >
          <Plus size={20} />
        </button>
      </div>

      {byTable.length === 0 && (
        <div style={{ textAlign: "center", color: "#9A9284", padding: "40px 0", fontSize: 14 }}>
          No se encontró ningún invitado.
        </div>
      )}

      {byTable.map(([table, list]) => (
        <div key={table} style={{ marginBottom: 22 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 15, color: "#8A7F6B", marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
            Mesa {table}
            <span style={{ fontSize: 12, fontWeight: 500, color: "#B8AFA0" }}>
              ({list.filter((g) => g.status === "arrived").length}/{list.length})
            </span>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {list.map((g) => (
              <GuestRow key={g.id} guest={g} onToggle={() => toggleArrived(g.id)} onEdit={() => openEdit(g.id)} />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function GuestRow({ guest, onToggle, onEdit }) {
  const arrived = guest.status === "arrived";
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 12,
        background: "#fff",
        border: `1.5px solid ${arrived ? "#4C7A5E30" : "#E5DFD3"}`,
        borderRadius: 13,
        padding: "10px 12px",
      }}
    >
      <div
        style={{
          width: 38,
          height: 38,
          borderRadius: "50%",
          background: arrived ? "#4C7A5E" : "#EFEAE0",
          color: arrived ? "#fff" : "#9A9284",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontWeight: 700,
          fontSize: 13,
          flexShrink: 0,
        }}
      >
        {initials(guest.name)}
      </div>

      <div style={{ flex: 1, minWidth: 0 }} onClick={onEdit}>
        <div style={{ fontWeight: 600, fontSize: 14.5, color: "#1B2430" }}>{guest.name}</div>
        {(guest.diet.length > 0 || guest.notes) && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 4 }}>
            {guest.diet.map((d) => (
              <span
                key={d}
                style={{
                  fontSize: 10.5,
                  fontWeight: 700,
                  padding: "2px 7px",
                  borderRadius: 6,
                  background: DIET_STYLES[d]?.bg || "#EEE",
                  color: DIET_STYLES[d]?.text || "#666",
                }}
              >
                {d}
              </span>
            ))}
            {guest.notes && (
              <span style={{ fontSize: 11.5, color: "#9A9284", fontStyle: "italic" }}>· {guest.notes}</span>
            )}
          </div>
        )}
      </div>

      <button
        onClick={onToggle}
        aria-label={arrived ? "Marcar como pendiente" : "Marcar como llegado"}
        style={{
          width: 40,
          height: 40,
          borderRadius: 10,
          border: "none",
          background: arrived ? "#4C7A5E" : "#F0ECE3",
          color: arrived ? "#fff" : "#8A7F6B",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexShrink: 0,
        }}
      >
        {arrived ? <Check size={19} /> : <Plus size={19} />}
      </button>
    </div>
  );
}

function StatPill({ label, value, tone }) {
  return (
    <div style={{ flex: 1, background: "#fff", border: "1.5px solid #E5DFD3", borderRadius: 12, padding: "10px 12px" }}>
      <div style={{ fontSize: 22, fontWeight: 700, fontFamily: "'Fraunces', serif", color: tone }}>{value}</div>
      <div style={{ fontSize: 11.5, color: "#9A9284", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.3 }}>{label}</div>
    </div>
  );
}

function AdminView({ guests, stats, tables, onImport }) {
  const [showImport, setShowImport] = useState(false);
  const pct = stats.total ? Math.round((stats.arrived / stats.total) * 100) : 0;
  const dietGroups = useMemo(() => {
    const map = {};
    for (const g of guests) {
      for (const d of g.diet) {
        if (!map[d]) map[d] = [];
        map[d].push(g);
      }
    }
    return map;
  }, [guests]);

  return (
    <div>
      <button
        onClick={() => setShowImport(true)}
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          gap: 8,
          width: "100%",
          padding: "12px",
          borderRadius: 12,
          border: "1.5px dashed #B8935F",
          background: "#FBF3E7",
          color: "#8A5A2A",
          fontWeight: 700,
          fontSize: 13.5,
          marginBottom: 18,
        }}
      >
        <FileSpreadsheet size={17} /> Importar invitados desde Excel
      </button>

      <div style={{ background: "#fff", border: "1.5px solid #E5DFD3", borderRadius: 16, padding: 18, marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 16 }}>Ingreso general</div>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#4C7A5E" }}>{pct}%</div>
        </div>
        <div style={{ height: 9, background: "#F0ECE3", borderRadius: 6, overflow: "hidden" }}>
          <div style={{ width: `${pct}%`, height: "100%", background: "#4C7A5E", borderRadius: 6, transition: "width 0.3s" }} />
        </div>
        <div style={{ display: "flex", gap: 16, marginTop: 12, fontSize: 12.5, color: "#8A7F6B" }}>
          <span><b style={{ color: "#1B2430" }}>{stats.arrived}</b> llegaron</span>
          <span><b style={{ color: "#1B2430" }}>{stats.pending}</b> faltan</span>
          <span><b style={{ color: "#1B2430" }}>{stats.total}</b> invitados</span>
        </div>
      </div>

      <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 15, color: "#8A7F6B", marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
        <Users size={15} /> Mapa de mesas
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(90px, 1fr))", gap: 10, marginBottom: 22 }}>
        {tables.map(([table, t]) => {
          const full = t.arrived === t.total;
          return (
            <div
              key={table}
              style={{
                aspectRatio: "1",
                borderRadius: 14,
                background: full ? "#4C7A5E" : t.arrived > 0 ? "#B8935F" : "#F0ECE3",
                color: full || t.arrived > 0 ? "#fff" : "#8A7F6B",
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 600, opacity: 0.85 }}>Mesa</div>
              <div style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700 }}>{table}</div>
              <div style={{ fontSize: 11, fontWeight: 600 }}>{t.arrived}/{t.total}</div>
            </div>
          );
        })}
      </div>

      <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 15, color: "#8A7F6B", marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
        <UtensilsCrossed size={15} /> Resumen para cocina
      </div>
      {Object.keys(dietGroups).length === 0 ? (
        <div style={{ fontSize: 13.5, color: "#9A9284", background: "#fff", border: "1.5px solid #E5DFD3", borderRadius: 12, padding: 14 }}>
          Nadie cargó restricciones alimentarias todavía.
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {Object.entries(dietGroups).map(([diet, list]) => (
            <div key={diet} style={{ background: "#fff", border: "1.5px solid #E5DFD3", borderRadius: 12, padding: "11px 14px" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 5 }}>
                <span
                  style={{
                    fontSize: 11.5,
                    fontWeight: 700,
                    padding: "2px 8px",
                    borderRadius: 6,
                    background: DIET_STYLES[diet]?.bg,
                    color: DIET_STYLES[diet]?.text,
                  }}
                >
                  {diet}
                </span>
                <span style={{ fontSize: 12, color: "#9A9284", fontWeight: 600 }}>{list.length}</span>
              </div>
              <div style={{ fontSize: 13, color: "#5B5348" }}>
                {list.map((g) => `${g.name} (Mesa ${g.table})`).join(" · ")}
              </div>
            </div>
          ))}
        </div>
      )}

      {showImport && <ImportModal onClose={() => setShowImport(false)} onImport={onImport} />}
    </div>
  );
}

function ImportModal({ onClose, onImport }) {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("idle"); // idle | importing | done | error
  const [errorMsg, setErrorMsg] = useState("");
  const [importedCount, setImportedCount] = useState(0);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setStatus("idle");
    setErrorMsg("");
    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseGuestsWorkbook(buffer);
      setRows(parsed);
    } catch (err) {
      setErrorMsg("No se pudo leer el archivo. Verificá que sea un Excel válido.");
      setRows(null);
    }
  };

  const handleImport = async () => {
    if (!rows || rows.length === 0) return;
    setStatus("importing");
    try {
      const count = await onImport(rows);
      setImportedCount(count);
      setStatus("done");
    } catch (err) {
      setErrorMsg(`No se pudo importar: ${err.message}`);
      setStatus("error");
    }
  };

  return (
    <div
      onClick={onClose}
      style={{ position: "fixed", inset: 0, background: "#1B243088", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{ background: "#FAF7F2", borderRadius: "20px 20px 0 0", width: "100%", maxWidth: 560, padding: "18px 18px 28px", maxHeight: "85vh", overflowY: "auto" }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <button onClick={onClose} style={{ border: "none", background: "none", color: "#8A7F6B", display: "flex", alignItems: "center", gap: 4, fontSize: 13, fontWeight: 600, padding: 4 }}>
            <ChevronLeft size={17} /> Cerrar
          </button>
          <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 15 }}>Importar invitados</div>
          <div style={{ width: 58 }} />
        </div>

        {status === "done" ? (
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 40, marginBottom: 10 }}>✅</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#1B2430", marginBottom: 6 }}>
              {importedCount} invitados importados
            </div>
            <div style={{ fontSize: 13, color: "#8A7F6B", marginBottom: 20 }}>Ya los podés ver en Recepción y Admin.</div>
            <button
              onClick={onClose}
              style={{ width: "100%", padding: "13px", borderRadius: 12, border: "none", background: "#1B2430", color: "#FAF7F2", fontWeight: 700, fontSize: 14.5 }}
            >
              Listo
            </button>
          </div>
        ) : (
          <>
            <div style={{ fontSize: 13, color: "#8A7F6B", marginBottom: 14, lineHeight: 1.5 }}>
              El archivo debe tener columnas <b>Nombre</b> y <b>Mesa</b> (obligatorias), y opcionalmente <b>Dieta</b> y <b>Notas</b>.
            </div>

            <label
              style={{
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                padding: "24px",
                borderRadius: 14,
                border: "1.5px dashed #E5DFD3",
                background: "#fff",
                marginBottom: 16,
                cursor: "pointer",
                textAlign: "center",
              }}
            >
              <Upload size={22} color="#8A7F6B" />
              <span style={{ fontSize: 13.5, fontWeight: 600, color: "#1B2430" }}>
                {fileName || "Tocá para elegir un archivo .xlsx"}
              </span>
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
            </label>

            {errorMsg && (
              <div style={{ fontSize: 13, color: "#A8493D", background: "#F4DEDC", borderRadius: 10, padding: "10px 12px", marginBottom: 16 }}>
                {errorMsg}
              </div>
            )}

            {rows && rows.length > 0 && (
              <div style={{ fontSize: 13.5, color: "#4C7A5E", background: "#E4EEE0", borderRadius: 10, padding: "10px 12px", marginBottom: 16, fontWeight: 600 }}>
                Se encontraron {rows.length} invitados listos para importar.
              </div>
            )}

            {rows && rows.length === 0 && (
              <div style={{ fontSize: 13.5, color: "#A8493D", background: "#F4DEDC", borderRadius: 10, padding: "10px 12px", marginBottom: 16 }}>
                No se encontraron filas válidas (revisá que existan las columnas "Nombre" y "Mesa").
              </div>
            )}

            <button
              onClick={handleImport}
              disabled={!rows || rows.length === 0 || status === "importing"}
              style={{
                width: "100%",
                padding: "13px",
                borderRadius: 12,
                border: "none",
                background: rows && rows.length > 0 ? "#1B2430" : "#D8D2C4",
                color: "#FAF7F2",
                fontWeight: 700,
                fontSize: 14.5,
              }}
            >
              {status === "importing" ? "Importando..." : "Importar"}
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function AddGuestModal({ onClose, onSave }) {
  const [name, setName] = useState("");
  const [table, setTable] = useState("");
  const [saving, setSaving] = useState(false);

  const canSave = name.trim().length > 0 && table !== "";

  const handleSave = async () => {
    if (!canSave || saving) return;
    setSaving(true);
    await onSave({ name: name.trim(), table: Number(table) });
    setSaving(false);
  };

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "#1B243088",
        display: "flex",
        alignItems: "flex-end",
        justifyContent: "center",
        zIndex: 50,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "#FAF7F2",
          borderRadius: "20px 20px 0 0",
          width: "100%",
          maxWidth: 560,
          padding: "18px 18px 28px",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <button onClick={onClose} style={{ border: "none", background: "none", color: "#8A7F6B", display: "flex", alignItems: "center", gap: 4, fontSize: 13, fontWeight: 600, padding: 4 }}>
            <ChevronLeft size={17} /> Cancelar
          </button>
          <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 15 }}>Nuevo invitado</div>
          <div style={{ width: 70 }} />
        </div>

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Nombre</label>
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Nombre y apellido"
          style={{ width: "100%", padding: "10px 13px", borderRadius: 10, border: "1.5px solid #E5DFD3", fontSize: 15, marginTop: 6, marginBottom: 16, outline: "none" }}
        />

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Mesa</label>
        <input
          type="number"
          value={table}
          onChange={(e) => setTable(e.target.value)}
          placeholder="Ej: 3"
          style={{ width: "100%", padding: "10px 13px", borderRadius: 10, border: "1.5px solid #E5DFD3", fontSize: 15, marginTop: 6, marginBottom: 20, outline: "none" }}
        />

        <button
          onClick={handleSave}
          disabled={!canSave || saving}
          style={{
            width: "100%",
            padding: "13px",
            borderRadius: 12,
            border: "none",
            background: canSave ? "#1B2430" : "#D8D2C4",
            color: "#FAF7F2",
            fontWeight: 700,
            fontSize: 14.5,
          }}
        >
          {saving ? "Guardando..." : "Agregar invitado"}
        </button>
      </div>
    </div>
  );
}

function EditModal({ guest, onClose, onSave }) {
  const [diet, setDiet] = useState(guest.diet);
  const [notes, setNotes] = useState(guest.notes);
  const [table, setTable] = useState(guest.table);

  const toggleDiet = (d) => {
    setDiet((cur) => (cur.includes(d) ? cur.filter((x) => x !== d) : [...cur, d]));
  };

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "#1B243088",
        display: "flex",
        alignItems: "flex-end",
        justifyContent: "center",
        zIndex: 50,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "#FAF7F2",
          borderRadius: "20px 20px 0 0",
          width: "100%",
          maxWidth: 560,
          padding: "18px 18px 28px",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <button onClick={onClose} style={{ border: "none", background: "none", color: "#8A7F6B", display: "flex", alignItems: "center", gap: 4, fontSize: 13, fontWeight: 600, padding: 4 }}>
            <ChevronLeft size={17} /> Cerrar
          </button>
          <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 15 }}>{guest.name}</div>
          <div style={{ width: 58 }} />
        </div>

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Mesa</label>
        <input
          type="number"
          value={table}
          onChange={(e) => setTable(Number(e.target.value))}
          style={{ width: "100%", padding: "10px 13px", borderRadius: 10, border: "1.5px solid #E5DFD3", fontSize: 15, marginTop: 6, marginBottom: 16, outline: "none" }}
        />

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Restricciones alimentarias</label>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginTop: 8, marginBottom: 16 }}>
          {DIET_OPTIONS.map((d) => {
            const active = diet.includes(d);
            return (
              <button
                key={d}
                onClick={() => toggleDiet(d)}
                style={{
                  padding: "7px 12px",
                  borderRadius: 9,
                  border: `1.5px solid ${active ? DIET_STYLES[d].text : "#E5DFD3"}`,
                  background: active ? DIET_STYLES[d].bg : "#fff",
                  color: active ? DIET_STYLES[d].text : "#8A7F6B",
                  fontSize: 13,
                  fontWeight: 600,
                }}
              >
                {d}
              </button>
            );
          })}
        </div>

        <label style={{ fontSize: 12, fontWeight: 700, color: "#8A7F6B", textTransform: "uppercase", letterSpacing: 0.3 }}>Observaciones</label>
        <textarea
          value={notes}
          onChange={(e) => setNotes(e.target.value)}
          placeholder="Ej: alergia a los frutos secos, va en silla de ruedas..."
          rows={3}
          style={{ width: "100%", padding: "10px 13px", borderRadius: 10, border: "1.5px solid #E5DFD3", fontSize: 14, marginTop: 6, marginBottom: 20, outline: "none", resize: "none", fontFamily: "inherit" }}
        />

        <button
          onClick={() => onSave({ table, diet, notes })}
          style={{
            width: "100%",
            padding: "13px",
            borderRadius: 12,
            border: "none",
            background: "#1B2430",
            color: "#FAF7F2",
            fontWeight: 700,
            fontSize: 14.5,
          }}
        >
          Guardar cambios
        </button>
      </div>
    </div>
  );
}

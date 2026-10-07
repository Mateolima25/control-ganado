import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";
import {
  Beef,
  CalendarDays,
  ChevronRight,
  ClipboardList,
  FileSpreadsheet,
  Gauge,
  LayoutDashboard,
  LogIn,
  LogOut,
  Tags,
  Plus,
  Search,
  Trash2,
  Upload,
  Wheat,
  X,
} from "lucide-react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

const STORAGE_KEY = "control-ganado-v1";

const normalize = (value = "") =>
  String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const numberValue = (value) => {
  if (typeof value === "number") {
    // True-Test puede entregar algunos pesos sin el separador decimal.
    // Ej.: 1734 = 173,4 kg
    if (value > 1000 && value < 10000) {
      return value / 10;
    }

    return value;
  }

  const s = String(value ?? "").trim().replace(/\s/g, "");

  if (!s) return null;

  let normalized;

  if (s.includes(",") && s.includes(".")) {
    normalized =
      s.lastIndexOf(",") > s.lastIndexOf(".")
        ? s.replace(/\./g, "").replace(",", ".")
        : s.replace(/,/g, "");
  } else if (s.includes(",")) {
    normalized = s.replace(",", ".");
  } else {
    normalized = s;
  }

  let n = Number(normalized);

  if (!Number.isFinite(n)) return null;

  // Corrige exportaciones de True-Test como:
  // 1734 -> 173,4
  // 2540 -> 254,0
  // 3125 -> 312,5
  if (n > 1000 && n < 10000) {
    n = n / 10;
  }

  return n;
};

const dateValue = (value) => {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value)) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const s = String(value).trim();
  const match = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (match) {
    let [, d, m, y] = match;
    if (y.length === 2) y = `20${y}`;
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  const parsed = new Date(s);
  return isNaN(parsed) ? null : parsed.toISOString().slice(0, 10);
};

const formatDate = (value) => {
  if (!value) return "-";
  const [y, m, d] = value.split("-");
  return `${d}/${m}/${y}`;
};

const daysBetween = (a, b) => {
  if (!a || !b) return 0;
  const ms = new Date(`${b}T12:00:00`) - new Date(`${a}T12:00:00`);
  return Math.max(0, Math.round(ms / 86400000));
};

const emptyData = {
  lots: [],
  weighings: [],
  feedings: [],
};

function readStored() {
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY));
    return {
      lots: Array.isArray(stored?.lots) ? stored.lots : [],
      weighings: Array.isArray(stored?.weighings) ? stored.weighings : [],
      feedings: Array.isArray(stored?.feedings) ? stored.feedings : [],
    };
  } catch {
    return { lots: [], weighings: [], feedings: [] };
  }
}

function App() {
  const [data, setData] = useState(emptyData);
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState("");
  const [role, setRole] = useState(null);
  const [dataStatus, setDataStatus] = useState("loading");
  const [dataError, setDataError] = useState("");
  const [syncError, setSyncError] = useState("");
  const [page, setPage] = useState("dashboard");
  const [selectedLotId, setSelectedLotId] = useState(null);
  const [selectedCaravana, setSelectedCaravana] = useState(null);
  const [animalReturnPage, setAnimalReturnPage] = useState("caravanas");
  const [animalReturnTab, setAnimalReturnTab] = useState("summary");
  const [search, setSearch] = useState("");
  const [modal, setModal] = useState(null);
  const [weightImportPreview, setWeightImportPreview] = useState(null);
  const lastSyncedData = useRef("");
  const isOwner = role === "owner";

  useEffect(() => {
    if (!supabase) {
      setAuthLoading(false);
      return;
    }

    let active = true;
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession);
      setRole(null);
      setData(emptyData);
      setDataStatus(nextSession ? "loading" : "idle");
      setDataError("");
      setSyncError("");
      lastSyncedData.current = "";
    });

    supabase.auth.getSession().then(({ data: result, error }) => {
      if (!active) return;
      if (error) setAuthError(error.message);
      setSession(result.session);
      setAuthLoading(false);
    }).catch((error) => {
      if (!active) return;
      setAuthError(error.message);
      setAuthLoading(false);
    });

    return () => {
      active = false;
      subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (!supabase || !session) return;
    let active = true;
    setDataStatus("loading");
    setDataError("");

    const loadSharedData = async () => {
      const { data: member, error: memberError } = await supabase
        .from("herd_members")
        .select("role")
        .eq("user_id", session.user.id)
        .maybeSingle();
      if (memberError) throw memberError;
      if (!member || !["owner", "viewer"].includes(member.role)) {
        if (active) setDataStatus("denied");
        return;
      }

      const { data: sharedState, error: stateError } = await supabase
        .from("herd_app_state")
        .select("data")
        .eq("id", 1)
        .maybeSingle();
      if (stateError) throw stateError;

      let initialData = sharedState?.data;
      if (!initialData && member.role === "owner") {
        initialData = readStored();
        const { error: insertError } = await supabase
          .from("herd_app_state")
          .insert({ id: 1, data: initialData });
        if (insertError) throw insertError;
      }

      if (!initialData) {
        if (active) {
          setRole(member.role);
          setDataStatus("missing");
        }
        return;
      }

      const nextData = {
        lots: Array.isArray(initialData.lots) ? initialData.lots : [],
        weighings: Array.isArray(initialData.weighings) ? initialData.weighings : [],
        feedings: Array.isArray(initialData.feedings) ? initialData.feedings : [],
      };
      if (active) {
        lastSyncedData.current = JSON.stringify(nextData);
        setRole(member.role);
        setData(nextData);
        setDataStatus("ready");
      }
    };

    loadSharedData().catch((error) => {
      if (!active) return;
      setDataError(error.message);
      setDataStatus("error");
    });

    return () => {
      active = false;
    };
  }, [session]);

  useEffect(() => {
    if (!supabase || !session || !isOwner || dataStatus !== "ready") return;
    const serialized = JSON.stringify(data);
    if (serialized === lastSyncedData.current) return;

    localStorage.setItem(STORAGE_KEY, serialized);
    const timeout = setTimeout(async () => {
      const { error } = await supabase
        .from("herd_app_state")
        .upsert({ id: 1, data, updated_at: new Date().toISOString() });
      if (error) {
        setSyncError(error.message);
        return;
      }
      lastSyncedData.current = serialized;
      setSyncError("");
    }, 400);

    return () => clearTimeout(timeout);
  }, [data, dataStatus, isOwner, session]);

  useEffect(() => {
    if (!supabase || role !== "viewer" || dataStatus !== "ready") return;
    let active = true;
    const refreshSharedData = async () => {
      const { data: sharedState, error } = await supabase
        .from("herd_app_state")
        .select("data")
        .eq("id", 1)
        .maybeSingle();
      if (!active || error || !sharedState?.data) return;
      const nextData = {
        lots: Array.isArray(sharedState.data.lots) ? sharedState.data.lots : [],
        weighings: Array.isArray(sharedState.data.weighings) ? sharedState.data.weighings : [],
        feedings: Array.isArray(sharedState.data.feedings) ? sharedState.data.feedings : [],
      };
      setData((current) => JSON.stringify(current) === JSON.stringify(nextData) ? current : nextData);
    };
    const interval = setInterval(refreshSharedData, 15000);
    return () => {
      active = false;
      clearInterval(interval);
    };
  }, [dataStatus, role]);

  useEffect(() => {
    if (isOwner && dataStatus === "ready") {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    }
  }, [data, dataStatus, isOwner]);

  const selectedLot = data.lots.find((l) => l.id === selectedLotId);

  const signOut = () => supabase?.auth.signOut();

  const goLot = (id) => {
    setSelectedLotId(id);
    setPage("lot");
  };

  const openCaravana = (caravana, returnPage = "caravanas", returnTab = "summary") => {
    setSelectedCaravana(caravana);
    setAnimalReturnPage(returnPage);
    setAnimalReturnTab(returnTab);
    setSearch("");
    setPage("animal");
  };

  const addLot = (lot) => {
    if (!isOwner) return;
    const id = crypto.randomUUID();
    setData((prev) => ({ ...prev, lots: [...prev.lots, { ...lot, id }] }));
    setModal(null);
    goLot(id);
  };

  const deleteLot = (id) => {
    if (!isOwner) return;
    if (!confirm("¿Eliminar este lote y todos sus datos?")) return;
    setData((prev) => ({
      lots: prev.lots.filter((l) => l.id !== id),
      weighings: prev.weighings.filter((w) => w.lotId !== id),
      feedings: prev.feedings.filter((f) => f.lotId !== id),
    }));
    setSelectedLotId(null);
    setPage("lots");
  };

  const deleteWeighing = (id) => {
    if (!isOwner) return;
    const weighing = data.weighings.find((row) => row.id === id);
    if (!weighing) return;
    if (!confirm(`¿Eliminar el pesaje de la caravana ${weighing.caravana} (${weighing.weight} kg)?`)) return;
    setData((prev) => ({
      ...prev,
      weighings: prev.weighings.filter((row) => row.id !== id),
    }));
  };

  const deleteWeightsByMonth = (lotId, month) => {
    if (!isOwner) return;
    const monthWeights = data.weighings.filter(
      (row) => row.lotId === lotId && row.date?.startsWith(`${month}-`)
    );
    if (!monthWeights.length) return;
    const monthLabel = new Date(`${month}-01T12:00:00`).toLocaleDateString("es-AR", {
      month: "long",
      year: "numeric",
    });
    if (!confirm(`¿Eliminar ${monthWeights.length} pesajes de ${monthLabel}? El lote y su alimentación se conservarán.`)) return;
    setData((prev) => ({
      ...prev,
      weighings: prev.weighings.filter(
        (row) => !(row.lotId === lotId && row.date?.startsWith(`${month}-`))
      ),
    }));
  };

  const importWeights = async (file, lotId) => {
    if (!isOwner) return;
    try {
      const rows = await readSpreadsheet(file);
      const mapped = mapWeightRows(rows);
      const uniqueRows = new Map();
      let duplicateCount = 0;
      const weighingKey = (row) => JSON.stringify([row.caravana, row.date]);

      mapped.forEach((row) => {
        const key = weighingKey(row);
        if (uniqueRows.has(key)) duplicateCount += 1;
        uniqueRows.set(key, row);
      });

      const existingRows = new Map(
        data.weighings
          .filter((row) => row.lotId === lotId)
          .map((row) => [weighingKey(row), row])
      );
      const previewRows = [...uniqueRows.values()].map((row) => {
        const existing = existingRows.get(weighingKey(row));
        return {
          ...row,
          id: existing?.id || crypto.randomUUID(),
          lotId,
          replacesExisting: Boolean(existing),
        };
      });

      setWeightImportPreview({
        fileName: file.name,
        lotId,
        lotName: data.lots.find((lot) => lot.id === lotId)?.name || "Lote",
        totalRows: rows.length,
        invalidCount: rows.length - mapped.length,
        duplicateCount,
        replacementCount: previewRows.filter((row) => row.replacesExisting).length,
        rows: previewRows,
      });
    } catch (error) {
      alert(`No se pudo leer el archivo: ${error.message}`);
    }
  };

  const confirmWeightImport = () => {
    if (!isOwner || !weightImportPreview?.rows.length) return;
    const keys = new Set(weightImportPreview.rows.map((row) => JSON.stringify([row.caravana, row.date])));
    const newRows = weightImportPreview.rows.map(({ replacesExisting, ...row }) => row);
    setData((prev) => ({
      ...prev,
      weighings: [
        ...prev.weighings.filter(
          (old) => !(old.lotId === weightImportPreview.lotId && keys.has(JSON.stringify([old.caravana, old.date])))
        ),
        ...newRows,
      ],
    }));
    setWeightImportPreview(null);
    alert(`Se importaron ${newRows.length} registros de pesaje.`);
  };

  const importFeed = async (file, lotId) => {
    if (!isOwner) return;
    const rows = await readSpreadsheet(file);
    const mapped = mapFeedRows(rows);
    const newRows = mapped.map((r) => ({
      ...r,
      id: crypto.randomUUID(),
      lotId,
    }));
    setData((prev) => ({
      ...prev,
      feedings: [...prev.feedings, ...newRows],
    }));
    alert(`Se importaron ${newRows.length} registros de alimentación.`);
  };

  if (!supabase) {
    return <InfoScreen title="Falta configurar Supabase" text="Configura VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY en el entorno de la aplicación." />;
  }
  if (authLoading) return <InfoScreen title="Conectando" text="Verificando la sesión…" />;
  if (!session) return <LoginScreen error={authError} onSignIn={async (email, password) => {
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setAuthError(error?.message || "");
    return error;
  }} />;
  if (dataStatus === "loading" || dataStatus === "idle") {
    return <InfoScreen title="Cargando datos" text="Conectando con los lotes compartidos…" />;
  }
  if (dataStatus === "denied") {
    return <InfoScreen title="Acceso no habilitado" text="Tu cuenta aún no está incluida en este grupo." action={signOut} actionLabel="Cerrar sesión" />;
  }
  if (dataStatus === "missing") {
    return <InfoScreen title="Datos aún no inicializados" text="El propietario debe iniciar sesión primero para preparar los datos compartidos." action={signOut} actionLabel="Cerrar sesión" />;
  }
  if (dataStatus === "error") {
    return <InfoScreen title="No se pudieron cargar los datos" text={dataError} action={signOut} actionLabel="Cerrar sesión" />;
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon"><Beef size={22} /></div>
          <div>
            <strong>Control Ganado</strong>
            <span>Pesajes & alimentación</span>
          </div>
        </div>

        <nav>
          <button className={page === "dashboard" ? "active" : ""} onClick={() => setPage("dashboard")}>
            <LayoutDashboard size={18} /> Inicio
          </button>
          <button className={page === "lots" || page === "lot" ? "active" : ""} onClick={() => setPage("lots")}>
            <ClipboardList size={18} /> Lotes
          </button>
          <button className={page === "caravanas" ? "active" : ""} onClick={() => setPage("caravanas")}>
            <Tags size={18} /> Caravanas
          </button>
        </nav>

        {isOwner && (
          <button className="new-lot-sidebar" onClick={() => setModal("lot")}>
            <Plus size={18} /> Nuevo lote
          </button>
        )}

        <div className="sidebar-footer">
          <span>{isOwner ? "Propietario · puede editar" : "Solo lectura"}</span>
          <button type="button" onClick={signOut}><LogOut size={15} /> Cerrar sesión</button>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <h1>
              {page === "dashboard" && "Dashboard"}
              {page === "lots" && "Lotes"}
              {page === "caravanas" && "Caravanas"}
              {page === "lot" && selectedLot?.name}
              {page === "animal" && `Caravana ${selectedCaravana || ""}`}
            </h1>
            <p>
              {page === "dashboard"
                ? "Resumen general de tus animales y alimentación."
                : page === "lots"
                ? "Administra los lotes y sus pesajes."
                : page === "caravanas"
                ? "Seguimiento individual de cada animal, aunque cambie de lote."
                    : page === "animal"
                    ? "Historial completo de pesajes y lotes de este animal."
                : "Pesajes, caravanas, alimentación y evolución."}
            </p>
          </div>
                  {page !== "dashboard" && page !== "animal" && (
            <div className="search-box">
              <Search size={17} />
              <input
                placeholder="Buscar caravana..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
          )}
        </header>

        {syncError && <div className="sync-error" role="alert">No se pudo guardar en la nube: {syncError}</div>}

        {page === "dashboard" && (
          <Dashboard data={data} goLot={goLot} setModal={setModal} canEdit={isOwner} />
        )}

        {page === "lots" && (
          <Lots data={data} goLot={goLot} setModal={setModal} deleteLot={deleteLot} canEdit={isOwner} />
        )}

        {page === "caravanas" && (
          <Caravanas
            data={data}
            search={search}
            goLot={goLot}
            openCaravana={openCaravana}
          />
        )}

        {page === "animal" && selectedCaravana && (
          <AnimalDetail
            data={data}
            caravana={selectedCaravana}
            goLot={goLot}
            onBack={() => setPage(animalReturnPage)}
          />
        )}

        {page === "lot" && selectedLot && (
          <LotDetail
            data={data}
            lot={selectedLot}
            search={search}
            importWeights={importWeights}
            importFeed={importFeed}
            deleteWeighing={deleteWeighing}
            deleteWeightsByMonth={deleteWeightsByMonth}
            deleteLot={deleteLot}
            setPage={setPage}
            goLot={goLot}
            openCaravana={openCaravana}
            initialTab={page === "lot" ? animalReturnTab : "summary"}
            canEdit={isOwner}
          />
        )}
      </main>

      {modal === "lot" && isOwner && (
        <LotModal onClose={() => setModal(null)} onSave={addLot} />
      )}
      {weightImportPreview && isOwner && (
        <WeightImportPreview
          preview={weightImportPreview}
          onCancel={() => setWeightImportPreview(null)}
          onConfirm={confirmWeightImport}
        />
      )}
    </div>
  );
}

function LoginScreen({ error, onSignIn }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [formError, setFormError] = useState("");

  const submit = async (event) => {
    event.preventDefault();
    setBusy(true);
    setFormError("");
    try {
      await onSignIn(email.trim(), password);
    } catch (submitError) {
      setFormError(submitError.message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <main className="auth-page">
      <form className="auth-panel" onSubmit={submit}>
        <div className="brand-icon"><Beef size={22} /></div>
        <h1>Control de Ganado</h1>
        <p>Inicia sesión para consultar los lotes compartidos.</p>
        <label>Correo electrónico
          <input type="email" autoComplete="username" value={email} onChange={(event) => setEmail(event.target.value)} required />
        </label>
        <label>Contraseña
          <input type="password" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value)} required />
        </label>
        {(error || formError) && <p className="auth-error" role="alert">{error || formError}</p>}
        <button className="primary" type="submit" disabled={busy}>
          <LogIn size={17} /> {busy ? "Ingresando…" : "Iniciar sesión"}
        </button>
      </form>
    </main>
  );
}

function InfoScreen({ title, text, action, actionLabel }) {
  return (
    <main className="auth-page">
      <section className="auth-panel">
        <div className="brand-icon"><Beef size={22} /></div>
        <h1>{title}</h1>
        <p>{text}</p>
        {action && <button className="secondary" type="button" onClick={action}>{actionLabel}</button>}
      </section>
    </main>
  );
}

function Dashboard({ data, goLot, setModal, canEdit }) {
  const totalAnimals = data.lots.reduce((sum, l) => {
    const latest = latestWeightByLot(data.weighings, l.id);
    return sum + latest.length;
  }, 0);

  const allLatest = data.lots.flatMap((l) => latestWeightByLot(data.weighings, l.id));
  const avgWeight = allLatest.length
    ? allLatest.reduce((s, x) => s + x.weight, 0) / allLatest.length
    : 0;

  const avgDaily = data.lots
    .map((l) => lotStats(data, l.id).avgDaily)
    .filter((x) => Number.isFinite(x) && x !== null);

  const avgDailyGlobal = avgDaily.length
    ? avgDaily.reduce((a, b) => a + b, 0) / avgDaily.length
    : 0;

  return (
    <div className="content">
      <div className="hero">
        <div>
          <span className="eyebrow">CONTROL DE GANADO</span>
          <h2>Todo el campo en un solo lugar.</h2>
          <p>Importá los pesajes de True-Test y relacioná la evolución con la alimentación de cada lote.</p>
        </div>
        {canEdit && <button className="primary" onClick={() => setModal("lot")}><Plus size={18} /> Crear lote</button>}
      </div>

      <div className="stats-grid">
        <Stat icon={<ClipboardList />} label="Lotes" value={data.lots.length} />
        <Stat icon={<Beef />} label="Animales pesados" value={totalAnimals} />
        <Stat icon={<Gauge />} label="Peso promedio actual" value={`${avgWeight.toFixed(1)} kg`} />
        <Stat icon={<Wheat />} label="Ganancia diaria promedio" value={`${avgDailyGlobal.toFixed(2)} kg/día`} />
      </div>

      <section className="section">
        <div className="section-title">
          <div><h3>Mis lotes</h3><span>Acceso rápido a la información</span></div>
          {canEdit && <button className="ghost" onClick={() => setModal("lot")}><Plus size={16} /> Nuevo lote</button>}
        </div>

        {data.lots.length === 0 ? (
          <Empty icon={<ClipboardList />} title="Todavía no hay lotes" text="Creá tu primer lote para comenzar a cargar pesajes." />
        ) : (
          <div className="lot-grid">
            {data.lots.map((lot) => {
              const s = lotStats(data, lot.id);
              return (
                <button className="lot-card" key={lot.id} onClick={() => goLot(lot.id)}>
                  <div className="lot-card-top">
                    <div className="lot-icon"><Beef size={20} /></div>
                    <ChevronRight size={19} />
                  </div>
                  <h3>{lot.name}</h3>
                  <span className="muted">{lot.category || "Sin categoría"}</span>
                  <div className="lot-card-metrics">
                    <div><b>{s.animals}</b><span>animales</span></div>
                    <div><b>{s.avgWeight.toFixed(1)} kg</b><span>peso promedio</span></div>
                    <div><b>{s.avgDaily.toFixed(2)}</b><span>kg/día</span></div>
                  </div>
                </button>
              );
            })}
          </div>
        )}
      </section>
    </div>
  );
}

function Lots({ data, goLot, setModal, deleteLot, canEdit }) {
  return (
    <div className="content">
      <div className="section-title">
        <div><h3>Todos los lotes</h3><span>Creá y administrá tus lotes de ganado.</span></div>
        {canEdit && <button className="primary" onClick={() => setModal("lot")}><Plus size={18} /> Nuevo lote</button>}
      </div>

      {data.lots.length === 0 ? (
        <Empty icon={<ClipboardList />} title="No hay lotes creados" text="Comenzá creando un lote." />
      ) : (
        <div className="table-wrap">
          <table>
            <thead><tr><th>Lote</th><th>Categoría</th><th>Animales</th><th>Peso promedio</th><th>Ganancia diaria</th><th></th></tr></thead>
            <tbody>
              {data.lots.map((lot) => {
                const s = lotStats(data, lot.id);
                return (
                  <tr key={lot.id}>
                    <td><button className="link-button" onClick={() => goLot(lot.id)}>{lot.name}</button></td>
                    <td>{lot.category || "-"}</td>
                    <td>{s.animals}</td>
                    <td>{s.avgWeight.toFixed(1)} kg</td>
                    <td>{s.avgDaily.toFixed(2)} kg/día</td>
                    <td className="actions">
                      <button title="Abrir" onClick={() => goLot(lot.id)}><ChevronRight size={17} /></button>
                      {canEdit && <button title="Eliminar" className="danger-icon" onClick={() => deleteLot(lot.id)}><Trash2 size={16} /></button>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function Caravanas({ data, search, goLot, openCaravana }) {
  const animals = useMemo(() => {
    const grouped = new Map();
    const lotNames = new Map(data.lots.map((lot) => [lot.id, lot.name]));

    data.weighings.forEach((weighing) => {
      if (!weighing.caravana) return;
      const rows = grouped.get(weighing.caravana) || [];
      rows.push(weighing);
      grouped.set(weighing.caravana, rows);
    });

    return [...grouped.entries()]
      .map(([caravana, rows]) => {
        const history = rows
          .slice()
          .sort((a, b) => (a.date || "").localeCompare(b.date || ""))
          .map((weighing) => ({
            ...weighing,
            lotName: lotNames.get(weighing.lotId) || "Lote no disponible",
          }));
        const latest = history[history.length - 1];

        return {
          ...animalStats(history, caravana),
          history,
          currentLot: latest?.lotName || "-",
          currentLotId: latest?.lotId,
        };
      })
      .filter((animal) => !search || normalize(animal.caravana).includes(normalize(search)))
      .sort((a, b) => b.currentWeight - a.currentWeight);
  }, [data.lots, data.weighings, search]);

  const avgWeight = animals.length
    ? animals.reduce((sum, animal) => sum + animal.currentWeight, 0) / animals.length
    : 0;
  const animalsWithDaily = animals.filter((animal) => animal.days > 0);
  const avgDaily = animalsWithDaily.length
    ? animalsWithDaily.reduce((sum, animal) => sum + animal.daily, 0) / animalsWithDaily.length
    : 0;
  return (
    <div className="content">
      <div className="section-title">
        <div><h3>Seguimiento individual</h3><span>Historial unificado de pesajes y lotes por caravana.</span></div>
      </div>

      <div className="stats-grid">
        <Stat icon={<Beef />} label="Caravanas" value={animals.length} />
        <Stat icon={<Gauge />} label="Peso actual promedio" value={`${avgWeight.toFixed(1)} kg`} />
        <Stat icon={<CalendarDays />} label="Ganancia diaria promedio" value={`${avgDaily.toFixed(2)} kg/día`} />
        <Stat icon={<FileSpreadsheet />} label="Pesajes encontrados" value={animals.reduce((sum, animal) => sum + animal.count, 0)} />
      </div>

      {animals.length === 0 ? (
        <Empty icon={<Beef />} title="No hay caravanas para mostrar" text="Importá pesajes para consultar la evolución global de tus animales." />
      ) : (
        <>
          <section className="panel caravan-list">
            <div className="panel-title"><h3>Animales</h3><span>{animals.length} caravanas</span></div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>Caravana</th><th>Pesajes</th><th>Desde</th><th>Peso inicial</th><th>Peso actual</th><th>Lote actual</th><th>Ganancia</th><th>Kg/día</th></tr></thead>
                <tbody>{animals.map((animal) => (
                  <tr key={animal.caravana}>
                    <td><button className="link-button" onClick={() => openCaravana(animal.caravana)}>{animal.caravana}</button></td>
                    <td>{animal.count}</td>
                    <td>{formatDate(animal.history[0]?.date)}</td>
                    <td>{animal.initialWeight.toFixed(1)} kg</td>
                    <td><b>{animal.currentWeight.toFixed(1)} kg</b></td>
                    <td>{animal.currentLotId ? <button className="link-button" onClick={() => goLot(animal.currentLotId)}>{animal.currentLot}</button> : animal.currentLot}</td>
                    <td className={animal.gain >= 0 ? "positive" : "negative"}>{animal.gain >= 0 ? "+" : ""}{animal.gain.toFixed(1)} kg</td>
                    <td>{animal.daily.toFixed(2)} kg/día</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          </section>
        </>
      )}
    </div>
  );
}

function AnimalDetail({ data, caravana, goLot, onBack }) {
  const lotNames = new Map(data.lots.map((lot) => [lot.id, lot.name]));
  const history = data.weighings
    .filter((weighing) => weighing.caravana === caravana)
    .slice()
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""))
    .map((weighing) => ({
      ...weighing,
      lotName: lotNames.get(weighing.lotId) || "Lote no disponible",
    }));
  const stats = animalStats(history, caravana);
  const latest = history[history.length - 1];
  const chartData = history.map((weighing) => ({
    date: formatDate(weighing.date),
    weight: weighing.weight,
    lotName: weighing.lotName,
  }));

  return (
    <div className="content">
      <button className="back animal-back" onClick={onBack}>← Volver</button>
      {history.length === 0 ? (
        <Empty icon={<Beef />} title="No hay pesajes" text="No encontramos historial para esta caravana." />
      ) : (
        <>
          <div className="stats-grid">
            <Stat icon={<Beef />} label="Lote actual" value={latest?.lotName || "-"} />
            <Stat icon={<FileSpreadsheet />} label="Pesajes" value={stats.count} />
            <Stat icon={<Gauge />} label="Peso actual" value={`${stats.currentWeight.toFixed(1)} kg`} />
            <Stat icon={<CalendarDays />} label="Ganancia total" value={`${stats.gain >= 0 ? "+" : ""}${stats.gain.toFixed(1)} kg`} />
          </div>

          <section className="panel caravan-evolution">
            <div className="panel-title">
              <div><h3>Evolución de {caravana}</h3><span>{stats.days} días de seguimiento</span></div>
              {latest?.lotId && <button className="link-button" onClick={() => goLot(latest.lotId)}>Abrir lote actual: {latest.lotName}</button>}
            </div>
            <div className="chart"><ResponsiveContainer width="100%" height={300}>
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="date" />
                <YAxis />
                <Tooltip formatter={(value) => [`${Number(value).toFixed(1)} kg`, "Peso"]} labelFormatter={(label, payload) => payload?.[0]?.payload?.lotName ? `${label} · ${payload[0].payload.lotName}` : label} />
                <Line type="monotone" dataKey="weight" stroke="#7ee29b" strokeWidth={2} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer></div>
          </section>

          <section className="panel caravan-history">
            <div className="panel-title"><h3>Historial y movimientos</h3><span>Cada pesaje conserva el lote donde se registró.</span></div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>Fecha</th><th>Peso</th><th>Lote</th><th>Cambio desde pesaje anterior</th></tr></thead>
                <tbody>{history.slice().reverse().map((weighing, index, descendingHistory) => {
                  const previous = descendingHistory[index + 1];
                  const change = previous ? weighing.weight - previous.weight : null;
                  return (
                    <tr key={weighing.id}>
                      <td>{formatDate(weighing.date)}</td>
                      <td><b>{weighing.weight.toFixed(1)} kg</b></td>
                      <td>{weighing.lotId ? <button className="link-button" onClick={() => goLot(weighing.lotId)}>{weighing.lotName}</button> : weighing.lotName}</td>
                      <td className={change === null ? "" : change >= 0 ? "positive" : "negative"}>{change === null ? "Primer registro" : `${change >= 0 ? "+" : ""}${change.toFixed(1)} kg`}</td>
                    </tr>
                  );
                })}</tbody>
              </table>
            </div>
          </section>
        </>
      )}
    </div>
  );
}

function WeightImportPreview({ preview, onCancel, onConfirm }) {
  return (
    <div className="modal-backdrop" onMouseDown={onCancel}>
      <section
        className="weight-preview-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="weight-preview-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="modal-head">
          <div>
            <h3 id="weight-preview-title">Revisar importación de pesajes</h3>
            <span>{preview.fileName} · {preview.lotName}</span>
          </div>
          <button type="button" title="Cerrar" aria-label="Cerrar vista previa" onClick={onCancel}><X size={19} /></button>
        </div>

        <div className="import-summary">
          <div><span>Filas leídas</span><b>{preview.totalRows}</b></div>
          <div><span>Se importarán</span><b>{preview.rows.length}</b></div>
          <div><span>Filas inválidas</span><b>{preview.invalidCount}</b></div>
          <div><span>Duplicadas en archivo</span><b>{preview.duplicateCount}</b></div>
        </div>

        {preview.replacementCount > 0 && (
          <p className="import-notice">
            {preview.replacementCount} pesajes coinciden con caravana y fecha ya guardadas en este lote; se reemplazarán al confirmar.
          </p>
        )}
        {preview.duplicateCount > 0 && (
          <p className="help">Si una caravana aparece más de una vez en el archivo para la misma fecha, se usará la última fila.</p>
        )}

        {preview.rows.length === 0 ? (
          <Empty icon={<FileSpreadsheet />} title="No se encontraron pesajes válidos" text="Revisa que el archivo tenga columnas de caravana y peso." />
        ) : (
          <div className="table-wrap weight-preview-table">
            <table>
              <thead><tr><th>Fecha</th><th>Caravana</th><th>Peso</th><th>Resultado</th></tr></thead>
              <tbody>{preview.rows.map((row) => (
                <tr key={row.id}>
                  <td>{formatDate(row.date)}</td>
                  <td><b>{row.caravana}</b></td>
                  <td>{row.weight.toFixed(1)} kg</td>
                  <td className={row.replacesExisting ? "negative" : "positive"}>
                    {row.replacesExisting ? "Reemplaza existente" : "Nuevo"}
                  </td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        )}

        <div className="modal-actions">
          <button type="button" className="secondary" onClick={onCancel}>Cancelar</button>
          <button type="button" className="primary" onClick={onConfirm} disabled={preview.rows.length === 0}>
            <Upload size={16} /> Importar {preview.rows.length} pesajes
          </button>
        </div>
      </section>
    </div>
  );
}

function LotDetail({ data, lot, search, importWeights, importFeed, deleteWeighing, deleteWeightsByMonth, deleteLot, setPage, goLot, openCaravana, initialTab, canEdit }) {
  const [tab, setTab] = useState(initialTab);
  const [selectedWeightMonth, setSelectedWeightMonth] = useState("");
  const weights = data.weighings.filter((w) => w.lotId === lot.id);
  const weightMonths = [...new Set(weights.map((weight) => weight.date?.slice(0, 7)).filter(Boolean))].sort().reverse();
  const activeWeightMonth = weightMonths.includes(selectedWeightMonth)
    ? selectedWeightMonth
    : weightMonths[0] || "";
  const feeds = data.feedings.filter((f) => f.lotId === lot.id);
  const stats = lotStats(data, lot.id);

  const animals = useMemo(() => {
    const histories = new Map();
    data.weighings.forEach((weighing) => {
      if (!weighing.caravana) return;
      const history = histories.get(weighing.caravana) || [];
      history.push(weighing);
      histories.set(weighing.caravana, history);
    });

    const lotNames = new Map(data.lots.map((item) => [item.id, item.name]));
    const caravanasInLot = new Set(weights.map((weighing) => weighing.caravana).filter(Boolean));

    return [...caravanasInLot]
      .map((caravana) => {
        const history = (histories.get(caravana) || [])
          .slice()
          .sort((a, b) => (a.date || "").localeCompare(b.date || ""));
        const latest = history[history.length - 1];
        return {
          ...animalStats(history, caravana),
          currentLotId: latest?.lotId,
          currentLot: lotNames.get(latest?.lotId) || "Lote no disponible",
        };
      })
      .filter((a) => !search || a.caravana.toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => b.currentWeight - a.currentWeight);
  }, [data.lots, data.weighings, weights, search]);

  const chartData = animals.slice(0, 15).map((a) => ({
    caravana: a.caravana.slice(-8),
    peso: Number(a.currentWeight.toFixed(1)),
  }));

  const totalFood = feeds.reduce((s, f) => s + (f.foodKg || 0), 0);
  const totalBags = feeds.reduce((s, f) => s + (f.bags || 0), 0);

  return (
    <div className="content">
      <div className="lot-header">
        <button className="back" onClick={() => setPage("lots")}>← Volver a lotes</button>
        <div className="lot-heading">
          <div className="lot-icon large"><Beef size={24} /></div>
          <div><h2>{lot.name}</h2><span>{lot.category || "Sin categoría"} {lot.notes ? `· ${lot.notes}` : ""}</span></div>
        </div>
        {canEdit && <div className="lot-actions">
          <label className="secondary">
            <Upload size={17} /> Importar True-Test
            <input type="file" accept=".xlsx,.xls,.csv" hidden onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) importWeights(file, lot.id);
              event.target.value = "";
            }} />
          </label>
          <label className="secondary">
            <Wheat size={17} /> Importar comida
            <input type="file" accept=".xlsx,.xls,.csv" hidden onChange={(e) => e.target.files[0] && importFeed(e.target.files[0], lot.id)} />
          </label>
        </div>}
      </div>

      <div className="tabs">
        <button className={tab === "summary" ? "active" : ""} onClick={() => setTab("summary")}>Resumen</button>
        <button className={tab === "animals" ? "active" : ""} onClick={() => setTab("animals")}>Animales</button>
        <button className={tab === "weights" ? "active" : ""} onClick={() => setTab("weights")}>Pesajes</button>
        <button className={tab === "feed" ? "active" : ""} onClick={() => setTab("feed")}>Alimentación</button>
      </div>

      {tab === "summary" && (
        <>
          <div className="stats-grid">
            <Stat icon={<Beef />} label="Animales" value={stats.animals} />
            <Stat icon={<Gauge />} label="Peso promedio actual" value={`${stats.avgWeight.toFixed(1)} kg`} />
            <Stat icon={<Gauge />} label="Ganancia promedio" value={`${stats.avgGain.toFixed(1)} kg`} />
            <Stat icon={<CalendarDays />} label="Ganancia diaria" value={`${stats.avgDaily.toFixed(2)} kg/día`} />
          </div>

          <div className="two-col">
            <section className="panel">
              <div className="panel-title"><h3>Pesos actuales</h3><span>Top 15 animales</span></div>
              {chartData.length ? (
                <div className="chart"><ResponsiveContainer width="100%" height={300}>
                  <BarChart data={chartData}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="caravana" /><YAxis /><Tooltip /><Bar dataKey="peso" /></BarChart>
                </ResponsiveContainer></div>
              ) : <Empty icon={<Gauge />} title="Sin pesajes" text="Importá un archivo de True-Test." />}
            </section>

            <section className="panel">
              <div className="panel-title"><h3>Alimentación</h3><span>Datos cargados en este lote</span></div>
              <div className="food-summary">
                <div><span>Comida total</span><b>{totalFood.toFixed(0)} kg</b></div>
                <div><span>Bolsas</span><b>{totalBags.toFixed(0)}</b></div>
                <div><span>Registros</span><b>{feeds.length}</b></div>
              </div>
              <p className="help">Para calcular consumo por animal/día, el archivo debe tener animales y días transcurridos.</p>
            </section>
          </div>
        </>
      )}

      {tab === "animals" && (
        <section className="panel">
          <div className="panel-title"><h3>Animales por caravana</h3><span>Ganancia e historial completo, incluso al cambiar de lote · {animals.length} encontrados</span></div>
          {animals.length === 0 ? (
            <Empty icon={<Beef />} title="No hay animales" text="Importá un archivo de True-Test." />
          ) : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Caravana</th><th>Lote actual</th><th>Pesajes</th><th>Inicial</th><th>Actual</th><th>Ganancia</th><th>Días</th><th>Ganancia diaria</th></tr></thead>
                <tbody>{animals.map((a) => (
                  <tr key={a.caravana}>
                    <td><button className="link-button" onClick={() => openCaravana(a.caravana, "lot", "animals")}>{a.caravana}</button></td>
                    <td>{a.currentLotId ? <button className="link-button" onClick={() => goLot(a.currentLotId)}>{a.currentLot}</button> : a.currentLot}</td>
                    <td>{a.count}</td>
                    <td>{a.initialWeight.toFixed(1)} kg</td>
                    <td>{a.currentWeight.toFixed(1)} kg</td>
                    <td className={a.gain >= 0 ? "positive" : "negative"}>{a.gain >= 0 ? "+" : ""}{a.gain.toFixed(1)} kg</td>
                    <td>{a.days}</td>
                    <td>{a.daily.toFixed(2)} kg/día</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
        </section>
      )}

      {tab === "weights" && (
        <section className="panel">
          <div className="panel-title">
            <div><h3>Historial de pesajes</h3><span>Comparación por número de caravana</span></div>
            {canEdit && weightMonths.length > 0 && (
              <div className="month-delete">
                <select
                  aria-label="Mes de pesaje que se eliminará"
                  value={activeWeightMonth}
                  onChange={(event) => setSelectedWeightMonth(event.target.value)}
                >
                  {weightMonths.map((month) => (
                    <option key={month} value={month}>
                      {new Date(`${month}-01T12:00:00`).toLocaleDateString("es-AR", {
                        month: "long",
                        year: "numeric",
                      })}
                    </option>
                  ))}
                </select>
                <button type="button" onClick={() => deleteWeightsByMonth(lot.id, activeWeightMonth)}>
                  <Trash2 size={15} /> Eliminar mes
                </button>
              </div>
            )}
          </div>
          {weights.length === 0 ? (
            <Empty icon={<FileSpreadsheet />} title="No hay pesajes" text="Importá el Excel/CSV de True-Test." />
          ) : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Fecha</th><th>Caravana</th><th>Peso</th>{canEdit && <th>Acciones</th>}</tr></thead>
                <tbody>{weights.slice().sort((a,b) => (b.date || "").localeCompare(a.date || "")).map((w) => (
                  <tr key={w.id}>
                    <td>{formatDate(w.date)}</td>
                    <td><button className="link-button" onClick={() => openCaravana(w.caravana, "lot", "weights")}>{w.caravana}</button></td>
                    <td>{w.weight.toFixed(1)} kg</td>
                    {canEdit && <td className="actions">
                      <button
                        type="button"
                        className="danger-icon"
                        title="Eliminar este pesaje"
                        aria-label={`Eliminar pesaje de la caravana ${w.caravana}`}
                        onClick={() => deleteWeighing(w.id)}
                      >
                        <Trash2 size={16} />
                      </button>
                    </td>}
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
        </section>
      )}

      {tab === "feed" && (
        <section className="panel">
          <div className="panel-title"><h3>Alimentación</h3><span>Registros importados</span></div>
          {feeds.length === 0 ? (
            <Empty icon={<Wheat />} title="No hay registros de comida" text="Importá tu planilla de alimentación." />
          ) : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Fecha</th><th>Comederos</th><th>Animales</th><th>Comida kg</th><th>Sal</th><th>Bolsas</th><th>Días</th><th>Kg/animal/día</th></tr></thead>
                <tbody>{feeds.slice().sort((a,b) => (b.date || "").localeCompare(a.date || "")).map((f) => {
                  const perDay = f.animals && f.days ? f.foodKg / f.animals / f.days : 0;
                  return <tr key={f.id}>
                    <td>{formatDate(f.date)}</td>
                    <td>{f.feeders || "-"}</td>
                    <td>{f.animals || "-"}</td>
                    <td>{(f.foodKg || 0).toFixed(0)}</td>
                    <td>{f.salt ?? "-"}</td>
                    <td>{f.bags ?? "-"}</td>
                    <td>{f.days ?? "-"}</td>
                    <td>{perDay ? perDay.toFixed(2) : "-"} kg</td>
                  </tr>;
                })}</tbody>
              </table>
            </div>
          )}
        </section>
      )}

      {canEdit && <button className="delete-lot" onClick={() => deleteLot(lot.id)}><Trash2 size={16} /> Eliminar lote</button>}
    </div>
  );
}

function LotModal({ onClose, onSave }) {
  const [name, setName] = useState("");
  const [category, setCategory] = useState("Novillos");
  const [notes, setNotes] = useState("");

  const submit = (e) => {
    e.preventDefault();
    if (!name.trim()) return;
    onSave({ name: name.trim(), category, notes: notes.trim(), createdAt: new Date().toISOString() });
  };

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <form className="modal" onSubmit={submit} onMouseDown={(e) => e.stopPropagation()}>
        <div className="modal-head"><div><h3>Crear nuevo lote</h3><span>Los pesajes y la alimentación quedarán vinculados a este lote.</span></div><button type="button" onClick={onClose}><X /></button></div>
        <label>Nombre del lote<input autoFocus value={name} onChange={(e) => setName(e.target.value)} placeholder="Ej. Novillos 2026" /></label>
        <label>Categoría<select value={category} onChange={(e) => setCategory(e.target.value)}><option>Novillos</option><option>Vaquillonas</option><option>Terneros</option><option>Terneras</option><option>Mixtos</option><option>Otro</option></select></label>
        <label>Observaciones<input value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Opcional" /></label>
        <div className="modal-actions"><button type="button" className="secondary" onClick={onClose}>Cancelar</button><button className="primary" type="submit">Crear lote</button></div>
      </form>
    </div>
  );
}

function Stat({ icon, label, value }) {
  return <div className="stat-card"><div className="stat-icon">{icon}</div><div><span>{label}</span><strong>{value}</strong></div></div>;
}

function Empty({ icon, title, text }) {
  return <div className="empty"><div className="empty-icon">{icon}</div><h3>{title}</h3><p>{text}</p></div>;
}

function latestWeightByLot(rows, lotId) {
  const grouped = {};
  rows.filter((w) => w.lotId === lotId).forEach((w) => {
    if (!grouped[w.caravana] || (w.date || "") > (grouped[w.caravana].date || "")) grouped[w.caravana] = w;
  });
  return Object.values(grouped);
}

function animalStats(rows, caravana) {
  const sorted = rows.slice().sort((a,b) => (a.date || "").localeCompare(b.date || ""));
  const first = sorted[0];
  const last = sorted[sorted.length - 1];
  const days = first && last ? daysBetween(first.date, last.date) : 0;
  const gain = first && last ? last.weight - first.weight : 0;
  return {
    caravana,
    count: sorted.length,
    initialWeight: first?.weight || 0,
    currentWeight: last?.weight || 0,
    gain,
    days,
    daily: days > 0 ? gain / days : 0,
  };
}

function lotStats(data, lotId) {
  const rows = data.weighings.filter((w) => w.lotId === lotId);
  const grouped = {};
  rows.forEach((w) => {
    if (!grouped[w.caravana]) grouped[w.caravana] = [];
    grouped[w.caravana].push(w);
  });
  const animals = Object.entries(grouped).map(([c, r]) => animalStats(r, c));
  const avgWeight = animals.length ? animals.reduce((s, a) => s + a.currentWeight, 0) / animals.length : 0;
  const avgGain = animals.length ? animals.reduce((s, a) => s + a.gain, 0) / animals.length : 0;
  const validDaily = animals.filter((a) => a.days > 0).map((a) => a.daily);
  const avgDaily = validDaily.length ? validDaily.reduce((s, x) => s + x, 0) / validDaily.length : 0;
  return { animals: animals.length, avgWeight, avgGain, avgDaily };
}

async function readSpreadsheet(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function findColumn(row, candidates) {
  const keys = Object.keys(row);
  const normalizedKeys = keys.map(normalize);
  for (const candidate of candidates) {
    const i = normalizedKeys.indexOf(normalize(candidate));
    if (i >= 0) return keys[i];
  }
  return null;
}

function mapWeightRows(rows) {
  return rows.map((row) => {
    const caravanaKey = findColumn(row, ["caravana", "eid", "rfid", "vid", "id", "tag", "ear tag", "numero de caravana"]);
    const weightKey = findColumn(row, ["peso", "weight", "kg", "peso vivo"]);
    const dateKey = findColumn(row, ["fecha", "date", "fecha pesaje", "weighing date"]);
    const caravana = caravanaKey ? String(row[caravanaKey]).trim() : "";
    const weight = weightKey ? numberValue(row[weightKey]) : null;
    return { caravana, weight, date: dateValue(dateKey ? row[dateKey] : null) };
  }).filter((r) => r.caravana && r.weight !== null);
}

function mapFeedRows(rows) {
  return rows.map((row) => {
    const dateKey = findColumn(row, ["fecha", "date"]);
    const feedersKey = findColumn(row, ["comederos", "comedero", "feeders"]);
    const animalsKey = findColumn(row, ["animales", "cantidad", "animals", "animales cantidad"]);
    const foodKey = findColumn(row, ["total de comida", "comida", "kg de comida", "total comida", "kg"]);
    const saltKey = findColumn(row, ["sal", "kg sal"]);
    const bagsKey = findColumn(row, ["bolsas de comida", "bolsas", "bags"]);
    const daysKey = findColumn(row, ["dias transcurridos", "dias", "days", "días transcurridos"]);
    return {
      date: dateValue(dateKey ? row[dateKey] : null),
      feeders: feedersKey ? String(row[feedersKey]).trim() : "",
      animals: animalsKey ? numberValue(row[animalsKey]) : null,
      foodKg: foodKey ? numberValue(row[foodKey]) : null,
      salt: saltKey ? numberValue(row[saltKey]) : null,
      bags: bagsKey ? numberValue(String(row[bagsKey]).replace(/[^\d,.-]/g, "")) : null,
      days: daysKey ? numberValue(row[daysKey]) : null,
    };
  }).filter((r) => r.date || r.foodKg !== null);
}

export default App;

import { useState, useEffect, useCallback, Fragment } from "react";
import Head from "next/head";

/* ─── API helpers ─────────────────────────────────────── */
async function apiGet(params = {}) {
  const qs = new URLSearchParams(params).toString();
  const res = await fetch(`/api/sheets?${qs}`);
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}
async function apiPost(body) {
  const res = await fetch("/api/sheets", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

/* ─── Icons (inline SVG) ──────────────────────────────── */
const Icon = {
  dash: () => (
    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-6 0a1 1 0 001-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 001 1m-6 0h6" />
    </svg>
  ),
  money: () => (
    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
  ),
  map: () => (
    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z" /><path strokeLinecap="round" strokeLinejoin="round" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" />
    </svg>
  ),
  plus: () => (
    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
    </svg>
  ),
  edit: () => (
    <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
    </svg>
  ),
  trash: () => (
    <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
    </svg>
  ),
  x: () => (
    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
    </svg>
  ),
  check: () => (
    <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
    </svg>
  ),
  refresh: () => (
    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
    </svg>
  ),
  search: () => (
    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
    </svg>
  ),
  warn: () => (
    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
    </svg>
  ),
};

/* ─── Status config ───────────────────────────────────── */
const COBRANZA_ESTADOS = ["Pendiente", "Pagado", "Vencido", "Parcial"];
const VISITA_ESTADOS   = ["Programada", "Realizada", "Cancelada", "Reprogramada"];

const estadoColor = {
  Pendiente:    "bg-amber-500/15 text-amber-400 border border-amber-500/30",
  Pagado:       "bg-emerald-500/15 text-emerald-400 border border-emerald-500/30",
  Vencido:      "bg-red-500/15 text-red-400 border border-red-500/30",
  Parcial:      "bg-blue-500/15 text-blue-400 border border-blue-500/30",
  Programada:   "bg-sky-500/15 text-sky-400 border border-sky-500/30",
  Realizada:    "bg-emerald-500/15 text-emerald-400 border border-emerald-500/30",
  Cancelada:    "bg-red-500/15 text-red-400 border border-red-500/30",
  Reprogramada: "bg-purple-500/15 text-purple-400 border border-purple-500/30",
};

/* ─── Utility ─────────────────────────────────────────── */
const fmt = (n) =>
  Number(n).toLocaleString("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 });

const today = () => new Date().toISOString().split("T")[0];

/* ─── StatCard ────────────────────────────────────────── */
function StatCard({ label, value, sub, color = "amber", icon }) {
  const colors = {
    amber: "from-amber-500/20 to-transparent border-amber-500/20",
    emerald: "from-emerald-500/20 to-transparent border-emerald-500/20",
    red: "from-red-500/20 to-transparent border-red-500/20",
    sky: "from-sky-500/20 to-transparent border-sky-500/20",
  };
  const textColors = { amber: "text-amber-400", emerald: "text-emerald-400", red: "text-red-400", sky: "text-sky-400" };
  return (
    <div className={`card p-5 bg-gradient-to-br ${colors[color]} relative overflow-hidden animate-slide-up`}>
      <div className="flex items-start justify-between mb-3">
        <span className="text-xs font-semibold text-slate-500 uppercase tracking-widest">{label}</span>
        <span className={`${textColors[color]} opacity-60`}>{icon}</span>
      </div>
      <p className={`font-display text-2xl font-bold ${textColors[color]} mb-0.5`}>{value}</p>
      {sub && <p className="text-xs text-slate-500">{sub}</p>}
    </div>
  );
}

/* ─── Modal ───────────────────────────────────────────── */
function Modal({ title, onClose, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-fade-in">
      <div className="card w-full max-w-lg max-h-[90vh] overflow-y-auto animate-slide-up">
        <div className="flex items-center justify-between p-5 border-b border-white/[0.07]">
          <h2 className="font-display text-base font-semibold text-white">{title}</h2>
          <button onClick={onClose} className="text-slate-500 hover:text-white transition-colors"><Icon.x /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

/* ─── Form field ──────────────────────────────────────── */
function Field({ label, children }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-slate-400 mb-1.5 uppercase tracking-wider">{label}</label>
      {children}
    </div>
  );
}

/* ─── CobranzasTab ────────────────────────────────────── */
function CobranzasTab() {
  const [rows, setRows]       = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState(null);
  const [search, setSearch]   = useState("");
  const [filterEst, setFilter]= useState("Todos");
  const [modal, setModal]     = useState(null); // null | "add" | row (edit)
  const [saving, setSaving]   = useState(false);
  const [form, setForm]       = useState({});
  const [deleting, setDeleting] = useState(null);

  const load = useCallback(async () => {
    setLoading(true); setError(null);
    try {
      const res = await apiGet({ action: "getData", sheet: "Cobranzas" });
      setRows(Array.isArray(res.data) ? res.data : []);
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  const openAdd = () => {
    setForm({ cliente: "", monto: "", fechaVencimiento: today(), estado: "Pendiente", notas: "" });
    setModal("add");
  };
  const openEdit = (row) => { setForm({ ...row }); setModal(row); };

  const handleSave = async () => {
    if (!form.cliente || !form.monto) return;
    setSaving(true);
    try {
      if (modal === "add") {
        await apiPost({ action: "addRow", sheet: "Cobranzas", data: form });
      } else {
        await apiPost({ action: "updateRow", sheet: "Cobranzas", id: modal.id, data: form });
      }
      setModal(null); await load();
    } catch (e) { alert("Error: " + e.message); }
    finally { setSaving(false); }
  };

  const handleDelete = async (id) => {
    if (!confirm("¿Eliminar esta cobranza?")) return;
    setDeleting(id);
    try { await apiPost({ action: "deleteRow", sheet: "Cobranzas", id }); await load(); }
    catch (e) { alert("Error: " + e.message); }
    finally { setDeleting(null); }
  };

  const handlePago = async (row) => {
    try { await apiPost({ action: "updateRow", sheet: "Cobranzas", id: row.id, data: { ...row, estado: "Pagado" } }); await load(); }
    catch (e) { alert("Error: " + e.message); }
  };

  const filtered = rows.filter((r) => {
    const q = search.toLowerCase();
    const match = r.cliente?.toLowerCase().includes(q) || r.notas?.toLowerCase().includes(q);
    const est = filterEst === "Todos" || r.estado === filterEst;
    return match && est;
  });

  // Stats
  const total    = rows.reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const cobrado  = rows.filter((r) => r.estado === "Pagado").reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const pendiente= rows.filter((r) => r.estado === "Pendiente" || r.estado === "Parcial").reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const vencidas = rows.filter((r) => r.estado === "Vencido").length;

  return (
    <div className="space-y-6">
      {/* Stats */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <StatCard label="Total asignado" value={fmt(total)} icon={<Icon.money />} color="amber" />
        <StatCard label="Cobrado"         value={fmt(cobrado)} sub={`${rows.filter(r=>r.estado==="Pagado").length} pagos`} icon={<Icon.check />} color="emerald" />
        <StatCard label="Pendiente"       value={fmt(pendiente)} icon={<Icon.warn />} color="sky" />
        <StatCard label="Vencidas"        value={vencidas} sub="cobranzas" icon={<Icon.warn />} color="red" />
      </div>

      {/* Toolbar */}
      <div className="flex flex-col sm:flex-row gap-3">
        <div className="relative flex-1">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500"><Icon.search /></span>
          <input className="input pl-9" placeholder="Buscar cliente, notas…" value={search} onChange={e=>setSearch(e.target.value)} />
        </div>
        <select className="input sm:w-40" value={filterEst} onChange={e=>setFilter(e.target.value)}>
          <option>Todos</option>
          {COBRANZA_ESTADOS.map(e=><option key={e}>{e}</option>)}
        </select>
        <button className="btn-ghost" onClick={load}><Icon.refresh /></button>
        <button className="btn-primary flex items-center gap-2" onClick={openAdd}><Icon.plus /><span>Nueva cobranza</span></button>
      </div>

      {/* Table */}
      <div className="card overflow-hidden">
        {loading ? (
          <div className="flex items-center justify-center py-16 text-slate-500 text-sm">Cargando datos…</div>
        ) : error ? (
          <div className="flex items-center justify-center py-16">
            <div className="text-center">
              <p className="text-red-400 text-sm mb-2">{error}</p>
              <button className="btn-ghost text-xs" onClick={load}>Reintentar</button>
            </div>
          </div>
        ) : filtered.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 text-slate-500">
            <p className="text-sm mb-3">No hay cobranzas {search ? "que coincidan" : "registradas"}</p>
            {!search && <button className="btn-primary flex items-center gap-2" onClick={openAdd}><Icon.plus />Agregar primera cobranza</button>}
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead className="border-b border-white/[0.07]">
                <tr>
                  {["Cliente","Monto","Vencimiento","Estado","Notas","Acciones"].map(h=>(
                    <th key={h} className="th">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-white/[0.04]">
                {filtered.map((row) => (
                  <tr key={row.id} className="hover:bg-white/[0.02] transition-colors">
                    <td className="td font-medium text-white">{row.cliente}</td>
                    <td className="td font-mono text-brand-400 font-semibold">{fmt(row.monto)}</td>
                    <td className="td text-slate-400">{row.fechaVencimiento}</td>
                    <td className="td">
                      <span className={`badge ${estadoColor[row.estado] || "bg-slate-500/20 text-slate-400"}`}>
                        {row.estado}
                      </span>
                    </td>
                    <td className="td text-slate-500 max-w-[140px] truncate">{row.notas || "—"}</td>
                    <td className="td">
                      <div className="flex items-center gap-1.5">
                        {row.estado !== "Pagado" && (
                          <button className="badge bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 hover:bg-emerald-500/20 transition-colors cursor-pointer" onClick={()=>handlePago(row)}>
                            <Icon.check /> Pagar
                          </button>
                        )}
                        <button className="btn-ghost p-1.5" onClick={()=>openEdit(row)}><Icon.edit /></button>
                        <button className="btn-danger p-1.5" disabled={deleting===row.id} onClick={()=>handleDelete(row.id)}>
                          <Icon.trash />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Modal */}
      {modal && (
        <Modal title={modal === "add" ? "Nueva cobranza" : "Editar cobranza"} onClose={()=>setModal(null)}>
          <div className="space-y-4">
            <Field label="Cliente *">
              <input className="input" value={form.cliente||""} onChange={e=>setForm(p=>({...p,cliente:e.target.value}))} placeholder="Nombre del cliente" />
            </Field>
            <Field label="Monto *">
              <input className="input" type="number" value={form.monto||""} onChange={e=>setForm(p=>({...p,monto:e.target.value}))} placeholder="0" />
            </Field>
            <Field label="Fecha de vencimiento">
              <input className="input" type="date" value={form.fechaVencimiento||""} onChange={e=>setForm(p=>({...p,fechaVencimiento:e.target.value}))} />
            </Field>
            <Field label="Estado">
              <select className="input" value={form.estado||"Pendiente"} onChange={e=>setForm(p=>({...p,estado:e.target.value}))}>
                {COBRANZA_ESTADOS.map(e=><option key={e}>{e}</option>)}
              </select>
            </Field>
            <Field label="Notas">
              <textarea className="input resize-none h-20" value={form.notas||""} onChange={e=>setForm(p=>({...p,notas:e.target.value}))} placeholder="Observaciones…" />
            </Field>
            <div className="flex gap-3 pt-2">
              <button className="btn-ghost flex-1" onClick={()=>setModal(null)}>Cancelar</button>
              <button className="btn-primary flex-1" disabled={saving||!form.cliente||!form.monto} onClick={handleSave}>
                {saving ? "Guardando…" : "Guardar"}
              </button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ─── VisitasTab ──────────────────────────────────────── */
function VisitasTab() {
  const [rows, setRows]       = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState(null);
  const [search, setSearch]   = useState("");
  const [filterEst, setFilter]= useState("Todos");
  const [modal, setModal]     = useState(null);
  const [saving, setSaving]   = useState(false);
  const [form, setForm]       = useState({});
  const [deleting, setDeleting] = useState(null);

  const load = useCallback(async () => {
    setLoading(true); setError(null);
    try {
      const res = await apiGet({ action: "getData", sheet: "Visitas" });
      setRows(Array.isArray(res.data) ? res.data : []);
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  const openAdd = () => {
    setForm({ cliente: "", direccion: "", fecha: today(), hora: "09:00", estado: "Programada", notas: "" });
    setModal("add");
  };
  const openEdit = (row) => { setForm({ ...row }); setModal(row); };

  const handleSave = async () => {
    if (!form.cliente) return;
    setSaving(true);
    try {
      if (modal === "add") {
        await apiPost({ action: "addRow", sheet: "Visitas", data: form });
      } else {
        await apiPost({ action: "updateRow", sheet: "Visitas", id: modal.id, data: form });
      }
      setModal(null); await load();
    } catch (e) { alert("Error: " + e.message); }
    finally { setSaving(false); }
  };

  const handleDelete = async (id) => {
    if (!confirm("¿Eliminar esta visita?")) return;
    setDeleting(id);
    try { await apiPost({ action: "deleteRow", sheet: "Visitas", id }); await load(); }
    catch (e) { alert("Error: " + e.message); }
    finally { setDeleting(null); }
  };

  const handleRealizar = async (row) => {
    try { await apiPost({ action: "updateRow", sheet: "Visitas", id: row.id, data: { ...row, estado: "Realizada" } }); await load(); }
    catch (e) { alert("Error: " + e.message); }
  };

  const filtered = rows.filter((r) => {
    const q = search.toLowerCase();
    const match = r.cliente?.toLowerCase().includes(q) || r.direccion?.toLowerCase().includes(q) || r.notas?.toLowerCase().includes(q);
    const est = filterEst === "Todos" || r.estado === filterEst;
    return match && est;
  });

  const programadas = rows.filter(r=>r.estado==="Programada").length;
  const realizadas  = rows.filter(r=>r.estado==="Realizada").length;
  const hoy = rows.filter(r=>r.fecha===today()).length;

  return (
    <div className="space-y-6">
      {/* Stats */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <StatCard label="Total visitas"  value={rows.length} icon={<Icon.map />} color="sky" />
        <StatCard label="Programadas"    value={programadas} icon={<Icon.map />} color="amber" />
        <StatCard label="Realizadas"     value={realizadas} icon={<Icon.check />} color="emerald" />
        <StatCard label="Hoy"            value={hoy} sub="visitas agendadas" icon={<Icon.map />} color="red" />
      </div>

      {/* Toolbar */}
      <div className="flex flex-col sm:flex-row gap-3">
        <div className="relative flex-1">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500"><Icon.search /></span>
          <input className="input pl-9" placeholder="Buscar cliente, dirección…" value={search} onChange={e=>setSearch(e.target.value)} />
        </div>
        <select className="input sm:w-40" value={filterEst} onChange={e=>setFilter(e.target.value)}>
          <option>Todos</option>
          {VISITA_ESTADOS.map(e=><option key={e}>{e}</option>)}
        </select>
        <button className="btn-ghost" onClick={load}><Icon.refresh /></button>
        <button className="btn-primary flex items-center gap-2" onClick={openAdd}><Icon.plus /><span>Nueva visita</span></button>
      </div>

      {/* Table */}
      <div className="card overflow-hidden">
        {loading ? (
          <div className="flex items-center justify-center py-16 text-slate-500 text-sm">Cargando datos…</div>
        ) : error ? (
          <div className="flex items-center justify-center py-16">
            <div className="text-center">
              <p className="text-red-400 text-sm mb-2">{error}</p>
              <button className="btn-ghost text-xs" onClick={load}>Reintentar</button>
            </div>
          </div>
        ) : filtered.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 text-slate-500">
            <p className="text-sm mb-3">No hay visitas {search ? "que coincidan" : "registradas"}</p>
            {!search && <button className="btn-primary flex items-center gap-2" onClick={openAdd}><Icon.plus />Agregar primera visita</button>}
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead className="border-b border-white/[0.07]">
                <tr>
                  {["Cliente","Dirección","Fecha","Hora","Estado","Notas","Acciones"].map(h=>(
                    <th key={h} className="th">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-white/[0.04]">
                {filtered.map((row) => (
                  <tr key={row.id} className="hover:bg-white/[0.02] transition-colors">
                    <td className="td font-medium text-white">{row.cliente}</td>
                    <td className="td text-slate-400 max-w-[160px] truncate">{row.direccion || "—"}</td>
                    <td className="td text-slate-400">{row.fecha}</td>
                    <td className="td font-mono text-slate-300">{row.hora || "—"}</td>
                    <td className="td">
                      <span className={`badge ${estadoColor[row.estado] || "bg-slate-500/20 text-slate-400"}`}>
                        {row.estado}
                      </span>
                    </td>
                    <td className="td text-slate-500 max-w-[120px] truncate">{row.notas || "—"}</td>
                    <td className="td">
                      <div className="flex items-center gap-1.5">
                        {row.estado === "Programada" && (
                          <button className="badge bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 hover:bg-emerald-500/20 transition-colors cursor-pointer" onClick={()=>handleRealizar(row)}>
                            <Icon.check /> Realizada
                          </button>
                        )}
                        <button className="btn-ghost p-1.5" onClick={()=>openEdit(row)}><Icon.edit /></button>
                        <button className="btn-danger p-1.5" disabled={deleting===row.id} onClick={()=>handleDelete(row.id)}>
                          <Icon.trash />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Modal */}
      {modal && (
        <Modal title={modal === "add" ? "Nueva visita" : "Editar visita"} onClose={()=>setModal(null)}>
          <div className="space-y-4">
            <Field label="Cliente *">
              <input className="input" value={form.cliente||""} onChange={e=>setForm(p=>({...p,cliente:e.target.value}))} placeholder="Nombre del cliente" />
            </Field>
            <Field label="Dirección">
              <input className="input" value={form.direccion||""} onChange={e=>setForm(p=>({...p,direccion:e.target.value}))} placeholder="Calle 123, Ciudad" />
            </Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Fecha">
                <input className="input" type="date" value={form.fecha||""} onChange={e=>setForm(p=>({...p,fecha:e.target.value}))} />
              </Field>
              <Field label="Hora">
                <input className="input" type="time" value={form.hora||""} onChange={e=>setForm(p=>({...p,hora:e.target.value}))} />
              </Field>
            </div>
            <Field label="Estado">
              <select className="input" value={form.estado||"Programada"} onChange={e=>setForm(p=>({...p,estado:e.target.value}))}>
                {VISITA_ESTADOS.map(e=><option key={e}>{e}</option>)}
              </select>
            </Field>
            <Field label="Notas">
              <textarea className="input resize-none h-20" value={form.notas||""} onChange={e=>setForm(p=>({...p,notas:e.target.value}))} placeholder="Observaciones…" />
            </Field>
            <div className="flex gap-3 pt-2">
              <button className="btn-ghost flex-1" onClick={()=>setModal(null)}>Cancelar</button>
              <button className="btn-primary flex-1" disabled={saving||!form.cliente} onClick={handleSave}>
                {saving ? "Guardando…" : "Guardar"}
              </button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ─── Main Page ───────────────────────────────────────── */
const TABS = [
  { id: "cobranzas", label: "Cobranzas", icon: <Icon.money /> },
  { id: "visitas",   label: "Visitas",   icon: <Icon.map /> },
];

export default function Home() {
  const [tab, setTab] = useState("cobranzas");

  return (
    <>
      <Head>
        <title>Gestión de Cobranzas y Visitas</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>💼</text></svg>" />
      </Head>

      {/* Ambient gradient background */}
      <div className="fixed inset-0 pointer-events-none overflow-hidden z-0">
        <div className="absolute -top-40 -left-40 w-96 h-96 bg-amber-500/5 rounded-full blur-3xl" />
        <div className="absolute top-1/3 right-0 w-80 h-80 bg-sky-500/5 rounded-full blur-3xl" />
      </div>

      <div className="relative z-10 min-h-screen">
        {/* Header */}
        <header className="glass border-b border-white/[0.07] sticky top-0 z-40">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="flex items-center justify-between h-16">
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 rounded-lg bg-brand-500/20 border border-brand-500/40 flex items-center justify-center text-brand-400">
                  <Icon.dash />
                </div>
                <div>
                  <h1 className="font-display font-bold text-white text-sm leading-none">GestiónApp</h1>
                  <p className="text-[10px] text-slate-500 mt-0.5">Cobranzas &amp; Visitas</p>
                </div>
              </div>
              <nav className="flex items-center">
                {TABS.map((t) => (
                  <button
                    key={t.id}
                    onClick={() => setTab(t.id)}
                    className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-lg transition-all duration-200 ${
                      tab === t.id
                        ? "bg-brand-500/15 text-brand-400 border border-brand-500/30"
                        : "text-slate-500 hover:text-slate-300"
                    }`}
                  >
                    {t.icon}
                    <span className="hidden sm:inline">{t.label}</span>
                  </button>
                ))}
              </nav>
            </div>
          </div>
        </header>

        {/* Main content */}
        <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
          <div className="mb-6">
            <h2 className="font-display text-2xl font-bold text-white mb-1">
              {tab === "cobranzas" ? "Cobranzas" : "Visitas"}
            </h2>
            <p className="text-sm text-slate-500">
              {tab === "cobranzas"
                ? "Gestión de pagos y deudas de clientes"
                : "Agenda y seguimiento de visitas a clientes"}
            </p>
          </div>

          <div className="animate-fade-in" key={tab}>
            {tab === "cobranzas" ? <CobranzasTab /> : <VisitasTab />}
          </div>
        </main>

        {/* Footer */}
        <footer className="border-t border-white/[0.05] mt-12">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
            <p className="text-center text-xs text-slate-600">
              GestiónApp · Sincronizado con Google Sheets
            </p>
          </div>
        </footer>
      </div>
    </>
  );
}

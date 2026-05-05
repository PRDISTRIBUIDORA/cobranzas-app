import { useState, useEffect, useCallback, useRef } from "react";
import Head from "next/head";
import * as XLSX from "xlsx";

/* ─── API ─────────────────────────────────────────────── */
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

/* ─── Constants ───────────────────────────────────────── */
const ADMIN_PASSWORD = "admin2024";
const today = () => new Date().toISOString().split("T")[0];
const fmt = (n) => Number(n || 0).toLocaleString("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 });
const COBRANZA_ESTADOS = ["Pendiente", "Pagado", "Vencido", "Parcial"];
const VISITA_ESTADOS = ["Programada", "Realizada", "Cancelada", "Reprogramada"];

/* ─── Styles ──────────────────────────────────────────── */
const S = {
  page:      { minHeight: "100vh", background: "#080b10", color: "#e2e8f0", fontFamily: "'DM Sans',system-ui,sans-serif" },
  card:      { background: "#161b22", border: "1px solid rgba(255,255,255,.08)", borderRadius: 14 },
  input:     { width: "100%", background: "#1e2530", border: "1px solid rgba(255,255,255,.1)", borderRadius: 8, padding: "8px 12px", color: "#e2e8f0", fontSize: 13, outline: "none", boxSizing: "border-box" },
  label:     { display: "block", fontSize: 11, fontWeight: 700, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 5 },
  btnPri:    { background: "#eab308", color: "#000", border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" },
  btnGhost:  { background: "transparent", color: "#94a3b8", border: "1px solid rgba(255,255,255,.1)", borderRadius: 8, padding: "7px 12px", fontSize: 12, cursor: "pointer" },
  btnDanger: { background: "#7f1d1d22", color: "#f87171", border: "1px solid #ef444433", borderRadius: 6, padding: "4px 8px", fontSize: 11, cursor: "pointer" },
  btnGreen:  { background: "#06522233", color: "#34d399", border: "1px solid #10b98133", borderRadius: 8, padding: "8px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" },
  th:        { textAlign: "left", padding: "10px 14px", fontSize: 11, color: "#475569", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em" },
  td:        { padding: "11px 14px", fontSize: 13, color: "#cbd5e1", borderBottom: "1px solid rgba(255,255,255,.04)" },
};

const ESTADO_COLOR = { Pendiente: "#fbbf24", Pagado: "#34d399", Vencido: "#f87171", Parcial: "#60a5fa", Programada: "#38bdf8", Realizada: "#34d399", Cancelada: "#f87171", Reprogramada: "#c084fc" };
const ESTADO_BG    = { Pendiente: "#92400e18", Pagado: "#06522218", Vencido: "#7f1d1d18", Parcial: "#1e3a5f18", Programada: "#0c344918", Realizada: "#06522218", Cancelada: "#7f1d1d18", Reprogramada: "#3b1f6618" };

function Badge({ estado }) {
  return <span style={{ background: ESTADO_BG[estado] || "#1f293718", color: ESTADO_COLOR[estado] || "#94a3b8", border: `1px solid ${ESTADO_COLOR[estado] || "#94a3b8"}33`, padding: "2px 10px", borderRadius: 99, fontSize: 11, fontWeight: 600 }}>{estado}</span>;
}
function StatCard({ label, value, sub, color = "amber" }) {
  const c = { amber: "#fbbf24", emerald: "#34d399", red: "#f87171", sky: "#38bdf8", purple: "#c084fc" }[color];
  return <div style={{ ...S.card, padding: "14px 16px", flex: 1, minWidth: 120, border: `1px solid ${c}22` }}>
    <div style={{ fontSize: 10, color: "#475569", textTransform: "uppercase", letterSpacing: "0.08em", fontWeight: 700, marginBottom: 6 }}>{label}</div>
    <div style={{ fontSize: 20, fontWeight: 800, color: c, marginBottom: 1 }}>{value}</div>
    {sub && <div style={{ fontSize: 11, color: "#475569" }}>{sub}</div>}
  </div>;
}
function Modal({ title, onClose, children }) {
  return <div style={{ position: "fixed", inset: 0, zIndex: 100, background: "rgba(0,0,0,.75)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
    <div style={{ ...S.card, width: "100%", maxWidth: 480, maxHeight: "90vh", overflowY: "auto" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px", borderBottom: "1px solid rgba(255,255,255,.07)" }}>
        <span style={{ fontWeight: 700, color: "#fff", fontSize: 14 }}>{title}</span>
        <button onClick={onClose} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 18 }}>✕</button>
      </div>
      <div style={{ padding: 18 }}>{children}</div>
    </div>
  </div>;
}
function Field({ label, children }) { return <div style={{ marginBottom: 13 }}><label style={S.label}>{label}</label>{children}</div>; }

/* ─── Excel Parsers ───────────────────────────────────── */
function parseClientesExcel(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const clientes = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const codigo = String(row[0] || "").trim();
    const nombre = String(row[1] || "").trim();
    const localidad = String(row[2] || "").trim();
    if (codigo && nombre && codigo !== "Código" && codigo !== "codigo" && !isNaN(Number(codigo))) {
      clientes.push({ codigo, nombre, localidad });
    }
  }
  return clientes;
}

function parseCobranzasExcel(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const deudores = [];
  const comprobantes = [];
  let currentCliente = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const colA = String(row[0] || "").trim();
    const colB = String(row[1] || "").trim();
    const colC = String(row[2] || "").trim();
    const colD = String(row[3] || "").trim();

    // Detect client header row: col A has name (non-numeric, non-"Nombre", non-"Cliente")
    if (colA && colA !== "Nombre" && colA !== "Cliente" && isNaN(Number(colA))) {
      // This is a client row
      const saldo = parseFloat(String(row[2] || "0").replace(/[^\d.,]/g, "").replace(",", ".")) || 0;
      const localidad = String(row[3] || "").trim();
      currentCliente = { nombre: colA.replace(/^\+\s*/, ""), saldo, localidad, codigo: null };
      deudores.push(currentCliente);
    }
    // Detect detail row: col A is numeric (client code)
    else if (colA && !isNaN(Number(colA)) && currentCliente) {
      const codigo = colA;
      if (!currentCliente.codigo) currentCliente.codigo = codigo;
      // Parse date - could be Excel serial or string
      let fecha = colB;
      if (!isNaN(Number(colB)) && Number(colB) > 40000) {
        const d = XLSX.SSF.parse_date_code(Number(colB));
        fecha = `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
      }
      const comprobante = String(row[3] || "").trim();
      const importe = parseFloat(String(row[4] || "0").replace(/[^\d.,]/g, "").replace(",", ".")) || 0;
      comprobantes.push({
        codigo,
        cliente: currentCliente.nombre,
        fecha,
        comprobante,
        importe,
      });
    }
  }
  return { deudores, comprobantes };
}

/* ─── Export Resumen to Excel ─────────────────────────── */
function exportResumenExcel(cobranzas, visitas, deudores) {
  const wb = XLSX.utils.book_new();
  const fechaHoy = new Date().toLocaleDateString("es-AR");

  // Resumen general
  const resumenData = [
    ["RESUMEN SEMANAL — " + fechaHoy],
    [],
    ["COBRANZAS"],
    ["Total cobrado", cobranzas.filter(r => r.estado === "Pagado").reduce((a, r) => a + (Number(r.monto) || 0), 0)],
    ["Pendiente",    cobranzas.filter(r => ["Pendiente","Parcial"].includes(r.estado)).reduce((a, r) => a + (Number(r.monto) || 0), 0)],
    ["Vencido",      cobranzas.filter(r => r.estado === "Vencido").reduce((a, r) => a + (Number(r.monto) || 0), 0)],
    [],
    ["VISITAS"],
    ["Realizadas",   visitas.filter(r => r.estado === "Realizada").length],
    ["Programadas",  visitas.filter(r => r.estado === "Programada").length],
    [],
    ["DEUDORES"],
    ["Total deuda",  deudores.reduce((a, r) => a + (Number(r.saldo) || 0), 0)],
    ["Cantidad",     deudores.length],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), "Resumen");

  // Por cobrador
  const cobradores = [...new Set(cobranzas.map(r => r.cobrador).filter(Boolean))];
  const porCobradorData = [["Cobrador", "Cobrado", "Pendiente", "Visitas realizadas"]];
  cobradores.forEach(nombre => {
    porCobradorData.push([
      nombre,
      cobranzas.filter(r => r.cobrador === nombre && r.estado === "Pagado").reduce((a, r) => a + (Number(r.monto) || 0), 0),
      cobranzas.filter(r => r.cobrador === nombre && ["Pendiente","Parcial"].includes(r.estado)).reduce((a, r) => a + (Number(r.monto) || 0), 0),
      visitas.filter(r => r.cobrador === nombre && r.estado === "Realizada").length,
    ]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(porCobradorData), "Por Cobrador");

  // Mayores deudores
  const deudoresData = [["Código", "Nombre", "Localidad", "Saldo"]];
  [...deudores].sort((a, b) => (Number(b.saldo) || 0) - (Number(a.saldo) || 0)).forEach(r => {
    deudoresData.push([r.codigo || "", r.nombre, r.localidad || "", Number(r.saldo) || 0]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(deudoresData), "Deudores");

  // Cobranzas detalle
  const cobData = [["Código","Cliente","Monto","Vencimiento","Estado","Cobrador","Notas"]];
  cobranzas.forEach(r => cobData.push([r.codigo||"", r.cliente, Number(r.monto)||0, r.fechaVencimiento, r.estado, r.cobrador||"", r.notas||""]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cobData), "Cobranzas");

  XLSX.writeFile(wb, `Resumen_${fechaHoy.replace(/\//g,"-")}.xlsx`);
}

/* ─── LOGIN ───────────────────────────────────────────── */
function Login({ onLogin }) {
  const [pw, setPw] = useState(""), [err, setErr] = useState(false);
  const handle = () => { if (pw === ADMIN_PASSWORD) { onLogin(); } else setErr(true); };
  return <div style={{ minHeight: "100vh", background: "#080b10", display: "flex", alignItems: "center", justifyContent: "center" }}>
    <div style={{ ...S.card, padding: 28, width: 320 }}>
      <div style={{ textAlign: "center", marginBottom: 20 }}>
        <div style={{ fontSize: 28, marginBottom: 6 }}>🔐</div>
        <div style={{ fontWeight: 800, color: "#fff", fontSize: 16 }}>Panel Administrador</div>
      </div>
      <Field label="Contraseña">
        <input style={S.input} type="password" value={pw} onChange={e => setPw(e.target.value)} onKeyDown={e => e.key === "Enter" && handle()} placeholder="••••••••" />
      </Field>
      {err && <p style={{ color: "#f87171", fontSize: 12, marginBottom: 8 }}>Contraseña incorrecta</p>}
      <button style={{ ...S.btnPri, width: "100%", padding: 10 }} onClick={handle}>Ingresar</button>
    </div>
  </div>;
}

/* ─── IMPORTAR TAB ────────────────────────────────────── */
function ImportarTab() {
  const [status, setStatus] = useState(null);
  const [loading, setLoading] = useState(false);
  const refClientes = useRef();
  const refCobranzas = useRef();

  const importarClientes = async (file) => {
    setLoading(true); setStatus(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const clientes = parseClientesExcel(wb);
      if (clientes.length === 0) throw new Error("No se encontraron clientes en el archivo");
      const res = await apiPost({ action: "bulkUpsert", sheet: "Clientes", rows: clientes });
      setStatus({ type: "success", msg: `✅ Clientes importados: ${res.added} nuevos, ${res.updated} actualizados` });
    } catch (e) { setStatus({ type: "error", msg: "❌ Error: " + e.message }); }
    finally { setLoading(false); }
  };

  const importarCobranzas = async (file) => {
    setLoading(true); setStatus(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const { deudores, comprobantes } = parseCobranzasExcel(wb);
      if (deudores.length === 0) throw new Error("No se encontraron clientes en el archivo");

      // Update deudores
      await apiPost({ action: "bulkUpsert", sheet: "Deudores", rows: deudores });
      // Replace comprobantes
      await apiPost({ action: "clearAndInsert", sheet: "Comprobantes", rows: comprobantes });

      setStatus({ type: "success", msg: `✅ Importado: ${deudores.length} clientes/deudores, ${comprobantes.length} comprobantes reemplazados` });
    } catch (e) { setStatus({ type: "error", msg: "❌ Error: " + e.message }); }
    finally { setLoading(false); }
  };

  return <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
    {/* Importar Clientes */}
    <div style={S.card}>
      <div style={{ padding: "14px 18px", borderBottom: "1px solid rgba(255,255,255,.07)" }}>
        <div style={{ fontWeight: 700, color: "#fff", fontSize: 14 }}>👥 Importar lista de Clientes</div>
        <div style={{ fontSize: 12, color: "#475569", marginTop: 3 }}>Excel con columnas: Código | Nombre | Localidad</div>
      </div>
      <div style={{ padding: 18 }}>
        <input ref={refClientes} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => e.target.files[0] && importarClientes(e.target.files[0])} />
        <button style={{ ...S.btnPri, opacity: loading ? 0.5 : 1 }} disabled={loading} onClick={() => refClientes.current?.click()}>
          📂 Seleccionar archivo Excel
        </button>
        <p style={{ fontSize: 11, color: "#475569", marginTop: 10, margin: "10px 0 0" }}>
          Si el cliente ya existe (mismo código), actualiza nombre y localidad. Los nuevos se agregan.
        </p>
      </div>
    </div>

    {/* Importar Cobranzas/Deudores */}
    <div style={S.card}>
      <div style={{ padding: "14px 18px", borderBottom: "1px solid rgba(255,255,255,.07)" }}>
        <div style={{ fontWeight: 700, color: "#fff", fontSize: 14 }}>💰 Importar Cobranzas / Deudores</div>
        <div style={{ fontSize: 12, color: "#475569", marginTop: 3 }}>Excel exportado de tu sistema — formato agrupado por cliente</div>
      </div>
      <div style={{ padding: 18 }}>
        <div style={{ ...S.card, padding: 12, marginBottom: 14, background: "#eab30808", border: "1px solid #eab30833" }}>
          <p style={{ fontSize: 12, color: "#fbbf24", margin: 0, fontWeight: 600 }}>⚠️ Importante: exportá el resumen ANTES de importar</p>
          <p style={{ fontSize: 11, color: "#92400e", margin: "4px 0 0" }}>Los comprobantes anteriores serán reemplazados. Los saldos se actualizan.</p>
        </div>
        <input ref={refCobranzas} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => e.target.files[0] && importarCobranzas(e.target.files[0])} />
        <button style={{ ...S.btnPri, opacity: loading ? 0.5 : 1 }} disabled={loading} onClick={() => refCobranzas.current?.click()}>
          📂 Seleccionar archivo Excel
        </button>
        <p style={{ fontSize: 11, color: "#475569", marginTop: 10, margin: "10px 0 0" }}>
          Actualiza saldos de deudores y reemplaza todos los comprobantes.
        </p>
      </div>
    </div>

    {loading && <div style={{ ...S.card, padding: 16, textAlign: "center", color: "#fbbf24" }}>⏳ Procesando archivo…</div>}
    {status && <div style={{ ...S.card, padding: 16, color: status.type === "success" ? "#34d399" : "#f87171", background: status.type === "success" ? "#06522210" : "#7f1d1d10" }}>{status.msg}</div>}
  </div>;
}

/* ─── CLIENTES TAB ────────────────────────────────────── */
function ClientesTab() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({});
  const [saving, setSaving] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    try { const r = await apiGet({ action: "getData", sheet: "Clientes" }); setRows(Array.isArray(r.data) ? r.data : []); }
    catch (e) { console.error(e); } finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);

  const filtered = rows.filter(r => r.nombre?.toLowerCase().includes(search.toLowerCase()) || r.localidad?.toLowerCase().includes(search.toLowerCase()) || String(r.codigo||"").includes(search));

  const save = async () => {
    if (!form.nombre) return; setSaving(true);
    try {
      if (modal === "add") await apiPost({ action: "addRow", sheet: "Clientes", data: form });
      else await apiPost({ action: "updateRow", sheet: "Clientes", id: modal.id, data: form });
      setModal(null); await load();
    } catch (e) { alert("Error: " + e.message); } finally { setSaving(false); }
  };

  return <div>
    <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
      <input style={{ ...S.input, flex: 1, minWidth: 180 }} placeholder="🔍 Buscar por nombre, localidad o código…" value={search} onChange={e => setSearch(e.target.value)} />
      <button style={S.btnPri} onClick={() => { setForm({ codigo: "", nombre: "", localidad: "" }); setModal("add"); }}>+ Nuevo cliente</button>
    </div>
    <div style={{ ...S.card, overflow: "hidden" }}>
      {loading ? <div style={{ padding: 40, textAlign: "center", color: "#475569" }}>Cargando…</div> :
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ borderBottom: "1px solid rgba(255,255,255,.07)" }}>
            {["Código", "Nombre", "Localidad", "Acciones"].map(h => <th key={h} style={S.th}>{h}</th>)}
          </tr></thead>
          <tbody>
            {filtered.length === 0 ? <tr><td colSpan={4} style={{ ...S.td, textAlign: "center", padding: 32, color: "#475569" }}>Sin clientes</td></tr>
              : filtered.map(r => <tr key={r.id}>
                <td style={{ ...S.td, color: "#fbbf24", fontFamily: "monospace" }}>{r.codigo || "—"}</td>
                <td style={{ ...S.td, color: "#fff", fontWeight: 600 }}>{r.nombre}</td>
                <td style={S.td}>{r.localidad || "—"}</td>
                <td style={S.td}><div style={{ display: "flex", gap: 5 }}>
                  <button style={S.btnGhost} onClick={() => { setForm({ ...r }); setModal(r); }}>✏️</button>
                  <button style={S.btnDanger} onClick={async () => { if (confirm("¿Eliminar?")) { await apiPost({ action: "deleteRow", sheet: "Clientes", id: r.id }); load(); } }}>🗑</button>
                </div></td>
              </tr>)}
          </tbody>
        </table>}
    </div>
    {modal && <Modal title={modal === "add" ? "Nuevo cliente" : "Editar cliente"} onClose={() => setModal(null)}>
      <Field label="Código"><input style={S.input} value={form.codigo || ""} onChange={e => setForm(p => ({ ...p, codigo: e.target.value }))} placeholder="Ej: 3156" /></Field>
      <Field label="Nombre *"><input style={S.input} value={form.nombre || ""} onChange={e => setForm(p => ({ ...p, nombre: e.target.value }))} placeholder="Nombre del cliente" /></Field>
      <Field label="Localidad"><input style={S.input} value={form.localidad || ""} onChange={e => setForm(p => ({ ...p, localidad: e.target.value }))} placeholder="Ciudad" /></Field>
      <div style={{ display: "flex", gap: 10 }}>
        <button style={{ ...S.btnGhost, flex: 1 }} onClick={() => setModal(null)}>Cancelar</button>
        <button style={{ ...S.btnPri, flex: 1 }} disabled={saving || !form.nombre} onClick={save}>{saving ? "Guardando…" : "Guardar"}</button>
      </div>
    </Modal>}
  </div>;
}

/* ─── DEUDORES TAB ────────────────────────────────────── */
function DeudoresTab() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    try { const r = await apiGet({ action: "getData", sheet: "Deudores" }); setRows(Array.isArray(r.data) ? r.data : []); }
    catch (e) { console.error(e); } finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);

  const filtered = rows.filter(r => r.nombre?.toLowerCase().includes(search.toLowerCase()) || r.localidad?.toLowerCase().includes(search.toLowerCase()) || String(r.codigo || "").includes(search));
  const total = rows.reduce((a, r) => a + (Number(r.saldo) || 0), 0);

  return <div>
    <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
      <StatCard label="Total deudores" value={rows.length} color="red" />
      <StatCard label="Deuda total" value={fmt(total)} color="amber" />
    </div>
    <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
      <input style={{ ...S.input, flex: 1 }} placeholder="🔍 Buscar…" value={search} onChange={e => setSearch(e.target.value)} />
      <button style={S.btnGhost} onClick={load}>↺</button>
    </div>
    <div style={{ ...S.card, overflow: "hidden" }}>
      {loading ? <div style={{ padding: 40, textAlign: "center", color: "#475569" }}>Cargando…</div> :
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr style={{ borderBottom: "1px solid rgba(255,255,255,.07)" }}>
              {["Código", "Nombre", "Localidad", "Saldo"].map(h => <th key={h} style={S.th}>{h}</th>)}
            </tr></thead>
            <tbody>
              {filtered.length === 0 ? <tr><td colSpan={4} style={{ ...S.td, textAlign: "center", padding: 32, color: "#475569" }}>Sin deudores</td></tr>
                : [...filtered].sort((a, b) => (Number(b.saldo) || 0) - (Number(a.saldo) || 0)).map(r => <tr key={r.id}>
                  <td style={{ ...S.td, color: "#fbbf24", fontFamily: "monospace" }}>{r.codigo || "—"}</td>
                  <td style={{ ...S.td, color: "#fff", fontWeight: 600 }}>{r.nombre}</td>
                  <td style={S.td}>{r.localidad || "—"}</td>
                  <td style={{ ...S.td, color: "#f87171", fontFamily: "monospace", fontWeight: 700 }}>{fmt(r.saldo)}</td>
                </tr>)}
            </tbody>
          </table>
        </div>}
    </div>
  </div>;
}

/* ─── COBRANZAS TAB ───────────────────────────────────── */
function CobranzasTab({ isAdmin, clientes }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("Todos");
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({});
  const [saving, setSaving] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    try { const r = await apiGet({ action: "getData", sheet: "Cobranzas" }); setRows(Array.isArray(r.data) ? r.data : []); }
    catch (e) { console.error(e); } finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);

  const filtered = rows.filter(r => (r.cliente?.toLowerCase().includes(search.toLowerCase()) || String(r.codigo||"").includes(search)) && (filter === "Todos" || r.estado === filter));

  const save = async () => {
    if (!form.cliente) return; setSaving(true);
    try {
      if (modal === "add") await apiPost({ action: "addRow", sheet: "Cobranzas", data: form });
      else await apiPost({ action: "updateRow", sheet: "Cobranzas", id: modal.id, data: form });
      setModal(null); await load();
    } catch (e) { alert("Error: " + e.message); } finally { setSaving(false); }
  };

  const total     = rows.reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const cobrado   = rows.filter(r => r.estado === "Pagado").reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const pendiente = rows.filter(r => ["Pendiente","Parcial"].includes(r.estado)).reduce((a, r) => a + (Number(r.monto) || 0), 0);

  return <div>
    <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
      <StatCard label="Total" value={fmt(total)} color="amber" />
      <StatCard label="Cobrado" value={fmt(cobrado)} sub={`${rows.filter(r => r.estado === "Pagado").length} pagos`} color="emerald" />
      <StatCard label="Pendiente" value={fmt(pendiente)} color="sky" />
      <StatCard label="Vencidas" value={rows.filter(r => r.estado === "Vencido").length} color="red" />
    </div>
    <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
      <input style={{ ...S.input, flex: 1, minWidth: 160 }} placeholder="🔍 Buscar cliente o código…" value={search} onChange={e => setSearch(e.target.value)} />
      <select style={{ ...S.input, width: 130 }} value={filter} onChange={e => setFilter(e.target.value)}>
        {["Todos", ...COBRANZA_ESTADOS].map(e => <option key={e}>{e}</option>)}
      </select>
      <button style={S.btnPri} onClick={() => { setForm({ codigo: "", cliente: "", monto: "", fechaVencimiento: today(), estado: "Pendiente", cobrador: "", notas: "" }); setModal("add"); }}>+ Nueva</button>
    </div>
    <div style={{ ...S.card, overflow: "hidden" }}>
      {loading ? <div style={{ padding: 40, textAlign: "center", color: "#475569" }}>Cargando…</div> :
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr style={{ borderBottom: "1px solid rgba(255,255,255,.07)" }}>
              {["Código","Cliente","Monto","Vence","Estado","Cobrador","Acciones"].map(h => <th key={h} style={S.th}>{h}</th>)}
            </tr></thead>
            <tbody>
              {filtered.length === 0 ? <tr><td colSpan={7} style={{ ...S.td, textAlign: "center", padding: 32, color: "#475569" }}>Sin resultados</td></tr>
                : filtered.map(r => <tr key={r.id}>
                  <td style={{ ...S.td, color: "#fbbf24", fontFamily: "monospace", fontSize: 11 }}>{r.codigo || "—"}</td>
                  <td style={{ ...S.td, color: "#fff", fontWeight: 600 }}>{r.cliente}</td>
                  <td style={{ ...S.td, color: "#fbbf24", fontFamily: "monospace", fontWeight: 700 }}>{fmt(r.monto)}</td>
                  <td style={{ ...S.td, color: "#64748b", fontSize: 12 }}>{r.fechaVencimiento}</td>
                  <td style={S.td}><Badge estado={r.estado} /></td>
                  <td style={{ ...S.td, color: "#94a3b8" }}>{r.cobrador || "—"}</td>
                  <td style={S.td}><div style={{ display: "flex", gap: 5 }}>
                    {r.estado !== "Pagado" && <button onClick={async () => { await apiPost({ action: "updateRow", sheet: "Cobranzas", id: r.id, data: { ...r, estado: "Pagado" } }); load(); }} style={{ background: "#06522222", color: "#34d399", border: "1px solid #10b98133", borderRadius: 6, padding: "3px 8px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>✓ Pagar</button>}
                    <button style={S.btnGhost} onClick={() => { setForm({ ...r }); setModal(r); }}>✏️</button>
                    {isAdmin && <button style={S.btnDanger} onClick={async () => { if (confirm("¿Eliminar?")) { await apiPost({ action: "deleteRow", sheet: "Cobranzas", id: r.id }); load(); } }}>🗑</button>}
                  </div></td>
                </tr>)}
            </tbody>
          </table>
        </div>}
    </div>
    {modal && <Modal title={modal === "add" ? "Nueva cobranza" : "Editar cobranza"} onClose={() => setModal(null)}>
      <Field label="Cliente *">
        {clientes.length > 0
          ? <select style={S.input} value={form.cliente || ""} onChange={e => {
              const c = clientes.find(x => x.nombre === e.target.value);
              setForm(p => ({ ...p, cliente: e.target.value, codigo: c?.codigo || p.codigo }));
            }}>
            <option value="">— Seleccionar —</option>
            {clientes.map(c => <option key={c.id} value={c.nombre}>{c.nombre} ({c.localidad})</option>)}
          </select>
          : <input style={S.input} value={form.cliente || ""} onChange={e => setForm(p => ({ ...p, cliente: e.target.value }))} placeholder="Nombre" />}
      </Field>
      <Field label="Monto"><input style={S.input} type="number" value={form.monto || ""} onChange={e => setForm(p => ({ ...p, monto: e.target.value }))} placeholder="0" /></Field>
      <Field label="Vencimiento"><input style={S.input} type="date" value={form.fechaVencimiento || ""} onChange={e => setForm(p => ({ ...p, fechaVencimiento: e.target.value }))} /></Field>
      <Field label="Estado"><select style={S.input} value={form.estado || "Pendiente"} onChange={e => setForm(p => ({ ...p, estado: e.target.value }))}>{COBRANZA_ESTADOS.map(e => <option key={e}>{e}</option>)}</select></Field>
      <Field label="Cobrador"><input style={S.input} value={form.cobrador || ""} onChange={e => setForm(p => ({ ...p, cobrador: e.target.value }))} placeholder="Juan / Anto / Sabri" /></Field>
      <Field label="Notas"><textarea style={{ ...S.input, height: 60, resize: "none" }} value={form.notas || ""} onChange={e => setForm(p => ({ ...p, notas: e.target.value }))} /></Field>
      <div style={{ display: "flex", gap: 10 }}>
        <button style={{ ...S.btnGhost, flex: 1 }} onClick={() => setModal(null)}>Cancelar</button>
        <button style={{ ...S.btnPri, flex: 1 }} disabled={saving || !form.cliente} onClick={save}>{saving ? "Guardando…" : "Guardar"}</button>
      </div>
    </Modal>}
  </div>;
}

/* ─── VISITAS TAB ─────────────────────────────────────── */
function VisitasTab({ isAdmin, clientes }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("Todos");
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({});
  const [saving, setSaving] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    try { const r = await apiGet({ action: "getData", sheet: "Visitas" }); setRows(Array.isArray(r.data) ? r.data : []); }
    catch (e) { console.error(e); } finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);

  const filtered = rows.filter(r => (r.cliente?.toLowerCase().includes(search.toLowerCase())) && (filter === "Todos" || r.estado === filter));

  const save = async () => {
    if (!form.cliente) return; setSaving(true);
    try {
      if (modal === "add") await apiPost({ action: "addRow", sheet: "Visitas", data: form });
      else await apiPost({ action: "updateRow", sheet: "Visitas", id: modal.id, data: form });
      setModal(null); await load();
    } catch (e) { alert("Error: " + e.message); } finally { setSaving(false); }
  };

  return <div>
    <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
      <StatCard label="Total" value={rows.length} color="sky" />
      <StatCard label="Programadas" value={rows.filter(r => r.estado === "Programada").length} color="amber" />
      <StatCard label="Realizadas" value={rows.filter(r => r.estado === "Realizada").length} color="emerald" />
      <StatCard label="Hoy" value={rows.filter(r => r.fecha === today()).length} sub="agendadas" color="red" />
    </div>
    <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
      <input style={{ ...S.input, flex: 1, minWidth: 160 }} placeholder="🔍 Buscar cliente…" value={search} onChange={e => setSearch(e.target.value)} />
      <select style={{ ...S.input, width: 140 }} value={filter} onChange={e => setFilter(e.target.value)}>
        {["Todos", ...VISITA_ESTADOS].map(e => <option key={e}>{e}</option>)}
      </select>
      <button style={S.btnPri} onClick={() => { setForm({ codigo: "", cliente: "", direccion: "", fecha: today(), hora: "09:00", estado: "Programada", cobrador: "", notas: "" }); setModal("add"); }}>+ Nueva</button>
    </div>
    <div style={{ ...S.card, overflow: "hidden" }}>
      {loading ? <div style={{ padding: 40, textAlign: "center", color: "#475569" }}>Cargando…</div> :
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr style={{ borderBottom: "1px solid rgba(255,255,255,.07)" }}>
              {["Cliente","Dirección","Fecha","Hora","Estado","Cobrador","Acciones"].map(h => <th key={h} style={S.th}>{h}</th>)}
            </tr></thead>
            <tbody>
              {filtered.length === 0 ? <tr><td colSpan={7} style={{ ...S.td, textAlign: "center", padding: 32, color: "#475569" }}>Sin resultados</td></tr>
                : filtered.map(r => <tr key={r.id}>
                  <td style={{ ...S.td, color: "#fff", fontWeight: 600 }}>{r.cliente}</td>
                  <td style={{ ...S.td, color: "#64748b", fontSize: 12 }}>{r.direccion || "—"}</td>
                  <td style={{ ...S.td, color: "#64748b", fontSize: 12 }}>{r.fecha}</td>
                  <td style={{ ...S.td, fontFamily: "monospace", fontSize: 12 }}>{r.hora || "—"}</td>
                  <td style={S.td}><Badge estado={r.estado} /></td>
                  <td style={{ ...S.td, color: "#94a3b8" }}>{r.cobrador || "—"}</td>
                  <td style={S.td}><div style={{ display: "flex", gap: 5 }}>
                    {r.estado === "Programada" && <button onClick={async () => { await apiPost({ action: "updateRow", sheet: "Visitas", id: r.id, data: { ...r, estado: "Realizada" } }); load(); }} style={{ background: "#06522222", color: "#34d399", border: "1px solid #10b98133", borderRadius: 6, padding: "3px 8px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>✓ Realizada</button>}
                    <button style={S.btnGhost} onClick={() => { setForm({ ...r }); setModal(r); }}>✏️</button>
                    {isAdmin && <button style={S.btnDanger} onClick={async () => { if (confirm("¿Eliminar?")) { await apiPost({ action: "deleteRow", sheet: "Visitas", id: r.id }); load(); } }}>🗑</button>}
                  </div></td>
                </tr>)}
            </tbody>
          </table>
        </div>}
    </div>
    {modal && <Modal title={modal === "add" ? "Nueva visita" : "Editar visita"} onClose={() => setModal(null)}>
      <Field label="Cliente *">
        {clientes.length > 0
          ? <select style={S.input} value={form.cliente || ""} onChange={e => { const c = clientes.find(x => x.nombre === e.target.value); setForm(p => ({ ...p, cliente: e.target.value, codigo: c?.codigo || "" })); }}>
            <option value="">— Seleccionar —</option>
            {clientes.map(c => <option key={c.id} value={c.nombre}>{c.nombre} ({c.localidad})</option>)}
          </select>
          : <input style={S.input} value={form.cliente || ""} onChange={e => setForm(p => ({ ...p, cliente: e.target.value }))} />}
      </Field>
      <Field label="Dirección"><input style={S.input} value={form.direccion || ""} onChange={e => setForm(p => ({ ...p, direccion: e.target.value }))} placeholder="Calle 123" /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Fecha"><input style={S.input} type="date" value={form.fecha || ""} onChange={e => setForm(p => ({ ...p, fecha: e.target.value }))} /></Field>
        <Field label="Hora"><input style={S.input} type="time" value={form.hora || ""} onChange={e => setForm(p => ({ ...p, hora: e.target.value }))} /></Field>
      </div>
      <Field label="Estado"><select style={S.input} value={form.estado || "Programada"} onChange={e => setForm(p => ({ ...p, estado: e.target.value }))}>{VISITA_ESTADOS.map(e => <option key={e}>{e}</option>)}</select></Field>
      <Field label="Cobrador"><input style={S.input} value={form.cobrador || ""} onChange={e => setForm(p => ({ ...p, cobrador: e.target.value }))} placeholder="Juan / Anto / Sabri" /></Field>
      <Field label="Notas"><textarea style={{ ...S.input, height: 60, resize: "none" }} value={form.notas || ""} onChange={e => setForm(p => ({ ...p, notas: e.target.value }))} /></Field>
      <div style={{ display: "flex", gap: 10 }}>
        <button style={{ ...S.btnGhost, flex: 1 }} onClick={() => setModal(null)}>Cancelar</button>
        <button style={{ ...S.btnPri, flex: 1 }} disabled={saving || !form.cliente} onClick={save}>{saving ? "Guardando…" : "Guardar"}</button>
      </div>
    </Modal>}
  </div>;
}

/* ─── RESUMEN TAB ─────────────────────────────────────── */
function ResumenTab() {
  const [cobranzas, setCobranzas] = useState([]);
  const [visitas, setVisitas] = useState([]);
  const [deudores, setDeudores] = useState([]);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [rc, rv, rd] = await Promise.all([
        apiGet({ action: "getData", sheet: "Cobranzas" }),
        apiGet({ action: "getData", sheet: "Visitas" }),
        apiGet({ action: "getData", sheet: "Deudores" }),
      ]);
      setCobranzas(Array.isArray(rc.data) ? rc.data : []);
      setVisitas(Array.isArray(rv.data) ? rv.data : []);
      setDeudores(Array.isArray(rd.data) ? rd.data : []);
    } catch (e) { console.error(e); } finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);

  if (loading) return <div style={{ padding: 40, textAlign: "center", color: "#475569" }}>Cargando resumen…</div>;

  const cobrado   = cobranzas.filter(r => r.estado === "Pagado").reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const pendiente = cobranzas.filter(r => ["Pendiente","Parcial"].includes(r.estado)).reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const vencido   = cobranzas.filter(r => r.estado === "Vencido").reduce((a, r) => a + (Number(r.monto) || 0), 0);
  const totalDeuda= deudores.reduce((a, r) => a + (Number(r.saldo) || 0), 0);
  const cobradores= [...new Set(cobranzas.map(r => r.cobrador).filter(Boolean))];

  return <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
    <div style={{ ...S.card, padding: 16, background: "#0d1117", border: "1px solid #eab30833", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
      <div>
        <div style={{ fontSize: 11, color: "#eab308", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 2 }}>📋 Resumen semanal</div>
        <div style={{ fontSize: 17, color: "#fff", fontWeight: 800 }}>{new Date().toLocaleDateString("es-AR", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}</div>
      </div>
      <button style={S.btnGreen} onClick={() => exportResumenExcel(cobranzas, visitas, deudores)}>
        📥 Exportar a Excel
      </button>
    </div>

    <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
      <StatCard label="Cobrado" value={fmt(cobrado)} sub={`${cobranzas.filter(r => r.estado === "Pagado").length} pagos`} color="emerald" />
      <StatCard label="Pendiente" value={fmt(pendiente)} color="amber" />
      <StatCard label="Vencido" value={fmt(vencido)} color="red" />
      <StatCard label="Deuda total" value={fmt(totalDeuda)} sub={`${deudores.length} deudores`} color="purple" />
    </div>
    <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
      <StatCard label="Visitas realizadas" value={visitas.filter(r => r.estado === "Realizada").length} color="sky" />
      <StatCard label="Visitas pendientes" value={visitas.filter(r => r.estado === "Programada").length} color="amber" />
      <StatCard label="Total cobranzas" value={cobranzas.length} color="sky" />
    </div>

    {cobradores.length > 0 && <div style={S.card}>
      <div style={{ padding: "12px 16px", borderBottom: "1px solid rgba(255,255,255,.07)", fontWeight: 700, color: "#fff", fontSize: 14 }}>👥 Por cobrador</div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr style={{ borderBottom: "1px solid rgba(255,255,255,.07)" }}>
          {["Cobrador","Cobrado","Pendiente","Visitas"].map(h => <th key={h} style={S.th}>{h}</th>)}
        </tr></thead>
        <tbody>
          {cobradores.map(nombre => <tr key={nombre}>
            <td style={{ ...S.td, color: "#fff", fontWeight: 600 }}>{nombre}</td>
            <td style={{ ...S.td, color: "#34d399", fontFamily: "monospace", fontWeight: 700 }}>{fmt(cobranzas.filter(r => r.cobrador === nombre && r.estado === "Pagado").reduce((a, r) => a + (Number(r.monto) || 0), 0))}</td>
            <td style={{ ...S.td, color: "#fbbf24", fontFamily: "monospace" }}>{fmt(cobranzas.filter(r => r.cobrador === nombre && ["Pendiente","Parcial"].includes(r.estado)).reduce((a, r) => a + (Number(r.monto) || 0), 0))}</td>
            <td style={{ ...S.td, color: "#38bdf8" }}>{visitas.filter(r => r.cobrador === nombre && r.estado === "Realizada").length}</td>
          </tr>)}
        </tbody>
      </table>
    </div>}

    {deudores.length > 0 && <div style={S.card}>
      <div style={{ padding: "12px 16px", borderBottom: "1px solid rgba(255,255,255,.07)", fontWeight: 700, color: "#fff", fontSize: 14 }}>🔴 Mayores deudores</div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr style={{ borderBottom: "1px solid rgba(255,255,255,.07)" }}>
          {["Código","Nombre","Localidad","Saldo"].map(h => <th key={h} style={S.th}>{h}</th>)}
        </tr></thead>
        <tbody>
          {[...deudores].sort((a, b) => (Number(b.saldo) || 0) - (Number(a.saldo) || 0)).slice(0, 15).map(r => <tr key={r.id}>
            <td style={{ ...S.td, color: "#fbbf24", fontFamily: "monospace", fontSize: 11 }}>{r.codigo || "—"}</td>
            <td style={{ ...S.td, color: "#fff", fontWeight: 600 }}>{r.nombre}</td>
            <td style={S.td}>{r.localidad || "—"}</td>
            <td style={{ ...S.td, color: "#f87171", fontFamily: "monospace", fontWeight: 700 }}>{fmt(r.saldo)}</td>
          </tr>)}
        </tbody>
      </table>
    </div>}
  </div>;
}

/* ─── MAIN APP ────────────────────────────────────────── */
export default function Home() {
  const [isAdmin, setIsAdmin] = useState(false);
  const [showLogin, setShowLogin] = useState(false);
  const [tab, setTab] = useState("cobranzas");
  const [clientes, setClientes] = useState([]);

  useEffect(() => {
    apiGet({ action: "getData", sheet: "Clientes" }).then(r => setClientes(Array.isArray(r.data) ? r.data : [])).catch(() => {});
  }, []);

  const TABS_CHICOS = [{ id: "cobranzas", icon: "💰", label: "Cobranzas" }, { id: "visitas", icon: "📍", label: "Visitas" }];
  const TABS_ADMIN  = [...TABS_CHICOS, { id: "importar", icon: "📂", label: "Importar" }, { id: "clientes", icon: "👥", label: "Clientes" }, { id: "deudores", icon: "🔴", label: "Deudores" }, { id: "resumen", icon: "📋", label: "Resumen" }];
  const tabs = isAdmin ? TABS_ADMIN : TABS_CHICOS;

  if (showLogin) return <Login onLogin={() => { setIsAdmin(true); setShowLogin(false); setTab("cobranzas"); }} />;

  return <>
    <Head>
      <title>GestiónApp — Cobranzas & Visitas</title>
      <meta name="viewport" content="width=device-width, initial-scale=1" />
    </Head>
    <div style={S.page}>
      {/* Header */}
      <div style={{ background: "rgba(13,17,23,.95)", borderBottom: "1px solid rgba(255,255,255,.07)", padding: "0 16px", position: "sticky", top: 0, zIndex: 40 }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 54, flexWrap: "wrap", gap: 6 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 30, height: 30, background: "#eab30822", border: "1px solid #eab30840", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>💼</div>
            <div>
              <div style={{ fontWeight: 800, color: "#fff", fontSize: 13, lineHeight: 1 }}>GestiónApp</div>
              <div style={{ fontSize: 10, color: "#475569", marginTop: 2 }}>{isAdmin ? "👑 Administrador" : "Cobranzas & Visitas"}</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 3, flexWrap: "wrap" }}>
            {tabs.map(t => <button key={t.id} onClick={() => setTab(t.id)} style={{ display: "flex", alignItems: "center", gap: 4, padding: "5px 10px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer", border: tab === t.id ? "1px solid #eab30840" : "1px solid transparent", background: tab === t.id ? "#eab30815" : "transparent", color: tab === t.id ? "#fbbf24" : "#64748b" }}>
              <span>{t.icon}</span><span>{t.label}</span>
            </button>)}
            {isAdmin
              ? <button onClick={() => { setIsAdmin(false); setTab("cobranzas"); }} style={{ ...S.btnGhost, fontSize: 11, padding: "5px 10px", marginLeft: 6 }}>Salir</button>
              : <button onClick={() => setShowLogin(true)} style={{ ...S.btnGhost, fontSize: 11, padding: "5px 10px", marginLeft: 6 }}>🔐 Admin</button>}
          </div>
        </div>
      </div>

      {/* Content */}
      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "20px 16px" }}>
        <div style={{ marginBottom: 18 }}>
          <h2 style={{ fontSize: 20, fontWeight: 800, color: "#fff", margin: 0, marginBottom: 3 }}>
            {tabs.find(t => t.id === tab)?.icon} {tabs.find(t => t.id === tab)?.label}
          </h2>
          <p style={{ fontSize: 12, color: "#475569", margin: 0 }}>
            {tab === "cobranzas" && "Gestión de pagos y deudas"}
            {tab === "visitas"   && "Agenda y seguimiento de visitas"}
            {tab === "importar"  && "Importar Excel desde tu sistema de gestión"}
            {tab === "clientes"  && "Lista de clientes activos"}
            {tab === "deudores"  && "Registro de deudores actualizado"}
            {tab === "resumen"   && "Resumen semanal — exportable a Excel"}
          </p>
        </div>
        {tab === "cobranzas" && <CobranzasTab isAdmin={isAdmin} clientes={clientes} />}
        {tab === "visitas"   && <VisitasTab isAdmin={isAdmin} clientes={clientes} />}
        {tab === "importar"  && isAdmin && <ImportarTab />}
        {tab === "clientes"  && isAdmin && <ClientesTab />}
        {tab === "deudores"  && isAdmin && <DeudoresTab />}
        {tab === "resumen"   && isAdmin && <ResumenTab />}
      </div>

      <div style={{ borderTop: "1px solid rgba(255,255,255,.05)", padding: "10px 16px", textAlign: "center" }}>
        <p style={{ fontSize: 11, color: "#1f2937", margin: 0 }}>GestiónApp · Sincronizado con Google Sheets</p>
      </div>
    </div>
  </>;
}

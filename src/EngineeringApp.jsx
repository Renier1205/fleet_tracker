import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Truck, Clock, AlertTriangle, ClipboardList,
  CalendarClock, ShieldCheck, CircleDot, Package, FileText,
  DollarSign, History, Factory, Download, Search, Menu, X
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from "recharts";

const NAVY = "#1F3864";
const GREEN = "#5FBF8F";
const MOBILE_BREAKPOINT = 720;

function useIsMobile() {
  const [isMobile, setIsMobile] = useState(
    typeof window !== "undefined" ? window.innerWidth < MOBILE_BREAKPOINT : false
  );
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < MOBILE_BREAKPOINT);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  return isMobile;
}

// ---------------------------------------------------------------
// Mock data — mirrors the Supabase schema/views. Swap for real
// Supabase queries once connected (same shape, same field names).
// ---------------------------------------------------------------
const assets = [
  { asset_id: "EQ-001", asset_name: "Haul Truck 01", make: "Caterpillar", model: "793F", fleet: "CAT 793F", serial_number: "CAT793F-88213", status: "Operating", current_hours: 1622.4 },
  { asset_id: "EQ-002", asset_name: "Excavator 01", make: "Caterpillar", model: "789D", fleet: "CAT 789D", serial_number: "CAT789D-44120", status: "Operating", current_hours: 986.1 },
  { asset_id: "EQ-003", asset_name: "Dozer 01", make: "Caterpillar", model: "D11", fleet: "CAT D11", serial_number: "CATD11-90341", status: "Under Maintenance", current_hours: 2140.8 },
  { asset_id: "EQ-004", asset_name: "Grader 01", make: "Caterpillar", model: "16M", fleet: "CAT 16M Grader", serial_number: "CAT16M-55210", status: "Operating", current_hours: 743.2 },
];

const dailyHours = [
  { log_date: "2026-08-05", asset_id: "EQ-001", opening_hours: 1598.2, closing_hours: 1610.4, hours_run: 12.2 },
  { log_date: "2026-08-06", asset_id: "EQ-001", opening_hours: 1610.4, closing_hours: 1622.4, hours_run: 12.0 },
  { log_date: "2026-08-06", asset_id: "EQ-002", opening_hours: 974.0, closing_hours: 986.1, hours_run: 12.1 },
];

const breakdowns = [
  { breakdown_date: "2026-08-03", asset_id: "EQ-001", wo_reference: "WO-00014", component_affected: "Starter Motor", cause_code: "Electrical Fault", severity: "Medium", repair_status: "Closed", downtime_hours: 2.3 },
  { breakdown_date: "2026-07-22", asset_id: "EQ-001", wo_reference: "WO-00009", component_affected: "Starter Motor", cause_code: "Electrical Fault", severity: "High", repair_status: "Closed", downtime_hours: 3.0 },
  { breakdown_date: "2026-08-06", asset_id: "EQ-003", wo_reference: "WO-00016", component_affected: "Hydraulic Pump", cause_code: "Hydraulic Failure", severity: "Critical", repair_status: "Open", downtime_hours: null },
];

const workOrders = [
  { wo_no: "WO-00016", asset_id: "EQ-003", work_type: "Corrective", priority: "Critical", status: "Open", request_date: "2026-08-06" },
  { wo_no: "WO-00015", asset_id: "EQ-004", work_type: "Preventive", priority: "Medium", status: "Planned", request_date: "2026-08-08" },
  { wo_no: "WO-00014", asset_id: "EQ-001", work_type: "Corrective", priority: "Medium", status: "Closed", request_date: "2026-08-03" },
];

const plannedMaintenance = [
  { asset_id: "EQ-003", asset_name: "Dozer 01", current_hours: 2140.8, next_service_due: 2150, remaining: 9.2, status: "DUE SOON" },
  { asset_id: "EQ-001", asset_name: "Haul Truck 01", current_hours: 1622.4, next_service_due: 1750, remaining: 127.6, status: "OK" },
  { asset_id: "EQ-002", asset_name: "Excavator 01", current_hours: 986.1, next_service_due: 950, remaining: -36.1, status: "OVERDUE" },
  { asset_id: "EQ-004", asset_name: "Grader 01", current_hours: 743.2, next_service_due: 750, remaining: 6.8, status: "DUE SOON" },
];

const inspections = [
  { log_date: "2026-08-06", asset_id: "EQ-003", inspection_type: "Pre-start", inspector: "J. Smith", result: "Fail", wo_reference: "WO-00016" },
  { log_date: "2026-08-06", asset_id: "EQ-001", inspection_type: "Pre-start", inspector: "T. Mokoena", result: "Pass", wo_reference: null },
];

const components = [
  { component_id: "CMP-001", asset_id: "EQ-001", component_type: "Engine", life_used_pct: 0.91, status: "PLAN CHANGE" },
  { component_id: "CMP-002", asset_id: "EQ-003", component_type: "Hydraulic Pump", life_used_pct: 1.0, status: "CHANGE OUT" },
];

const tyres = [
  { asset_id: "EQ-001", position: "Front Left", remaining_life: 78, replacement_due: "DUE SOON" },
  { asset_id: "EQ-002", position: "Rear Right", remaining_life: 1240, replacement_due: "OK" },
];

const parts = [
  { part_no: "HYD-4521", description: "Hydraulic Filter Kit", qty_in_stock: 2, minimum_qty: 5, reorder_status: "REORDER" },
  { part_no: "OIL-1180", description: "Engine Oil 15W-40 (20L)", qty_in_stock: 14, minimum_qty: 6, reorder_status: "OK" },
];

const warrantyDocs = [
  { asset_id: "EQ-001", type: "Warranty — Engine", reference: "ENG-2019-4471", expiry: "2026-08-15", status: "EXPIRING SOON" },
  { asset_id: "EQ-001", type: "Document — Inspection Cert", reference: "COC-2026-0142", expiry: "2026-08-20", status: "EXPIRING SOON" },
];

const costLedger = [
  { cost_date: "2026-08-06", asset_id: "EQ-003", cost_type: "Parts", description: "Hydraulic pump rebuild kit", total_cost: 8400 },
  { cost_date: "2026-08-03", asset_id: "EQ-001", cost_type: "Labour", description: "Starter motor replacement", total_cost: 1850 },
  { cost_date: "2026-08-01", asset_id: "EQ-002", cost_type: "Fuel", description: "Monthly fuel fills", total_cost: 24600 },
];

const auditLog = [
  { changed_at: "2026-08-06 14:32", changed_by_email: "j.smith@company.com", action: "INSERT", table_name: "breakdown_log", record_id: "142" },
  { changed_at: "2026-08-06 09:10", changed_by_email: "t.mokoena@company.com", action: "UPDATE", table_name: "work_orders", record_id: "16" },
  { changed_at: "2026-08-05 17:02", changed_by_email: "j.smith@company.com", action: "INSERT", table_name: "daily_hours", record_id: "981" },
];

const costByType = [
  { type: "Parts", cost: 112000 },
  { type: "Labour", cost: 78000 },
  { type: "Fuel", cost: 56000 },
  { type: "External", cost: 24000 },
  { type: "Tyres", cost: 10000 },
  { type: "Oil", cost: 4000 },
];

// ---------------------------------------------------------------
const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "assets", label: "Assets", icon: Truck },
  { key: "daily_hours", label: "Daily Hours", icon: Clock },
  { key: "breakdowns", label: "Breakdowns", icon: AlertTriangle },
  { key: "work_orders", label: "Work Orders", icon: ClipboardList },
  { key: "planned_maintenance", label: "Planned Maintenance", icon: CalendarClock },
  { key: "inspections", label: "Inspections", icon: ShieldCheck },
  { key: "components", label: "Components", icon: CircleDot },
  { key: "tyres", label: "Tyres", icon: CircleDot },
  { key: "parts", label: "Parts Inventory", icon: Package },
  { key: "warranty_docs", label: "Warranty & Documents", icon: FileText },
  { key: "cost_ledger", label: "Cost Ledger", icon: DollarSign },
  { key: "audit", label: "Audit Trail", icon: History },
];

function statusColor(status) {
  const s = (status || "").toUpperCase();
  if (["OVERDUE", "CRITICAL", "CHANGE OUT", "REORDER", "REPLACE NOW", "EXPIRED", "FAIL", "OPEN"].includes(s))
    return { bg: "#FCEBEB", text: "#791F1F" };
  if (["DUE SOON", "PLAN CHANGE", "EXPIRING SOON", "MEDIUM", "PLANNED"].includes(s))
    return { bg: "#FAEEDA", text: "#633806" };
  if (["OK", "CLOSED", "PASS", "ACTIVE", "VALID"].includes(s))
    return { bg: "#EAF3DE", text: "#27500A" };
  return { bg: "#F1EFE8", text: "#444441" };
}

function Badge({ value }) {
  if (value === null || value === undefined || value === "") return <span style={{ color: "#B4B2A9" }}>—</span>;
  const c = statusColor(value);
  return (
    <span style={{ background: c.bg, color: c.text, fontSize: 12, fontWeight: 600, padding: "2px 9px", borderRadius: 6, whiteSpace: "nowrap" }}>
      {value}
    </span>
  );
}

const STATUS_FIELDS = new Set(["status", "repair_status", "result", "replacement_due", "reorder_status", "severity", "priority", "action"]);

function DataTable({ columns, rows, exportName }) {
  const [query, setQuery] = useState("");
  const isMobile = useIsMobile();

  const filtered = useMemo(() => {
    if (!query.trim()) return rows;
    const q = query.toLowerCase();
    return rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
  }, [rows, query]);

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(filtered);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, exportName.slice(0, 31));
    XLSX.writeFile(wb, `${exportName.replace(/\s+/g, "_")}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", justifyContent: "space-between", alignItems: isMobile ? "stretch" : "center", marginBottom: 12, gap: 10 }}>
        <div style={{ position: "relative", flex: 1, maxWidth: isMobile ? "none" : 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 13, color: "#898781" }} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search this table"
            style={{ width: "100%", padding: "10px 10px 10px 32px", fontSize: 14, border: "1px solid #D3D1C7", borderRadius: 8, outline: "none", boxSizing: "border-box", minHeight: 42 }}
          />
        </div>
        <button
          onClick={exportToExcel}
          style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 14, fontWeight: 600, padding: "10px 14px", borderRadius: 8, cursor: "pointer", minHeight: 42, width: isMobile ? "100%" : "auto" }}
        >
          <Download size={14} /> Export to Excel
        </button>
      </div>
      <div style={{ overflowX: "auto", border: "1px solid #E4E2D8", borderRadius: 10, WebkitOverflowScrolling: "touch" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: isMobile ? 12.5 : 13 }}>
          <thead>
            <tr style={{ background: "#F7F6F1" }}>
              {columns.map((c) => (
                <th key={c.key} style={{ textAlign: "left", padding: isMobile ? "9px 10px" : "9px 12px", fontWeight: 600, color: "#5F5E5A", whiteSpace: "nowrap", borderBottom: "1px solid #E4E2D8" }}>
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr key={i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                {columns.map((c) => (
                  <td key={c.key} style={{ padding: isMobile ? "9px 10px" : "9px 12px", color: "#2C2C2A", whiteSpace: "nowrap" }}>
                    {STATUS_FIELDS.has(c.key) ? <Badge value={row[c.key]} /> : (row[c.key] ?? <span style={{ color: "#B4B2A9" }}>—</span>)}
                  </td>
                ))}
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length} style={{ padding: 20, textAlign: "center", color: "#898781" }}>No rows match your search.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      {isMobile && <p style={{ fontSize: 11.5, color: "#898781", margin: "8px 2px 0" }}>Swipe the table sideways to see more columns.</p>}
    </div>
  );
}

function MetricCard({ label, value }) {
  return (
    <div style={{ background: "#F7F6F1", borderRadius: 10, padding: "14px 16px" }}>
      <p style={{ fontSize: 12, color: "#5F5E5A", margin: "0 0 4px" }}>{label}</p>
      <p style={{ fontSize: 22, fontWeight: 700, margin: 0, color: "#2C2C2A" }}>{value}</p>
    </div>
  );
}

function Dashboard() {
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12, marginBottom: 24 }}>
        <MetricCard label="Total assets" value={assets.length} />
        <MetricCard label="Open work orders" value={workOrders.filter((w) => w.status !== "Closed").length} />
        <MetricCard label="Services due soon / overdue" value={plannedMaintenance.filter((p) => p.status !== "OK").length} />
        <MetricCard label="Cost this month" value="R284k" />
      </div>

      <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Alerts</h3>
      <div style={{ border: "1px solid #E4E2D8", borderRadius: 10, overflow: "hidden", marginBottom: 24 }}>
        {[
          ["Services overdue / due soon", plannedMaintenance.filter((p) => p.status !== "OK").length, "warn"],
          ["Repeat failures (2+ times)", 1, "danger"],
          ["Components requiring action", components.filter((c) => c.status !== "OK").length, "warn"],
          ["Parts below reorder point", parts.filter((p) => p.reorder_status === "REORDER").length, "warn"],
          ["Failed inspections", inspections.filter((i) => i.result === "Fail").length, "danger"],
        ].map(([label, count], i) => (
          <div key={label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none" }}>
            <span style={{ fontSize: 13.5, color: "#2C2C2A" }}>{label}</span>
            <Badge value={count > 0 ? String(count) : "0"} />
          </div>
        ))}
      </div>

      <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Cost by type, this month</h3>
      <div style={{ height: 220, marginBottom: 8 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={costByType}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EFEEE7" vertical={false} />
            <XAxis dataKey="type" tick={{ fontSize: 12, fill: "#5F5E5A" }} axisLine={{ stroke: "#E4E2D8" }} tickLine={false} />
            <YAxis tick={{ fontSize: 12, fill: "#5F5E5A" }} axisLine={false} tickLine={false} tickFormatter={(v) => `R${v / 1000}k`} />
            <Tooltip formatter={(v) => `R${v.toLocaleString()}`} />
            <Bar dataKey="cost" fill={NAVY} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

const TABLE_CONFIG = {
  assets: { title: "Assets", cols: [["asset_id","Equipment #"],["asset_name","Name"],["make","Make"],["model","Model"],["fleet","Fleet"],["serial_number","Serial number"],["status","Status"],["current_hours","Current hours"]], data: assets },
  daily_hours: { title: "Daily Hours", cols: [["log_date","Date"],["asset_id","Equipment #"],["opening_hours","Opening hours"],["closing_hours","Closing hours"],["hours_run","Hours run"]], data: dailyHours },
  breakdowns: { title: "Breakdowns", cols: [["breakdown_date","Date"],["asset_id","Equipment #"],["wo_reference","Work order #"],["component_affected","Component"],["cause_code","Cause"],["severity","Severity"],["repair_status","Status"],["downtime_hours","Downtime (hrs)"]], data: breakdowns },
  work_orders: { title: "Work Orders", cols: [["wo_no","Work order #"],["asset_id","Equipment #"],["work_type","Type"],["priority","Priority"],["status","Status"],["request_date","Requested"]], data: workOrders },
  planned_maintenance: { title: "Planned Maintenance", cols: [["asset_id","Equipment #"],["asset_name","Name"],["current_hours","Current hours"],["next_service_due","Next service due"],["remaining","Hours remaining"],["status","Status"]], data: plannedMaintenance },
  inspections: { title: "Inspections", cols: [["log_date","Date"],["asset_id","Equipment #"],["inspection_type","Type"],["inspector","Inspector"],["result","Result"],["wo_reference","Work order #"]], data: inspections },
  components: { title: "Components", cols: [["component_id","Component #"],["asset_id","Equipment #"],["component_type","Type"],["life_used_pct","Life used"],["status","Status"]], data: components.map(c => ({...c, life_used_pct: `${Math.round(c.life_used_pct*100)}%`})) },
  tyres: { title: "Tyres", cols: [["asset_id","Equipment #"],["position","Position"],["remaining_life","Remaining life (hrs)"],["replacement_due","Status"]], data: tyres },
  parts: { title: "Parts Inventory", cols: [["part_no","Part #"],["description","Description"],["qty_in_stock","Qty in stock"],["minimum_qty","Minimum qty"],["reorder_status","Status"]], data: parts },
  warranty_docs: { title: "Warranty & Documents", cols: [["asset_id","Equipment #"],["type","Type"],["reference","Reference"],["expiry","Expiry"],["status","Status"]], data: warrantyDocs },
  cost_ledger: { title: "Cost Ledger", cols: [["cost_date","Date"],["asset_id","Equipment #"],["cost_type","Cost type"],["description","Description"],["total_cost","Total cost"]], data: costLedger.map(c => ({...c, total_cost: `R${c.total_cost.toLocaleString()}`})) },
  audit: { title: "Audit Trail", cols: [["changed_at","Timestamp"],["changed_by_email","User"],["action","Action"],["table_name","Table"],["record_id","Record #"]], data: auditLog },
};

export default function App() {
  const [active, setActive] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const isMobile = useIsMobile();

  // On mobile the sidebar starts closed (it's an overlay drawer, not a
  // permanent column); on desktop it starts open as a fixed column.
  useEffect(() => {
    setSidebarOpen(!isMobile);
  }, [isMobile]);

  const activeConfig = TABLE_CONFIG[active];

  const handleNavClick = (key) => {
    setActive(key);
    if (isMobile) setSidebarOpen(false);
  };

  const sidebarWidth = isMobile ? 240 : (sidebarOpen ? 220 : 56);

  return (
    <div style={{ position: "relative", display: "flex", minHeight: 600, fontFamily: "Arial, sans-serif", background: "#fff", overflow: "hidden" }}>
      {/* Backdrop, mobile only, shown when drawer is open */}
      {isMobile && sidebarOpen && (
        <div
          onClick={() => setSidebarOpen(false)}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 20 }}
        />
      )}

      <div
        style={{
          width: sidebarWidth,
          background: NAVY,
          transition: "width 0.15s, transform 0.2s",
          display: "flex",
          flexDirection: "column",
          flexShrink: 0,
          ...(isMobile
            ? { position: "fixed", top: 0, bottom: 0, left: 0, zIndex: 30, transform: sidebarOpen ? "translateX(0)" : "translateX(-100%)" }
            : {}),
        }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: sidebarOpen ? "space-between" : "center", padding: "16px 14px", borderBottom: "1px solid rgba(255,255,255,0.12)" }}>
          {sidebarOpen && <span style={{ color: GREEN, fontWeight: 700, fontSize: 15 }}>Fleet Tracker</span>}
          {!isMobile && (
            <button onClick={() => setSidebarOpen(!sidebarOpen)} style={{ background: "none", border: "none", color: "#fff", cursor: "pointer", padding: 4 }}>
              {sidebarOpen ? <X size={16} /> : <Menu size={16} />}
            </button>
          )}
          {isMobile && (
            <button onClick={() => setSidebarOpen(false)} aria-label="Close menu" style={{ background: "none", border: "none", color: "#fff", cursor: "pointer", padding: 4 }}>
              <X size={18} />
            </button>
          )}
        </div>

        <div style={{ overflowY: "auto", flex: 1 }}>
          {sidebarOpen && <p style={{ color: "rgba(255,255,255,0.5)", fontSize: 11, fontWeight: 700, letterSpacing: 0.5, margin: "14px 16px 6px", textTransform: "uppercase" }}>Engineering</p>}
          {NAV.map(({ key, label, icon: Icon }) => (
            <button
              key={key}
              onClick={() => handleNavClick(key)}
              title={label}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: isMobile ? "12px 16px" : "9px 16px", margin: "1px 8px", borderRadius: 8,
                background: active === key ? "rgba(255,255,255,0.12)" : "transparent",
                border: "none", color: active === key ? "#fff" : "rgba(255,255,255,0.72)",
                cursor: "pointer", fontSize: 14, textAlign: "left", justifyContent: sidebarOpen ? "flex-start" : "center",
                width: sidebarOpen ? "calc(100% - 16px)" : "auto", minHeight: 40,
              }}
            >
              <Icon size={16} style={{ flexShrink: 0 }} />
              {sidebarOpen && <span>{label}</span>}
            </button>
          ))}

          {sidebarOpen && <p style={{ color: "rgba(255,255,255,0.5)", fontSize: 11, fontWeight: 700, letterSpacing: 0.5, margin: "18px 16px 6px", textTransform: "uppercase" }}>Production</p>}
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: isMobile ? "12px 16px" : "9px 16px", margin: "1px 8px", borderRadius: 8, color: "rgba(255,255,255,0.32)", fontSize: 14, cursor: "not-allowed", minHeight: 40 }}>
            <Factory size={16} style={{ flexShrink: 0 }} />
            {sidebarOpen && <span>Coming soon</span>}
          </div>
        </div>
      </div>

      <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        {/* Mobile top bar */}
        {isMobile && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px", borderBottom: "1px solid #E4E2D8", position: "sticky", top: 0, background: "#fff", zIndex: 10 }}>
            <button onClick={() => setSidebarOpen(true)} aria-label="Open menu" style={{ background: "none", border: "none", color: NAVY, cursor: "pointer", padding: 6, display: "flex" }}>
              <Menu size={20} />
            </button>
            <span style={{ fontSize: 16, fontWeight: 700, color: NAVY }}>
              {active === "dashboard" ? "Dashboard" : activeConfig.title}
            </span>
          </div>
        )}

        <div style={{ padding: isMobile ? "16px" : "24px 28px", overflow: "auto", flex: 1 }}>
          {!isMobile && (
            <h2 style={{ fontSize: 19, fontWeight: 700, margin: "0 0 18px", color: NAVY }}>
              {active === "dashboard" ? "Dashboard" : activeConfig.title}
            </h2>
          )}
          {active === "dashboard" ? (
            <Dashboard />
          ) : (
            <DataTable
              columns={activeConfig.cols.map(([key, label]) => ({ key, label }))}
              rows={activeConfig.data}
              exportName={activeConfig.title}
            />
          )}
        </div>
      </div>
    </div>
  );
}

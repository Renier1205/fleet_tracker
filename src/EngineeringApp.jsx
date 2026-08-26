import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Truck, Clock, AlertTriangle, ClipboardList,
  CalendarClock, ShieldCheck, CircleDot, Package, FileText,
  DollarSign, History, Factory, Download, Search, Menu, X,
  Layers, ChevronRight, GitCompare, Fuel, Droplet, LogOut
} from "lucide-react";
import { supabase } from "./lib/supabaseClient";
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
const MOCK_ASSETS = [
  { asset_id: "EQ-001", asset_name: "Haul Truck 01", make: "Caterpillar", model: "793F", fleet: "CAT 793F", serial_number: "CAT793F-88213", status: "Operating", current_hours: 1622.4 },
  { asset_id: "EQ-005", asset_name: "Haul Truck 02", make: "Caterpillar", model: "793F", fleet: "CAT 793F", serial_number: "CAT793F-88214", status: "Operating", current_hours: 2044.7 },
  { asset_id: "EQ-006", asset_name: "Haul Truck 03", make: "Caterpillar", model: "793F", fleet: "CAT 793F", serial_number: "CAT793F-88215", status: "Breakdown", current_hours: 1387.9 },
  { asset_id: "EQ-002", asset_name: "Excavator 01", make: "Caterpillar", model: "789D", fleet: "CAT 789D", serial_number: "CAT789D-44120", status: "Operating", current_hours: 986.1 },
  { asset_id: "EQ-007", asset_name: "Excavator 02", make: "Caterpillar", model: "789D", fleet: "CAT 789D", serial_number: "CAT789D-44121", status: "Operating", current_hours: 1755.3 },
  { asset_id: "EQ-003", asset_name: "Dozer 01", make: "Caterpillar", model: "D11", fleet: "CAT D11", serial_number: "CATD11-90341", status: "Under Maintenance", current_hours: 2140.8 },
  { asset_id: "EQ-004", asset_name: "Grader 01", make: "Caterpillar", model: "16M", fleet: "CAT 16M Grader", serial_number: "CAT16M-55210", status: "Operating", current_hours: 743.2 },
];

// Per-equipment performance metrics — same fields the plant_performance_kpi()
// SQL function returns (hours, availability, utilisation, breakdowns, MTBF,
// MTTR). Deliberately no cost fields here.
const assetMetrics = {
  "EQ-001": { hours_worked: 622.4, availability: 0.942, utilisation: 0.781, breakdown_count: 2, mtbf: 311.2, mttr: 2.3 },
  "EQ-005": { hours_worked: 704.1, availability: 0.968, utilisation: 0.845, breakdown_count: 1, mtbf: 704.1, mttr: 1.8 },
  "EQ-006": { hours_worked: 388.0, availability: 0.734, utilisation: 0.602, breakdown_count: 5, mtbf: 77.6, mttr: 6.1 },
  "EQ-002": { hours_worked: 486.1, availability: 0.917, utilisation: 0.760, breakdown_count: 2, mtbf: 243.1, mttr: 2.7 },
  "EQ-007": { hours_worked: 655.3, availability: 0.961, utilisation: 0.832, breakdown_count: 1, mtbf: 655.3, mttr: 1.5 },
  "EQ-003": { hours_worked: 340.8, availability: 0.688, utilisation: 0.512, breakdown_count: 4, mtbf: 85.2, mttr: 5.4 },
  "EQ-004": { hours_worked: 543.2, availability: 0.895, utilisation: 0.702, breakdown_count: 2, mtbf: 271.6, mttr: 2.9 },
};

const METRIC_DEFS = [
  { key: "hours_worked", label: "Hours worked", fmt: (v) => v.toLocaleString(undefined, { maximumFractionDigits: 1 }) },
  { key: "availability", label: "Availability", fmt: (v) => `${Math.round(v * 100)}%` },
  { key: "utilisation", label: "Utilisation", fmt: (v) => `${Math.round(v * 100)}%` },
  { key: "breakdown_count", label: "Breakdowns", fmt: (v) => String(v) },
  { key: "mtbf", label: "MTBF (hrs)", fmt: (v) => v.toLocaleString(undefined, { maximumFractionDigits: 1 }) },
  { key: "mttr", label: "MTTR (hrs)", fmt: (v) => v.toLocaleString(undefined, { maximumFractionDigits: 1 }) },
];

const MOCK_DAILY_HOURS = [
  { log_date: "2026-08-05", asset_id: "EQ-001", opening_hours: 1598.2, closing_hours: 1610.4, hours_run: 12.2 },
  { log_date: "2026-08-06", asset_id: "EQ-001", opening_hours: 1610.4, closing_hours: 1622.4, hours_run: 12.0 },
  { log_date: "2026-08-06", asset_id: "EQ-002", opening_hours: 974.0, closing_hours: 986.1, hours_run: 12.1 },
];

const MOCK_FUEL_LOG = [
  { fill_date: "2026-08-10", asset_id: "EQ-001", hour_meter: 1610.4, litres: 620, cost: 12400, recorded_by: "J. Smith", consumption_rate: 5.1 },
  { fill_date: "2026-08-09", asset_id: "EQ-005", hour_meter: 2020.1, litres: 705, cost: 14100, recorded_by: "T. Mokoena", consumption_rate: 4.9 },
  { fill_date: "2026-08-08", asset_id: "EQ-006", hour_meter: 1370.5, litres: 588, cost: 11760, recorded_by: "J. Smith", consumption_rate: 5.6 },
  { fill_date: "2026-08-07", asset_id: "EQ-002", hour_meter: 970.2, litres: 340, cost: 6800, recorded_by: "T. Mokoena", consumption_rate: 3.8 },
  { fill_date: "2026-08-06", asset_id: "EQ-007", hour_meter: 1740.0, litres: 355, cost: 7100, recorded_by: "J. Smith", consumption_rate: 3.6 },
  { fill_date: "2026-08-05", asset_id: "EQ-004", hour_meter: 730.6, litres: 210, cost: 4200, recorded_by: "T. Mokoena", consumption_rate: 2.9 },
];

const MOCK_OIL_CONSUMPTION = [
  { fill_date: "2026-08-09", asset_id: "EQ-001", hour_meter: 1608.0, oil_type: "Engine Oil 15W-40", litres: 45, fill_reason: "Scheduled Change", recorded_by: "J. Smith", consumption_rate: 0.18 },
  { fill_date: "2026-08-08", asset_id: "EQ-003", hour_meter: 2135.2, oil_type: "Hydraulic Oil", litres: 60, fill_reason: "Top-Up", recorded_by: "T. Mokoena", consumption_rate: 0.32 },
  { fill_date: "2026-08-06", asset_id: "EQ-006", hour_meter: 1365.0, oil_type: "Engine Oil 15W-40", litres: 45, fill_reason: "Top-Up", recorded_by: "J. Smith", consumption_rate: 1.24 },
  { fill_date: "2026-08-04", asset_id: "EQ-005", hour_meter: 2010.5, oil_type: "Engine Oil 15W-40", litres: 45, fill_reason: "Scheduled Change", recorded_by: "T. Mokoena", consumption_rate: 0.15 },
];

const MOCK_BREAKDOWNS = [
  { breakdown_date: "2026-08-10", asset_id: "EQ-001", wo_reference: "WO-00021", component_affected: "Brake System", cause_code: "Wear & Tear", severity: "Medium", repair_status: "Open", downtime_hours: 1.8 },
  { breakdown_date: "2026-08-03", asset_id: "EQ-001", wo_reference: "WO-00014", component_affected: "Starter Motor", cause_code: "Electrical Fault", severity: "Medium", repair_status: "Closed", downtime_hours: 2.3 },
  { breakdown_date: "2026-07-22", asset_id: "EQ-001", wo_reference: "WO-00009", component_affected: "Starter Motor", cause_code: "Electrical Fault", severity: "High", repair_status: "Closed", downtime_hours: 3.0 },
  { breakdown_date: "2026-08-06", asset_id: "EQ-003", wo_reference: "WO-00016", component_affected: "Hydraulic Pump", cause_code: "Hydraulic Failure", severity: "Critical", repair_status: "Open", downtime_hours: null },
  { breakdown_date: "2026-06-18", asset_id: "EQ-003", wo_reference: "WO-00002", component_affected: "Track Roller", cause_code: "Wear & Tear", severity: "Low", repair_status: "Closed", downtime_hours: 1.2 },
  { breakdown_date: "2026-08-09", asset_id: "EQ-006", wo_reference: "WO-00019", component_affected: "Gearbox", cause_code: "Mechanical Failure", severity: "Critical", repair_status: "Open", downtime_hours: 8.5 },
  { breakdown_date: "2026-08-01", asset_id: "EQ-006", wo_reference: "WO-00011", component_affected: "Tyre/Track", cause_code: "Tyre/Track", severity: "Medium", repair_status: "Closed", downtime_hours: 1.5 },
  { breakdown_date: "2026-07-05", asset_id: "EQ-006", wo_reference: "WO-00004", component_affected: "Gearbox", cause_code: "Mechanical Failure", severity: "High", repair_status: "Closed", downtime_hours: 4.2 },
  { breakdown_date: "2026-07-30", asset_id: "EQ-002", wo_reference: "WO-00010", component_affected: "Bucket Pin", cause_code: "Wear & Tear", severity: "Low", repair_status: "Closed", downtime_hours: 0.9 },
  { breakdown_date: "2026-08-11", asset_id: "EQ-004", wo_reference: "WO-00022", component_affected: "Blade Cylinder", cause_code: "Hydraulic Failure", severity: "Medium", repair_status: "Open", downtime_hours: 2.0 },
];

// Reference "now" for time-frame filtering — matches the current mock data window
const NOW = new Date("2026-08-11T16:00:00");

function pad2(n) { return String(n).padStart(2, "0"); }

// Format a Date for an <input type="datetime-local"> value
function toDatetimeLocalValue(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatRangeForDisplay(fromStr, toStr) {
  const opts = { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" };
  const from = new Date(fromStr);
  const to = new Date(toStr);
  return `${from.toLocaleString("en-ZA", opts)} – ${to.toLocaleString("en-ZA", opts)}`;
}

// Default range on first load: start of this month through now
const DEFAULT_FROM = toDatetimeLocalValue(new Date(NOW.getFullYear(), NOW.getMonth(), 1));
const DEFAULT_TO = toDatetimeLocalValue(NOW);

function isInRange(dateStr, fromStr, toStr) {
  const d = new Date(dateStr + "T12:00:00");
  const from = new Date(fromStr);
  const to = new Date(toStr);
  return d >= from && d <= to;
}

function breakdownsInRange(breakdownsData, assetIds, fromStr, toStr) {
  const ids = Array.isArray(assetIds) ? assetIds : [assetIds];
  return breakdownsData.filter((b) => ids.includes(b.asset_id) && isInRange(b.breakdown_date, fromStr, toStr));
}

const MOCK_WORK_ORDERS = [
  { wo_no: "WO-00016", asset_id: "EQ-003", work_type: "Corrective", priority: "Critical", status: "Open", request_date: "2026-08-06" },
  { wo_no: "WO-00015", asset_id: "EQ-004", work_type: "Preventive", priority: "Medium", status: "Planned", request_date: "2026-08-08" },
  { wo_no: "WO-00014", asset_id: "EQ-001", work_type: "Corrective", priority: "Medium", status: "Closed", request_date: "2026-08-03" },
];

const MOCK_PLANNED_MAINTENANCE = [
  { asset_id: "EQ-003", asset_name: "Dozer 01", current_hours: 2140.8, next_service_due: 2150, remaining: 9.2, status: "DUE SOON" },
  { asset_id: "EQ-001", asset_name: "Haul Truck 01", current_hours: 1622.4, next_service_due: 1750, remaining: 127.6, status: "OK" },
  { asset_id: "EQ-002", asset_name: "Excavator 01", current_hours: 986.1, next_service_due: 950, remaining: -36.1, status: "OVERDUE" },
  { asset_id: "EQ-004", asset_name: "Grader 01", current_hours: 743.2, next_service_due: 750, remaining: 6.8, status: "DUE SOON" },
];

const MOCK_INSPECTIONS = [
  { log_date: "2026-08-06", asset_id: "EQ-003", inspection_type: "Pre-start", inspector: "J. Smith", result: "Fail", wo_reference: "WO-00016" },
  { log_date: "2026-08-06", asset_id: "EQ-001", inspection_type: "Pre-start", inspector: "T. Mokoena", result: "Pass", wo_reference: null },
];

const MOCK_COMPONENTS = [
  { component_id: "CMP-001", asset_id: "EQ-001", component_type: "Engine", life_used_pct: 0.91, status: "PLAN CHANGE" },
  { component_id: "CMP-002", asset_id: "EQ-003", component_type: "Hydraulic Pump", life_used_pct: 1.0, status: "CHANGE OUT" },
];

const MOCK_TYRES = [
  { asset_id: "EQ-001", position: "Front Left", remaining_life: 78, replacement_due: "DUE SOON" },
  { asset_id: "EQ-002", position: "Rear Right", remaining_life: 1240, replacement_due: "OK" },
];

const MOCK_PARTS = [
  { part_no: "HYD-4521", description: "Hydraulic Filter Kit", qty_in_stock: 2, minimum_qty: 5, reorder_status: "REORDER" },
  { part_no: "OIL-1180", description: "Engine Oil 15W-40 (20L)", qty_in_stock: 14, minimum_qty: 6, reorder_status: "OK" },
];

const MOCK_WARRANTY_DOCS = [
  { asset_id: "EQ-001", type: "Warranty — Engine", reference: "ENG-2019-4471", expiry: "2026-08-15", status: "EXPIRING SOON" },
  { asset_id: "EQ-001", type: "Document — Inspection Cert", reference: "COC-2026-0142", expiry: "2026-08-20", status: "EXPIRING SOON" },
];

const costLedger = [
  { cost_date: "2026-08-06", asset_id: "EQ-003", cost_type: "Parts", description: "Hydraulic pump rebuild kit", total_cost: 8400 },
  { cost_date: "2026-08-03", asset_id: "EQ-001", cost_type: "Labour", description: "Starter motor replacement", total_cost: 1850 },
  { cost_date: "2026-08-01", asset_id: "EQ-002", cost_type: "Fuel", description: "Monthly fuel fills", total_cost: 24600 },
];

const MOCK_AUDIT_LOG = [
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
  { key: "fleet_performance", label: "Fleet Performance", icon: Layers },
  { key: "assets", label: "Assets", icon: Truck },
  { key: "daily_hours", label: "Daily Hours", icon: Clock },
  { key: "fuel_log", label: "Fuel Log", icon: Fuel },
  { key: "oil_consumption", label: "Oil Consumption", icon: Droplet },
  { key: "breakdowns", label: "Breakdowns", icon: AlertTriangle },
  { key: "work_orders", label: "Work Orders", icon: ClipboardList },
  { key: "planned_maintenance", label: "Planned Maintenance", icon: CalendarClock },
  { key: "inspections", label: "Inspections", icon: ShieldCheck },
  { key: "components", label: "Components", icon: CircleDot },
  { key: "tyres", label: "Tyres", icon: CircleDot },
  { key: "parts", label: "Parts Inventory", icon: Package },
  { key: "warranty_docs", label: "Warranty & Documents", icon: FileText },
  // Cost Ledger hidden for now per request — data/config still here, just
  // not in the nav, so it's a one-line change to bring back.
  // { key: "cost_ledger", label: "Cost Ledger", icon: DollarSign },
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
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });

    const headerRow = columns.map((c) => c.label);
    const dataRows = filtered.map((row) => columns.map((c) => row[c.key] ?? ""));

    // Title + export timestamp on top, then a blank spacer row, then the
    // real header row using the same friendly labels shown on screen —
    // not the raw database column names.
    const aoa = [
      [exportName],
      [`Exported: ${timestamp}`],
      [],
      headerRow,
      ...dataRows,
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Merge the title and timestamp rows across the full table width
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(columns.length - 1, 0) } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(columns.length - 1, 0) } },
    ];

    // Reasonable column widths so it doesn't open looking cramped
    ws["!cols"] = columns.map((c) => ({ wch: Math.max(c.label.length + 2, 14) }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, exportName.slice(0, 31));
    const dateForFilename = now.toISOString().slice(0, 10);
    XLSX.writeFile(wb, `${exportName.replace(/\s+/g, "_")}_${dateForFilename}.xlsx`);
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

function KpiRow({ metrics }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 12, marginBottom: 20 }}>
      {METRIC_DEFS.map((m) => (
        <MetricCard key={m.key} label={m.label} value={metrics[m.key] != null ? m.fmt(metrics[m.key]) : "—"} />
      ))}
    </div>
  );
}

function aggregateMetrics(breakdownsData, assetIds, fromStr, toStr) {
  const ms = assetIds.map((id) => assetMetrics[id]).filter(Boolean);
  if (ms.length === 0) return {};
  const sum = (key) => ms.reduce((a, m) => a + m[key], 0);
  return {
    hours_worked: sum("hours_worked"),
    availability: sum("availability") / ms.length,
    utilisation: sum("utilisation") / ms.length,
    breakdown_count: breakdownsInRange(breakdownsData, assetIds, fromStr, toStr).length,
    mtbf: sum("mtbf") / ms.length,
    mttr: sum("mttr") / ms.length,
  };
}

function DateRangePicker({ fromDateTime, toDateTime, setFromDateTime, setToDateTime }) {
  const invalid = new Date(fromDateTime) > new Date(toDateTime);
  const inputStyle = {
    fontSize: 13, padding: "7px 10px", border: `1px solid ${invalid ? "#C0392B" : "#D3D1C7"}`,
    borderRadius: 8, fontFamily: "Arial, sans-serif", color: "#2C2C2A",
  };
  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <span style={{ fontSize: 12.5, color: "#5F5E5A", fontWeight: 600 }}>Show data from:</span>
        <input
          type="datetime-local"
          value={fromDateTime}
          onChange={(e) => setFromDateTime(e.target.value)}
          style={inputStyle}
        />
        <span style={{ fontSize: 12.5, color: "#5F5E5A", fontWeight: 600 }}>to:</span>
        <input
          type="datetime-local"
          value={toDateTime}
          onChange={(e) => setToDateTime(e.target.value)}
          style={inputStyle}
        />
      </div>
      {invalid && (
        <p style={{ fontSize: 12, color: "#C0392B", margin: "6px 0 0" }}>The "from" date/time must be before the "to" date/time.</p>
      )}
    </div>
  );
}

function FleetPerformance({ assets, breakdowns }) {
  const [fromDateTime, setFromDateTime] = useState(DEFAULT_FROM);
  const [toDateTime, setToDateTime] = useState(DEFAULT_TO);
  const [selectedFleet, setSelectedFleet] = useState(null);
  const [selectedEquipment, setSelectedEquipment] = useState(null);
  const [compareMode, setCompareMode] = useState(false);
  const [compareIds, setCompareIds] = useState([]);

  // Note: the date range is intentionally NOT reset by any of the
  // drill-down navigation below — it's set once and stays applied at
  // every level (fleet list, fleet detail, equipment detail, compare)
  // until the user changes it themselves.

  const fleets = useMemo(() => {
    const names = [...new Set(assets.map((a) => a.fleet))];
    return names.map((fleet) => {
      const fleetAssetIds = assets.filter((a) => a.fleet === fleet).map((a) => a.asset_id);
      return { fleet, count: fleetAssetIds.length, metrics: aggregateMetrics(breakdowns, fleetAssetIds, fromDateTime, toDateTime) };
    });
  }, [assets, breakdowns, fromDateTime, toDateTime]);

  const fleetAssets = selectedFleet ? assets.filter((a) => a.fleet === selectedFleet) : [];

  const resetToFleetList = () => { setSelectedFleet(null); setSelectedEquipment(null); setCompareMode(false); setCompareIds([]); };
  const backToFleet = () => { setSelectedEquipment(null); setCompareMode(false); setCompareIds([]); };

  const toggleCompare = (id) => {
    setCompareIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  };

  const crumbs = (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "#5F5E5A", marginBottom: 16, flexWrap: "wrap" }}>
      <button onClick={resetToFleetList} style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: selectedFleet ? NAVY : "#5F5E5A", fontWeight: selectedFleet ? 400 : 700, fontSize: 13 }}>All fleets</button>
      {selectedFleet && (
        <>
          <ChevronRight size={13} />
          <button onClick={backToFleet} style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: selectedEquipment ? NAVY : "#5F5E5A", fontWeight: selectedEquipment ? 400 : 700, fontSize: 13 }}>{selectedFleet}</button>
        </>
      )}
      {selectedEquipment && (
        <>
          <ChevronRight size={13} />
          <span style={{ fontWeight: 700, color: "#5F5E5A" }}>{selectedEquipment}</span>
        </>
      )}
    </div>
  );

  const rangeLabel = formatRangeForDisplay(fromDateTime, toDateTime);

  // ---------------------------------------------------------------
  // Build the content for whichever level is active — one render tree,
  // with the time frame selector rendered exactly once, above it, so it
  // never disappears or resets while drilling down or backing out.
  // ---------------------------------------------------------------
  let content;

  if (!selectedFleet) {
    // Level 1: fleet cards
    content = (
      <div>
        <p style={{ fontSize: 13, color: "#5F5E5A", margin: "0 0 16px" }}>Select a fleet to see its performance, then drill into an individual machine. Figures below are for {rangeLabel}.</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 14 }}>
          {fleets.map((f) => (
            <button
              key={f.fleet}
              onClick={() => setSelectedFleet(f.fleet)}
              style={{ textAlign: "left", background: "#F7F6F1", border: "1px solid #E4E2D8", borderRadius: 10, padding: 16, cursor: "pointer" }}
            >
              <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>{f.fleet}</p>
              <p style={{ fontSize: 12, color: "#5F5E5A", margin: "0 0 10px" }}>{f.count} equipment</p>
              <div style={{ display: "flex", gap: 14 }}>
                <span style={{ fontSize: 13 }}><b>{Math.round(f.metrics.availability * 100)}%</b> avail.</span>
                <span style={{ fontSize: 13 }}><b>{Math.round(f.metrics.utilisation * 100)}%</b> util.</span>
                <span style={{ fontSize: 13 }}><b>{f.metrics.breakdown_count}</b> breakdowns</span>
              </div>
            </button>
          ))}
        </div>
      </div>
    );
  } else if (selectedEquipment) {
    // Level 3: single equipment detail
    const asset = assets.find((a) => a.asset_id === selectedEquipment);
    const metrics = { ...assetMetrics[selectedEquipment], breakdown_count: breakdownsInRange(breakdowns, selectedEquipment, fromDateTime, toDateTime).length };
    const eqBreakdowns = breakdownsInRange(breakdowns, selectedEquipment, fromDateTime, toDateTime);
    content = (
      <div>
        {crumbs}
        <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 4px", color: NAVY }}>{asset.asset_name}</h3>
        <p style={{ fontSize: 13, color: "#5F5E5A", margin: "0 0 16px" }}>{asset.make} {asset.model} · Serial {asset.serial_number} · <Badge value={asset.status} /></p>
        <KpiRow metrics={metrics} />
        <h4 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px", color: NAVY }}>Breakdowns — {rangeLabel}</h4>
        {eqBreakdowns.length === 0 ? (
          <p style={{ fontSize: 13, color: "#898781" }}>No breakdowns for this machine in the selected date range.</p>
        ) : (
          <DataTable
            columns={[["breakdown_date","Date"],["wo_reference","Work order #"],["component_affected","Component"],["severity","Severity"],["repair_status","Status"],["downtime_hours","Downtime (hrs)"]].map(([key,label])=>({key,label}))}
            rows={eqBreakdowns}
            exportName={`${asset.asset_id}_Breakdowns`}
          />
        )}
      </div>
    );
  } else {
    // Level 2: fleet-level view with equipment list + compare
    const fleetMetrics = aggregateMetrics(breakdowns, fleetAssets.map((a) => a.asset_id), fromDateTime, toDateTime);
    content = (
      <div>
        {crumbs}
        <KpiRow metrics={fleetMetrics} />

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <h4 style={{ fontSize: 14, fontWeight: 700, margin: 0, color: NAVY }}>Equipment in this fleet</h4>
          <button
            onClick={() => { setCompareMode(!compareMode); setCompareIds([]); }}
            style={{ display: "flex", alignItems: "center", gap: 6, background: compareMode ? NAVY : "#fff", color: compareMode ? "#fff" : NAVY, border: `1px solid ${NAVY}`, fontSize: 13, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer" }}
          >
            <GitCompare size={14} /> {compareMode ? "Cancel compare" : "Compare equipment"}
          </button>
        </div>

        <div style={{ border: "1px solid #E4E2D8", borderRadius: 10, overflow: "hidden", marginBottom: compareMode && compareIds.length >= 2 ? 24 : 0 }}>
          {fleetAssets.map((a, i) => {
            const m = assetMetrics[a.asset_id] || {};
            const bcount = breakdownsInRange(breakdowns, a.asset_id, fromDateTime, toDateTime).length;
            return (
              <div
                key={a.asset_id}
                onClick={() => (compareMode ? toggleCompare(a.asset_id) : setSelectedEquipment(a.asset_id))}
                style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none", cursor: "pointer", background: compareIds.includes(a.asset_id) ? "#EAF1FB" : "#fff" }}
              >
                {compareMode && (
                  <input type="checkbox" checked={compareIds.includes(a.asset_id)} onChange={() => toggleCompare(a.asset_id)} onClick={(e) => e.stopPropagation()} style={{ width: 16, height: 16 }} />
                )}
                <div style={{ flex: 1 }}>
                  <p style={{ fontSize: 13.5, fontWeight: 600, margin: 0, color: "#2C2C2A" }}>{a.asset_id} — {a.asset_name}</p>
                  <p style={{ fontSize: 12, color: "#898781", margin: "2px 0 0" }}>{a.serial_number}</p>
                </div>
                <span style={{ fontSize: 13 }}>{Math.round((m.availability || 0) * 100)}% avail.</span>
                <span style={{ fontSize: 13 }}>{Math.round((m.utilisation || 0) * 100)}% util.</span>
                <span style={{ fontSize: 13 }}>{bcount} breakdowns</span>
                <Badge value={a.status} />
                {!compareMode && <ChevronRight size={15} style={{ color: "#B4B2A9" }} />}
              </div>
            );
          })}
        </div>

        {compareMode && compareIds.length >= 2 && (
          <div>
            <h4 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px", color: NAVY }}>Comparison — {rangeLabel}</h4>
            <div style={{ overflowX: "auto", border: "1px solid #E4E2D8", borderRadius: 10 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: "#F7F6F1" }}>
                    <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#5F5E5A", borderBottom: "1px solid #E4E2D8" }}>Metric</th>
                    {compareIds.map((id) => (
                      <th key={id} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#5F5E5A", borderBottom: "1px solid #E4E2D8", whiteSpace: "nowrap" }}>{id}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {METRIC_DEFS.map((met, i) => (
                    <tr key={met.key} style={{ borderTop: i > 0 ? "1px solid #EFEEE7" : "none" }}>
                      <td style={{ padding: "9px 12px", fontWeight: 600 }}>{met.label}</td>
                      {compareIds.map((id) => {
                        const val = met.key === "breakdown_count" ? breakdownsInRange(breakdowns, id, fromDateTime, toDateTime).length : assetMetrics[id]?.[met.key];
                        return <td key={id} style={{ padding: "9px 12px" }}>{val != null ? met.fmt(val) : "—"}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
        {compareMode && compareIds.length === 1 && (
          <p style={{ fontSize: 13, color: "#898781" }}>Select at least one more machine to compare.</p>
        )}
      </div>
    );
  }

  return (
    <div>
      <DateRangePicker
        fromDateTime={fromDateTime}
        toDateTime={toDateTime}
        setFromDateTime={setFromDateTime}
        setToDateTime={setToDateTime}
      />
      {content}
    </div>
  );
}
function Dashboard({ assets, workOrders, plannedMaintenance, components, parts, inspections }) {
  const fleetAvailability = useMemo(() => {
    const byFleet = {};
    assets.forEach((a) => {
      const m = assetMetrics[a.asset_id];
      if (!m) return;
      if (!byFleet[a.fleet]) byFleet[a.fleet] = [];
      byFleet[a.fleet].push(m.availability);
    });
    return Object.entries(byFleet).map(([fleet, vals]) => ({
      fleet, availability: Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100),
    }));
  }, [assets]);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12, marginBottom: 24 }}>
        <MetricCard label="Total assets" value={assets.length} />
        <MetricCard label="Open work orders" value={workOrders.filter((w) => w.status !== "Closed").length} />
        <MetricCard label="Services due soon / overdue" value={plannedMaintenance.filter((p) => p.status !== "OK").length} />
        <MetricCard label="Fleet avg availability" value={`${Math.round(fleetAvailability.reduce((a, f) => a + f.availability, 0) / fleetAvailability.length)}%`} />
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

      <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Availability by fleet, this month</h3>
      <div style={{ height: 220, marginBottom: 8 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={fleetAvailability}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EFEEE7" vertical={false} />
            <XAxis dataKey="fleet" tick={{ fontSize: 12, fill: "#5F5E5A" }} axisLine={{ stroke: "#E4E2D8" }} tickLine={false} />
            <YAxis tick={{ fontSize: 12, fill: "#5F5E5A" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}%`} domain={[0, 100]} />
            <Tooltip formatter={(v) => `${v}%`} />
            <Bar dataKey="availability" fill={NAVY} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function buildTableConfig(assets, dailyHours, breakdowns, fuelLog, oilConsumption, workOrders, plannedMaintenance, inspections, components, tyres, parts, warrantyDocs, auditLog) {
  return {
    assets: { title: "Assets", cols: [["asset_id","Equipment #"],["asset_name","Name"],["make","Make"],["model","Model"],["fleet","Fleet"],["serial_number","Serial number"],["status","Status"],["current_hours","Current hours"]], data: assets },
    daily_hours: { title: "Daily Hours", cols: [["log_date","Date"],["asset_id","Equipment #"],["opening_hours","Opening hours"],["closing_hours","Closing hours"],["hours_run","Hours run"]], data: dailyHours },
    fuel_log: { title: "Fuel Log", cols: [["fill_date","Date"],["asset_id","Equipment #"],["hour_meter","Hour meter"],["litres","Litres"],["consumption_rate","Rate (L/hr)"],["recorded_by","Recorded by"]], data: fuelLog },
    oil_consumption: { title: "Oil Consumption", cols: [["fill_date","Date"],["asset_id","Equipment #"],["hour_meter","Hour meter"],["oil_type","Oil type"],["litres","Litres"],["fill_reason","Reason"],["consumption_rate","Rate (L/hr)"],["recorded_by","Recorded by"]], data: oilConsumption },
    breakdowns: { title: "Breakdowns", cols: [["breakdown_date","Date"],["asset_id","Equipment #"],["wo_reference","Work order #"],["component_affected","Component"],["cause_code","Cause"],["severity","Severity"],["repair_status","Status"],["downtime_hours","Downtime (hrs)"]], data: breakdowns },
    work_orders: { title: "Work Orders", cols: [["wo_no","Work order #"],["asset_id","Equipment #"],["work_type","Type"],["priority","Priority"],["status","Status"],["request_date","Requested"]], data: workOrders },
    planned_maintenance: { title: "Planned Maintenance", cols: [["asset_id","Equipment #"],["asset_name","Name"],["current_hours","Current hours"],["next_service_due","Next service due"],["remaining","Hours remaining"],["status","Status"]], data: plannedMaintenance },
    inspections: { title: "Inspections", cols: [["log_date","Date"],["asset_id","Equipment #"],["inspection_type","Type"],["inspector","Inspector"],["result","Result"],["wo_reference","Work order #"]], data: inspections },
    components: { title: "Components", cols: [["component_id","Component #"],["asset_id","Equipment #"],["component_type","Type"],["life_used_pct","Life used"],["status","Status"]], data: components },
    tyres: { title: "Tyres", cols: [["asset_id","Equipment #"],["position","Position"],["remaining_life","Remaining life (hrs)"],["replacement_due","Status"]], data: tyres },
    parts: { title: "Parts Inventory", cols: [["part_no","Part #"],["description","Description"],["qty_in_stock","Qty in stock"],["minimum_qty","Minimum qty"],["reorder_status","Status"]], data: parts },
    warranty_docs: { title: "Warranty & Documents", cols: [["asset_id","Equipment #"],["type","Type"],["reference","Reference"],["expiry","Expiry"],["status","Status"]], data: warrantyDocs },
    // Cost Ledger stays on mock data — hidden from the nav per standing request, no point wiring it live yet
    cost_ledger: { title: "Cost Ledger", cols: [["cost_date","Date"],["asset_id","Equipment #"],["cost_type","Cost type"],["description","Description"],["total_cost","Total cost"]], data: costLedger.map(c => ({...c, total_cost: `R${c.total_cost.toLocaleString()}`})) },
    audit: { title: "Audit Trail", cols: [["changed_at","Timestamp"],["changed_by_email","User"],["action","Action"],["table_name","Table"],["record_id","Record #"]], data: auditLog },
  };
}

export default function App() {
  const [active, setActive] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const isMobile = useIsMobile();

  // Live data from Supabase — Assets, Daily Hours, Breakdowns (the
  // foundation Dashboard and Fleet Performance are built on).
  const [assets, setAssets] = useState([]);
  const [dailyHours, setDailyHours] = useState([]);
  const [breakdowns, setBreakdowns] = useState([]);
  const [coreLoading, setCoreLoading] = useState(true);
  const [coreError, setCoreError] = useState(null);

  // Live data for every other tab. Cost Ledger stays on mock data (hidden
  // from the nav per standing request).
  const [fuelLog, setFuelLog] = useState([]);
  const [oilConsumption, setOilConsumption] = useState([]);
  const [workOrders, setWorkOrders] = useState([]);
  const [plannedMaintenance, setPlannedMaintenance] = useState([]);
  const [inspections, setInspections] = useState([]);
  const [components, setComponents] = useState([]);
  const [tyres, setTyres] = useState([]);
  const [parts, setParts] = useState([]);
  const [warrantyDocs, setWarrantyDocs] = useState([]);
  const [auditLog, setAuditLog] = useState([]);
  const [restLoading, setRestLoading] = useState(true);
  const [restError, setRestError] = useState(null);

  useEffect(() => {
    let cancelled = false;

    async function loadCoreData() {
      setCoreLoading(true);
      setCoreError(null);
      try {
        const [assetsRes, hoursRes, breakdownsRes, currentHoursRes] = await Promise.all([
          supabase.from("assets").select("*"),
          supabase.from("daily_hours_calc").select("*"),
          supabase.from("breakdown_log_calc").select("*"),
          supabase.from("current_hours").select("*"),
        ]);
        if (assetsRes.error) throw assetsRes.error;
        if (hoursRes.error) throw hoursRes.error;
        if (breakdownsRes.error) throw breakdownsRes.error;
        if (currentHoursRes.error) throw currentHoursRes.error;

        if (cancelled) return;

        // assets table has no current_hours column of its own — it comes
        // from the current_hours view, merged in here by asset_id
        const currentHoursByAsset = Object.fromEntries(
          (currentHoursRes.data || []).map((r) => [r.asset_id, r.current_hours])
        );
        const mergedAssets = (assetsRes.data || []).map((a) => ({
          ...a,
          current_hours: currentHoursByAsset[a.asset_id] ?? null,
        }));

        setAssets(mergedAssets);
        setDailyHours(hoursRes.data || []);
        setBreakdowns(breakdownsRes.data || []);
      } catch (err) {
        if (!cancelled) setCoreError(err.message || String(err));
      } finally {
        if (!cancelled) setCoreLoading(false);
      }
    }

    async function loadRestOfData() {
      setRestLoading(true);
      setRestError(null);
      try {
        const [
          fuelRes, oilRes, woRes, serviceRes, inspRes,
          compRes, tyreRes, partsRes, warrRes, docRes, auditRes,
        ] = await Promise.all([
          supabase.from("fuel_log_calc").select("*"),
          supabase.from("oil_consumption_calc").select("*"),
          supabase.from("work_orders_calc").select("*"),
          supabase.from("service_schedule_calc").select("*"),
          supabase.from("inspections").select("*"),
          supabase.from("component_status").select("*"),
          supabase.from("tyre_tracking_calc").select("*"),
          supabase.from("parts_inventory_calc").select("*"),
          supabase.from("warranty_register_calc").select("*"),
          supabase.from("document_register_calc").select("*"),
          supabase.from("audit_report").select("*").limit(200),
        ]);
        const results = { fuelRes, oilRes, woRes, serviceRes, inspRes, compRes, tyreRes, partsRes, warrRes, docRes, auditRes };
        for (const [name, res] of Object.entries(results)) {
          if (res.error) throw new Error(`${name}: ${res.error.message}`);
        }

        if (cancelled) return;

        // Field names below are reshaped to match what the existing table
        // columns already expect, so DataTable/export don't need changes.
        setFuelLog((fuelRes.data || []).map((r) => ({ ...r, consumption_rate: r.consumption_rate_l_per_hr })));
        setOilConsumption((oilRes.data || []).map((r) => ({ ...r, consumption_rate: r.consumption_rate_l_per_hr })));
        setWorkOrders(woRes.data || []);
        setPlannedMaintenance((serviceRes.data || []).map((r) => ({
          ...r,
          remaining: r.current_hours != null ? r.next_service_due - r.current_hours : null,
        })));
        setInspections(inspRes.data || []);
        setComponents((compRes.data || []).map((r) => ({
          ...r,
          component_id: r.component_code,
          component_type: r.component_name,
          life_used_pct: r.expected_life_hours
            ? `${Math.round((r.current_comp_hrs / r.expected_life_hours) * 100)}%`
            : "—",
        })));
        setTyres(tyreRes.data || []);
        setParts(partsRes.data || []);

        const warrantyRows = (warrRes.data || []).map((r) => ({
          asset_id: r.asset_id, type: `Warranty — ${r.component || ""}`.trim(),
          reference: r.serial_number, expiry: r.warranty_end, status: r.status,
        }));
        const docRows = (docRes.data || []).map((r) => ({
          asset_id: r.asset_id, type: r.document_type,
          reference: r.reference, expiry: r.expiry_date, status: r.status,
        }));
        setWarrantyDocs([...warrantyRows, ...docRows]);

        setAuditLog(auditRes.data || []);
      } catch (err) {
        if (!cancelled) setRestError(err.message || String(err));
      } finally {
        if (!cancelled) setRestLoading(false);
      }
    }

    loadCoreData();
    loadRestOfData();
    return () => { cancelled = true; };
  }, []);

  // On mobile the sidebar starts closed (it's an overlay drawer, not a
  // permanent column); on desktop it starts open as a fixed column.
  useEffect(() => {
    setSidebarOpen(!isMobile);
  }, [isMobile]);

  const tableConfig = useMemo(
    () => buildTableConfig(assets, dailyHours, breakdowns, fuelLog, oilConsumption, workOrders, plannedMaintenance, inspections, components, tyres, parts, warrantyDocs, auditLog),
    [assets, dailyHours, breakdowns, fuelLog, oilConsumption, workOrders, plannedMaintenance, inspections, components, tyres, parts, warrantyDocs, auditLog]
  );
  const activeConfig = tableConfig[active];

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

          <button
            onClick={() => supabase.auth.signOut()}
            style={{ display: "flex", alignItems: "center", gap: 10, padding: isMobile ? "12px 16px" : "9px 16px", margin: "12px 8px 8px", borderRadius: 8, background: "transparent", border: "none", color: "rgba(255,255,255,0.72)", fontSize: 14, cursor: "pointer", minHeight: 40, justifyContent: sidebarOpen ? "flex-start" : "center" }}
          >
            <LogOut size={16} style={{ flexShrink: 0 }} />
            {sidebarOpen && <span>Sign out</span>}
          </button>
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
              {active === "dashboard" ? "Dashboard" : active === "fleet_performance" ? "Fleet Performance" : activeConfig.title}
            </span>
          </div>
        )}

        <div style={{ padding: isMobile ? "16px" : "24px 28px", overflow: "auto", flex: 1 }}>
          {!isMobile && (
            <h2 style={{ fontSize: 19, fontWeight: 700, margin: "0 0 18px", color: NAVY }}>
              {active === "dashboard" ? "Dashboard" : active === "fleet_performance" ? "Fleet Performance" : activeConfig.title}
            </h2>
          )}

          {(coreError || restError) && (
            <div style={{ background: "#FCEBEB", border: "1px solid #E3A8A8", borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13, color: "#791F1F" }}>
              Couldn't load live data from Supabase: {coreError || restError}. Check your .env values and that the SQL files ran successfully.
            </div>
          )}

          {(coreLoading && ["dashboard", "fleet_performance", "assets", "daily_hours", "breakdowns"].includes(active)) ||
           (restLoading && ["dashboard", "fuel_log", "oil_consumption", "work_orders", "planned_maintenance", "inspections", "components", "tyres", "parts", "warranty_docs", "audit"].includes(active)) ? (
            <p style={{ fontSize: 13, color: "#898781" }}>Loading…</p>
          ) : active === "dashboard" ? (
            <Dashboard assets={assets} workOrders={workOrders} plannedMaintenance={plannedMaintenance} components={components} parts={parts} inspections={inspections} />
          ) : active === "fleet_performance" ? (
            <FleetPerformance assets={assets} breakdowns={breakdowns} />
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

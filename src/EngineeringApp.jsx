import React, { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Truck, Clock, AlertTriangle, ClipboardList,
  CalendarClock, ShieldCheck, CircleDot, Package, FileText,
  DollarSign, History, Factory, Download, Search, Menu, X,
  Layers, ChevronRight, GitCompare, Fuel, Droplet, LogOut, Upload, Trash2, Printer, FileBarChart, ChevronDown, Activity, MapPin
} from "lucide-react";
import { supabase } from "./lib/supabaseClient";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend, ReferenceLine, LabelList,
} from "recharts";

const NAVY = "#1F6668";
const SIDEBAR = "#203B46";
const ACCENT = "#3D7379";
const ACCENT_LIGHT = "#DCEFED";
const GREEN = "#DCEFED";
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
// Mock data - mirrors the Supabase schema/views. Swap for real
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

// Real per-equipment performance metrics now come from the
// plant_performance_kpi() SQL function via RPC (see FleetPerformance),
// not from a hardcoded object - this used to be a placeholder here,
// which is why Fleet Performance showed nothing once real assets
// replaced the mock EQ-001..EQ-007 IDs it was keyed on.

const METRIC_DEFS = [
  { key: "hours_worked", label: "Hours worked", fmt: (v) => Number(v).toLocaleString(undefined, { maximumFractionDigits: 1 }) },
  { key: "availability", label: "Availability", fmt: (v) => `${Math.round(v * 100)}%` },
  { key: "utilisation", label: "Utilisation", fmt: (v) => `${Math.round(v * 100)}%` },
  { key: "breakdown_count", label: "Events", fmt: (v) => String(v) },
  { key: "mtbf", label: "MTBF (hrs)", fmt: (v) => Number(v).toLocaleString(undefined, { maximumFractionDigits: 1 }) },
  { key: "mttr", label: "MTTR (hrs)", fmt: (v) => Number(v).toLocaleString(undefined, { maximumFractionDigits: 1 }) },
  { key: "availability_index", label: "Availability Index", fmt: (v) => `${Number(v).toFixed(1)}%` },
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

// Reference "now" for time-frame filtering - matches the current mock data window
// This used to be a hardcoded string from the mock-data phase of this
// project ("2026-08-11T16:00:00") and never got updated when the app was
// wired to real data - meaning Fleet Performance's default date range was
// frozen at a fixed moment in the past, while Dashboard computed the
// actual current time fresh. That's exactly why their Availability
// numbers didn't match: two different date ranges, not two different
// calculations. Fixed by using the real current time.
const NOW = new Date();

function pad2(n) { return String(n).padStart(2, "0"); }

// Format a Date for an <input type="datetime-local"> value
function toDatetimeLocalValue(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatRangeForDisplay(fromStr, toStr) {
  const opts = { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" };
  const from = new Date(fromStr);
  const to = new Date(toStr);
  return `${from.toLocaleString("en-ZA", opts)} - ${to.toLocaleString("en-ZA", opts)}`;
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
  { asset_id: "EQ-001", type: "Warranty - Engine", reference: "ENG-2019-4471", expiry: "2026-08-15", status: "EXPIRING SOON" },
  { asset_id: "EQ-001", type: "Document - Inspection Cert", reference: "COC-2026-0142", expiry: "2026-08-20", status: "EXPIRING SOON" },
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
  { key: "daily_hours", label: "Daily Hours", icon: Clock, operatorVisible: true },
  { key: "fuel_log", label: "Fuel Log", icon: Fuel, operatorVisible: true },
  { key: "oil_consumption", label: "Oil Consumption", icon: Droplet, operatorVisible: true },
  { key: "breakdowns", label: "Events", icon: AlertTriangle, operatorVisible: true },
  { key: "work_orders", label: "Work Orders", icon: ClipboardList, operatorVisible: true },
  { key: "downtime_summary", label: "Downtime Summary", icon: FileBarChart },
  { key: "mtbf_mttr", label: "MTBF / MTTR Report", icon: Activity },
  { key: "planned_maintenance", label: "Planned Maintenance", icon: CalendarClock },
  { key: "inspections", label: "Inspections", icon: ShieldCheck, operatorVisible: true },
  // Backlogs is rendered as its own collapsible group in the sidebar
  // (see engineeringExpanded/backlogsExpanded below) rather than as flat
  // NAV entries, but the two keys still need to exist here so the rest of
  // the app (page routing, operator visibility filtering) treats them
  // like any other nav destination.
  { key: "backlog_report", label: "Backlog Report", icon: ClipboardList, group: "backlogs", operatorVisible: true },
  { key: "daily_service", label: "Daily Service", icon: CalendarClock, group: "backlogs", operatorVisible: true },
  { key: "components", label: "Components", icon: CircleDot },
  { key: "tyres", label: "Tyres", icon: CircleDot },
  { key: "parts", label: "Parts Inventory", icon: Package },
  { key: "warranty_docs", label: "Warranty & Documents", icon: FileText },
  // Cost Ledger hidden for now per request - data/config still here, just
  // not in the nav, so it's a one-line change to bring back.
  // { key: "cost_ledger", label: "Cost Ledger", icon: DollarSign },
  { key: "site_management", label: "Site Management", icon: MapPin },
  { key: "component_codes", label: "Component Codes", icon: CircleDot },
  { key: "audit", label: "Audit Trail", icon: History },
];

function statusColor(status) {
  const s = (status || "").toUpperCase();
  if (["OVERDUE", "CRITICAL", "CHANGE OUT", "REORDER", "REPLACE NOW", "EXPIRED", "FAIL", "OPEN", "OVER SCHEDULED HOURS"].includes(s))
    return { bg: "#F6E2E0", text: "#7A3330" };
  if (["DUE SOON", "PLAN CHANGE", "EXPIRING SOON", "MEDIUM", "PLANNED", "APPROACHING LIMIT"].includes(s))
    return { bg: "#F5E9D8", text: "#7A5320" };
  if (["OK", "CLOSED", "PASS", "ACTIVE", "VALID"].includes(s))
    return { bg: "#E2EFE9", text: "#2C5646" };
  return { bg: "#F1EFE8", text: "#444441" };
}

function Badge({ value }) {
  if (value === null || value === undefined || value === "") return <span style={{ color: "#B4B2A9" }}>-</span>;
  const c = statusColor(value);
  return (
    <span style={{ background: c.bg, color: c.text, fontSize: 12, fontWeight: 600, padding: "2px 9px", borderRadius: 6, whiteSpace: "nowrap" }}>
      {value}
    </span>
  );
}

const STATUS_FIELDS = new Set(["status", "repair_status", "result", "replacement_due", "reorder_status", "severity", "priority", "action"]);

class PageErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("Page crashed:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 10, padding: 20, color: "#7A3330" }}>
          <p style={{ fontWeight: 700, fontSize: 14, margin: "0 0 6px" }}>This page hit an error and couldn't display.</p>
          <p style={{ fontSize: 13, margin: "0 0 10px" }}>{String(this.state.error.message || this.state.error)}</p>
          <button
            onClick={() => this.setState({ error: null })}
            style={{ background: "#fff", border: "1px solid #DDB6B2", color: "#7A3330", padding: "7px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
          >
            Try again
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

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
    // real header row using the same friendly labels shown on screen -
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
          <Search size={15} style={{ position: "absolute", left: 10, top: 13, color: "#859195" }} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search this table"
            style={{ width: "100%", padding: "10px 10px 10px 32px", fontSize: 14, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", boxSizing: "border-box", minHeight: 42 }}
          />
        </div>
        <button
          onClick={exportToExcel}
          style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 14, fontWeight: 600, padding: "10px 14px", borderRadius: 8, cursor: "pointer", minHeight: 42, width: isMobile ? "100%" : "auto" }}
        >
          <Download size={14} /> Export to Excel
        </button>
      </div>
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10, WebkitOverflowScrolling: "touch" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: isMobile ? 12.5 : 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => (
                <th key={c.key} style={{ textAlign: "left", padding: isMobile ? "9px 10px" : "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr key={i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                {columns.map((c) => {
                  const val = row[c.key];
                  const display = STATUS_FIELDS.has(c.key) ? <Badge value={val} />
                    : typeof val === "number" && !Number.isInteger(val) ? val.toFixed(2)
                    : (val ?? <span style={{ color: "#B4B2A9" }}>-</span>);
                  return (
                    <td key={c.key} style={{ padding: isMobile ? "9px 10px" : "9px 12px", color: "#183642", whiteSpace: "nowrap" }}>
                      {display}
                    </td>
                  );
                })}
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No rows match your search.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      {isMobile && <p style={{ fontSize: 11.5, color: "#859195", margin: "8px 2px 0" }}>Swipe the table sideways to see more columns.</p>}
    </div>
  );
}

function MetricCard({ label, value, icon: Icon, accentColor, onClick }) {
  if (!Icon) {
    return (
      <div onClick={onClick} style={{ background: "#F7F8F6", borderRadius: 10, padding: "14px 16px", cursor: onClick ? "pointer" : "default" }}>
        <p style={{ fontSize: 12, color: "#4B5659", margin: "0 0 4px" }}>{label}</p>
        <p style={{ fontSize: 22, fontWeight: 700, margin: 0, color: "#183642" }}>{value}</p>
      </div>
    );
  }
  const color = accentColor || NAVY;
  return (
    <div
      onClick={onClick}
      style={{ background: "#fff", border: "1px solid #E2E6E3", borderRadius: 12, padding: "16px 18px", display: "flex", alignItems: "center", gap: 14, cursor: onClick ? "pointer" : "default" }}
    >
      <div style={{ width: 42, height: 42, borderRadius: 10, background: `${color}18`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Icon size={20} style={{ color }} />
      </div>
      <div style={{ minWidth: 0 }}>
        <p style={{ fontSize: 12, color: "#4B5659", margin: "0 0 2px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{label}</p>
        <p style={{ fontSize: 21, fontWeight: 700, margin: 0, color: "#183642" }}>{value}</p>
      </div>
      {onClick && <ChevronRight size={16} style={{ color: "#B4B2A9", flexShrink: 0 }} />}
    </div>
  );
}

const CAUSE_CODES = ["Mechanical Failure", "Electrical Fault", "Hydraulic Failure", "Tyre/Track",
  "Operator Error", "Wear & Tear", "Structural", "Overheating", "Lubrication/Fluid", "Weather", "Unknown", "Other"];
const SEVERITIES = ["Low", "Medium", "High", "Critical"];
const REPAIR_STATUSES = ["In Progress", "Awaiting Parts", "Closed"];

function nowForInput() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function isoToInputValue(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// A datetime-local input plus a one-click "Now" button - the native
// picker stays available for backdating, but the overwhelming majority
// case (logging something as it happens) becomes a single click instead
// of manually setting date and time by hand.
// Mining equipment can only genuinely be in one state at a time - broken
// down, or being worked on, not both. If two open events exist for the
// same asset simultaneously, their downtime windows overlap and get
// double-counted in Availability, Utilisation, MTBF and MTTR. This checks
// for that before letting a new active breakdown or work order through.
async function checkAssetOverlap(assetId, excludeTable, excludeId) {
  const [breakdownRes, woRes] = await Promise.all([
    supabase.from("breakdown_log")
      .select("id, component_affected, breakdown_date")
      .eq("asset_id", assetId)
      .neq("repair_status", "Closed"),
    supabase.from("work_orders")
      .select("id, problem_scope, request_date")
      .eq("asset_id", assetId)
      .neq("status", "Closed")
      .not("actual_start", "is", null)
      .is("actual_finish", null),
  ]);

  if (breakdownRes.error || woRes.error) {
    return { error: breakdownRes.error?.message || woRes.error?.message };
  }

  const conflicts = [
    ...(breakdownRes.data || [])
      .filter((r) => !(excludeTable === "breakdown_log" && r.id === excludeId))
      .map((r) => `an open breakdown (${r.component_affected || "no component set"}, logged ${r.breakdown_date})`),
    ...(woRes.data || [])
      .filter((r) => !(excludeTable === "work_orders" && r.id === excludeId))
      .map((r) => `an active work order (${r.problem_scope || "no scope set"}, requested ${r.request_date})`),
  ];

  return { conflict: conflicts[0] || null };
}

// An hour meter can only ever increase, and can only increase by as many
// hours as have actually elapsed in the real world - a machine running
// two 12-hour shifts a day can rack up at most 24 engine-hours per
// calendar day, never more. This compares a new Fuel Log / Oil
// Consumption reading against the most recent known Daily Hours reading
// for that asset and flags anything that breaks either rule. It's
// advisory, not a hard block - a genuinely unusual but correct reading
// shouldn't be impossible to save, just worth a second look.
function checkHourMeterPlausibility(assetId, newHourMeter, newDateStr, dailyHours) {
  if (newHourMeter === "" || newHourMeter == null || !newDateStr) return null;

  const priorReadings = dailyHours
    .filter((d) => d.asset_id === assetId && d.log_date <= newDateStr && d.closing_hours != null)
    .sort((a, b) => b.log_date.localeCompare(a.log_date) || (b.shift || "").localeCompare(a.shift || ""));
  const baseline = priorReadings[0];
  if (!baseline) return null; // nothing recorded yet for this asset to compare against

  const elapsedDays = Math.max(0, (new Date(newDateStr) - new Date(baseline.log_date)) / 86400000);
  const maxPossibleHours = (elapsedDays + 1) * 24; // generous - full 24hrs allowed on both the baseline day and the new day, to avoid false positives from shift-timing ambiguity
  const hoursGained = Number(newHourMeter) - Number(baseline.closing_hours);

  if (hoursGained < 0) {
    return `This reading (${newHourMeter}) is lower than the last recorded hours for this machine (${baseline.closing_hours} on ${baseline.log_date}). Hour meters don't run backward - worth double-checking.`;
  }
  if (hoursGained > maxPossibleHours) {
    return `This is ${hoursGained.toFixed(1)} hours higher than the last recorded reading (${baseline.closing_hours} on ${baseline.log_date}), but only about ${Math.ceil(elapsedDays) || 1} day(s) have passed - more than the machine could physically run at 24 hrs/day across both shifts. Worth double-checking the reading or the date.`;
  }
  return null;
}

function DateTimeField({ value, onChange, max, required, disabled }) {
  const baseFieldStyle = { padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const fieldStyle = disabled ? { ...baseFieldStyle, background: "#F2F1EA", color: "#859195" } : baseFieldStyle;

  const [datePart, timePart] = value ? value.split("T") : ["", ""];
  const [maxDate, maxTimeRaw] = max ? max.split("T") : [undefined, undefined];
  // Only restrict the time picker's max when the selected date IS the max
  // date - on any earlier date, any time of day is still valid.
  const maxTime = datePart && datePart === maxDate ? maxTimeRaw : undefined;

  const handleDateChange = (newDate) => onChange(newDate ? `${newDate}T${timePart || "00:00"}` : "");
  const handleTimeChange = (newTime) => onChange(datePart ? `${datePart}T${newTime}` : "");

  return (
    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
      <input type="date" value={datePart} onChange={(e) => handleDateChange(e.target.value)} max={maxDate} required={required} disabled={disabled} style={{ ...fieldStyle, flex: "1 1 130px", minWidth: 0 }} />
      <input type="time" value={timePart} onChange={(e) => handleTimeChange(e.target.value)} max={maxTime} required={required} disabled={disabled} style={{ ...fieldStyle, flex: "1 1 100px", minWidth: 0 }} />
      <button
        type="button"
        onClick={() => onChange(nowForInput())}
        disabled={disabled}
        title="Set to right now"
        style={{ background: disabled ? "#F2F1EA" : "#F2F1EA", border: "1px solid #E2E6E3", color: disabled ? "#B4B2A9" : NAVY, fontSize: 12.5, fontWeight: 600, padding: "0 12px", borderRadius: 8, cursor: disabled ? "default" : "pointer", whiteSpace: "nowrap" }}
      >
        Now
      </button>
    </div>
  );
}

// Auto-handles date/time: defaults "Downtime Start" to right now (so nobody
// has to think about timestamps when logging a breakdown as it happens),
// and auto-fills "Downtime End" with now the moment Status is set to
// Closed. Both stay editable for logging something that happened earlier.
// These are exactly the fields Availability, Utilisation, MTBF and MTTR
// are calculated from downstream.
const SHIFTS = ["Day", "Night"];

function todayForInput() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// Finds what Opening Hours SHOULD be for a given asset/date/shift, by
// looking at the closing hours of the immediately preceding shift already
// on record. This is a live preview only - the database (via daily_hours_calc's
// LAG-based view) is the actual source of truth once saved.
function findExpectedOpeningHours(dailyHours, assetId, logDate, shift) {
  const rows = dailyHours
    .filter((r) => r.asset_id === assetId)
    .slice()
    .sort((a, b) => (a.log_date + a.shift).localeCompare(b.log_date + b.shift));
  const targetKey = logDate + shift;
  let candidate = null;
  for (const r of rows) {
    if ((r.log_date + r.shift) < targetKey) candidate = r;
    else break;
  }
  return candidate ? candidate.closing_hours : null;
}

function DailyHoursForm({ assets, dailyHours, existing, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [logDate, setLogDate] = useState(existing?.log_date || todayForInput());
  const [shift, setShift] = useState(existing?.shift || "Day");
  const [closingHours, setClosingHours] = useState(existing?.closing_hours ?? "");
  const [status, setStatus] = useState(existing?.status || "Operating");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const openingHours = isEdit && existing?.opening_hours != null
    ? existing.opening_hours
    : findExpectedOpeningHours(dailyHours, assetId, logDate, shift);

  const hoursRun = openingHours != null && closingHours !== ""
    ? Number(closingHours) - Number(openingHours)
    : null;
  const exceedsShiftLimit = hoursRun != null && hoursRun > 12;

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab before logging hours.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            Close
          </button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");

    if (closingHours === "") {
      setError("Enter Closing Hours.");
      return;
    }
    if (openingHours != null && Number(closingHours) < Number(openingHours)) {
      setError("Closing Hours can't be less than Opening Hours.");
      return;
    }

    setSaving(true);
    const payload = {
      asset_id: assetId, log_date: logDate, shift,
      closing_hours: Number(closingHours), status, notes: notes || null,
    };

    try {
      const { error: dbError } = isEdit
        ? await supabase.from("daily_hours").update(payload).eq("id", existing.id)
        : await supabase.from("daily_hours").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 460, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>
          {isEdit ? "Edit Daily Hours" : "Log Daily Hours"}
        </h3>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Date</label>
            <input type="date" value={logDate} onChange={(e) => setLogDate(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }} />
          </div>
          <div>
            <label style={labelStyle}>Shift</label>
            <select value={shift} onChange={(e) => setShift(e.target.value)} disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
              {SHIFTS.map((s) => <option key={s} value={s}>{s} {s === "Day" ? "(07:00-19:00)" : "(19:00-07:00)"}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 4 }}>
          <div>
            <label style={labelStyle}>Opening Hours</label>
            <input type="text" value={openingHours != null ? openingHours : "-"} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
          </div>
          <div>
            <label style={labelStyle}>Closing Hours</label>
            <input type="number" step="0.1" value={closingHours} onChange={(e) => setClosingHours(e.target.value)} required style={fieldStyle} />
          </div>
        </div>
        <p style={{ fontSize: 11.5, color: "#859195", margin: "0 0 14px" }}>
          Opening Hours fills in automatically from the previous shift - just enter Closing Hours.
        </p>

        {hoursRun != null && (
          <div style={{
            background: exceedsShiftLimit ? "#F6E2E0" : "#F7F8F6",
            border: `1px solid ${exceedsShiftLimit ? "#DDB6B2" : "#E2E6E3"}`,
            borderRadius: 8, padding: "10px 12px", marginBottom: 14, fontSize: 13,
          }}>
            Hours run this shift: <b>{hoursRun.toFixed(1)}</b>
            {exceedsShiftLimit && (
              <div style={{ color: "#7A3330", marginTop: 4 }}>
                ⚠ This exceeds 12 hours for a single shift - double check the reading before saving.
              </div>
            )}
          </div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Status</label>
            <select value={status} onChange={(e) => setStatus(e.target.value)} style={fieldStyle}>
              {["Operating", "Under Maintenance", "Breakdown", "Standby"].map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Notes</label>
            <input type="text" value={notes} onChange={(e) => setNotes(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>
            Cancel
          </button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Log Hours"}
          </button>
        </div>
      </form>
    </div>
  );
}

function ScheduledHoursForm({ fleets, existing, onClose, onSaved }) {
  const [fleet, setFleet] = useState(existing?.fleet || fleets[0] || "");
  const [scheduledHours, setScheduledHours] = useState(existing?.scheduled_hours ?? "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const monthStart = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}-01`;

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    if (scheduledHours === "") { setError("Enter a scheduled hours value."); return; }
    setSaving(true);
    try {
      const { error: dbError } = await supabase.from("scheduled_hours")
        .upsert({ fleet, month: monthStart, scheduled_hours: Number(scheduledHours) }, { onConflict: "fleet,month" });
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>Set Monthly Scheduled Hours</h3>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Fleet</label>
          <select value={fleet} onChange={(e) => setFleet(e.target.value)} style={fieldStyle}>
            {fleets.map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Scheduled Hours (this month)</label>
          <input type="number" step="1" value={scheduledHours} onChange={(e) => setScheduledHours(e.target.value)} style={fieldStyle} />
        </div>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </form>
    </div>
  );
}

function DailyHoursPage({ assets, dailyHours, userEmail, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [showBudgetForm, setShowBudgetForm] = useState(false);
  const [query, setQuery] = useState("");
  const [budget, setBudget] = useState([]);
  const [budgetLoading, setBudgetLoading] = useState(true);
  const [deleting, setDeleting] = useState(null);
  const [templateMonth, setTemplateMonth] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  });
  const [bulkImporting, setBulkImporting] = useState(false);
  const [bulkMessage, setBulkMessage] = useState(null);
  const bulkFileInputRef = React.useRef(null);

  const fleets = useMemo(() => [...new Set(assets.map((a) => a.fleet))], [assets]);

  const loadBudget = React.useCallback(async () => {
    setBudgetLoading(true);
    const monthStart = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}-01`;
    const { data, error } = await supabase.rpc("monthly_hours_vs_scheduled", { target_month: monthStart });
    if (!error) setBudget(data || []);
    setBudgetLoading(false);
  }, []);

  useEffect(() => { loadBudget(); }, [loadBudget]);

  const columns = [
    ["log_date", "Date"], ["shift", "Shift"], ["asset_id", "Equipment #"],
    ["opening_hours", "Opening hours"], ["closing_hours", "Closing hours"], ["hours_run", "Hours run"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    let rows = dailyHours;
    if (query.trim()) {
      const q = query.toLowerCase();
      rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return [...rows].sort((a, b) => (b.log_date || "").localeCompare(a.log_date || "") || (a.shift || "").localeCompare(b.shift || ""));
  }, [dailyHours, query]);

  const handleSaved = () => {
    setShowForm(false);
    setEditing(null);
    onRefresh();
  };

  const handleBudgetSaved = () => {
    setShowBudgetForm(false);
    loadBudget();
  };

  const handleDelete = async (reason) => {
    await deleteWithReason("daily_hours", deleting.id, "id", reason, userEmail);
    setDeleting(null);
    onRefresh();
  };

  const BULK_HEADERS = ["Equipment #", "Date", "Shift", "Closing Hours"];

  const downloadBulkTemplate = () => {
    const [year, month] = templateMonth.split("-").map(Number);
    const daysInMonth = new Date(year, month, 0).getDate();
    const existingByKey = new Map(dailyHours.map((r) => [`${r.asset_id}|${r.log_date}|${r.shift}`, r.closing_hours]));

    const dataRows = [];
    assets.forEach((a) => {
      for (let d = 1; d <= daysInMonth; d++) {
        const dateStr = `${year}-${String(month).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
        ["Day", "Night"].forEach((shift) => {
          const existing = existingByKey.get(`${a.asset_id}|${dateStr}|${shift}`);
          dataRows.push([a.asset_id, dateStr, shift, existing ?? ""]);
        });
      }
    });

    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const aoa = [
      [`Daily Hours Bulk Template - ${templateMonth}`],
      [`Exported: ${timestamp}. Fill in Closing Hours only - Opening Hours and everything else is calculated automatically. Leave a row blank to skip it.`],
      [],
      BULK_HEADERS, ...dataRows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: BULK_HEADERS.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: BULK_HEADERS.length - 1 } },
    ];
    ws["!cols"] = BULK_HEADERS.map((h) => ({ wch: Math.max(h.length + 2, 16) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Daily Hours");
    XLSX.writeFile(wb, `Daily_Hours_Template_${templateMonth}.xlsx`);
  };

  const handleBulkFileSelected = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBulkImporting(true);
    setBulkMessage(null);

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headerRowIndex = raw.findIndex((row) => row[0] === "Equipment #");
      const dataStart = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;

      const validAssetIds = new Set(assets.map((a) => a.asset_id));
      const parsedRows = raw.slice(dataStart)
        .filter((row) => row[0] !== "" && row[0] != null && row[3] !== "" && row[3] != null)
        .map((row, i) => ({
          rowNum: dataStart + i + 1,
          asset_id: String(row[0]).trim(),
          log_date: row[1] instanceof Date ? row[1].toISOString().slice(0, 10) : String(row[1]).trim(),
          shift: String(row[2]).trim(),
          closing_hours: Number(row[3]),
        }));

      const errors = [];
      const validRows = [];

      // Check each row is individually sane first.
      parsedRows.forEach((r) => {
        if (!validAssetIds.has(r.asset_id)) { errors.push(`Row ${r.rowNum}: "${r.asset_id}" isn't a known Equipment #.`); return; }
        if (r.shift !== "Day" && r.shift !== "Night") { errors.push(`Row ${r.rowNum}: Shift must be "Day" or "Night", got "${r.shift}".`); return; }
        if (isNaN(r.closing_hours) || r.closing_hours < 0) { errors.push(`Row ${r.rowNum}: Closing Hours "${r.closing_hours}" isn't a valid number.`); return; }
        if (!/^\d{4}-\d{2}-\d{2}$/.test(r.log_date)) { errors.push(`Row ${r.rowNum}: Date "${r.log_date}" isn't in YYYY-MM-DD format.`); return; }
        validRows.push(r);
      });

      // Hours can only go up over time - check each asset's rows are
      // non-decreasing in chronological order, both against each other in
      // this upload and against whatever's already on record.
      const byAsset = {};
      validRows.forEach((r) => { (byAsset[r.asset_id] ||= []).push(r); });
      Object.entries(byAsset).forEach(([assetId, rows]) => {
        rows.sort((a, b) => a.log_date.localeCompare(b.log_date) || a.shift.localeCompare(b.shift));
        const priorExisting = dailyHours
          .filter((d) => d.asset_id === assetId && d.log_date < rows[0].log_date)
          .sort((a, b) => b.log_date.localeCompare(a.log_date))[0];
        let lastKnown = priorExisting?.closing_hours ?? null;
        rows.forEach((r) => {
          if (lastKnown != null && r.closing_hours < lastKnown) {
            errors.push(`Row ${r.rowNum}: ${assetId} on ${r.log_date} (${r.shift}) - ${r.closing_hours} is less than the previous reading of ${lastKnown}. Hours can't go backward.`);
          } else {
            lastKnown = r.closing_hours;
          }
        });
      });

      const cleanRows = validRows.filter((r) => !errors.some((e) => e.startsWith(`Row ${r.rowNum}:`)));

      if (cleanRows.length > 0) {
        const payload = cleanRows.map((r) => ({ asset_id: r.asset_id, log_date: r.log_date, shift: r.shift, closing_hours: r.closing_hours }));
        const { error } = await supabase.from("daily_hours").upsert(payload, { onConflict: "asset_id,log_date,shift" });
        if (error) throw error;
        onRefresh();
      }

      if (errors.length === 0) {
        setBulkMessage({ type: "success", text: `Imported ${cleanRows.length} row${cleanRows.length === 1 ? "" : "s"} successfully. Opening Hours and running totals are calculated automatically.` });
      } else {
        setBulkMessage({
          type: cleanRows.length > 0 ? "warning" : "error",
          text: `Imported ${cleanRows.length} row${cleanRows.length === 1 ? "" : "s"}, skipped ${errors.length} with errors:\n${errors.slice(0, 15).join("\n")}${errors.length > 15 ? `\n…and ${errors.length - 15} more.` : ""}`,
        });
      }
    } catch (err) {
      setBulkMessage({ type: "error", text: err.message || String(err) });
    } finally {
      setBulkImporting(false);
      if (bulkFileInputRef.current) bulkFileInputRef.current.value = "";
    }
  };

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = columns.map((c) => c.label);
    const dataRows = filtered.map((row) => columns.map((c) => row[c.key] ?? ""));
    const aoa = [["Daily Hours"], [`Exported: ${timestamp}`], [], headerRow, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: columns.length - 1 } },
    ];
    ws["!cols"] = columns.map((c) => ({ wch: Math.max(c.label.length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Daily Hours");
    XLSX.writeFile(wb, `Daily_Hours_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div>
      <h4 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px", color: NAVY }}>Monthly Scheduled Hours</h4>
      {budgetLoading ? (
        <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
      ) : (
        <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "hidden", marginBottom: 20 }}>
          {budget.map((row, i) => (
            <div key={row.fleet} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none" }}>
              <span style={{ flex: 1, fontSize: 13.5, fontWeight: 600 }}>{row.fleet}</span>
              <span style={{ fontSize: 13 }}>{Number(row.actual_hours).toFixed(1)} / {row.scheduled_hours != null ? Number(row.scheduled_hours).toFixed(0) : "-"} hrs</span>
              <Badge value={row.status} />
            </div>
          ))}
          {budget.length === 0 && <p style={{ padding: 14, fontSize: 13, color: "#859195" }}>No fleets yet.</p>}
        </div>
      )}
      <div style={{ marginBottom: 20 }}>
        <button onClick={() => setShowBudgetForm(true)} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "7px 14px", borderRadius: 8, cursor: "pointer" }}>
          Set Monthly Scheduled Hours
        </button>
      </div>

      <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, background: "#F7F8F6", padding: "14px 16px", marginBottom: 24 }}>
        <h4 style={{ fontSize: 13.5, fontWeight: 700, margin: "0 0 4px", color: NAVY }}>Bulk Upload - a faster way to load a month of hours</h4>
        <p style={{ fontSize: 12, color: "#4B5659", margin: "0 0 12px" }}>
          Download the template for a month, type in Closing Hours in Excel, then upload it back. Opening Hours and running totals are calculated automatically - you only ever need to fill in the reading at the end of each shift. Existing entries are pre-filled in the template so you can see gaps at a glance.
        </p>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="month"
            value={templateMonth}
            onChange={(e) => setTemplateMonth(e.target.value)}
            style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' }}
          />
          <button onClick={downloadBulkTemplate} disabled={assets.length === 0} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: assets.length === 0 ? "default" : "pointer", opacity: assets.length === 0 ? 0.5 : 1 }}>
            <Download size={14} /> Download Template
          </button>
          <button onClick={() => bulkFileInputRef.current?.click()} disabled={bulkImporting} style={{ display: "flex", alignItems: "center", gap: 6, background: NAVY, border: "none", color: "#fff", fontSize: 13, fontWeight: 700, padding: "8px 14px", borderRadius: 8, cursor: bulkImporting ? "default" : "pointer", opacity: bulkImporting ? 0.6 : 1 }}>
            <Upload size={14} /> {bulkImporting ? "Uploading…" : "Upload Filled Template"}
          </button>
          <input ref={bulkFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleBulkFileSelected} style={{ display: "none" }} />
        </div>
        {bulkMessage && (
          <div style={{
            marginTop: 12, padding: "10px 12px", borderRadius: 8, fontSize: 12.5, whiteSpace: "pre-line",
            background: bulkMessage.type === "error" ? "#F6E2E0" : bulkMessage.type === "warning" ? "#F5E9D8" : "#E2EFE9",
            border: `1px solid ${bulkMessage.type === "error" ? "#DDB6B2" : bulkMessage.type === "warning" ? "#E3C79B" : "#B7D89A"}`,
            color: bulkMessage.type === "error" ? "#7A3330" : bulkMessage.type === "warning" ? "#7A5320" : "#2C5646",
          }}>
            {bulkMessage.text}
          </div>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search daily hours"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
            <Download size={14} /> Export to Excel
          </button>
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Log Daily Hours
          </button>
        </div>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => <th key={c.key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{c.label}</th>)}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => {
              const prevDate = i > 0 ? filtered[i - 1].log_date : null;
              const showHeader = row.log_date !== prevDate;
              return (
                <React.Fragment key={row.id ?? i}>
                  {showHeader && (
                    <tr>
                      <td colSpan={columns.length + 1} style={{ padding: "8px 12px", background: "#F2F1EA", fontWeight: 700, fontSize: 12.5, color: NAVY, borderTop: i > 0 ? "1px solid #E2E6E3" : "none" }}>
                        {row.log_date ? new Date(row.log_date + "T12:00:00").toLocaleDateString("en-ZA", { weekday: "long", day: "2-digit", month: "short", year: "numeric" }) : "No date"}
                      </td>
                    </tr>
                  )}
                  <tr style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none", background: row.exceeds_shift_limit ? "#F6E2E0" : "transparent" }}>
                    {columns.map((c) => (
                      <td key={c.key} onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                    ))}
                    <td style={{ padding: "9px 12px" }}>
                      <button onClick={() => setDeleting(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                        <Trash2 size={15} />
                      </button>
                    </td>
                  </tr>
                </React.Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {dailyHours.length === 0 ? "No hours logged yet." : "No entries match your search."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <DailyHoursForm assets={assets} dailyHours={dailyHours} existing={editing}
          onClose={() => { setShowForm(false); setEditing(null); }} onSaved={handleSaved} />
      )}
      {deleting && (
        <DeleteConfirmModal
          itemLabel={`${deleting.asset_id} - ${deleting.shift} shift on ${deleting.log_date}`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
      {showBudgetForm && (
        <ScheduledHoursForm fleets={fleets} onClose={() => setShowBudgetForm(false)} onSaved={handleBudgetSaved} />
      )}
    </div>
  );
}


function BreakdownForm({ assets, existing, activatingWorkOrder, onClose, onSaved, userEmail, myFullName, workOrders, parts, componentCodes, onRefresh }) {
  const [assetId, setAssetId] = useState(existing?.asset_id || activatingWorkOrder?.asset_id || assets[0]?.asset_id || "");
  const [causeCode, setCauseCode] = useState(existing?.cause_code || CAUSE_CODES[0]);
  const [severity, setSeverity] = useState(existing?.severity || "Medium");
  const [componentAffected, setComponentAffected] = useState(existing?.component_affected || activatingWorkOrder?.component || "");
  const [componentCode, setComponentCode] = useState(existing?.component_affected || activatingWorkOrder?.component || "");
  const [description, setDescription] = useState(existing?.description || activatingWorkOrder?.problem_scope || "");
  const [status, setStatus] = useState(existing?.repair_status || "In Progress");
  const [downtimeStart, setDowntimeStart] = useState(
    existing?.downtime_start ? isoToInputValue(existing.downtime_start) : nowForInput()
  );
  const [downtimeEnd, setDowntimeEnd] = useState(
    existing?.downtime_end ? isoToInputValue(existing.downtime_end) : ""
  );
  const [expectedUpTime, setExpectedUpTime] = useState(
    existing?.expected_up_time ? isoToInputValue(existing.expected_up_time) : ""
  );
  const [reportedBy] = useState(existing?.reported_by || myFullName || userEmail || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  // Once a brand-new event has been saved once, we keep the modal open
  // (rather than closing it) so a Work Order - and parts against it -
  // can be added immediately, without having to close and reopen. From
  // that point on this form behaves like it's editing savedRecord.
  const [savedRecord, setSavedRecord] = useState(existing || null);
  const currentId = savedRecord?.id;
  const hasSavedRecord = !!savedRecord;

  // Downtime End only makes sense once the breakdown is actually Closed -
  // an "in progress" event by definition doesn't have an end yet. Auto-fills
  // when closing, and auto-clears if status is changed back to active,
  // rather than leaving a stale end time sitting alongside an open status.
  useEffect(() => {
    if (status === "Closed" && !downtimeEnd) {
      setDowntimeEnd(nowForInput());
    } else if (status !== "Closed" && downtimeEnd) {
      setDowntimeEnd("");
    }
  }, [status]);

  useEffect(() => {
    if (status === "Closed" && expectedUpTime) {
      setExpectedUpTime("");
    }
  }, [status]); // eslint-disable-line react-hooks/exhaustive-deps

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab before logging a breakdown.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            Close
          </button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");

    if (!description.trim()) {
      setError("Description is required.");
      return;
    }
    if (new Date(downtimeStart) > new Date()) {
      setError("Downtime Start can't be in the future - only past or present times are allowed.");
      return;
    }
    if (downtimeEnd && new Date(downtimeEnd) > new Date()) {
      setError("Downtime End can't be in the future.");
      return;
    }
    if (downtimeEnd && new Date(downtimeEnd) < new Date(downtimeStart)) {
      setError("Downtime End can't be before Downtime Start.");
      return;
    }

    // Only relevant when this breakdown will be active after saving - a
    // breakdown being closed out isn't creating a new overlap.
    if (status !== "Closed") {
      const { conflict, error: checkError } = await checkAssetOverlap(assetId, "breakdown_log", currentId);
      if (checkError) {
        setError(`Couldn't verify this asset is free: ${checkError}`);
        return;
      }
      if (conflict) {
        setError(`${assetId} already has ${conflict}. Close that out first - overlapping active events on the same machine would double-count downtime in Availability, MTBF and MTTR.`);
        return;
      }
    }

    // The event can't be Closed while any Work Order linked to it is
    // still open - every linked Work Order has to be Closed first.
    if (status === "Closed" && currentId) {
      const { data: openWos, error: woCheckError } = await supabase
        .from("work_orders")
        .select("wo_no, status")
        .eq("event_id", currentId)
        .neq("status", "Closed");
      if (woCheckError) {
        setError(`Couldn't verify linked Work Orders: ${woCheckError.message}`);
        return;
      }
      if (openWos && openWos.length > 0) {
        const list = openWos.map((w) => `${w.wo_no} (${w.status})`).join(", ");
        setError(`This event can't be Closed while linked Work Orders are still open: ${list}. Close those first.`);
        return;
      }
    }

    setSaving(true);
    const startDate = new Date(downtimeStart);
    const pad = (n) => String(n).padStart(2, "0");
    const breakdownDate = `${startDate.getFullYear()}-${pad(startDate.getMonth() + 1)}-${pad(startDate.getDate())}`;

    const payload = {
      asset_id: assetId,
      breakdown_date: breakdownDate,
      time_reported: `${pad(startDate.getHours())}:${pad(startDate.getMinutes())}`,
      description,
      cause_code: causeCode,
      severity,
      component_affected: componentAffected,
      repair_status: status,
      downtime_start: startDate.toISOString(),
      downtime_end: downtimeEnd ? new Date(downtimeEnd).toISOString() : null,
      expected_up_time: expectedUpTime ? new Date(expectedUpTime).toISOString() : null,
      // wo_reference is intentionally omitted on new entries - the database
      // auto-generates it (format FT-MM-NNNNN) via a trigger. On edits, we
      // simply don't touch it, so the original number is preserved.
      reported_by: reportedBy || null,
    };

    try {
      const eventSummaryParts = [description];
      if (componentAffected) eventSummaryParts.push(`Component: ${componentAffected}`);
      const eventSummary = eventSummaryParts.join(" - ");

      if (hasSavedRecord) {
        const { error: dbError } = await supabase.from("breakdown_log").update(payload).eq("id", currentId);
        if (dbError) throw dbError;
        setSavedRecord({ ...savedRecord, ...payload });
        logActivity("Events", currentId, status === "Closed" ? "closed" : "updated", eventSummary);
      } else {
        const { data, error: dbError } = await supabase.from("breakdown_log").insert(payload).select().single();
        if (dbError) throw dbError;
        // Stay open on the just-created record instead of closing - this
        // is what lets a Work Order (and parts against it) be added right
        // away, in the same session, instead of having to reopen it.
        setSavedRecord(data);
        logActivity("Events", data.id, "created", eventSummary);

        // This event was created by booking down a scheduled Planned
        // Maintenance job, not logged from scratch - link that same Work
        // Order to the new event and stamp its actual start, so what was
        // a placeholder row in the Events list becomes this real one,
        // with the Work Order now showing up under Linked Work Orders.
        if (activatingWorkOrder) {
          const { error: linkErr } = await supabase.from("work_orders")
            .update({
              event_id: data.id,
              actual_start: startDate.toISOString(),
              status: (activatingWorkOrder.status === "Open" || activatingWorkOrder.status === "Planned") ? "In Progress" : activatingWorkOrder.status,
            })
            .eq("id", activatingWorkOrder.id);
          if (linkErr) throw linkErr;
        }
      }

      onRefresh?.();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 560, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>
          {hasSavedRecord ? "Edit Event" : activatingWorkOrder ? "Book Machine Down" : "Log Event"}
        </h3>

        {activatingWorkOrder && !hasSavedRecord && (
          <p style={{ fontSize: 12.5, color: "#7A5320", background: "#F3E4C8", border: "1px solid #E8A33D", borderRadius: 8, padding: "8px 10px", margin: "0 0 16px" }}>
            This was scheduled as {activatingWorkOrder.wo_no}. Confirm the details below to book it down as a real event - the Work Order stays linked, now with its actual start time.
          </p>
        )}

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required style={fieldStyle}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Downtime</label>
          <DateTimeField value={downtimeStart} onChange={setDowntimeStart} max={nowForInput()} required />
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Uptime {status !== "Closed" && <span style={{ fontWeight: 400, color: "#859195" }}>(set Status to Closed first)</span>}</label>
          <DateTimeField value={downtimeEnd} onChange={setDowntimeEnd} max={nowForInput()} disabled={status !== "Closed"} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Expected Up Time {status === "Closed" && <span style={{ fontWeight: 400, color: "#859195" }}>(not needed once Closed)</span>}</label>
          <DateTimeField value={expectedUpTime} onChange={setExpectedUpTime} disabled={status === "Closed"} />
        </div>
        <p style={{ fontSize: 11.5, color: "#859195", margin: "-6px 0 14px" }}>
          Downtime defaults to right now, but you can set it earlier if you're logging something that happened previously (e.g. catching up after time away from the system) - future dates and times aren't allowed. Uptime only opens up once Status is Closed - an event that's still active doesn't have an end yet, and having both set at once would throw off MTBF, MTTR, Availability and Utilisation. Expected Up Time is your best estimate of when the machine will be back - unlike the other two, future dates are fine here, since that's the whole point of it.
        </p>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Cause</label>
            <select value={causeCode} onChange={(e) => setCauseCode(e.target.value)} style={fieldStyle}>
              {CAUSE_CODES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Severity</label>
            <select value={severity} onChange={(e) => setSeverity(e.target.value)} style={fieldStyle}>
              {SEVERITIES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Component Code</label>
          <select
            value={componentCode}
            onChange={(e) => { setComponentCode(e.target.value); if (e.target.value) setComponentAffected(e.target.value); }}
            style={fieldStyle}
          >
            <option value="">- Select a component -</option>
            {componentCodes.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}
          </select>
          <p style={{ fontSize: 11, color: "#859195", margin: "4px 0 0" }}>Selecting one fills in Component / System Affected below - you can still edit it after.</p>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Component / System Affected</label>
          <input type="text" value={componentAffected} onChange={(e) => setComponentAffected(e.target.value)} placeholder="e.g. Starter Motor" style={fieldStyle} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Description *</label>
          <textarea value={description} onChange={(e) => setDescription(e.target.value)} rows={2} required style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Status</label>
            <select value={status} onChange={(e) => setStatus(e.target.value)} style={fieldStyle}>
              {REPAIR_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Work Order #</label>
            <input type="text" value={savedRecord?.wo_reference || "Generated automatically on save"} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
          </div>
        </div>

        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Reported By</label>
          <input type="text" value={reportedBy} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
          <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
            {existing ? "Original reporter - kept as-is when editing." : "Automatically captured from your signed-in account."}
          </p>
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={() => { if (hasSavedRecord) onSaved(); else onClose(); }} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>
            {hasSavedRecord ? "Close" : "Cancel"}
          </button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : hasSavedRecord ? "Save Changes" : activatingWorkOrder ? "Book Machine Down" : "Log Event"}
          </button>
        </div>

        {!hasSavedRecord && (
          <p style={{ fontSize: 11.5, color: "#859195", margin: "10px 0 0", textAlign: "center" }}>
            Save the event first - Work Orders and parts can be added right here once it's saved.
          </p>
        )}

        {hasSavedRecord && (
          <EventWorkOrdersPanel event={savedRecord} assets={assets} parts={parts} onRefresh={onRefresh} />
        )}
      </form>
    </div>
  );
}

function BreakdownsPage({ assets, breakdowns, onRefresh, userEmail, myFullName, workOrders, parts, componentCodes }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [activating, setActivating] = useState(null); // scheduled Work Order being booked down
  const [query, setQuery] = useState("");
  const [deleting, setDeleting] = useState(null); // row pending delete confirmation
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [sortKey, setSortKey] = useState("downtime_start");
  const [sortDir, setSortDir] = useState("desc");

  const columns = [
    ["asset_id", "Equipment #"], ["component_affected", "Event"], ["cause_code", "Reason"],
    ["description", "Description"], ["wo_reference", "WO-number"],
    ["downtime_start", "Downtime Start"], ["downtime_end", "Downtime End"],
    ["downtime_hours", "Downtime Hours"], ["repair_status", "Status"],
  ].map(([key, label]) => ({ key, label }));

  const formatDateTime = (iso) => iso
    ? new Date(iso).toLocaleString("en-ZA", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).replace(",", "")
    : "";

  // Same comparator drives every column - clicking Status groups equal
  // values together as a side effect of sorting, same as any other
  // column; no special-case grouping logic needed beyond that.
  const sortValue = (row, key) => {
    if (key === "downtime_start" || key === "downtime_end") {
      return row[key] ? new Date(row[key]).getTime() : Infinity; // still-open events sort to the end
    }
    if (key === "downtime_hours") return Number(row[key] ?? 0);
    return String(row[key] ?? "").toLowerCase();
  };

  const toggleSort = (key) => {
    if (key === sortKey) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("asc"); }
  };

  // Planned Maintenance Work Orders that have a planned date but haven't
  // been booked down yet - shown as placeholder rows so they're visible
  // in the Events list ahead of time, same idea as the Timeline markers.
  // Shaped with the same field names as a real breakdown_log row so all
  // the existing column/sort/render logic below just works unmodified.
  const scheduledRows = useMemo(() => workOrders
    .filter((w) => w.work_type === "Preventive" && !w.actual_start && w.planned_start && !w.event_id)
    .map((w) => ({
      id: `sched-${w.id}`,
      isScheduledPlaceholder: true,
      sourceWorkOrder: w,
      asset_id: w.asset_id,
      component_affected: w.component || "Planned Service",
      cause_code: null,
      description: w.problem_scope || "",
      wo_reference: w.wo_no,
      downtime_start: w.planned_start,
      downtime_end: null,
      downtime_hours: null,
      repair_status: "Planned",
    })), [workOrders]);

  const allRows = useMemo(() => [...breakdowns, ...scheduledRows], [breakdowns, scheduledRows]);

  const filtered = useMemo(() => {
    let rows = allRows;
    if (selectedAsset) {
      rows = rows.filter((r) => r.asset_id === selectedAsset);
    } else if (selectedFleet) {
      rows = rows.filter((r) => {
        const asset = assets.find((a) => a.asset_id === r.asset_id);
        return asset && asset.fleet === selectedFleet;
      });
    }
    if (query.trim()) {
      const q = query.toLowerCase();
      rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return [...rows].sort((a, b) => {
      const va = sortValue(a, sortKey), vb = sortValue(b, sortKey);
      if (va < vb) return sortDir === "asc" ? -1 : 1;
      if (va > vb) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
  }, [allRows, assets, query, selectedFleet, selectedAsset, sortKey, sortDir]);

  const handleSaved = () => {
    setShowForm(false);
    setEditing(null);
    setActivating(null);
    onRefresh();
  };

  const handleDelete = async (reason) => {
    await deleteWithReason("breakdown_log", deleting.id, "id", reason, userEmail);
    onRefresh();
    setDeleting(null);
  };

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = columns.map((c) => c.label);
    const dataRows = filtered.map((row) => columns.map((c) =>
      c.key === "downtime_start" ? formatDateTime(row.downtime_start) + (row.isScheduledPlaceholder ? " (planned)" : "")
      : c.key === "downtime_end" ? (formatDateTime(row.downtime_end) || (row.isScheduledPlaceholder ? "Not booked down" : "Still down"))
      : (row[c.key] ?? "")
    ));
    const aoa = [["Events"], [`Exported: ${timestamp}`], [], headerRow, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: columns.length - 1 } },
    ];
    ws["!cols"] = columns.map((c) => ({ wch: Math.max(c.label.length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Events");
    XLSX.writeFile(wb, `Events_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  const EVENT_FIELDS = [
    { key: "downtime_start", header: "Downtime Start (YYYY-MM-DD HH:MM)", type: "datetime" },
    { key: "downtime_end", header: "Downtime End (YYYY-MM-DD HH:MM, blank if still open)", type: "datetime" },
    { key: "cause_code", header: "Cause", type: "text" },
    { key: "severity", header: "Severity", type: "text" },
    { key: "component_affected", header: "Component", type: "text" },
    { key: "description", header: "Description", type: "text" },
    { key: "repair_status", header: "Status", type: "text" },
    { key: "expected_up_time", header: "Expected Up Time (YYYY-MM-DD HH:MM)", type: "datetime" },
  ];

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search events"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }}
          />
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button
            onClick={exportToExcel}
            style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}
          >
            <Download size={14} /> Export to Excel
          </button>
          <button
            onClick={() => { setEditing(null); setActivating(null); setShowForm(true); }}
            style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}
          >
            + Log Event
          </button>
        </div>
      </div>

      <details style={{ marginBottom: 16 }}>
        <summary style={{ cursor: "pointer", fontSize: 12.5, fontWeight: 600, color: NAVY, marginBottom: 8 }}>Offline backup (Excel)</summary>
        <div style={{ marginTop: 8 }}>
          <ExcelSync
            data={breakdowns} assets={assets} fields={EVENT_FIELDS} tableName="breakdown_log"
            sheetTitle="Events Backup" filenamePrefix="Events_Backup" onRefresh={onRefresh}
            extraOnSave={(obj) => {
              const extra = { reported_by: obj.reported_by || userEmail || null };
              if (obj.downtime_start) {
                const d = new Date(obj.downtime_start);
                const pad = (n) => String(n).padStart(2, "0");
                extra.breakdown_date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
                extra.time_reported = `${pad(d.getHours())}:${pad(d.getMinutes())}`;
              }
              return extra;
            }}
          />
        </div>
      </details>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => (
                <th
                  key={c.key}
                  onClick={() => toggleSort(c.key)}
                  title="Click to sort"
                  style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3", cursor: "pointer", userSelect: "none" }}
                >
                  {c.label}
                  <span style={{ display: "inline-block", width: 12, textAlign: "center", color: sortKey === c.key ? NAVY : "#C7C5BB" }}>
                    {sortKey === c.key ? (sortDir === "asc" ? "▲" : "▼") : "↕"}
                  </span>
                </th>
              ))}
              <th style={{ padding: "9px 12px", borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => {
              const isRepeat = (row.repeat_count ?? 0) >= 2;
              const openRow = () => {
                if (row.isScheduledPlaceholder) { setActivating(row.sourceWorkOrder); setShowForm(true); }
                else { setEditing(row); setShowForm(true); }
              };
              return (
                <tr
                  key={row.id ?? i}
                  style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none", background: row.isScheduledPlaceholder ? "#FFF8EB" : isRepeat ? "#F5E9D8" : "transparent" }}
                >
                  {columns.map((c) => (
                    <td key={c.key} onClick={openRow} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>
                      {c.key === "repair_status" ? <Badge value={row[c.key]} />
                        : c.key === "downtime_start" ? (row.isScheduledPlaceholder ? `${formatDateTime(row.downtime_start)} (planned)` : formatDateTime(row.downtime_start))
                        : c.key === "downtime_end" ? (formatDateTime(row.downtime_end) || <span style={{ color: "#B4B2A9" }}>{row.isScheduledPlaceholder ? "Not booked down" : "Still down"}</span>)
                        : c.key === "component_affected" && isRepeat ? (
                          <span title={`${row.repeat_count} breakdowns on this component for this machine`} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            {row[c.key] || <span style={{ color: "#B4B2A9" }}>-</span>}
                            <span style={{ fontSize: 10.5, fontWeight: 700, color: "#7A5320", background: "#FCE2B8", padding: "1px 6px", borderRadius: 6 }}>REPEAT ×{row.repeat_count}</span>
                          </span>
                        ) : typeof row[c.key] === "number" && !Number.isInteger(row[c.key]) ? row[c.key].toFixed(2)
                        : (row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>)}
                    </td>
                  ))}
                  <td style={{ padding: "9px 12px" }}>
                    {row.isScheduledPlaceholder ? (
                      <span title="Scheduled from Planned Maintenance - manage or remove it from that tab" style={{ fontSize: 10.5, fontWeight: 700, color: "#C58A32", background: "#F3E4C8", padding: "3px 8px", borderRadius: 6, whiteSpace: "nowrap" }}>
                        SCHEDULED
                      </span>
                    ) : (
                      <button onClick={() => setDeleting(row)} title="Delete this entry" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "flex" }}>
                        <Trash2 size={15} />
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {allRows.length === 0 ? "No events logged yet." : "No events match your search."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <BreakdownForm
          assets={assets}
          existing={editing}
          activatingWorkOrder={activating}
          onClose={() => { setShowForm(false); setEditing(null); setActivating(null); }}
          onSaved={handleSaved}
          userEmail={userEmail}
          myFullName={myFullName}
          workOrders={workOrders}
          parts={parts}
          componentCodes={componentCodes}
          onRefresh={onRefresh}
        />
      )}

      {deleting && (
        <DeleteConfirmModal
          itemLabel={`event on ${deleting.asset_id} (${deleting.component_affected || "no component set"}, ${deleting.breakdown_date})`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

const WORK_ORDER_TYPES = ["Corrective", "Preventive"];
const WORK_ORDER_PRIORITIES = ["Low", "Medium", "High", "Critical"];
const WORK_ORDER_STATUSES = ["Open", "Planned", "In Progress", "Awaiting Parts", "Closed"];

function WorkOrderForm({ assets, existing, defaultWorkType, defaultAssetId, eventId, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || defaultAssetId || assets[0]?.asset_id || "");
  const [workType, setWorkType] = useState(existing?.work_type || defaultWorkType || "Corrective");
  const [priority, setPriority] = useState(existing?.priority || "Medium");
  const [problemScope, setProblemScope] = useState(existing?.problem_scope || "");
  const [component, setComponent] = useState(existing?.component || "");
  const [status, setStatus] = useState(existing?.status || "Open");
  const [plannedStart, setPlannedStart] = useState(existing?.planned_start || "");
  const [actualStart, setActualStart] = useState(existing?.actual_start ? isoToInputValue(existing.actual_start) : "");
  const [actualFinish, setActualFinish] = useState(existing?.actual_finish ? isoToInputValue(existing.actual_finish) : "");
  const [technicianVendor, setTechnicianVendor] = useState(existing?.technician_vendor || "");
  const [closeoutNotes, setCloseoutNotes] = useState(existing?.closeout_notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  // Same principle as Downtime End on breakdowns: Actual Finish only makes
  // sense once the work order is Closed - auto-clears if status moves back
  // to active, so it can't sit alongside an "in progress" status.
  useEffect(() => {
    if (status !== "Closed" && actualFinish) {
      setActualFinish("");
    }
  }, [status]);

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");

    if (actualStart && new Date(actualStart) > new Date()) {
      setError("Actual Start can't be in the future - only Planned Start can be scheduled ahead.");
      return;
    }
    if (actualFinish && new Date(actualFinish) > new Date()) {
      setError("Actual Finish can't be in the future.");
      return;
    }
    if (actualFinish && actualStart && new Date(actualFinish) < new Date(actualStart)) {
      setError("Actual Finish can't be before Actual Start.");
      return;
    }

    // Only relevant when this work order will be actively "in progress"
    // after saving - actual_start set, no actual_finish yet, and not
    // already closed.
    if (actualStart && !actualFinish && status !== "Closed") {
      const { conflict, error: checkError } = await checkAssetOverlap(assetId, "work_orders", existing?.id);
      if (checkError) {
        setError(`Couldn't verify this asset is free: ${checkError}`);
        return;
      }
      if (conflict) {
        setError(`${assetId} already has ${conflict}. Close that out first - overlapping active events on the same machine would double-count downtime in Availability, MTBF and MTTR.`);
        return;
      }
    }

    setSaving(true);
    const payload = {
      asset_id: assetId, work_type: workType, priority, problem_scope: problemScope,
      component: component || null, status,
      planned_start: plannedStart || null,
      actual_start: actualStart ? new Date(actualStart).toISOString() : null,
      actual_finish: actualFinish ? new Date(actualFinish).toISOString() : null,
      technician_vendor: technicianVendor || null,
      closeout_notes: closeoutNotes || null,
      ...(isEdit ? {} : { request_date: todayForInput(), event_id: eventId || null }),
    };

    try {
      const woSummaryParts = [problemScope];
      if (component) woSummaryParts.push(`Component: ${component}`);
      const woSummary = woSummaryParts.join(" - ");

      const { data, error: dbError } = isEdit
        ? await supabase.from("work_orders").update(payload).eq("id", existing.id).select().single()
        : await supabase.from("work_orders").insert(payload).select().single();
      if (dbError) throw dbError;
      logActivity("Work Orders", (isEdit ? existing.id : data?.id), status === "Closed" ? "closed" : (isEdit ? "updated" : "created"), woSummary);
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 480, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>
          {isEdit ? "Edit Work Order" : workType === "Preventive" ? "Add Planned Maintenance" : "Add Work Order"}
        </h3>
        {isEdit && existing?.wo_no && (
          <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 16px" }}>{existing.wo_no}</p>
        )}
        {!isEdit && <div style={{ marginBottom: 16 }} />}

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit || !!eventId} style={{ ...fieldStyle, ...((isEdit || eventId) ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Type</label>
            <select value={workType} onChange={(e) => setWorkType(e.target.value)} style={fieldStyle}>
              {WORK_ORDER_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Priority</label>
            <select value={priority} onChange={(e) => setPriority(e.target.value)} style={fieldStyle}>
              {WORK_ORDER_PRIORITIES.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Problem / Scope of Work</label>
          <textarea value={problemScope} onChange={(e) => setProblemScope(e.target.value)} rows={2} required style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Component / System</label>
          <input type="text" value={component} onChange={(e) => setComponent(e.target.value)} style={fieldStyle} />
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Status</label>
            <select value={status} onChange={(e) => setStatus(e.target.value)} style={fieldStyle}>
              {WORK_ORDER_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Assigned Technician</label>
            <input type="text" value={technicianVendor} onChange={(e) => setTechnicianVendor(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Planned Start <span style={{ fontWeight: 400, color: "#859195" }}>(this is where you schedule ahead - future dates are fine here)</span></label>
          <input type="date" value={plannedStart} onChange={(e) => setPlannedStart(e.target.value)} style={fieldStyle} />
        </div>

        {isEdit && (
          <>
            <p style={{ fontSize: 11.5, color: "#859195", margin: "0 0 8px" }}>
              The two fields below are for logging when the work actually happened - leave them blank until it's done. They only accept past or present times, not future ones. Actual Finish only opens up once Status is Closed.
            </p>
            <div style={{ marginBottom: 12 }}>
              <label style={labelStyle}>Actual Start</label>
              <DateTimeField value={actualStart} onChange={setActualStart} max={nowForInput()} />
            </div>
            <div style={{ marginBottom: 12 }}>
              <label style={labelStyle}>Actual Finish {status !== "Closed" && <span style={{ fontWeight: 400, color: "#859195" }}>(set Status to Closed first)</span>}</label>
              <DateTimeField value={actualFinish} onChange={setActualFinish} max={nowForInput()} disabled={status !== "Closed"} />
            </div>
          </>
        )}

        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Close-out Notes</label>
          <textarea value={closeoutNotes} onChange={(e) => setCloseoutNotes(e.target.value)} rows={2} style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Create"}
          </button>
        </div>
      </form>
    </div>
  );
}

function PartUsedForm({ parts, workOrder, onClose, onSaved }) {
  const [partId, setPartId] = useState(parts[0]?.id || "");
  const [qty, setQty] = useState(1);
  const [notes, setNotes] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const selectedPart = parts.find((p) => String(p.id) === String(partId));

  if (parts.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 70, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No parts in inventory yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add parts on the Parts Inventory tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    const qtyNum = Number(qty);
    if (!qtyNum || qtyNum <= 0) { setError("Enter a quantity greater than 0."); return; }
    if (selectedPart && qtyNum > (selectedPart.qty_in_stock ?? 0)) {
      setError(`Only ${selectedPart.qty_in_stock ?? 0} in stock for ${selectedPart.part_no}.`);
      return;
    }
    setSaving(true);
    try {
      const { error: insertErr } = await supabase.from("parts_used").insert({
        work_order_id: workOrder.id,
        part_id: Number(partId),
        qty_used: qtyNum,
        notes: notes || null,
      });
      if (insertErr) throw insertErr;

      // Stock is reduced immediately (automatic) - a future Stores module
      // is what will let a store manager confirm/adjust the final count;
      // for now "Confirmed" is just a status flag on the parts_used row.
      const { error: stockErr } = await supabase.from("parts_inventory")
        .update({ qty_in_stock: (selectedPart?.qty_in_stock ?? 0) - qtyNum })
        .eq("id", partId);
      if (stockErr) throw stockErr;

      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: "inherit" };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 70, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 400, maxWidth: "100%" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Add Part Used</h3>
        <p style={{ fontSize: 12, color: "#859195", margin: "0 0 16px" }}>Against Work Order {workOrder.wo_no}</p>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Part</label>
          <select value={partId} onChange={(e) => setPartId(e.target.value)} required style={fieldStyle}>
            {parts.map((p) => <option key={p.id} value={p.id}>{p.part_no} - {p.description} ({p.qty_in_stock ?? 0} in stock)</option>)}
          </select>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Quantity</label>
          <input type="number" min="1" step="1" value={qty} onChange={(e) => setQty(e.target.value)} required style={fieldStyle} />
        </div>

        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Notes</label>
          <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows={2} style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : "Add Part"}
          </button>
        </div>
      </form>
    </div>
  );
}

function PartsUsedList({ workOrder, parts }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showAdd, setShowAdd] = useState(false);

  const partById = useMemo(() => Object.fromEntries(parts.map((p) => [p.id, p])), [parts]);

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("parts_used").select("*").eq("work_order_id", workOrder.id).order("created_at", { ascending: false });
    if (!error) setRows(data || []);
    setLoading(false);
  }, [workOrder.id]);

  useEffect(() => { load(); }, [load]);

  const confirmRow = async (row) => {
    await supabase.from("parts_used").update({ status: "Confirmed", confirmed_at: new Date().toISOString() }).eq("id", row.id);
    load();
  };

  return (
    <div style={{ background: "#F9F8F4", borderRadius: 8, padding: "10px 12px", marginTop: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <span style={{ fontSize: 11.5, fontWeight: 700, color: "#4B5659", textTransform: "uppercase", letterSpacing: 0.3 }}>Parts Used</span>
        <button onClick={() => setShowAdd(true)} style={{ background: "none", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 11.5, fontWeight: 600, padding: "3px 10px", borderRadius: 6, cursor: "pointer" }}>
          + Add Part
        </button>
      </div>
      {loading ? (
        <p style={{ fontSize: 12, color: "#859195", margin: 0 }}>Loading…</p>
      ) : rows.length === 0 ? (
        <p style={{ fontSize: 12, color: "#859195", margin: 0 }}>No parts logged against this Work Order yet.</p>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          {rows.map((r) => {
            const p = partById[r.part_id];
            return (
              <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5, background: "#fff", border: "1px solid #E2E6E3", borderRadius: 6, padding: "5px 9px" }}>
                <span>{p ? `${p.part_no} - ${p.description}` : "Part"} × {r.qty_used}</span>
                <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <Badge value={r.status} />
                  {r.status === "Pending" && (
                    <button onClick={() => confirmRow(r)} style={{ background: "none", border: "none", color: "#2C5646", fontSize: 11.5, fontWeight: 600, cursor: "pointer" }}>Confirm</button>
                  )}
                </span>
              </div>
            );
          })}
        </div>
      )}
      {showAdd && (
        <PartUsedForm parts={parts} workOrder={workOrder} onClose={() => setShowAdd(false)} onSaved={() => { setShowAdd(false); load(); }} />
      )}
    </div>
  );
}

function EventWorkOrdersPanel({ event, assets, parts, onRefresh }) {
  const [linked, setLinked] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showWoForm, setShowWoForm] = useState(false);
  const [editingWo, setEditingWo] = useState(null);
  const [expanded, setExpanded] = useState(null); // wo.id currently showing its Parts Used

  // Queried directly against work_orders (not the work_orders_calc view
  // used elsewhere) - event_id is a brand-new column and the calc view
  // predates it, so this is the one place guaranteed to have it without
  // needing that view's definition touched.
  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("work_orders").select("*").eq("event_id", event.id).order("id", { ascending: false });
    if (!error) setLinked(data || []);
    setLoading(false);
  }, [event.id]);

  useEffect(() => { load(); }, [load]);

  const handleWoSaved = () => {
    setShowWoForm(false);
    setEditingWo(null);
    load();
    onRefresh?.();
  };

  return (
    <div style={{ marginTop: 18, paddingTop: 16, borderTop: "1px solid #E2E6E3" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <p style={{ fontSize: 13, fontWeight: 700, color: NAVY, margin: 0 }}>Linked Work Orders</p>
        <button type="button" onClick={() => { setEditingWo(null); setShowWoForm(true); }} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "6px 12px", borderRadius: 8, cursor: "pointer" }}>
          + Add Work Order
        </button>
      </div>

      {loading ? (
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>Loading…</p>
      ) : linked.length === 0 ? (
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>No Work Orders linked to this event yet - each one gets its own work order number, and you can log parts used against it once it's created.</p>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {linked.map((w) => (
            <div key={w.id} style={{ border: "1px solid #E2E6E3", borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }} onClick={() => { setEditingWo(w); setShowWoForm(true); }}>
                  <span style={{ fontSize: 13, fontWeight: 700, color: NAVY }}>{w.wo_no}</span>
                  <span style={{ fontSize: 12.5, color: "#4B5659" }}>{w.component || "No component set"}</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <Badge value={w.priority} />
                  <Badge value={w.status} />
                  <button onClick={() => setExpanded(expanded === w.id ? null : w.id)} style={{ background: "none", border: "none", color: NAVY, cursor: "pointer", padding: 2, display: "inline-flex" }}>
                    <ChevronDown size={15} style={{ transform: expanded === w.id ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                  </button>
                </div>
              </div>
              {expanded === w.id && <PartsUsedList workOrder={w} parts={parts} />}
            </div>
          ))}
        </div>
      )}

      {showWoForm && (
        <WorkOrderForm
          assets={assets}
          existing={editingWo}
          defaultAssetId={event.asset_id}
          eventId={event.id}
          onClose={() => { setShowWoForm(false); setEditingWo(null); }}
          onSaved={handleWoSaved}
        />
      )}
    </div>
  );
}

// Print isolation: only .job-card-print and its children stay visible when
// printing - everything else on the page (sidebar, other content) is
// hidden. window.print() also covers "export" since every browser's print
// dialog offers Save as PDF, so this one button serves both print and export.
function JobCardPrintModal({ workOrder, asset, onClose, onUploaded }) {
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState("");
  const [scans, setScans] = useState([]); // full history, newest first, each with a signed view URL
  const [scansLoading, setScansLoading] = useState(true);
  const fileInputRef = React.useRef(null);

  const loadScans = useCallback(async () => {
    setScansLoading(true);
    const { data, error } = await supabase
      .from("work_order_scans")
      .select("*")
      .eq("work_order_id", workOrder.id)
      .order("uploaded_at", { ascending: false });
    if (error) {
      console.error("Failed to load scan history:", error.message);
      setScans([]);
      setScansLoading(false);
      return;
    }
    const withUrls = await Promise.all(
      (data || []).map(async (row) => {
        const { data: signed } = await supabase.storage
          .from("job-card-scans")
          .createSignedUrl(row.storage_path, 3600);
        return { ...row, url: signed?.signedUrl || null };
      })
    );
    setScans(withUrls);
    setScansLoading(false);
  }, [workOrder.id]);

  useEffect(() => {
    loadScans();
  }, [loadScans]);

  const handleUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    setUploadError("");
    try {
      // Timestamp-prefixed path - every upload gets its own file, nothing
      // is overwritten, so the full history stays available in Storage too.
      const path = `${workOrder.wo_no || workOrder.id}/${Date.now()}-${file.name}`;
      const { error: uploadErr } = await supabase.storage.from("job-card-scans").upload(path, file);
      if (uploadErr) throw uploadErr;
      const uploadedAt = new Date().toISOString();
      const { error: historyErr } = await supabase.from("work_order_scans").insert({
        work_order_id: workOrder.id,
        storage_path: path,
        file_name: file.name,
        uploaded_at: uploadedAt,
      });
      if (historyErr) throw historyErr;
      // Kept in sync purely for convenience - anything elsewhere that reads
      // "the latest scan" straight off work_orders still works; the
      // work_order_scans table above is the source of truth for history.
      await supabase.from("work_orders")
        .update({ job_card_scan_path: path, job_card_scan_uploaded_at: uploadedAt })
        .eq("id", workOrder.id);
      await loadScans();
      onUploaded();
    } catch (err) {
      setUploadError(err.message || String(err));
    } finally {
      setUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const row = (label, value) => (
    <tr>
      <td style={{ padding: "6px 10px", fontWeight: 700, fontSize: 12.5, color: "#183642", width: 160, border: "1px solid #ccc" }}>{label}</td>
      <td style={{ padding: "6px 10px", fontSize: 12.5, color: "#183642", border: "1px solid #ccc" }}>{value || "-"}</td>
    </tr>
  );

  return (
    <div className="job-card-modal-backdrop" style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60, padding: 16 }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          .job-card-print, .job-card-print * { visibility: visible; }
          .job-card-print { position: absolute; top: 0; left: 0; width: 100%; }
          .job-card-modal-backdrop { position: static !important; background: none !important; padding: 0 !important; }
          .job-card-no-print { display: none !important; }
        }
      `}</style>
      <div style={{ background: "#fff", borderRadius: 12, padding: 28, width: 640, maxWidth: "100%", maxHeight: "92vh", overflowY: "auto" }}>
        <div className="job-card-no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <p style={{ fontSize: 15, fontWeight: 700, color: NAVY, margin: 0 }}>Job Card</p>
          <div style={{ display: "flex", gap: 10 }}>
            <button onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "8px 14px", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>Close</button>
            <button onClick={() => window.print()} style={{ display: "flex", alignItems: "center", gap: 6, background: NAVY, border: "none", color: "#fff", padding: "8px 14px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
              <Printer size={14} /> Print / Save as PDF
            </button>
          </div>
        </div>

        <div className="job-card-print">
          <div style={{ textAlign: "center", marginBottom: 16 }}>
            <p style={{ fontSize: 18, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Job Card</p>
            <p style={{ fontSize: 13, color: "#4B5659", margin: 0 }}>{workOrder.wo_no}</p>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 16 }}>
            <tbody>
              {row("Equipment #", asset?.asset_id)}
              {row("Equipment Name", asset?.asset_name)}
              {row("Make / Model", `${asset?.make || ""} ${asset?.model || ""}`.trim())}
              {row("Fleet", asset?.fleet)}
              {row("Current Hours", asset?.current_hours)}
              {row("Work Type", workOrder.work_type)}
              {row("Priority", workOrder.priority)}
              {row("Status", workOrder.status)}
              {row("Requested", workOrder.request_date)}
              {row("Planned Start", workOrder.planned_start)}
              {row("Assigned Technician", workOrder.technician_vendor)}
              {row("Problem / Scope of Work", workOrder.problem_scope)}
              {row("Component / System", workOrder.component)}
            </tbody>
          </table>

          <p style={{ fontSize: 12.5, fontWeight: 700, color: "#183642", margin: "0 0 6px" }}>Work Done</p>
          <div style={{ border: "1px solid #ccc", height: 70, marginBottom: 14 }} />

          <p style={{ fontSize: 12.5, fontWeight: 700, color: "#183642", margin: "0 0 6px" }}>Parts Used</p>
          <div style={{ border: "1px solid #ccc", height: 50, marginBottom: 14 }} />

          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20 }}>
            <tbody>
              <tr>
                <td style={{ padding: "6px 10px", fontSize: 12.5, border: "1px solid #ccc", width: "50%" }}>Hours Worked: _______________</td>
                <td style={{ padding: "6px 10px", fontSize: 12.5, border: "1px solid #ccc" }}>Date Completed: _______________</td>
              </tr>
            </tbody>
          </table>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginTop: 30 }}>
            <div>
              <div style={{ borderTop: "1px solid #183642", paddingTop: 6, fontSize: 12 }}>Technician Signature / Date</div>
            </div>
            <div>
              <div style={{ borderTop: "1px solid #183642", paddingTop: 6, fontSize: 12 }}>Supervisor Signature / Date</div>
            </div>
          </div>
        </div>

        <div className="job-card-no-print" style={{ marginTop: 24, paddingTop: 16, borderTop: "1px solid #E2E6E3" }}>
          <p style={{ fontSize: 13, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>Completed document scans</p>

          {scansLoading ? (
            <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 10px" }}>Loading history…</p>
          ) : scans.length === 0 ? (
            <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 10px" }}>Once the technician has completed and signed this document, scan it and upload it here. Every service done against this Work Order can be scanned and kept here - nothing gets overwritten.</p>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 4 }}>
              {scans.map((s) => (
                <a
                  key={s.id}
                  href={s.url || undefined}
                  target="_blank"
                  rel="noreferrer"
                  style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
                    fontSize: 13, color: s.url ? NAVY : "#B0AEA6",
                    background: "#F7F6F2", border: "1px solid #E2E6E3", borderRadius: 8,
                    padding: "8px 12px", textDecoration: "none",
                    pointerEvents: s.url ? "auto" : "none",
                  }}
                >
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {s.file_name || "Scanned document"}
                  </span>
                  <span style={{ fontSize: 11.5, color: "#859195", flexShrink: 0 }}>
                    {new Date(s.uploaded_at).toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" })}
                  </span>
                </a>
              ))}
            </div>
          )}

          <div style={{ marginTop: 10 }}>
            <button onClick={() => fileInputRef.current?.click()} disabled={uploading} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "7px 14px", borderRadius: 8, cursor: uploading ? "default" : "pointer", opacity: uploading ? 0.6 : 1 }}>
              <Upload size={14} /> {uploading ? "Uploading…" : "Upload completed document"}
            </button>
            <input ref={fileInputRef} type="file" accept="image/*,.pdf" onChange={handleUpload} style={{ display: "none" }} />
          </div>
          {uploadError && <p style={{ color: "#B85450", fontSize: 12.5, margin: "8px 0 0" }}>{uploadError}</p>}
        </div>
      </div>
    </div>
  );
}

function buildDowntimeRows(breakdowns, workOrders, fromDateTime, toDateTime) {
  const from = new Date(fromDateTime);
  const to = new Date(toDateTime);
  const overlaps = (start, end) => {
    if (!start) return false;
    const s = new Date(start);
    // Still open/ongoing - always show it regardless of the selected "to"
    // cutoff. The date picker's "to" is a snapshot from whenever the page
    // loaded and doesn't auto-extend as time passes, but a currently
    // active breakdown or planned job is operationally relevant right
    // now no matter what range someone happened to have selected.
    if (!end) return true;
    const e = new Date(end);
    return s < to && e > from;
  };

  const breakdownRows = breakdowns
    .filter((b) => overlaps(b.downtime_start, b.downtime_end))
    .map((b) => ({
      asset_id: b.asset_id, event: "Breakdown", reason: b.cause_code, description: b.description,
      wo_number: b.wo_reference, downtime_start: b.downtime_start, downtime_end: b.downtime_end,
      downtime_hours: b.downtime_hours, status: b.repair_status,
    }));

  const plannedRows = workOrders
    .filter((w) => w.work_type === "Preventive" && overlaps(w.actual_start, w.actual_finish))
    .map((w) => ({
      asset_id: w.asset_id, event: "Planned", reason: w.component, description: w.problem_scope,
      wo_number: w.wo_no, downtime_start: w.actual_start, downtime_end: w.actual_finish,
      downtime_hours: w.downtime_hours, status: w.status,
    }));

  return [...breakdownRows, ...plannedRows].sort((a, b) => new Date(a.downtime_start) - new Date(b.downtime_start));
}

const DOWNTIME_SUMMARY_COLUMNS = [
  ["asset_id", "Equipment"], ["event", "Event"], ["reason", "Reason"], ["description", "Description"],
  ["wo_number", "WO Number"], ["downtime_start", "Downtime Start"], ["downtime_end", "Downtime End"],
  ["downtime_hours", "Downtime (hrs)"], ["status", "Status"],
];

function FuelLogForm({ assets, existing, userEmail, dailyHours, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [fillDate, setFillDate] = useState(existing?.fill_date || todayForInput());
  const [hourMeter, setHourMeter] = useState(existing?.hour_meter ?? "");
  const [litres, setLitres] = useState(existing?.litres ?? "");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const recordedBy = existing?.recorded_by || userEmail || "";

  const hourMeterWarning = useMemo(
    () => checkHourMeterPlausibility(assetId, hourMeter, fillDate, dailyHours),
    [assetId, hourMeter, fillDate, dailyHours]
  );

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    if (litres === "" || hourMeter === "") { setError("Enter Hour Meter and Litres."); return; }
    setSaving(true);
    const payload = {
      asset_id: assetId, fill_date: fillDate, hour_meter: Number(hourMeter),
      litres: Number(litres), recorded_by: recordedBy || null, notes: notes || null,
    };
    try {
      const { error: dbError } = isEdit
        ? await supabase.from("fuel_log").update(payload).eq("id", existing.id)
        : await supabase.from("fuel_log").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 420, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>{isEdit ? "Edit Fuel Entry" : "Log Fuel"}</h3>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Date</label>
            <input type="date" value={fillDate} onChange={(e) => setFillDate(e.target.value)} required max={todayForInput()} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Hour Meter</label>
            <input type="number" step="0.1" value={hourMeter} onChange={(e) => setHourMeter(e.target.value)} required style={fieldStyle} />
          </div>
        </div>
        {hourMeterWarning && (
          <div style={{ background: "#F5E9D8", border: "1px solid #E3C79B", borderRadius: 8, padding: "8px 10px", marginBottom: 12, fontSize: 12, color: "#7A5320" }}>
            {hourMeterWarning}
          </div>
        )}
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Litres</label>
          <input type="number" step="0.1" value={litres} onChange={(e) => setLitres(e.target.value)} required style={fieldStyle} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Recorded By</label>
          <input type="text" value={recordedBy} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
        </div>
        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Notes</label>
          <input type="text" value={notes} onChange={(e) => setNotes(e.target.value)} style={fieldStyle} />
        </div>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Log Fuel"}
          </button>
        </div>
      </form>
    </div>
  );
}

const INSPECTION_TYPES = ["Pre-Start", "Daily", "Weekly", "Monthly", "Statutory"];
const INSPECTION_RESULTS = ["Pass", "Fail", "Pass with Defects"];
const RISK_RATINGS = ["Low", "Medium", "High", "Critical"];

function InspectionForm({ assets, existing, userEmail, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [logDate, setLogDate] = useState(existing?.log_date || todayForInput());
  const [inspectionType, setInspectionType] = useState(existing?.inspection_type || INSPECTION_TYPES[0]);
  const [hourMeter, setHourMeter] = useState(existing?.hour_meter ?? "");
  const [result, setResult] = useState(existing?.result || "Pass");
  const [defectFinding, setDefectFinding] = useState(existing?.defect_finding || "");
  const [riskRating, setRiskRating] = useState(existing?.risk_rating || "Low");
  const [immediateAction, setImmediateAction] = useState(existing?.immediate_action || "");
  const [nextInspectionDate, setNextInspectionDate] = useState(existing?.next_inspection_date || "");
  const [signedOff, setSignedOff] = useState(existing?.signed_off || false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const inspector = existing?.inspector || userEmail || "";

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSaving(true);
    const payload = {
      asset_id: assetId, log_date: logDate, inspection_type: inspectionType,
      inspector: inspector || null, hour_meter: hourMeter === "" ? null : Number(hourMeter),
      result, defect_finding: defectFinding || null,
      risk_rating: result === "Pass" ? null : riskRating,
      immediate_action: immediateAction || null,
      next_inspection_date: nextInspectionDate || null,
      signed_off: signedOff,
    };
    try {
      const { error: dbError } = isEdit
        ? await supabase.from("inspections").update(payload).eq("id", existing.id)
        : await supabase.from("inspections").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: 'inherit' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 460, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>{isEdit ? "Edit Inspection" : "Log Inspection"}</h3>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Date</label>
            <input type="date" value={logDate} onChange={(e) => setLogDate(e.target.value)} required max={todayForInput()} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Type</label>
            <select value={inspectionType} onChange={(e) => setInspectionType(e.target.value)} style={fieldStyle}>
              {INSPECTION_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Hour Meter</label>
            <input type="number" step="0.1" value={hourMeter} onChange={(e) => setHourMeter(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Result</label>
            <select value={result} onChange={(e) => setResult(e.target.value)} style={fieldStyle}>
              {INSPECTION_RESULTS.map((r) => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
        </div>

        {result !== "Pass" && (
          <>
            <div style={{ marginBottom: 12 }}>
              <label style={labelStyle}>Defect Finding</label>
              <textarea value={defectFinding} onChange={(e) => setDefectFinding(e.target.value)} rows={2} style={{ ...fieldStyle, resize: "vertical" }} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
              <div>
                <label style={labelStyle}>Risk Rating</label>
                <select value={riskRating} onChange={(e) => setRiskRating(e.target.value)} style={fieldStyle}>
                  {RISK_RATINGS.map((r) => <option key={r} value={r}>{r}</option>)}
                </select>
              </div>
              <div>
                <label style={labelStyle}>Immediate Action</label>
                <input type="text" value={immediateAction} onChange={(e) => setImmediateAction(e.target.value)} style={fieldStyle} />
              </div>
            </div>
          </>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Inspector</label>
            <input type="text" value={inspector} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
          </div>
          <div>
            <label style={labelStyle}>Next Inspection Due</label>
            <input type="date" value={nextInspectionDate} onChange={(e) => setNextInspectionDate(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#183642", marginBottom: 18, cursor: "pointer" }}>
          <input type="checkbox" checked={signedOff} onChange={(e) => setSignedOff(e.target.checked)} style={{ width: 16, height: 16 }} />
          Signed off
        </label>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Log Inspection"}
          </button>
        </div>
      </form>
    </div>
  );
}

function InspectionsPage({ assets, inspections, userEmail, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const columns = [
    ["log_date", "Date"], ["asset_id", "Equipment #"], ["inspection_type", "Type"],
    ["inspector", "Inspector"], ["result", "Result"], ["wo_reference", "Work order #"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    let rows = inspections;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    return [...rows].sort((a, b) => (b.log_date || "").localeCompare(a.log_date || ""));
  }, [inspections, assets, query, selectedFleet, selectedAsset]);

  const handleDelete = async (reason) => {
    await deleteWithReason("inspections", deleting.id, "id", reason, userEmail);
    onRefresh();
    setDeleting(null);
  };

  const INSPECTION_FIELDS = [
    { key: "log_date", header: "Date (YYYY-MM-DD)", type: "date" },
    { key: "inspection_type", header: "Inspection Type", type: "text" },
    { key: "result", header: "Result (Pass/Fail)", type: "text" },
    { key: "defect_finding", header: "Defect Finding", type: "text" },
    { key: "risk_rating", header: "Risk Rating", type: "text" },
    { key: "immediate_action", header: "Immediate Action", type: "text" },
    { key: "next_inspection_date", header: "Next Inspection Date (YYYY-MM-DD)", type: "date" },
  ];

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search inspections" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
          + Log Inspection
        </button>
      </div>
      <ExcelSync
        data={inspections} assets={assets} fields={INSPECTION_FIELDS} tableName="inspections"
        sheetTitle="Inspections" filenamePrefix="Inspections" onRefresh={onRefresh}
        extraOnSave={(obj) => ({ inspector: obj.inspector || userEmail || null })}
      />
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => <th key={c.key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{c.label}</th>)}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr key={row.id ?? i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                {columns.map((c) => (
                  <td key={c.key} onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>
                    {c.key === "result" ? <Badge value={row[c.key]} /> : (row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>)}
                  </td>
                ))}
                <td style={{ padding: "9px 12px" }}>
                  <button onClick={() => setDeleting(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {inspections.length === 0 ? "No inspections logged yet." : "No entries match your filters."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      {showForm && (
        <InspectionForm assets={assets} existing={editing} userEmail={userEmail}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); onRefresh(); }} />
      )}
      {deleting && (
        <DeleteConfirmModal
          itemLabel={`inspection on ${deleting.asset_id} (${deleting.log_date})`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

const BACKLOG_SOURCES = [
  { value: "daily_service", label: "Daily Service" },
  { value: "monthly_inspection", label: "Monthly Inspection" },
  { value: "service_card", label: "Service Card" },
];
const backlogSourceLabel = (v) => BACKLOG_SOURCES.find((s) => s.value === v)?.label || v;
const BACKLOG_PRIORITIES = ["Low", "Medium", "High", "Critical"];

function ageDays(dateStr) {
  if (!dateStr) return null;
  const ms = new Date().setHours(0, 0, 0, 0) - new Date(dateStr).setHours(0, 0, 0, 0);
  return Math.max(0, Math.round(ms / 86400000));
}

function BacklogForm({ assets, workOrders, existing, userEmail, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [sourceType, setSourceType] = useState(existing?.source_type || "daily_service");
  const [componentCode, setComponentCode] = useState(existing?.component_code || "");
  const [description, setDescription] = useState(existing?.description || "");
  const [priority, setPriority] = useState(existing?.priority || "Medium");
  const [status, setStatus] = useState(existing?.status || "Open");
  const [dateNotified, setDateNotified] = useState(existing?.date_notified || todayForInput());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const woOptionsForAsset = useMemo(
    () => workOrders.filter((w) => w.asset_id === assetId),
    [workOrders, assetId]
  );

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    if (!description.trim()) { setError("Enter a description."); return; }
    setSaving(true);
    try {
      let woId = existing?.work_order_id || null;

      // Every backlog item gets its own Work Order, auto-numbered the
      // same way as everywhere else in the app - not linked to an
      // existing one, so whoever's fixing it has a number to quote for
      // parts regardless of where the backlog came from. Only created
      // once, on first save - editing an existing item never creates a
      // second one.
      if (!isEdit) {
        const { data: newWo, error: woErr } = await supabase.from("work_orders").insert({
          asset_id: assetId,
          work_type: "Corrective",
          priority,
          problem_scope: description.trim(),
          component: componentCode || null,
          status: "Open",
          request_date: todayForInput(),
        }).select().single();
        if (woErr) throw woErr;
        woId = newWo.id;
      }

      const payload = {
        asset_id: assetId,
        source_type: sourceType,
        work_order_id: woId,
        component_code: componentCode || null,
        description: description.trim(),
        priority,
        status,
        date_notified: dateNotified,
        closed_date: status === "Closed" ? (existing?.closed_date || todayForInput()) : null,
      };

      const { data, error: dbError } = isEdit
        ? await supabase.from("backlogs").update(payload).eq("id", existing.id).select().single()
        : await supabase.from("backlogs").insert(payload).select().single();
      if (dbError) throw dbError;
      const backlogSummaryParts = [description.trim()];
      if (componentCode) backlogSummaryParts.push(`Component: ${componentCode}`);
      logActivity("Backlogs", (isEdit ? existing.id : data?.id), status === "Closed" ? "closed" : (isEdit ? "updated" : "created"), backlogSummaryParts.join(" - "));
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: "inherit" };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 460, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>{isEdit ? "Edit Backlog Item" : "Add Backlog Item"}</h3>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required style={fieldStyle}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Source</label>
            <select value={sourceType} onChange={(e) => setSourceType(e.target.value)} style={fieldStyle}>
              {BACKLOG_SOURCES.map((s) => <option key={s.value} value={s.value}>{s.label}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Date Notified</label>
            <input type="date" value={dateNotified} onChange={(e) => setDateNotified(e.target.value)} required max={todayForInput()} style={fieldStyle} />
          </div>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Work Order #</label>
          <input type="text" value={existing?.work_order_id ? (woOptionsForAsset.find((w) => w.id === existing.work_order_id)?.wo_no || "Assigned") : "Generated automatically on save"} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Component Code</label>
          <input type="text" value={componentCode} onChange={(e) => setComponentCode(e.target.value)} placeholder="e.g. 6305.00.GR.0" style={fieldStyle} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Description</label>
          <textarea value={description} onChange={(e) => setDescription(e.target.value)} rows={3} required style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 18 }}>
          <div>
            <label style={labelStyle}>Priority</label>
            <select value={priority} onChange={(e) => setPriority(e.target.value)} style={fieldStyle}>
              {BACKLOG_PRIORITIES.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Status</label>
            <select value={status} onChange={(e) => setStatus(e.target.value)} style={fieldStyle}>
              <option value="Open">Open</option>
              <option value="Closed">Closed</option>
            </select>
          </div>
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Add Backlog Item"}
          </button>
        </div>
      </form>
    </div>
  );
}

function BacklogsPage({ assets, backlogs, workOrders, userEmail, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");
  const [sourceFilter, setSourceFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("Open");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const woNoById = useMemo(() => Object.fromEntries(workOrders.map((w) => [w.id, w.wo_no])), [workOrders]);

  const filtered = useMemo(() => {
    let rows = backlogs;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (sourceFilter) rows = rows.filter((r) => r.source_type === sourceFilter);
    if (statusFilter) rows = rows.filter((r) => r.status === statusFilter);
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    return [...rows].sort((a, b) => (b.date_notified || "").localeCompare(a.date_notified || ""));
  }, [backlogs, assets, query, sourceFilter, statusFilter, selectedFleet, selectedAsset]);

  const handleDelete = async (reason) => {
    await deleteWithReason("backlogs", deleting.id, "id", reason, userEmail);
    onRefresh();
    setDeleting(null);
  };

  const quickClose = async (row) => {
    await supabase.from("backlogs").update({ status: "Closed", closed_date: todayForInput() }).eq("id", row.id);
    onRefresh();
  };

  const columns = [
    ["asset_id", "Equipment #"], ["source_type", "Source"], ["wo_no", "Work order #"],
    ["component_code", "Component"], ["description", "Description"], ["priority", "Priority"],
    ["date_notified", "Date Notified"], ["age_days", "Age (days)"], ["status", "Status"],
  ];

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = columns.map((c) => c[1]);
    const dataRows = filtered.map((row) => [
      row.asset_id, backlogSourceLabel(row.source_type), row.work_order_id ? (woNoById[row.work_order_id] || "") : "",
      row.component_code ?? "", row.description ?? "", row.priority, row.date_notified, ageDays(row.date_notified), row.status,
    ]);
    const aoa = [["Backlog Management"], [`Exported: ${timestamp}`], [], headerRow, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: columns.length - 1 } }];
    ws["!cols"] = columns.map((c) => ({ wch: Math.max(c[1].length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Backlogs");
    XLSX.writeFile(wb, `Backlog_Management_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ position: "relative", flex: 1, minWidth: 200, maxWidth: 280 }}>
            <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search backlogs" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", boxSizing: "border-box" }} />
          </div>
          <select value={sourceFilter} onChange={(e) => setSourceFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
            <option value="">All sources</option>
            {BACKLOG_SOURCES.map((s) => <option key={s.value} value={s.value}>{s.label}</option>)}
          </select>
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
            <option value="">All statuses</option>
            <option value="Open">Open</option>
            <option value="Closed">Closed</option>
          </select>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
            <Download size={14} /> Export to Excel
          </button>
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Add Backlog
          </button>
        </div>
      </div>
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map(([key, label]) => <th key={key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{label}</th>)}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr key={row.id ?? i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{row.asset_id}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{backlogSourceLabel(row.source_type)}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{row.work_order_id ? (woNoById[row.work_order_id] || "-") : <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", cursor: "pointer" }}>{row.component_code || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", cursor: "pointer", maxWidth: 320 }}>{row.description}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", cursor: "pointer" }}><Badge value={row.priority} /></td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{row.date_notified}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{ageDays(row.date_notified)}</td>
                <td onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", cursor: "pointer" }}><Badge value={row.status} /></td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {row.status === "Open" && (
                    <button onClick={() => quickClose(row)} title="Mark Closed" style={{ background: "none", border: "none", color: "#2C5646", cursor: "pointer", padding: 4, fontSize: 12, fontWeight: 600 }}>
                      Close
                    </button>
                  )}
                  <button onClick={() => setDeleting(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {backlogs.length === 0 ? "No backlog items logged yet." : "No entries match your filters."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      {showForm && (
        <BacklogForm assets={assets} workOrders={workOrders} existing={editing} userEmail={userEmail}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); onRefresh(); }} />
      )}
      {deleting && (
        <DeleteConfirmModal
          itemLabel={`backlog item on ${deleting.asset_id} (${deleting.description?.slice(0, 40)}${deleting.description?.length > 40 ? "…" : ""})`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

// Runs entirely in the browser - no server call, no API key, no cost.
// Downscales the image, converts to grayscale, then flags it as 'dark'
// if the average brightness is low, or 'blurry' using the variance of
// a Laplacian edge filter (a standard, cheap sharpness metric - sharp
// photos have high-variance edges, blurry ones don't). Thresholds are
// a reasonable starting point, not a guarantee - the raw numbers are
// kept in the notes so they can be tuned later against real photos.
function analyzeImageQuality(file) {
  return new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      try {
        const maxDim = 400;
        const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
        const w = Math.max(1, Math.round(img.width * scale));
        const h = Math.max(1, Math.round(img.height * scale));
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, w, h);
        const { data } = ctx.getImageData(0, 0, w, h);

        const gray = new Float32Array(w * h);
        let sum = 0;
        for (let i = 0, p = 0; i < data.length; i += 4, p++) {
          const g = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
          gray[p] = g;
          sum += g;
        }
        const brightness = sum / gray.length;

        let lapSum = 0, lapSumSq = 0, count = 0;
        for (let y = 1; y < h - 1; y++) {
          for (let x = 1; x < w - 1; x++) {
            const idx = y * w + x;
            const lap = 4 * gray[idx] - gray[idx - 1] - gray[idx + 1] - gray[idx - w] - gray[idx + w];
            lapSum += lap;
            lapSumSq += lap * lap;
            count++;
          }
        }
        const lapMean = lapSum / count;
        const lapVariance = lapSumSq / count - lapMean * lapMean;

        URL.revokeObjectURL(url);

        const detail = `brightness ${brightness.toFixed(0)}/255, sharpness ${lapVariance.toFixed(0)}`;
        if (brightness < 55) resolve({ status: "dark", notes: `Looks too dark to read clearly (${detail})` });
        else if (lapVariance < 120) resolve({ status: "blurry", notes: `Looks blurry or out of focus (${detail})` });
        else resolve({ status: "ok", notes: detail });
      } catch (err) {
        URL.revokeObjectURL(url);
        resolve({ status: "unchecked", notes: "Quality check couldn't run on this file" });
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve({ status: "unchecked", notes: "Couldn't load this file as an image" });
    };
    img.src = url;
  });
}

const PHOTO_QUALITY_LABEL = { ok: "Looks clear", blurry: "May be blurry", dark: "May be too dark", unchecked: "Not checked" };

function DailyServiceForm({ assets, defaultAssetId, defaultDate, existing, userEmail, onClose, onSaved }) {
  const [assetId, setAssetId] = useState(existing?.asset_id || defaultAssetId || assets[0]?.asset_id || "");
  const [serviceDate, setServiceDate] = useState(existing?.service_date || defaultDate || todayForInput());
  const [completedBy, setCompletedBy] = useState(existing?.completed_by || userEmail || "");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const [existingPhotoUrl, setExistingPhotoUrl] = useState(null);
  const [newFile, setNewFile] = useState(null);
  const [checkingQuality, setCheckingQuality] = useState(false);
  const [qualityResult, setQualityResult] = useState(null); // { status, notes } for newFile
  const fileInputRef = React.useRef(null);

  useEffect(() => {
    let cancelled = false;
    async function loadExistingPhoto() {
      if (!existing?.photo_path) { setExistingPhotoUrl(null); return; }
      const { data } = await supabase.storage.from("job-card-scans").createSignedUrl(existing.photo_path, 3600);
      if (!cancelled) setExistingPhotoUrl(data?.signedUrl || null);
    }
    loadExistingPhoto();
    return () => { cancelled = true; };
  }, [existing?.photo_path]);

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setNewFile(file);
    setQualityResult(null);
    setCheckingQuality(true);
    const result = await analyzeImageQuality(file);
    setQualityResult(result);
    setCheckingQuality(false);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSaving(true);
    try {
      const payload = { asset_id: assetId, service_date: serviceDate, completed_by: completedBy || null, notes: notes || null };

      if (newFile) {
        const path = `daily-service/${assetId}/${serviceDate}-${Date.now()}-${newFile.name}`;
        const { error: uploadErr } = await supabase.storage.from("job-card-scans").upload(path, newFile);
        if (uploadErr) throw uploadErr;
        payload.photo_path = path;
        payload.photo_uploaded_at = new Date().toISOString();
        payload.photo_quality_status = qualityResult?.status || "unchecked";
        payload.photo_quality_notes = qualityResult?.notes || null;
      }

      const { error: dbError } = await supabase.from("daily_service_checklist").upsert(payload, { onConflict: "asset_id,service_date" });
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: "inherit" };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };
  const flagged = qualityResult && qualityResult.status !== "ok" && qualityResult.status !== "unchecked";

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 420, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>Log Daily Service</h3>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required style={fieldStyle}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Date</label>
          <input type="date" value={serviceDate} onChange={(e) => setServiceDate(e.target.value)} required max={todayForInput()} style={fieldStyle} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Completed By</label>
          <input type="text" value={completedBy} onChange={(e) => setCompletedBy(e.target.value)} style={fieldStyle} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Notes</label>
          <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows={2} style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Photo of Completed Service</label>

          {existingPhotoUrl && !newFile && (
            <a href={existingPhotoUrl} target="_blank" rel="noreferrer" style={{ display: "block", fontSize: 12.5, color: NAVY, marginBottom: 8 }}>
              View attached photo {existing?.photo_quality_status && existing.photo_quality_status !== "unchecked" && (
                <span style={{ marginLeft: 6 }}><Badge value={existing.photo_quality_status === "ok" ? "OK" : existing.photo_quality_status.toUpperCase()} /></span>
              )}
            </a>
          )}

          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "7px 14px", borderRadius: 8, cursor: "pointer" }}
          >
            <Upload size={14} /> {existingPhotoUrl ? "Replace Photo" : "Take / Attach Photo"}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            capture="environment"
            onChange={handleFileChange}
            style={{ display: "none" }}
          />

          {newFile && (
            <p style={{ fontSize: 12, color: "#4B5659", margin: "8px 0 0" }}>{newFile.name}</p>
          )}
          {checkingQuality && (
            <p style={{ fontSize: 12, color: "#859195", margin: "6px 0 0" }}>Checking photo quality…</p>
          )}
          {qualityResult && !checkingQuality && (
            <div style={{
              marginTop: 8, padding: "8px 10px", borderRadius: 8, fontSize: 12,
              background: flagged ? "#F6E2E0" : "#E2EFE9", color: flagged ? "#7A3330" : "#2C5646",
            }}>
              {flagged
                ? `${PHOTO_QUALITY_LABEL[qualityResult.status]} - consider retaking this photo before saving. (${qualityResult.notes})`
                : `${PHOTO_QUALITY_LABEL[qualityResult.status]}.`}
            </div>
          )}
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : "Log Daily Service"}
          </button>
        </div>
      </form>
    </div>
  );
}

// DS = a daily_service_checklist entry exists for that asset+day.
// BD = no DS entry, but a breakdown fully covers the day (down at the
//      start of day through at least the end of day) - no penalty.
// NU = no DS entry, not a full-day breakdown, but Daily Hours shows the
//      asset logged with zero hours run that day - not utilised, no
//      penalty.
// gap = none of the above - the cell needing attention, highlighted.
function dailyServiceStatus(assetId, dateStr, checklistByKey, breakdowns, dailyHours) {
  if (checklistByKey.has(`${assetId}|${dateStr}`)) return "DS";

  const dayStart = new Date(`${dateStr}T00:00:00`);
  const dayEnd = new Date(`${dateStr}T23:59:59.999`);
  const fullDayDown = breakdowns.some((b) => {
    if (b.asset_id !== assetId) return false;
    const start = b.downtime_start ? new Date(b.downtime_start) : null;
    const end = b.downtime_end ? new Date(b.downtime_end) : null;
    if (!start || start > dayStart) return false;
    return !end || end >= dayEnd;
  });
  if (fullDayDown) return "BD";

  const dayHours = dailyHours.filter((h) => h.asset_id === assetId && h.log_date === dateStr);
  const notUtilised = dayHours.length > 0 && dayHours.every((h) => Number(h.hours_run || 0) === 0);
  if (notUtilised) return "NU";

  return null; // gap
}

function DailyServicePage({ assets, dailyServiceChecklist, breakdowns, dailyHours, userEmail, onRefresh }) {
  const [monthValue, setMonthValue] = useState(() => todayForInput().slice(0, 7)); // "YYYY-MM"
  const [showForm, setShowForm] = useState(false);
  const [formDefaults, setFormDefaults] = useState({ assetId: "", date: "" });
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const checklistByKey = useMemo(
    () => new Map(dailyServiceChecklist.map((r) => [`${r.asset_id}|${r.service_date}`, r])),
    [dailyServiceChecklist]
  );

  const filteredAssets = useMemo(() => {
    let rows = assets;
    if (selectedAsset) rows = rows.filter((a) => a.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((a) => a.fleet === selectedFleet);
    return rows;
  }, [assets, selectedFleet, selectedAsset]);

  const [year, month] = monthValue.split("-").map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const days = Array.from({ length: daysInMonth }, (_, i) => i + 1);
  const todayStr = todayForInput();

  const cellStyle = (label) => {
    if (label === "DS") return { background: "#E2EFE9", color: "#2C5646" };
    if (label === "BD") return { background: "#F6E2E0", color: "#7A3330" };
    if (label === "NU") return { background: "#F1EFE8", color: "#4B5659" };
    return { background: "#F9D8D8", color: "#7A3330" }; // gap - needs attention
  };

  const openLogForm = (assetId, dateStr) => {
    setFormDefaults({ assetId, date: dateStr });
    setShowForm(true);
  };

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <input type="month" value={monthValue} onChange={(e) => setMonthValue(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 10, fontSize: 12, color: "#4B5659" }}>
            <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#E2EFE9", borderRadius: 2, marginRight: 4 }} />DS - serviced</span>
            <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#F6E2E0", borderRadius: 2, marginRight: 4 }} />BD - breakdown</span>
            <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#F1EFE8", borderRadius: 2, marginRight: 4 }} />NU - not utilised</span>
            <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#F9D8D8", borderRadius: 2, marginRight: 4 }} />missing</span>
          </div>
          <button onClick={() => openLogForm(filteredAssets[0]?.asset_id || "", todayStr)} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Log Daily Service
          </button>
        </div>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ borderCollapse: "collapse", fontSize: 11.5 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              <th style={{ textAlign: "left", padding: "8px 10px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", position: "sticky", left: 0, background: "#F7F8F6", minWidth: 140 }}>Equipment</th>
              {days.map((d) => (
                <th key={d} style={{ textAlign: "center", padding: "6px 4px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", minWidth: 26 }}>{d}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredAssets.map((a, ri) => (
              <tr key={a.asset_id} style={{ borderBottom: ri < filteredAssets.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "6px 10px", whiteSpace: "nowrap", fontWeight: 600, position: "sticky", left: 0, background: "#fff" }}>{a.asset_id}</td>
                {days.map((d) => {
                  const dateStr = `${year}-${String(month).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
                  const isFuture = dateStr > todayStr;
                  const label = isFuture ? "" : (dailyServiceStatus(a.asset_id, dateStr, checklistByKey, breakdowns, dailyHours) || "");
                  return (
                    <td key={d} style={{ padding: 0, textAlign: "center", borderLeft: "1px solid #F2F1EA" }}>
                      <button
                        onClick={() => !isFuture && openLogForm(a.asset_id, dateStr)}
                        disabled={isFuture}
                        title={isFuture ? "" : `${a.asset_id} - ${dateStr}${label ? `: ${label}` : ": no daily service logged"}`}
                        style={{
                          width: "100%", minHeight: 26, border: "none", cursor: isFuture ? "default" : "pointer",
                          fontSize: 10.5, fontWeight: 700, ...(isFuture ? { background: "transparent" } : cellStyle(label)),
                        }}
                      >
                        {label}
                      </button>
                    </td>
                  );
                })}
              </tr>
            ))}
            {filteredAssets.length === 0 && (
              <tr><td colSpan={days.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No equipment to show.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <DailyServiceForm
          assets={assets}
          defaultAssetId={formDefaults.assetId}
          defaultDate={formDefaults.date}
          existing={checklistByKey.get(`${formDefaults.assetId}|${formDefaults.date}`) || null}
          userEmail={userEmail}
          onClose={() => setShowForm(false)}
          onSaved={() => { setShowForm(false); onRefresh(); }}
        />
      )}
    </div>
  );
}

function FuelLogPage({ assets, fuelLog, userEmail, dailyHours, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const columns = [
    ["fill_date", "Date"], ["asset_id", "Equipment #"], ["hour_meter", "Hour meter"],
    ["litres", "Litres"], ["consumption_rate", "Rate (L/hr)"], ["recorded_by", "Recorded by"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    let rows = fuelLog;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    return [...rows].sort((a, b) => (b.fill_date || "").localeCompare(a.fill_date || ""));
  }, [fuelLog, assets, query, selectedFleet, selectedAsset]);

  const handleDelete = async (reason) => {
    await deleteWithReason("fuel_log", deleting.id, "id", reason, userEmail);
    onRefresh();
    setDeleting(null);
  };

  const FUEL_LOG_FIELDS = [
    { key: "fill_date", header: "Date (YYYY-MM-DD)", type: "date" },
    { key: "hour_meter", header: "Hour Meter", type: "number" },
    { key: "litres", header: "Litres", type: "number" },
    { key: "notes", header: "Notes", type: "text" },
  ];

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search fuel log" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
          + Log Fuel
        </button>
      </div>
      <ExcelSync
        data={fuelLog} assets={assets} fields={FUEL_LOG_FIELDS} tableName="fuel_log"
        sheetTitle="Fuel Log" filenamePrefix="Fuel_Log" onRefresh={onRefresh}
        extraOnSave={(obj) => ({ recorded_by: obj.recorded_by || userEmail || null })}
      />
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => <th key={c.key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{c.label}</th>)}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => {
              const prevDate = i > 0 ? filtered[i - 1].fill_date : null;
              const showHeader = row.fill_date !== prevDate;
              return (
                <React.Fragment key={row.id ?? i}>
                  {showHeader && (
                    <tr>
                      <td colSpan={columns.length + 1} style={{ padding: "8px 12px", background: "#F2F1EA", fontWeight: 700, fontSize: 12.5, color: NAVY, borderTop: i > 0 ? "1px solid #E2E6E3" : "none" }}>
                        {row.fill_date ? new Date(row.fill_date + "T12:00:00").toLocaleDateString("en-ZA", { weekday: "long", day: "2-digit", month: "short", year: "numeric" }) : "No date"}
                      </td>
                    </tr>
                  )}
                  <tr style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                    {columns.map((c) => (
                      <td key={c.key} onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>
                        {row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>}
                      </td>
                    ))}
                    <td style={{ padding: "9px 12px" }}>
                      <button onClick={() => setDeleting(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                        <Trash2 size={15} />
                      </button>
                    </td>
                  </tr>
                </React.Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {fuelLog.length === 0 ? "No fuel entries yet." : "No entries match your filters."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      {showForm && (
        <FuelLogForm assets={assets} existing={editing} userEmail={userEmail} dailyHours={dailyHours}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); onRefresh(); }} />
      )}
      {deleting && (
        <DeleteConfirmModal
          itemLabel={`fuel entry on ${deleting.asset_id} (${deleting.fill_date})`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

const OIL_TYPES = ["Engine Oil 15W-40", "Hydraulic Oil", "Transmission Oil", "Gear Oil", "Other"];
const OIL_FILL_REASONS = ["Scheduled Change", "Top-Up", "Leak Repair"];

function OilConsumptionForm({ assets, existing, userEmail, dailyHours, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [fillDate, setFillDate] = useState(existing?.fill_date || todayForInput());
  const [hourMeter, setHourMeter] = useState(existing?.hour_meter ?? "");
  const [oilType, setOilType] = useState(existing?.oil_type || OIL_TYPES[0]);
  const [litres, setLitres] = useState(existing?.litres ?? "");
  const [fillReason, setFillReason] = useState(existing?.fill_reason || OIL_FILL_REASONS[0]);
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const recordedBy = existing?.recorded_by || userEmail || "";

  const hourMeterWarning = useMemo(
    () => checkHourMeterPlausibility(assetId, hourMeter, fillDate, dailyHours),
    [assetId, hourMeter, fillDate, dailyHours]
  );

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    if (litres === "" || hourMeter === "") { setError("Enter Hour Meter and Litres."); return; }
    setSaving(true);
    const payload = {
      asset_id: assetId, fill_date: fillDate, hour_meter: Number(hourMeter), oil_type: oilType,
      litres: Number(litres), fill_reason: fillReason, recorded_by: recordedBy || null, notes: notes || null,
    };
    try {
      const { error: dbError } = isEdit
        ? await supabase.from("oil_consumption").update(payload).eq("id", existing.id)
        : await supabase.from("oil_consumption").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 420, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>{isEdit ? "Edit Oil Entry" : "Log Oil"}</h3>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Date</label>
            <input type="date" value={fillDate} onChange={(e) => setFillDate(e.target.value)} required max={todayForInput()} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Hour Meter</label>
            <input type="number" step="0.1" value={hourMeter} onChange={(e) => setHourMeter(e.target.value)} required style={fieldStyle} />
          </div>
        </div>
        {hourMeterWarning && (
          <div style={{ background: "#F5E9D8", border: "1px solid #E3C79B", borderRadius: 8, padding: "8px 10px", marginBottom: 12, fontSize: 12, color: "#7A5320" }}>
            {hourMeterWarning}
          </div>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Oil Type</label>
            <select value={oilType} onChange={(e) => setOilType(e.target.value)} style={fieldStyle}>
              {OIL_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Reason</label>
            <select value={fillReason} onChange={(e) => setFillReason(e.target.value)} style={fieldStyle}>
              {OIL_FILL_REASONS.map((r) => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Litres</label>
          <input type="number" step="0.1" value={litres} onChange={(e) => setLitres(e.target.value)} required style={fieldStyle} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Recorded By</label>
          <input type="text" value={recordedBy} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
        </div>
        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Notes</label>
          <input type="text" value={notes} onChange={(e) => setNotes(e.target.value)} style={fieldStyle} />
        </div>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Log Oil"}
          </button>
        </div>
      </form>
    </div>
  );
}

function OilConsumptionPage({ assets, oilConsumption, userEmail, dailyHours, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const columns = [
    ["fill_date", "Date"], ["asset_id", "Equipment #"], ["hour_meter", "Hour meter"], ["oil_type", "Oil type"],
    ["litres", "Litres"], ["fill_reason", "Reason"], ["consumption_rate", "Rate (L/hr)"], ["recorded_by", "Recorded by"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    let rows = oilConsumption;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    return [...rows].sort((a, b) => (b.fill_date || "").localeCompare(a.fill_date || ""));
  }, [oilConsumption, assets, query, selectedFleet, selectedAsset]);

  const handleDelete = async (reason) => {
    await deleteWithReason("oil_consumption", deleting.id, "id", reason, userEmail);
    onRefresh();
    setDeleting(null);
  };

  const OIL_FIELDS = [
    { key: "fill_date", header: "Date (YYYY-MM-DD)", type: "date" },
    { key: "hour_meter", header: "Hour Meter", type: "number" },
    { key: "oil_type", header: "Oil Type", type: "text" },
    { key: "litres", header: "Litres", type: "number" },
    { key: "fill_reason", header: "Fill Reason", type: "text" },
    { key: "notes", header: "Notes", type: "text" },
  ];

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search oil consumption" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
          + Log Oil
        </button>
      </div>
      <ExcelSync
        data={oilConsumption} assets={assets} fields={OIL_FIELDS} tableName="oil_consumption"
        sheetTitle="Oil Consumption" filenamePrefix="Oil_Consumption" onRefresh={onRefresh}
        extraOnSave={(obj) => ({ recorded_by: obj.recorded_by || userEmail || null })}
      />
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => <th key={c.key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{c.label}</th>)}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => {
              const prevDate = i > 0 ? filtered[i - 1].fill_date : null;
              const showHeader = row.fill_date !== prevDate;
              return (
                <React.Fragment key={row.id ?? i}>
                  {showHeader && (
                    <tr>
                      <td colSpan={columns.length + 1} style={{ padding: "8px 12px", background: "#F2F1EA", fontWeight: 700, fontSize: 12.5, color: NAVY, borderTop: i > 0 ? "1px solid #E2E6E3" : "none" }}>
                        {row.fill_date ? new Date(row.fill_date + "T12:00:00").toLocaleDateString("en-ZA", { weekday: "long", day: "2-digit", month: "short", year: "numeric" }) : "No date"}
                      </td>
                    </tr>
                  )}
                  <tr style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                    {columns.map((c) => (
                      <td key={c.key} onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>
                        {row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>}
                      </td>
                    ))}
                    <td style={{ padding: "9px 12px" }}>
                      <button onClick={() => setDeleting(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                        <Trash2 size={15} />
                      </button>
                    </td>
                  </tr>
                </React.Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {oilConsumption.length === 0 ? "No oil entries yet." : "No entries match your filters."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      {showForm && (
        <OilConsumptionForm assets={assets} existing={editing} userEmail={userEmail} dailyHours={dailyHours}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); onRefresh(); }} />
      )}
      {deleting && (
        <DeleteConfirmModal
          itemLabel={`oil entry on ${deleting.asset_id} (${deleting.fill_date})`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

const MTBF_MTTR_COLUMNS = [
  ["asset_id", "Equipment #"], ["asset_name", "Name"], ["fleet", "Fleet"],
  ["num_unplanned_events", "Breakdowns"], ["mtbf", "MTBF (hrs)"], ["mttr", "MTTR (hrs)"],
  ["availability", "Availability"], ["availability_index", "Availability Index"],
];

function MtbfMttrReportPage({ assets }) {
  const [fromDateTime, setFromDateTime] = useState(DEFAULT_FROM);
  const [toDateTime, setToDateTime] = useState(DEFAULT_TO);
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [kpiData, setKpiData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [sortKey, setSortKey] = useState("num_unplanned_events");
  const [sortDir, setSortDir] = useState("desc");

  useEffect(() => {
    let cancelled = false;
    async function load() {
      setLoading(true);
      setError(null);
      const { data, error: rpcError } = await supabase.rpc("plant_performance_kpi", {
        period_start: new Date(fromDateTime).toISOString(),
        period_end: new Date(toDateTime).toISOString(),
      });
      if (cancelled) return;
      if (rpcError) setError(rpcError.message);
      else setKpiData(data || []);
      setLoading(false);
    }
    load();
    return () => { cancelled = true; };
  }, [fromDateTime, toDateTime]);

  const rows = useMemo(() => {
    let r = kpiData;
    if (selectedAsset) r = r.filter((x) => x.asset_id === selectedAsset);
    else if (selectedFleet) r = r.filter((x) => x.fleet === selectedFleet);

    return [...r].sort((a, b) => {
      const av = a[sortKey], bv = b[sortKey];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      const cmp = typeof av === "string" ? av.localeCompare(bv) : av - bv;
      return sortDir === "asc" ? cmp : -cmp;
    });
  }, [kpiData, selectedFleet, selectedAsset, sortKey, sortDir]);

  const handleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("desc"); }
  };

  const fmt = (key, val) => {
    if (val == null) return "-";
    if (key === "availability") return `${Math.round(val * 100)}%`;
    if (key === "availability_index") return `${Number(val).toFixed(1)}%`;
    if (key === "mtbf" || key === "mttr") return Number(val).toFixed(1);
    return val;
  };

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = MTBF_MTTR_COLUMNS.map((c) => c[1]);
    const dataRows = rows.map((row) => MTBF_MTTR_COLUMNS.map(([key]) => fmt(key, row[key])));
    const aoa = [
      ["MTBF / MTTR Report"],
      [`Date Range: ${formatRangeForDisplay(fromDateTime, toDateTime)}`],
      [`Exported: ${timestamp}`],
      [],
      headerRow, ...dataRows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: MTBF_MTTR_COLUMNS.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: MTBF_MTTR_COLUMNS.length - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: MTBF_MTTR_COLUMNS.length - 1 } },
    ];
    ws["!cols"] = MTBF_MTTR_COLUMNS.map((c) => ({ wch: Math.max(c[1].length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MTBF MTTR");
    XLSX.writeFile(wb, `MTBF_MTTR_Report_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div>
      <DateRangePicker fromDateTime={fromDateTime} toDateTime={toDateTime} setFromDateTime={setFromDateTime} setToDateTime={setToDateTime} />
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />

      {!loading && rows.length > 0 && (() => {
        const mtbfTarget = 40, mttrTarget = 4, utilTarget = 85;
        const mtbfMax = Math.max(mtbfTarget, ...rows.map((r) => r.mtbf || 0)) * 1.15;
        const mttrMax = Math.max(mttrTarget, ...rows.map((r) => r.mttr || 0)) * 1.15;
        const utilRows = rows.map((r) => ({ ...r, utilisationPct: r.utilisation != null ? Math.round(r.utilisation * 100) : null }));
        return (
        <div style={{ marginBottom: 8 }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 16, marginBottom: 8 }}>
            <KpiBarChart
              title="MTBF by equipment (hrs)" data={rows} xKey="asset_id" dataKey="mtbf"
              target={mtbfTarget} domainMax={mtbfMax} unitSuffix="h"
              valueFormatter={(v) => Number(v).toFixed(1)}
              meetsTarget={(r) => (r.mtbf || 0) >= mtbfTarget}
            />
            <KpiBarChart
              title="MTTR by equipment (hrs)" data={rows} xKey="asset_id" dataKey="mttr"
              target={mttrTarget} domainMax={mttrMax} unitSuffix="h"
              valueFormatter={(v) => Number(v).toFixed(1)}
              meetsTarget={(r) => (r.mttr || 0) <= mttrTarget}
            />
            <KpiBarChart
              title="Utilisation by equipment" data={utilRows} xKey="asset_id" dataKey="utilisationPct"
              target={utilTarget} domainMax={100} unitSuffix="%"
              valueFormatter={(v) => `${v}%`}
              meetsTarget={(r) => (r.utilisationPct || 0) >= utilTarget}
            />
          </div>
          <KpiLegend targetLabel="Target" />
        </div>
        );
      })()}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <p style={{ fontSize: 13, color: "#4B5659", margin: 0 }}>
          Sorted by most breakdowns first, so problem equipment surfaces automatically. Click any column to sort by it instead.
        </p>
        <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer", whiteSpace: "nowrap" }}>
          <Download size={14} /> Export to Excel
        </button>
      </div>

      {error && (
        <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13, color: "#7A3330" }}>
          Couldn't load report: {error}
        </div>
      )}

      {loading ? (
        <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
      ) : (
        <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#F7F8F6" }}>
                {MTBF_MTTR_COLUMNS.map(([key, label]) => (
                  <th
                    key={key}
                    onClick={() => handleSort(key)}
                    style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3", cursor: "pointer", userSelect: "none" }}
                  >
                    {label} {sortKey === key ? (sortDir === "asc" ? "▲" : "▼") : ""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, i) => (
                <tr key={row.asset_id ?? i} style={{ borderBottom: i < rows.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                  {MTBF_MTTR_COLUMNS.map(([key]) => (
                    <td key={key} style={{ padding: "9px 12px", whiteSpace: "nowrap", fontWeight: key === "num_unplanned_events" && row[key] >= 3 ? 700 : 400, color: key === "num_unplanned_events" && row[key] >= 3 ? "#B85450" : "#183642" }}>
                      {fmt(key, row[key])}
                    </td>
                  ))}
                </tr>
              ))}
              {rows.length === 0 && (
                <tr><td colSpan={MTBF_MTTR_COLUMNS.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No data for this date range.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function GanttTooltip({ active, payload }) {
  if (!active || !payload || !payload.length) return null;
  const item = payload.find((p) => p.dataKey === "durationHrs")?.payload;
  if (!item) return null;
  const fmt = (d) => d ? d.toLocaleString("en-ZA", { hour: "2-digit", minute: "2-digit", day: "2-digit", month: "short" }) : "ongoing";
  const statusLabel = item.kind === "scheduled" ? "Scheduled - not booked down yet" : item.completed ? "Closed" : "In progress";
  return (
    <div style={{ background: "#fff", border: "1px solid #E2E6E3", borderRadius: 8, padding: "9px 12px", fontSize: 12, boxShadow: "0 2px 8px rgba(0,0,0,0.1)" }}>
      <p style={{ fontWeight: 700, margin: "0 0 3px", color: "#183642" }}>{item.label}</p>
      <p style={{ margin: 0, color: "#4B5659" }}>{item.category} · {statusLabel}</p>
      {item.kind === "scheduled" ? (
        <p style={{ margin: "4px 0 0", color: "#4B5659" }}>Planned for {fmt(item.start)}</p>
      ) : (
        <p style={{ margin: "4px 0 0", color: "#4B5659" }}>{fmt(item.start)} → {fmt(item.end)}</p>
      )}
      {!item.completed && item.expectedUp && (
        <p style={{ margin: "4px 0 0", color: NAVY, fontWeight: 600 }}>Expected up: {fmt(item.expectedUp)}</p>
      )}
    </div>
  );
}

// A live Gantt-style timeline for breakdowns and planned events, in the
// spirit of a project-plan Gantt chart: each event as a horizontal bar
// positioned along a time axis, with a "now" marker. Built on Recharts'
// BarChart using the standard technique for faking a Gantt with a
// generic bar library - a transparent "offset" segment positions each
// bar's start, stacked with a visible "duration" segment for its length.
function EventTimeline({ breakdowns, workOrders, fromDateTime, toDateTime }) {
  const [statusFilter, setStatusFilter] = useState("all");
  const [nowTick, setNowTick] = useState(() => new Date());
  // Self-contained rather than threaded down as a prop, since this chart
  // gets embedded in more than one page - the equipment-name column eats
  // most of a phone-width chart at the same fixed width used on desktop,
  // squeezing the actual timeline bars into a thin sliver.
  const [narrow, setNarrow] = useState(() => typeof window !== "undefined" && window.innerWidth < 560);

  useEffect(() => {
    const interval = setInterval(() => setNowTick(new Date()), 60000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    const onResize = () => setNarrow(window.innerWidth < 560);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  // Driven by the page's own date range picker when provided (Downtime
  // Summary) - selecting a range changes what the timeline shows, same
  // as it changes the table below it. Falls back to a live ±6h window
  // around "now" when no range is passed in (Dashboard's usage).
  const windowStart = useMemo(() => fromDateTime ? new Date(fromDateTime) : new Date(nowTick.getTime() - 6 * 3600000), [fromDateTime, nowTick]);
  const windowEnd = useMemo(() => toDateTime ? new Date(toDateTime) : new Date(nowTick.getTime() + 6 * 3600000), [toDateTime, nowTick]);
  const windowHrs = Math.max(0.1, (windowEnd.getTime() - windowStart.getTime()) / 3600000);
  const nowInWindow = nowTick >= windowStart && nowTick <= windowEnd;

  const events = useMemo(() => {
    const bEvents = breakdowns
      .filter((b) => b.downtime_start)
      .map((b) => ({
        id: `b-${b.id}`, label: `${b.asset_id} - ${b.component_affected || b.cause_code || "Breakdown"}`,
        start: new Date(b.downtime_start), end: b.downtime_end ? new Date(b.downtime_end) : null,
        completed: b.repair_status === "Closed", kind: "actual", category: "Breakdown",
        expectedUp: b.expected_up_time ? new Date(b.expected_up_time) : null,
      }));
    const wStarted = workOrders
      .filter((w) => w.work_type === "Preventive" && w.actual_start)
      .map((w) => ({
        id: `w-${w.id}`, label: `${w.asset_id} - ${w.problem_scope || "Planned Maintenance"}`,
        start: new Date(w.actual_start), end: w.actual_finish ? new Date(w.actual_finish) : null,
        completed: w.status === "Closed", kind: "actual", category: "Planned Maintenance",
        expectedUp: null,
      }));
    // Not booked down yet - shown as a slim marker at its planned time so
    // people on the ground can see a machine is coming up for service,
    // before anyone has actually taken it out of service.
    const wScheduled = workOrders
      .filter((w) => w.work_type === "Preventive" && !w.actual_start && w.planned_start)
      .map((w) => ({
        id: `w-sched-${w.id}`, label: `${w.asset_id} - ${w.problem_scope || "Planned Maintenance"} (scheduled)`,
        start: new Date(w.planned_start), end: null, completed: false, kind: "scheduled", category: "Scheduled",
        expectedUp: null,
      }));

    const actual = [...bEvents, ...wStarted]
      .filter((e) => statusFilter === "all" || (statusFilter === "completed" ? e.completed : !e.completed))
      .filter((e) => e.start < windowEnd && (e.end || nowTick) > windowStart)
      .map((e) => {
        // Solid bar = actual elapsed time so far (start to now, clipped to
        // the window). Still-open events with an Expected Up Time also
        // get a thin projected line continuing on from "now" to that
        // estimate - no Expected Up Time set means no line at all, since
        // there's nothing to project towards.
        const effectiveEnd = e.end || nowTick;
        const clippedStartMs = Math.max(e.start.getTime(), windowStart.getTime());
        const clippedEndMs = Math.min(effectiveEnd.getTime(), windowEnd.getTime());
        const offsetHrs = (clippedStartMs - windowStart.getTime()) / 3600000;
        const durationHrs = Math.max(windowHrs * 0.004, (clippedEndMs - clippedStartMs) / 3600000);

        let continuationHrs = 0;
        if (!e.completed && !e.end && e.expectedUp && e.expectedUp > nowTick && nowTick < windowEnd) {
          const continuationStartMs = Math.max(clippedEndMs, windowStart.getTime());
          const projectedEndMs = Math.min(e.expectedUp.getTime(), windowEnd.getTime());
          continuationHrs = Math.max(0, (projectedEndMs - continuationStartMs) / 3600000);
        }
        return { ...e, offsetHrs, durationHrs, continuationHrs };
      });

    const scheduled = wScheduled
      .filter((e) => statusFilter !== "completed")
      .filter((e) => e.start >= windowStart && e.start <= windowEnd)
      .map((e) => ({
        ...e,
        offsetHrs: (e.start.getTime() - windowStart.getTime()) / 3600000,
        durationHrs: Math.max(windowHrs * 0.006, windowHrs * 0.006),
        continuationHrs: 0,
      }));

    return [...actual, ...scheduled].sort((a, b) => a.offsetHrs - b.offsetHrs);
  }, [breakdowns, workOrders, statusFilter, nowTick, windowStart, windowEnd, windowHrs]);

  const nowOffsetHrs = (nowTick.getTime() - windowStart.getTime()) / 3600000;
  const formatTick = (h) => new Date(windowStart.getTime() + h * 3600000).toLocaleString("en-ZA", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });
  const tickCount = narrow ? 3 : 6;
  const ticks = Array.from({ length: tickCount + 1 }, (_, i) => Math.round((windowHrs / tickCount) * i * 100) / 100);

  // Draws the solid "actual so far" bar, and - for a still-open event -
  // a thinner projected line continuing on from there. Scheduled (not
  // yet booked down) jobs render as a small amber diamond marker instead
  // of a bar, since there's no duration yet to show.
  const DurationBarShape = (props) => {
    const { x, y, width, height, payload, fill } = props;
    if (payload.kind === "scheduled") {
      const cx = x + width / 2, cy = y + height / 2, r = Math.max(5, height / 2.4);
      return <polygon points={`${cx},${cy - r} ${cx + r},${cy} ${cx},${cy + r} ${cx - r},${cy}`} fill="#E8A33D" stroke="#C58A32" strokeWidth={1} />;
    }
    const pxPerHour = payload.durationHrs > 0 ? width / payload.durationHrs : 0;
    const continuationWidth = pxPerHour > 0 ? payload.continuationHrs * pxPerHour : 0;
    return (
      <g>
        <rect x={x} y={y} width={width} height={height} fill={fill} rx={4} ry={4} />
        {continuationWidth > 0 && (
          <rect x={x + width} y={y + height / 2 - 1.5} width={continuationWidth} height={3} fill="#3F7D65" rx={1.5} ry={1.5} />
        )}
      </g>
    );
  };

  return (
    <div style={{ marginBottom: 28 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <h3 style={{ fontSize: 15, fontWeight: 700, margin: 0, color: NAVY }}>Event Timeline</h3>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          style={{ padding: "7px 10px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', background: "#fff" }}
        >
          <option value="all">All events</option>
          <option value="completed">Completed only</option>
          <option value="in_progress">In progress only</option>
        </select>
      </div>
      <p style={{ fontSize: 12, color: "#859195", margin: "0 0 12px" }}>
        {formatTick(0)} to {formatTick(windowHrs)}{fromDateTime && toDateTime ? " - matches the date range selected below." : " - 6 hours either side of now."} Includes upcoming scheduled jobs, not just what's already down.
      </p>
      {events.length === 0 ? (
        <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, background: "#fff", padding: 24, textAlign: "center" }}>
          <p style={{ fontSize: 13, color: "#859195", margin: 0 }}>No events in this date range.</p>
        </div>
      ) : (
        <div style={{ height: Math.max(180, events.length * 38 + 50), background: "#fff", border: "1px solid #E2E6E3", borderRadius: 10, padding: "10px 8px" }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={events} layout="vertical" margin={{ left: 4, right: 12, top: 6, bottom: 6 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EFEEE7" horizontal={false} />
              <XAxis type="number" domain={[0, windowHrs]} ticks={ticks} tickFormatter={formatTick} tick={{ fontSize: narrow ? 9.5 : 11, fill: "#4B5659" }} axisLine={{ stroke: "#E2E6E3" }} tickLine={false} />
              <YAxis type="category" dataKey="label" width={narrow ? 110 : 220} tick={{ fontSize: narrow ? 9.5 : 11, fill: "#183642" }} axisLine={false} tickLine={false} />
              <Tooltip content={<GanttTooltip />} />
              {nowInWindow && (
                <ReferenceLine x={nowOffsetHrs} stroke="#B85450" strokeWidth={1.5} strokeDasharray="4 4" label={{ value: "Now", position: "top", fill: "#B85450", fontSize: 11, fontWeight: 700 }} />
              )}
              <Bar dataKey="offsetHrs" stackId="a" fill="transparent" isAnimationActive={false} />
              <Bar dataKey="durationHrs" stackId="a" radius={[4, 4, 4, 4]} isAnimationActive={false} shape={DurationBarShape}>
                {events.map((e) => <Cell key={e.id} fill={e.kind === "scheduled" ? "transparent" : e.completed ? "#3F7D65" : "#B85450"} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
      <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 11.5, color: "#4B5659", flexWrap: "wrap" }}>
        <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#B85450", borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />In progress</span>
        <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#3F7D65", borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Completed</span>
        <span><span style={{ display: "inline-block", width: 12, height: 3, background: "#3F7D65", borderRadius: 2, marginRight: 6, verticalAlign: "middle" }} />Projected to Expected Up Time</span>
        <span><span style={{ display: "inline-block", width: 9, height: 9, background: "#E8A33D", transform: "rotate(45deg)", marginRight: 6, verticalAlign: "middle" }} />Scheduled, not booked down yet</span>
      </div>
    </div>
  );
}

function DowntimeSummaryPage({ assets, breakdowns, workOrders }) {
  // Lazy initializers so "now" is computed fresh at mount time, not once
  // at app load - since this page fully unmounts when you navigate away
  // (the sidebar swaps which page renders), coming back always resets to
  // a fresh window rather than remembering whatever was picked last time.
  // Centred on now (12h back, 12h ahead) rather than only trailing, so
  // upcoming scheduled jobs are visible without having to change the
  // range first.
  const [fromDateTime, setFromDateTime] = useState(() => toDatetimeLocalValue(new Date(Date.now() - 12 * 3600000)));
  const [toDateTime, setToDateTime] = useState(() => toDatetimeLocalValue(new Date(Date.now() + 12 * 3600000)));
  const [showPrint, setShowPrint] = useState(false);
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const rows = useMemo(() => {
    const all = buildDowntimeRows(breakdowns, workOrders, fromDateTime, toDateTime);
    return all.filter((r) => {
      if (selectedAsset && r.asset_id !== selectedAsset) return false;
      if (selectedFleet && !selectedAsset) {
        const asset = assets.find((a) => a.asset_id === r.asset_id);
        if (!asset || asset.fleet !== selectedFleet) return false;
      }
      return true;
    });
  }, [breakdowns, workOrders, fromDateTime, toDateTime, assets, selectedFleet, selectedAsset]);

  const formatDT = (v) => v ? new Date(v).toLocaleString("en-ZA", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "-";

  const [exporting, setExporting] = useState(false);

  const exportToExcel = async () => {
    setExporting(true);
    try {
      // Loaded on demand rather than bundled into the main app - this
      // library is only needed the moment someone clicks Export, and
      // pulling it in upfront would add real weight to every page load,
      // including on a slow site connection.
      const { default: ExcelJS } = await import("exceljs");
      const wb = new ExcelJS.Workbook();
      wb.creator = "Fleet Tracker";
      wb.created = new Date();

      const colCount = DOWNTIME_SUMMARY_COLUMNS.length;
      const BRAND = "FF1F6668";
      const HEADER_FILL = "FFD9D9D9";
      const HEADER_TEXT = "FF000000";
      const BAND = "FFF7F8F6";
      const BORDER = "FF808080";
      const thinBorder = {
        top: { style: "thin", color: { argb: BORDER } },
        bottom: { style: "thin", color: { argb: BORDER } },
        left: { style: "thin", color: { argb: BORDER } },
        right: { style: "thin", color: { argb: BORDER } },
      };

      const ws = wb.addWorksheet("Downtime Summary");

      // Title
      ws.mergeCells(1, 1, 1, colCount);
      const titleCell = ws.getCell(1, 1);
      titleCell.value = "Shift Downtime Summary";
      titleCell.font = { size: 20, bold: true, color: { argb: BRAND } };
      titleCell.alignment = { vertical: "middle" };
      ws.getRow(1).height = 32;

      // Info block - equipment scope, date range, export timestamp
      const scopeLabel = selectedAsset || selectedFleet || "All equipment";
      const infoRows = [
        ["Fleet / Equipment", scopeLabel],
        ["Date Range", formatRangeForDisplay(fromDateTime, toDateTime)],
        ["Exported", new Date().toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" })],
      ];
      let r = 2;
      infoRows.forEach(([label, value]) => {
        ws.getCell(r, 1).value = label;
        ws.getCell(r, 1).font = { bold: true, color: { argb: "FF4B5659" } };
        ws.mergeCells(r, 2, r, colCount);
        ws.getCell(r, 2).value = value;
        ws.getCell(r, 2).font = { color: { argb: "FF183642" } };
        r++;
      });

      // Thick underline separating the info block from the table, echoing
      // the double blue rule under the title on the reference report.
      for (let c = 1; c <= colCount; c++) {
        ws.getCell(r, c).border = { bottom: { style: "medium", color: { argb: BRAND } } };
      }
      r += 2;

      const headerRowIdx = r;
      DOWNTIME_SUMMARY_COLUMNS.forEach(([, label], i) => {
        const cell = ws.getCell(headerRowIdx, i + 1);
        cell.value = label;
        cell.font = { bold: true, color: { argb: HEADER_TEXT } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
        cell.alignment = { vertical: "middle" };
        cell.border = thinBorder;
      });
      ws.getRow(headerRowIdx).height = 20;

      const STATUS_COLOR = { Closed: "FF2C5646", Completed: "FF2C5646", "In Progress": "FF7A3330", "Awaiting Parts": "FF7A5320" };
      rows.forEach((row, idx) => {
        const rowIdx = headerRowIdx + 1 + idx;
        DOWNTIME_SUMMARY_COLUMNS.forEach(([key], colIdx) => {
          const cell = ws.getCell(rowIdx, colIdx + 1);
          cell.value =
            key === "downtime_start" || key === "downtime_end" ? formatDT(row[key])
            : key === "downtime_hours" ? (row[key] != null ? Number(row[key]).toFixed(2) : "-")
            : (row[key] ?? "-");
          cell.border = thinBorder;
          if (idx % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BAND } };
          if (key === "status" && STATUS_COLOR[row[key]]) cell.font = { bold: true, color: { argb: STATUS_COLOR[row[key]] } };
          if (key === "event") cell.font = { bold: true, color: { argb: row[key] === "Breakdown" ? "FF7A3330" : "FF1F6668" } };
        });
      });

      DOWNTIME_SUMMARY_COLUMNS.forEach(([key, label], i) => {
        ws.getColumn(i + 1).width = key === "description" ? 34 : Math.max(label.length + 4, 16);
      });
      ws.views = [{ state: "frozen", ySplit: headerRowIdx }];

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Shift_Downtime_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } finally {
      setExporting(false);
    }
  };

  return (
    <div>
      <EventTimeline breakdowns={breakdowns} workOrders={workOrders} fromDateTime={fromDateTime} toDateTime={toDateTime} />

      <DateRangePicker fromDateTime={fromDateTime} toDateTime={toDateTime} setFromDateTime={setFromDateTime} setToDateTime={setToDateTime} />
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginBottom: 12 }}>
        <button onClick={exportToExcel} disabled={exporting} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: exporting ? "default" : "pointer", opacity: exporting ? 0.6 : 1 }}>
          <Download size={14} /> {exporting ? "Preparing…" : "Export to Excel"}
        </button>
        <button onClick={() => setShowPrint(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: NAVY, border: "none", color: "#fff", fontSize: 13, fontWeight: 700, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
          <Printer size={14} /> Print for Sign-off
        </button>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {DOWNTIME_SUMMARY_COLUMNS.map(([key, label]) => (
                <th key={key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} style={{ borderBottom: i < rows.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                {DOWNTIME_SUMMARY_COLUMNS.map(([key]) => (
                  <td key={key} style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                    {key === "event" || key === "status" ? <Badge value={r[key]} />
                      : key === "downtime_start" || key === "downtime_end" ? formatDT(r[key])
                      : key === "downtime_hours" ? (r[key] != null ? Number(r[key]).toFixed(2) : "-")
                      : (r[key] ?? <span style={{ color: "#B4B2A9" }}>-</span>)}
                  </td>
                ))}
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={DOWNTIME_SUMMARY_COLUMNS.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No downtime events in this date range.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showPrint && (
        <div className="job-card-modal-backdrop" style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60, padding: 16 }}>
          <style>{`
            @media print {
              body * { visibility: hidden; }
              .job-card-print, .job-card-print * { visibility: visible; }
              .job-card-print { position: absolute; top: 0; left: 0; width: 100%; }
              .job-card-modal-backdrop { position: static !important; background: none !important; padding: 0 !important; }
              .job-card-no-print { display: none !important; }
            }
          `}</style>
          <div style={{ background: "#fff", borderRadius: 12, padding: 28, width: 900, maxWidth: "100%", maxHeight: "92vh", overflowY: "auto" }}>
            <div className="job-card-no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <p style={{ fontSize: 15, fontWeight: 700, color: NAVY, margin: 0 }}>Shift Downtime Summary</p>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => setShowPrint(false)} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "8px 14px", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>Close</button>
                <button onClick={() => window.print()} style={{ display: "flex", alignItems: "center", gap: 6, background: NAVY, border: "none", color: "#fff", padding: "8px 14px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  <Printer size={14} /> Print / Save as PDF
                </button>
              </div>
            </div>

            <div className="job-card-print">
              <p style={{ fontSize: 18, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Shift Downtime Summary</p>
              <p style={{ fontSize: 12.5, color: "#4B5659", margin: "0 0 16px" }}>Date Range: {formatRangeForDisplay(fromDateTime, toDateTime)}</p>

              <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 30, fontSize: 11.5 }}>
                <thead>
                  <tr>
                    {DOWNTIME_SUMMARY_COLUMNS.map(([key, label]) => (
                      <th key={key} style={{ textAlign: "left", padding: "5px 8px", border: "1px solid #ccc", background: "#F2F1EA" }}>{label}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={i}>
                      {DOWNTIME_SUMMARY_COLUMNS.map(([key]) => (
                        <td key={key} style={{ padding: "5px 8px", border: "1px solid #ccc" }}>
                          {key === "downtime_start" || key === "downtime_end" ? formatDT(r[key])
                            : key === "downtime_hours" ? (r[key] != null ? Number(r[key]).toFixed(2) : "-")
                            : (r[key] ?? "-")}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>

              <p style={{ fontSize: 14, fontWeight: 700, color: "#183642", margin: "0 0 20px" }}>Sign Off</p>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 40 }}>
                <div>
                  <p style={{ fontSize: 12, margin: "0 0 30px" }}>Maintenance (Name): _______________________</p>
                  <p style={{ fontSize: 12, margin: 0 }}>Maintenance (Signature): _______________________</p>
                </div>
                <div>
                  <p style={{ fontSize: 12, margin: "0 0 30px" }}>Production (Name): _______________________</p>
                  <p style={{ fontSize: 12, margin: 0 }}>Production (Signature): _______________________</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function WorkOrdersPage({ assets, workOrders, userEmail, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [printing, setPrinting] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");

  const columns = [
    ["wo_no", "Work order #"], ["asset_id", "Equipment #"], ["work_type", "Type"],
    ["priority", "Priority"], ["status", "Status"], ["request_date", "Requested"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    let rows = workOrders;
    if (query.trim()) {
      const q = query.toLowerCase();
      rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return [...rows].sort((a, b) => (b.request_date || "").localeCompare(a.request_date || ""));
  }, [workOrders, query]);

  const handleSaved = () => { setShowForm(false); setEditing(null); onRefresh(); };

  const handleDelete = async (reason) => {
    await deleteWithReason("work_orders", deleting.id, "id", reason, userEmail);
    onRefresh();
    setDeleting(null);
  };

  const WORK_ORDER_FIELDS = [
    { key: "work_type", header: "Work Type", type: "text" },
    { key: "priority", header: "Priority", type: "text" },
    { key: "problem_scope", header: "Problem / Scope", type: "text" },
    { key: "component", header: "Component", type: "text" },
    { key: "status", header: "Status", type: "text" },
    { key: "planned_start", header: "Planned Start (YYYY-MM-DD HH:MM)", type: "datetime" },
    { key: "actual_start", header: "Actual Start (YYYY-MM-DD HH:MM)", type: "datetime" },
    { key: "actual_finish", header: "Actual Finish (YYYY-MM-DD HH:MM)", type: "datetime" },
    { key: "technician_vendor", header: "Technician / Vendor", type: "text" },
    { key: "closeout_notes", header: "Closeout Notes", type: "text" },
  ];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search work orders"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
          + Add Work Order
        </button>
      </div>
      <ExcelSync
        data={workOrders} assets={assets} fields={WORK_ORDER_FIELDS} tableName="work_orders"
        sheetTitle="Work Orders" filenamePrefix="Work_Orders" onRefresh={onRefresh}
        extraOnSave={(obj) => (obj.id ? {} : { request_date: todayForInput() })}
      />

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => <th key={c.key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{c.label}</th>)}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr key={row.id ?? i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                {columns.map((c) => (
                  <td key={c.key} onClick={() => { setEditing(row); setShowForm(true); }} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>
                    {c.key === "status" || c.key === "priority" ? <Badge value={row[c.key]} /> : (row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>)}
                  </td>
                ))}
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  <button onClick={() => setPrinting(row)} title="Print Job Card" style={{ background: "none", border: "none", color: NAVY, cursor: "pointer", padding: 4, display: "inline-flex", marginRight: 4 }}>
                    <Printer size={15} />
                  </button>
                  <button onClick={() => setDeleting(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {workOrders.length === 0 ? "No work orders yet." : "No work orders match your search."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <WorkOrderForm assets={assets} existing={editing} onClose={() => { setShowForm(false); setEditing(null); }} onSaved={handleSaved} />
      )}

      {printing && (
        <JobCardPrintModal
          workOrder={printing}
          asset={assets.find((a) => a.asset_id === printing.asset_id)}
          onClose={() => setPrinting(null)}
          onUploaded={() => { onRefresh(); }}
        />
      )}

      {deleting && (
        <DeleteConfirmModal
          itemLabel={deleting.wo_no}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

function ServiceScheduleForm({ assets, existing, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [assignedTechnician, setAssignedTechnician] = useState(existing?.assigned_technician || "");
  const [scheduledDate, setScheduledDate] = useState(existing?.scheduled_date || "");
  const [autoGenerate, setAutoGenerate] = useState(existing?.auto_generate ?? true);
  const [defaultPriority, setDefaultPriority] = useState(existing?.default_priority || "Medium");
  const [defaultScope, setDefaultScope] = useState(existing?.default_scope || "");

  const [useHours, setUseHours] = useState(existing ? existing.service_interval != null : true);
  const [serviceInterval, setServiceInterval] = useState(existing?.service_interval ?? 250);

  const [useCalendar, setUseCalendar] = useState(existing?.calendar_interval_days != null);
  const [calendarIntervalDays, setCalendarIntervalDays] = useState(existing?.calendar_interval_days ?? 90);
  const [lastServiceDate, setLastServiceDate] = useState(existing?.last_service_date || todayForInput());

  const [useCounter, setUseCounter] = useState(existing?.counter_interval != null);
  const [counterLabel, setCounterLabel] = useState(existing?.counter_label || "");
  const [counterInterval, setCounterInterval] = useState(existing?.counter_interval ?? "");
  const [currentCounter, setCurrentCounter] = useState(existing?.current_counter ?? "");
  const [lastServiceCounter, setLastServiceCounter] = useState(existing?.last_service_counter ?? "");

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  if (assets.length === 0) {
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 360, maxWidth: "100%", textAlign: "center" }}>
          <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>No equipment added yet</p>
          <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>Add at least one asset on the Assets tab first.</p>
          <button onClick={onClose} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Close</button>
        </div>
      </div>
    );
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");

    if (!useHours && !useCalendar && !useCounter) {
      setError("Turn on at least one trigger - Hours, Calendar, or Usage Counter.");
      return;
    }

    setSaving(true);
    const asset = assets.find((a) => a.asset_id === assetId);
    const payload = {
      asset_id: assetId,
      assigned_technician: assignedTechnician || null,
      scheduled_date: scheduledDate || null,
      auto_generate: autoGenerate,
      default_priority: defaultPriority,
      default_scope: defaultScope || null,

      service_interval: useHours ? Number(serviceInterval) : null,
      ...(isEdit ? {} : { last_service_hours: asset?.current_hours ?? 0 }),

      calendar_interval_days: useCalendar ? Number(calendarIntervalDays) : null,
      last_service_date: useCalendar ? lastServiceDate : null,

      counter_label: useCounter ? (counterLabel || null) : null,
      counter_interval: useCounter && counterInterval !== "" ? Number(counterInterval) : null,
      current_counter: useCounter && currentCounter !== "" ? Number(currentCounter) : null,
      last_service_counter: useCounter && lastServiceCounter !== "" ? Number(lastServiceCounter) : (useCounter ? 0 : null),
    };

    try {
      const { error: dbError } = isEdit
        ? await supabase.from("service_schedule").update(payload).eq("id", existing.id)
        : await supabase.from("service_schedule").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };
  const triggerBox = { border: "1px solid #E2E6E3", borderRadius: 8, padding: "10px 12px", marginBottom: 12 };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 480, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>
          {isEdit ? "Edit Service Schedule" : "Add Service Schedule"}
        </h3>
        <p style={{ fontSize: 12, color: "#859195", margin: "0 0 16px" }}>
          Turn on any combination of triggers below - whichever one is reached first generates the work order.
        </p>

        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>

        <div style={triggerBox}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: NAVY, marginBottom: useHours ? 10 : 0, cursor: "pointer" }}>
            <input type="checkbox" checked={useHours} onChange={(e) => setUseHours(e.target.checked)} style={{ width: 16, height: 16 }} />
            Runtime Hours Trigger
          </label>
          {useHours && (
            <div>
              <label style={labelStyle}>Service Interval (hours)</label>
              <input type="number" step="1" value={serviceInterval} onChange={(e) => setServiceInterval(e.target.value)} required={useHours} style={fieldStyle} />
            </div>
          )}
        </div>

        <div style={triggerBox}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: NAVY, marginBottom: useCalendar ? 10 : 0, cursor: "pointer" }}>
            <input type="checkbox" checked={useCalendar} onChange={(e) => setUseCalendar(e.target.checked)} style={{ width: 16, height: 16 }} />
            Calendar Trigger
          </label>
          {useCalendar && (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <div>
                <label style={labelStyle}>Interval (days)</label>
                <input type="number" step="1" value={calendarIntervalDays} onChange={(e) => setCalendarIntervalDays(e.target.value)} required={useCalendar} style={fieldStyle} />
              </div>
              <div>
                <label style={labelStyle}>Last Serviced</label>
                <input type="date" value={lastServiceDate} onChange={(e) => setLastServiceDate(e.target.value)} required={useCalendar} style={fieldStyle} />
              </div>
            </div>
          )}
        </div>

        <div style={triggerBox}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: NAVY, marginBottom: useCounter ? 10 : 0, cursor: "pointer" }}>
            <input type="checkbox" checked={useCounter} onChange={(e) => setUseCounter(e.target.checked)} style={{ width: 16, height: 16 }} />
            Usage Counter Trigger <span style={{ fontWeight: 400, color: "#859195" }}>(cycles, rotations, anything numeric besides hours)</span>
          </label>
          {useCounter && (
            <div>
              <div style={{ marginBottom: 10 }}>
                <label style={labelStyle}>Counter Name</label>
                <input type="text" value={counterLabel} onChange={(e) => setCounterLabel(e.target.value)} placeholder="e.g. Bucket Cycles" required={useCounter} style={fieldStyle} />
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <div>
                  <label style={labelStyle}>Interval</label>
                  <input type="number" step="1" value={counterInterval} onChange={(e) => setCounterInterval(e.target.value)} required={useCounter} style={fieldStyle} />
                </div>
                <div>
                  <label style={labelStyle}>Current Reading</label>
                  <input type="number" step="1" value={currentCounter} onChange={(e) => setCurrentCounter(e.target.value)} required={useCounter} style={fieldStyle} />
                </div>
              </div>
            </div>
          )}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Assigned Technician</label>
            <input type="text" value={assignedTechnician} onChange={(e) => setAssignedTechnician(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Manually Scheduled Date <span style={{ fontWeight: 400, color: "#859195" }}>(optional)</span></label>
            <input type="date" value={scheduledDate} onChange={(e) => setScheduledDate(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Auto-generated WO Priority</label>
            <select value={defaultPriority} onChange={(e) => setDefaultPriority(e.target.value)} style={fieldStyle}>
              {WORK_ORDER_PRIORITIES.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Scope Description</label>
            <input type="text" value={defaultScope} onChange={(e) => setDefaultScope(e.target.value)} placeholder="e.g. 250hr Service" style={fieldStyle} />
          </div>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#183642", marginBottom: 18, cursor: "pointer" }}>
          <input type="checkbox" checked={autoGenerate} onChange={(e) => setAutoGenerate(e.target.checked)} style={{ width: 16, height: 16 }} />
          Automatically create a work order when due
        </label>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Add Schedule"}
          </button>
        </div>
      </form>
    </div>
  );
}

function PlannedMaintenancePage({ assets, plannedMaintenance, workOrders, userEmail, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [servicing, setServicing] = useState(null); // row being confirmed for "Mark Serviced"

  const [showJobForm, setShowJobForm] = useState(false);
  const [editingJob, setEditingJob] = useState(null);
  const [printingJob, setPrintingJob] = useState(null);
  const [deletingJob, setDeletingJob] = useState(null);

  const [pmCheckMessage, setPmCheckMessage] = useState(null);
  const [pmChecking, setPmChecking] = useState(false);

  const plannedJobs = useMemo(
    () => [...workOrders].filter((w) => w.work_type === "Preventive").sort((a, b) => (b.request_date || "").localeCompare(a.request_date || "")),
    [workOrders]
  );

  const checkForDuePM = React.useCallback(async (silent) => {
    setPmChecking(true);
    const { data, error } = await supabase.rpc("check_and_generate_pm_work_orders");
    setPmChecking(false);
    if (error) {
      if (!silent) setPmCheckMessage({ type: "error", text: error.message });
      return;
    }
    const generated = data || [];
    if (generated.length > 0) {
      setPmCheckMessage({
        type: "success",
        text: `${generated.length} work order${generated.length === 1 ? "" : "s"} auto-generated: ${generated.map((g) => `${g.wo_no} (${g.asset_id})`).join(", ")}`,
      });
      onRefresh();
    } else if (!silent) {
      setPmCheckMessage({ type: "info", text: "Nothing due right now." });
    }
  }, [onRefresh]);

  // Check automatically whenever this page loads - this is what makes PM
  // generation feel automatic rather than something someone has to
  // remember to trigger. The manual button below covers checking again
  // without leaving and re-entering the page.
  useEffect(() => { checkForDuePM(true); }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const handleMarkServiced = async (row) => {
    const asset = assets.find((a) => a.asset_id === row.asset_id);
    if (row.service_interval != null && (!asset || asset.current_hours == null)) {
      alert("Can't mark this serviced - no current hours on record for this asset yet.");
      return;
    }
    const update = {};
    if (row.service_interval != null) update.last_service_hours = asset.current_hours;
    if (row.calendar_interval_days != null) update.last_service_date = todayForInput();
    if (row.counter_interval != null) update.last_service_counter = row.current_counter;

    const { error } = await supabase.from("service_schedule").update(update).eq("id", row.id);
    if (error) {
      alert("Couldn't save: " + error.message);
    } else {
      onRefresh();
    }
    setServicing(null);
  };

  const handleDeleteJob = async (reason) => {
    await deleteWithReason("work_orders", deletingJob.id, "id", reason, userEmail);
    onRefresh();
    setDeletingJob(null);
  };

  return (
    <div>
      {pmCheckMessage && (
        <div style={{
          background: pmCheckMessage.type === "error" ? "#F6E2E0" : pmCheckMessage.type === "success" ? "#E2EFE9" : "#F2F1EA",
          border: `1px solid ${pmCheckMessage.type === "error" ? "#DDB6B2" : pmCheckMessage.type === "success" ? "#B7D89A" : "#E2E6E3"}`,
          color: pmCheckMessage.type === "error" ? "#7A3330" : pmCheckMessage.type === "success" ? "#2C5646" : "#4B5659",
          borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13,
        }}>
          {pmCheckMessage.text}
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <h3 style={{ fontSize: 15, fontWeight: 700, margin: 0, color: NAVY }}>Service Intervals</h3>
        <button onClick={() => checkForDuePM(false)} disabled={pmChecking} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "6px 12px", borderRadius: 8, cursor: pmChecking ? "default" : "pointer" }}>
          {pmChecking ? "Checking…" : "Check for Due Maintenance"}
        </button>
      </div>
      <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 16px" }}>
        Overdue hours keep counting even when past due - nothing is capped or hidden. Due maintenance is checked automatically every time this page loads, and generates a work order under Planned Maintenance Jobs below. Use "Mark Serviced" once work is done to reset the interval and start the next count from now.
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
          + Add Service Schedule
        </button>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10, marginBottom: 32 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Equipment #", "Name", "Current hours", "Triggers", "Hours remaining", "Status", "Technician", ""].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {plannedMaintenance.map((row, i) => (
              <tr key={row.id ?? i} style={{ borderBottom: i < plannedMaintenance.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px", cursor: "pointer" }} onClick={() => { setEditing(row); setShowForm(true); }}>{row.asset_id}</td>
                <td style={{ padding: "9px 12px" }}>{row.asset_name}</td>
                <td style={{ padding: "9px 12px" }}>{row.current_hours ?? "-"}</td>
                <td style={{ padding: "9px 12px", fontSize: 12 }}>
                  {[
                    row.service_interval != null ? `Hours: every ${row.service_interval}` : null,
                    row.calendar_interval_days != null ? `Calendar: every ${row.calendar_interval_days}d` : null,
                    row.counter_interval != null ? `${row.counter_label || "Counter"}: every ${row.counter_interval}` : null,
                  ].filter(Boolean).join(" · ") || <span style={{ color: "#B4B2A9" }}>-</span>}
                </td>
                <td style={{ padding: "9px 12px", color: row.remaining != null && row.remaining < 0 ? "#B85450" : "#183642", fontWeight: row.remaining != null && row.remaining < 0 ? 700 : 400 }}>
                  {row.remaining != null ? row.remaining.toFixed(1) : "-"}
                </td>
                <td style={{ padding: "9px 12px" }}><Badge value={row.status} /></td>
                <td style={{ padding: "9px 12px" }}>{row.assigned_technician ?? "-"}</td>
                <td style={{ padding: "9px 12px" }}>
                  <button onClick={() => setServicing(row)} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12, fontWeight: 600, padding: "5px 10px", borderRadius: 6, cursor: "pointer", whiteSpace: "nowrap" }}>
                    Mark Serviced
                  </button>
                </td>
              </tr>
            ))}
            {plannedMaintenance.length === 0 && (
              <tr><td colSpan={8} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No service schedules set up yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 6px", color: NAVY }}>Planned Maintenance Jobs</h3>
      <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 16px" }}>
        Add a specific planned maintenance task, print its Job Card for the technician to take on site, then upload the completed (signed) scan back here once it's done.
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <button onClick={() => { setEditingJob(null); setShowJobForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
          + Add Planned Maintenance
        </button>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Work order #", "Equipment #", "Scope", "Priority", "Status", "Planned Start", ""].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {plannedJobs.map((row, i) => (
              <tr key={row.id ?? i} style={{ borderBottom: i < plannedJobs.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px", cursor: "pointer" }} onClick={() => { setEditingJob(row); setShowJobForm(true); }}>{row.wo_no}</td>
                <td style={{ padding: "9px 12px" }}>{row.asset_id}</td>
                <td style={{ padding: "9px 12px", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.problem_scope}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={row.priority} /></td>
                <td style={{ padding: "9px 12px" }}><Badge value={row.status} /></td>
                <td style={{ padding: "9px 12px" }}>{row.planned_start ?? "-"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  <button onClick={() => setPrintingJob(row)} title="Print Job Card" style={{ background: "none", border: "none", color: NAVY, cursor: "pointer", padding: 4, display: "inline-flex", marginRight: 4 }}>
                    <Printer size={15} />
                  </button>
                  <button onClick={() => setDeletingJob(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
            {plannedJobs.length === 0 && (
              <tr><td colSpan={7} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No planned maintenance jobs yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <ServiceScheduleForm assets={assets} existing={editing}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); onRefresh(); }} />
      )}

      {showJobForm && (
        <WorkOrderForm assets={assets} existing={editingJob} defaultWorkType="Preventive"
          onClose={() => { setShowJobForm(false); setEditingJob(null); }}
          onSaved={() => { setShowJobForm(false); setEditingJob(null); onRefresh(); }} />
      )}

      {printingJob && (
        <JobCardPrintModal
          workOrder={printingJob}
          asset={assets.find((a) => a.asset_id === printingJob.asset_id)}
          onClose={() => setPrintingJob(null)}
          onUploaded={() => onRefresh()}
        />
      )}

      {deletingJob && (
        <DeleteConfirmModal
          itemLabel={deletingJob.wo_no}
          userEmail={userEmail}
          onCancel={() => setDeletingJob(null)}
          onConfirm={handleDeleteJob}
        />
      )}

      {servicing && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 380, maxWidth: "100%" }}>
            <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>Mark {servicing.asset_id} as serviced?</p>
            <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>
              This resets the interval - Last Service Hours becomes the asset's current hours ({assets.find((a) => a.asset_id === servicing.asset_id)?.current_hours ?? "-"}), and the next due point starts counting from here.
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button onClick={() => setServicing(null)} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
              <button onClick={() => handleMarkServiced(servicing)} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Confirm</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function AssetForm({ existing, selectedSiteId, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || "");
  const [assetName, setAssetName] = useState(existing?.asset_name || "");
  const [make, setMake] = useState(existing?.make || "");
  const [model, setModel] = useState(existing?.model || "");
  const [fleet, setFleet] = useState(existing?.fleet || "");
  const [serialNumber, setSerialNumber] = useState(existing?.serial_number || "");
  const [status, setStatus] = useState(existing?.status || "Operating");
  const [siteLocation, setSiteLocation] = useState(existing?.site_location || "");
  const [openingHours, setOpeningHours] = useState(existing?.opening_hours ?? "");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSaving(true);

    const payload = {
      asset_id: assetId, asset_name: assetName, make, model, fleet,
      serial_number: serialNumber || null, status, site_location: siteLocation || null,
      opening_hours: openingHours === "" ? null : Number(openingHours),
      notes: notes || null,
      ...(isEdit ? {} : { site_id: selectedSiteId }),
    };

    try {
      const { error: dbError } = isEdit
        ? await supabase.from("assets").update(payload).eq("asset_id", existing.asset_id)
        : await supabase.from("assets").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 480, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>
          {isEdit ? "Edit Asset" : "Add Asset"}
        </h3>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Equipment #</label>
            <input type="text" value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} placeholder="EQ-008" style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }} />
          </div>
          <div>
            <label style={labelStyle}>Name</label>
            <input type="text" value={assetName} onChange={(e) => setAssetName(e.target.value)} required style={fieldStyle} />
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Make</label>
            <input type="text" value={make} onChange={(e) => setMake(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Model</label>
            <input type="text" value={model} onChange={(e) => setModel(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Fleet</label>
            <input type="text" value={fleet} onChange={(e) => setFleet(e.target.value)} required placeholder="e.g. CAT 793F" style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Serial Number</label>
            <input type="text" value={serialNumber} onChange={(e) => setSerialNumber(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 6 }}>
          <div>
            <label style={labelStyle}>Status</label>
            <select value={status} onChange={(e) => setStatus(e.target.value)} style={fieldStyle}>
              {["Operating", "Under Maintenance", "Breakdown", "Standby", "Disposed"].map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Site / Location</label>
            <input type="text" value={siteLocation} onChange={(e) => setSiteLocation(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Current Hours (at intake)</label>
          <input type="number" step="0.1" value={openingHours} onChange={(e) => setOpeningHours(e.target.value)} placeholder="e.g. 12450" style={fieldStyle} />
          <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
            The hour-meter reading when this machine was loaded into the system. Shows as Current Hours immediately - once Daily Hours entries start coming in for this asset, those take over automatically.
          </p>
        </div>

        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Notes</label>
          <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows={2} style={{ ...fieldStyle, resize: "vertical" }} />
        </div>

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>
            Cancel
          </button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Add Asset"}
          </button>
        </div>
      </form>
    </div>
  );
}

function PartForm({ existing, selectedSiteId, onClose, onSaved }) {
  const isEdit = !!existing;
  const [partNo, setPartNo] = useState(existing?.part_no || "");
  const [description, setDescription] = useState(existing?.description || "");
  const [qtyInStock, setQtyInStock] = useState(existing?.qty_in_stock ?? 0);
  const [minimumQty, setMinimumQty] = useState(existing?.minimum_qty ?? 0);
  const [supplier, setSupplier] = useState(existing?.supplier || "");
  const [unitCost, setUnitCost] = useState(existing?.unit_cost ?? "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSaving(true);
    const payload = {
      part_no: partNo, description, qty_in_stock: Number(qtyInStock),
      minimum_qty: Number(minimumQty), supplier: supplier || null,
      unit_cost: unitCost === "" ? 0 : Number(unitCost),
      ...(isEdit ? {} : { site_id: selectedSiteId }),
    };
    try {
      const { error: dbError } = isEdit
        ? await supabase.from("parts_inventory").update(payload).eq("id", existing.id)
        : await supabase.from("parts_inventory").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 420, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>{isEdit ? "Edit Part" : "Add Part"}</h3>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Part No</label>
          <input type="text" value={partNo} onChange={(e) => setPartNo(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Description</label>
          <input type="text" value={description} onChange={(e) => setDescription(e.target.value)} style={fieldStyle} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Qty In Stock</label>
            <input type="number" step="1" value={qtyInStock} onChange={(e) => setQtyInStock(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Minimum Qty</label>
            <input type="number" step="1" value={minimumQty} onChange={(e) => setMinimumQty(e.target.value)} style={fieldStyle} />
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 18 }}>
          <div>
            <label style={labelStyle}>Supplier</label>
            <input type="text" value={supplier} onChange={(e) => setSupplier(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Unit Cost</label>
            <input type="number" step="0.01" value={unitCost} onChange={(e) => setUnitCost(e.target.value)} style={fieldStyle} />
          </div>
        </div>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : isEdit ? "Save Changes" : "Add Part"}
          </button>
        </div>
      </form>
    </div>
  );
}

const PARTS_IMPORT_HEADERS = ["Part No", "Description", "Qty In Stock", "Minimum Qty", "Supplier", "Unit Cost"];
const PARTS_IMPORT_KEYS = ["part_no", "description", "qty_in_stock", "minimum_qty", "supplier", "unit_cost"];

const TABLE_NAME_LABELS = {
  assets: "Assets", daily_hours: "Daily Hours", fuel_log: "Fuel Log", oil_consumption: "Oil Consumption",
  breakdown_log: "Events", work_orders: "Work Orders", service_schedule: "Planned Maintenance",
  inspections: "Inspections", component_types: "Components", component_changeouts: "Components",
  tyre_tracking: "Tyres", parts_inventory: "Parts Inventory", warranty_register: "Warranty & Documents",
  document_register: "Warranty & Documents", scheduled_hours: "Monthly Scheduled Hours",
  equipment_class_cost_rates: "Cost Rates", shift_reports: "Shift Reports", technician_job_cards: "Job Cards",
  backlogs: "Backlogs", daily_service_checklist: "Daily Service", quote_prices: "Quote Price List",
};

function ComponentCodesAdminPage({ componentCodes, isAdmin, onRefresh }) {
  const [newName, setNewName] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [deleting, setDeleting] = useState(null);

  const handleAdd = async (e) => {
    e.preventDefault();
    setError("");
    if (!newName.trim()) return;
    setSaving(true);
    try {
      const { error: dbError } = await supabase.from("component_codes").insert({ name: newName.trim() });
      if (dbError) throw dbError;
      setNewName("");
      onRefresh();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (id) => {
    setDeleting(id);
    await supabase.from("component_codes").delete().eq("id", id);
    onRefresh();
    setDeleting(null);
  };

  if (!isAdmin) {
    return <p style={{ fontSize: 13.5, color: "#4B5659" }}>Only an admin can manage component codes.</p>;
  }

  return (
    <div style={{ maxWidth: 480 }}>
      <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 16px" }}>
        These names populate the Component Code dropdown on the Log Event form. Add whatever the team uses - Engine, Transmission, LH Wheel Station, and so on.
      </p>
      <form onSubmit={handleAdd} style={{ display: "flex", gap: 8, marginBottom: 18 }}>
        <input
          value={newName}
          onChange={(e) => setNewName(e.target.value)}
          placeholder="e.g. Torque Converter"
          style={{ flex: 1, padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box" }}
        />
        <button type="submit" disabled={saving || !newName.trim()} style={{ background: NAVY, border: "none", color: "#fff", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
          {saving ? "Adding…" : "+ Add"}
        </button>
      </form>
      {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

      <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "hidden" }}>
        {componentCodes.length === 0 ? (
          <p style={{ padding: 16, fontSize: 13, color: "#859195", margin: 0 }}>No component codes yet - add the first one above.</p>
        ) : componentCodes.map((c, i) => (
          <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 14px", borderBottom: i < componentCodes.length - 1 ? "1px solid #EFEEE7" : "none" }}>
            <span style={{ fontSize: 13.5, color: "#183642" }}>{c.name}</span>
            <button onClick={() => handleDelete(c.id)} disabled={deleting === c.id} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "flex" }}>
              <Trash2 size={15} />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function SiteManagementPage({ isAdmin, onSitesChanged, onNameSaved }) {
  const [sites, setSites] = useState([]);
  const [users, setUsers] = useState([]);
  const [access, setAccess] = useState([]);
  const [roles, setRoles] = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [newSiteName, setNewSiteName] = useState("");
  const [newSiteLocation, setNewSiteLocation] = useState("");
  const [savingSite, setSavingSite] = useState(false);
  const [message, setMessage] = useState(null);
  const [managingUser, setManagingUser] = useState(null);
  const [editingNameFor, setEditingNameFor] = useState(null);
  const [nameInput, setNameInput] = useState("");
  const [savingName, setSavingName] = useState(false);

  const loadAll = React.useCallback(async () => {
    setLoading(true);
    const [sitesRes, usersRes, accessRes, rolesRes, profilesRes] = await Promise.all([
      supabase.rpc("list_all_sites_for_admin"),
      supabase.rpc("list_users_for_admin"),
      supabase.rpc("list_site_access_for_admin"),
      supabase.rpc("list_user_roles_for_admin"),
      supabase.from("profiles").select("*"),
    ]);
    if (!sitesRes.error) setSites(sitesRes.data || []);
    if (!usersRes.error) setUsers(usersRes.data || []);
    if (!accessRes.error) setAccess(accessRes.data || []);
    if (!rolesRes.error) setRoles(rolesRes.data || []);
    if (!profilesRes.error) setProfiles(profilesRes.data || []);
    setLoading(false);
  }, []);

  useEffect(() => { if (isAdmin) loadAll(); }, [isAdmin, loadAll]);

  if (!isAdmin) {
    return (
      <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 10, padding: 28, textAlign: "center" }}>
        <p style={{ fontWeight: 700, fontSize: 14, color: "#7A3330", margin: "0 0 6px" }}>Admins only</p>
        <p style={{ fontSize: 13, color: "#7A3330", margin: 0 }}>Site management is restricted to administrators.</p>
      </div>
    );
  }

  const handleAddSite = async (e) => {
    e.preventDefault();
    if (!newSiteName.trim()) return;
    setSavingSite(true);
    const { error } = await supabase.from("sites").insert({ site_name: newSiteName.trim(), location: newSiteLocation.trim() || null });
    setSavingSite(false);
    if (error) {
      setMessage({ type: "error", text: error.message });
    } else {
      setNewSiteName("");
      setNewSiteLocation("");
      setMessage({ type: "success", text: `Site "${newSiteName.trim()}" created. Grant users access to it below - it starts with no data and no one able to see it until you do.` });
      loadAll();
      onSitesChanged?.();
    }
  };

  const hasAccess = (userId, siteId) => access.some((a) => a.user_id === userId && a.site_id === siteId);

  const toggleAccess = async (userId, siteId) => {
    if (hasAccess(userId, siteId)) {
      await supabase.rpc("revoke_site_access", { target_user_id: userId, target_site_id: siteId });
    } else {
      await supabase.rpc("grant_site_access", { target_user_id: userId, target_site_id: siteId });
    }
    loadAll();
    onSitesChanged?.();
  };

  const getRole = (userId) => roles.find((r) => r.user_id === userId)?.role || "manager";

  const getName = (userId) => profiles.find((p) => p.user_id === userId)?.full_name || "";

  const handleSaveName = async (userId) => {
    setSavingName(true);
    const { error } = await supabase.from("profiles").upsert({ user_id: userId, full_name: nameInput.trim() || null, updated_at: new Date().toISOString() });
    setSavingName(false);
    if (error) {
      setMessage({ type: "error", text: error.message });
    } else {
      setEditingNameFor(null);
      loadAll();
      onNameSaved?.(); // in case the admin just edited their own name - refreshes the header immediately
    }
  };

  const handleRoleChange = async (userId, newRole) => {
    const { error } = await supabase.rpc("set_user_role", { target_user_id: userId, new_role: newRole });
    if (error) {
      setMessage({ type: "error", text: error.message });
    } else {
      loadAll();
    }
  };

  const fieldStyle = { padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };

  return (
    <div>
      {message && (
        <div style={{
          background: message.type === "error" ? "#F6E2E0" : "#E2EFE9",
          border: `1px solid ${message.type === "error" ? "#DDB6B2" : "#B7D89A"}`,
          color: message.type === "error" ? "#7A3330" : "#2C5646",
          borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13,
        }}>
          {message.text}
        </div>
      )}

      <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Sites</h3>
      <form onSubmit={handleAddSite} style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
        <div>
          <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#183642", margin: "0 0 4px" }}>Site Name</label>
          <input type="text" value={newSiteName} onChange={(e) => setNewSiteName(e.target.value)} placeholder="e.g. Maaden Mahd" required style={fieldStyle} />
        </div>
        <div>
          <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#183642", margin: "0 0 4px" }}>Location</label>
          <input type="text" value={newSiteLocation} onChange={(e) => setNewSiteLocation(e.target.value)} placeholder="Optional" style={fieldStyle} />
        </div>
        <button type="submit" disabled={savingSite} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: savingSite ? "default" : "pointer" }}>
          {savingSite ? "Adding…" : "+ Add Site"}
        </button>
      </form>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10, marginBottom: 28 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3" }}>Site</th>
              <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3" }}>Location</th>
              <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3" }}>Status</th>
            </tr>
          </thead>
          <tbody>
            {sites.map((s, i) => (
              <tr key={s.id} style={{ borderBottom: i < sites.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px", fontWeight: 600 }}>{s.site_name}</td>
                <td style={{ padding: "9px 12px" }}>{s.location || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={s.active ? "Active" : "Inactive"} /></td>
              </tr>
            ))}
            {sites.length === 0 && !loading && (
              <tr><td colSpan={3} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No sites yet - add one above.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 6px", color: NAVY }}>User Access</h3>
      <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 12px" }}>
        Tick a box to grant a user access to that site. A user with only one site never sees a site switcher - it only appears once someone has more than one, which is what makes them a manager in practice. A newly granted site starts completely blank for that user until data gets logged against it. Names are set here by you only - people can't edit their own, by design. Role controls which tabs they see at all: Manager sees everything, Operator sees only the data-entry tabs (Daily Hours, Fuel Log, Oil Consumption, Breakdowns, Work Orders, Inspections) - no dashboards or reports. Use "Manage pages" for finer control than that - tick exactly which pages a specific person can see, on top of their role. Admin is granted separately via SQL, not from here, to avoid accidentally locking yourself out.
      </p>
      {loading ? (
        <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
      ) : (
        <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#F7F8F6" }}>
                <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" }}>User</th>
                <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" }}>Name</th>
                <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" }}>Role</th>
                {sites.map((s) => (
                  <th key={s.id} style={{ textAlign: "center", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" }}>{s.site_name}</th>
                ))}
                <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" }}>Pages</th>
              </tr>
            </thead>
            <tbody>
              {users.map((u, i) => {
                const role = getRole(u.id);
                return (
                  <tr key={u.id} style={{ borderBottom: i < users.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{u.email}</td>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                      {editingNameFor === u.id ? (
                        <span style={{ display: "flex", gap: 4, alignItems: "center" }}>
                          <input
                            type="text"
                            autoFocus
                            value={nameInput}
                            onChange={(e) => setNameInput(e.target.value)}
                            placeholder="Full name and surname"
                            style={{ padding: "4px 7px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6, width: 150 }}
                          />
                          <button onClick={() => handleSaveName(u.id)} disabled={savingName} style={{ fontSize: 11.5, color: NAVY, background: "none", border: "none", cursor: "pointer", fontWeight: 600 }}>Save</button>
                          <button onClick={() => setEditingNameFor(null)} style={{ fontSize: 11.5, color: "#859195", background: "none", border: "none", cursor: "pointer" }}>Cancel</button>
                        </span>
                      ) : (
                        <span
                          onClick={() => { setEditingNameFor(u.id); setNameInput(getName(u.id)); }}
                          style={{ cursor: "pointer", textDecoration: "underline", textDecorationStyle: "dotted", color: getName(u.id) ? "#183642" : "#859195" }}
                        >
                          {getName(u.id) || "Set name"}
                        </span>
                      )}
                    </td>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                      {role === "admin" ? (
                        <Badge value="Admin" />
                      ) : (
                        <select
                          value={role}
                          onChange={(e) => handleRoleChange(u.id, e.target.value)}
                          style={{ padding: "5px 8px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', background: "#fff" }}
                        >
                          <option value="manager">Manager</option>
                          <option value="operator">Operator</option>
                        </select>
                      )}
                    </td>
                    {sites.map((s) => (
                      <td key={s.id} style={{ padding: "9px 12px", textAlign: "center" }}>
                        <input type="checkbox" checked={hasAccess(u.id, s.id)} onChange={() => toggleAccess(u.id, s.id)} style={{ width: 16, height: 16, cursor: "pointer" }} />
                      </td>
                    ))}
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                      {role === "admin" ? (
                        <span style={{ fontSize: 12, color: "#859195" }}>Full access</span>
                      ) : (
                        <button onClick={() => setManagingUser(u)} style={{ fontSize: 12, color: NAVY, background: "none", border: `1px solid ${NAVY}`, borderRadius: 6, padding: "4px 9px", cursor: "pointer", fontWeight: 600 }}>
                          Manage pages
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
              {users.length === 0 && (
                <tr><td colSpan={sites.length + 4} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No users found.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {managingUser && (
        <PageAccessModal user={managingUser} onClose={() => setManagingUser(null)} />
      )}
    </div>
  );
}

function PageAccessModal({ user, onClose }) {
  const tickablePages = NAV.filter((n) => !n.group && n.key !== "audit" && n.key !== "site_management" && n.key !== "component_codes");
  const [checked, setChecked] = useState(new Set(tickablePages.map((p) => p.key))); // default: everything ticked, until we know otherwise
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error: err } = await supabase.from("user_page_access").select("page_key").eq("user_id", user.id);
      if (cancelled) return;
      if (!err && data && data.length > 0) {
        setChecked(new Set(data.map((r) => r.page_key)));
      }
      // No rows = no explicit restriction yet, so leave everything ticked
      // (that's this person's actual current access, via their role).
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, [user.id]);

  const toggle = (key) => {
    setChecked((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  };

  const handleSave = async () => {
    setSaving(true);
    setError("");
    try {
      const { error: delErr } = await supabase.from("user_page_access").delete().eq("user_id", user.id);
      if (delErr) throw delErr;
      // Ticking every page back on is the same as having no restriction
      // at all, so that state is stored as no rows rather than a
      // redundant "everything allowed" list.
      if (checked.size < tickablePages.length) {
        const rows = [...checked].map((page_key) => ({ user_id: user.id, page_key }));
        const { error: insErr } = await supabase.from("user_page_access").insert(rows);
        if (insErr) throw insErr;
      }
      onClose();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 420, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Manage Pages</h3>
        <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 16px" }}>{user.email}</p>

        {loading ? (
          <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 4, marginBottom: 18 }}>
            {tickablePages.map((p) => (
              <label key={p.key} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, color: "#183642", padding: "5px 4px", cursor: "pointer" }}>
                <input type="checkbox" checked={checked.has(p.key)} onChange={() => toggle(p.key)} style={{ width: 15, height: 15, cursor: "pointer" }} />
                {p.label}
              </label>
            ))}
          </div>
        )}

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="button" onClick={handleSave} disabled={saving || loading} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}

const ACTIVITY_VERB = { created: "created", updated: "updated", closed: "closed", deleted: "deleted" };
const ACTIVITY_ARTICLE = { Events: "an", Backlogs: "a", "Work Orders": "a", "Daily Service": "a" };

function activitySentence(row, nameByUser) {
  const name = nameByUser.get(row.user_id) || "Someone";
  const verb = ACTIVITY_VERB[row.action] || row.action;
  const article = ACTIVITY_ARTICLE[row.table_name] || "a";
  const table = row.table_name || "record";
  const date = row.created_at ? new Date(row.created_at).toLocaleDateString("en-ZA", { day: "2-digit", month: "short", year: "numeric" }) : "";
  const time = row.created_at ? new Date(row.created_at).toLocaleTimeString("en-ZA", { hour: "2-digit", minute: "2-digit" }) : "";
  return `${name} ${verb} ${article} ${table.replace(/s$/, "")} on ${date} at ${time}${row.summary ? ` - ${row.summary}` : ""}`;
}

const ACTIVITY_COLUMNS = [
  ["created_at", "Date & Time"], ["user_name", "User"], ["table_name", "Tab"],
  ["action", "Action"], ["summary", "Details"],
];

function AuditTrailPage({ activityLog, profiles, isAdmin }) {
  const [tableFilter, setTableFilter] = useState("");
  const [actionFilter, setActionFilter] = useState("");
  const [userFilter, setUserFilter] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [showPrint, setShowPrint] = useState(false);

  if (!isAdmin) {
    return (
      <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 10, padding: 28, textAlign: "center" }}>
        <p style={{ fontWeight: 700, fontSize: 14, color: "#7A3330", margin: "0 0 6px" }}>Admins only</p>
        <p style={{ fontSize: 13, color: "#7A3330", margin: 0 }}>The Audit Trail is restricted to administrators.</p>
      </div>
    );
  }

  const nameByUser = useMemo(() => new Map(profiles.map((p) => [p.user_id, p.full_name])), [profiles]);
  const tables = useMemo(() => [...new Set(activityLog.map((r) => r.table_name))].filter(Boolean).sort(), [activityLog]);
  const users = useMemo(() => [...new Set(activityLog.map((r) => nameByUser.get(r.user_id)))].filter(Boolean).sort(), [activityLog, nameByUser]);

  const filtered = useMemo(() => {
    return activityLog.filter((r) => {
      if (tableFilter && r.table_name !== tableFilter) return false;
      if (actionFilter && r.action !== actionFilter) return false;
      if (userFilter && nameByUser.get(r.user_id) !== userFilter) return false;
      if (fromDate && new Date(r.created_at) < new Date(fromDate)) return false;
      if (toDate && new Date(r.created_at) > new Date(toDate + "T23:59:59")) return false;
      return true;
    });
  }, [activityLog, tableFilter, actionFilter, userFilter, fromDate, toDate, nameByUser]);

  const formatDT = (v) => v ? new Date(v).toLocaleString("en-ZA", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "-";

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = ACTIVITY_COLUMNS.map((c) => c[1]);
    const dataRows = filtered.map((r) => [
      formatDT(r.created_at), nameByUser.get(r.user_id) || "-", r.table_name, r.action, r.summary || "",
    ]);
    const aoa = [["Audit Trail"], [`Exported: ${timestamp}`], [], headerRow, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: ACTIVITY_COLUMNS.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: ACTIVITY_COLUMNS.length - 1 } },
    ];
    ws["!cols"] = ACTIVITY_COLUMNS.map((c) => ({ wch: c[0] === "summary" ? 50 : Math.max(c[1].length + 2, 16) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Audit Trail");
    XLSX.writeFile(wb, `Audit_Trail_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  const selectStyle = { padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', background: "#fff" };

  return (
    <div>
      <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 14px" }}>
        Covers Events, Work Orders and Backlogs, plus every deletion anywhere in the app. Other tabs (Daily Hours, Fuel Log, etc.) aren't summarised here yet - a fast follow-up once you confirm this is the level of detail you want.
      </p>

      <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
        <select value={tableFilter} onChange={(e) => setTableFilter(e.target.value)} style={selectStyle}>
          <option value="">All tabs</option>
          {tables.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
        <select value={actionFilter} onChange={(e) => setActionFilter(e.target.value)} style={selectStyle}>
          <option value="">All actions</option>
          <option value="created">Created</option>
          <option value="updated">Updated</option>
          <option value="closed">Closed</option>
          <option value="deleted">Deleted</option>
        </select>
        <select value={userFilter} onChange={(e) => setUserFilter(e.target.value)} style={selectStyle}>
          <option value="">All users</option>
          {users.map((u) => <option key={u} value={u}>{u}</option>)}
        </select>
        <input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} style={selectStyle} />
        <input type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} style={selectStyle} />
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <p style={{ fontSize: 13, color: "#4B5659", margin: 0 }}>{filtered.length} of {activityLog.length} entries shown</p>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
            <Download size={14} /> Export to Excel
          </button>
          <button onClick={() => setShowPrint(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: NAVY, border: "none", color: "#fff", fontSize: 13, fontWeight: 700, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
            <Printer size={14} /> Print Report
          </button>
        </div>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
        {filtered.map((r, i) => (
          <div key={r.id ?? i} style={{ border: "1px solid #E2E6E3", borderRadius: 10, padding: "12px 16px", background: "#fff" }}>
            <p style={{ fontSize: 13.5, color: "#183642", margin: 0, lineHeight: 1.5 }}>{activitySentence(r, nameByUser)}</p>
          </div>
        ))}
        {filtered.length === 0 && (
          <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, padding: 24, textAlign: "center", color: "#859195", fontSize: 13 }}>
            No audit entries match your filters.
          </div>
        )}
      </div>

      {showPrint && (
        <div className="job-card-modal-backdrop" style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60, padding: 16 }}>
          <style>{`
            @media print {
              body * { visibility: hidden; }
              .job-card-print, .job-card-print * { visibility: visible; }
              .job-card-print { position: absolute; top: 0; left: 0; width: 100%; }
              .job-card-modal-backdrop { position: static !important; background: none !important; padding: 0 !important; }
              .job-card-no-print { display: none !important; }
            }
          `}</style>
          <div style={{ background: "#fff", borderRadius: 12, padding: 28, width: 900, maxWidth: "100%", maxHeight: "92vh", overflowY: "auto" }}>
            <div className="job-card-no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <p style={{ fontSize: 15, fontWeight: 700, color: NAVY, margin: 0 }}>Audit Report</p>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => setShowPrint(false)} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "8px 14px", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>Close</button>
                <button onClick={() => window.print()} style={{ display: "flex", alignItems: "center", gap: 6, background: NAVY, border: "none", color: "#fff", padding: "8px 14px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  <Printer size={14} /> Print / Save as PDF
                </button>
              </div>
            </div>

            <div className="job-card-print">
              <p style={{ fontSize: 18, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Audit Report</p>
              <p style={{ fontSize: 12.5, color: "#4B5659", margin: "0 0 16px" }}>
                {fromDate || toDate ? `${fromDate || "earliest"} to ${toDate || "latest"}` : "All recorded activity"} · Generated {new Date().toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" })}
              </p>

              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {filtered.map((r, i) => (
                  <p key={r.id ?? i} style={{ fontSize: 11.5, padding: "6px 8px", border: "1px solid #ccc", margin: 0 }}>
                    {activitySentence(r, nameByUser)}
                  </p>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// Turns a raw Excel cell value into the right JS type for a given
// field. Excel/xlsx hands back native Date objects for date-formatted
// cells and strings/numbers for everything else, so this normalizes
// both cases per field type.
function excelCoerce(type, raw) {
  if (raw === "" || raw == null) return null;
  if (type === "number") {
    const n = Number(raw);
    return Number.isNaN(n) ? null : n;
  }
  if (type === "datetime") {
    if (raw instanceof Date) return raw.toISOString();
    const d = new Date(String(raw).trim().replace(" ", "T"));
    return isNaN(d.getTime()) ? null : d.toISOString();
  }
  if (type === "date") {
    if (raw instanceof Date) {
      const pad = (n) => String(n).padStart(2, "0");
      return `${raw.getFullYear()}-${pad(raw.getMonth() + 1)}-${pad(raw.getDate())}`;
    }
    return String(raw).trim();
  }
  return String(raw).trim(); // text
}

// Shared offline-backup Export/Upload bar, used identically across Fuel
// Log, Oil Consumption, Events, Inspections, and Work Orders: export the
// current data to Excel (with a hidden ID column), fill in new rows or
// edit existing ones while offline, then upload the same file once back
// online. Rows with an ID update that entry; rows without one get added
// as new - nothing is ever deleted by an upload.
function ExcelSync({ data, assets, fields, tableName, sheetTitle, filenamePrefix, onRefresh, extraOnSave }) {
  const [importing, setImporting] = useState(false);
  const [message, setMessage] = useState(null);
  const fileInputRef = React.useRef(null);

  const headers = ["ID", "Equipment #", ...fields.map((f) => f.header)];

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const dataRows = data.map((row) => [row.id ?? "", row.asset_id ?? "", ...fields.map((f) => row[f.key] ?? "")]);
    const aoa = [[sheetTitle], [`Exported: ${timestamp}`], [], headers, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
    ];
    ws["!cols"] = headers.map((h) => ({ wch: Math.max(h.length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetTitle.slice(0, 31));
    XLSX.writeFile(wb, `${filenamePrefix}_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  const handleFileSelected = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImporting(true);
    setMessage(null);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headerRowIndex = raw.findIndex((row) => row[0] === "ID" && row[1] === "Equipment #");
      const dataStart = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;

      const assetIds = new Set(assets.map((a) => a.asset_id));
      let skipped = 0;
      const rows = [];
      raw.slice(dataStart).forEach((row) => {
        const assetId = String(row[1] ?? "").trim();
        if (!assetId) return; // blank row
        if (!assetIds.has(assetId)) { skipped++; return; }
        const obj = { asset_id: assetId };
        const idVal = row[0];
        if (idVal !== "" && idVal != null && !Number.isNaN(Number(idVal))) obj.id = Number(idVal);
        fields.forEach((f, i) => { obj[f.key] = excelCoerce(f.type, row[2 + i]); });
        if (extraOnSave) Object.assign(obj, extraOnSave(obj));
        rows.push(obj);
      });

      if (rows.length === 0) {
        setMessage({ type: "error", text: "No valid rows found - make sure Equipment # matches an existing asset ID exactly." });
        setImporting(false);
        return;
      }

      const { error } = await supabase.from(tableName).upsert(rows, { onConflict: "id" });
      if (error) throw error;

      setMessage({
        type: "success",
        text: `Imported ${rows.length} row${rows.length === 1 ? "" : "s"}${skipped ? ` (${skipped} skipped - Equipment # didn't match an existing asset)` : ""}. Rows with an ID updated that entry; blank-ID rows were added as new.`,
      });
      onRefresh();
    } catch (err) {
      setMessage({ type: "error", text: err.message || String(err) });
    } finally {
      setImporting(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  return (
    <div style={{ marginBottom: 12 }}>
      {message && (
        <div style={{ marginBottom: 10, padding: "8px 12px", borderRadius: 8, fontSize: 12.5, background: message.type === "error" ? "#F6E2E0" : "#E2EFE9", color: message.type === "error" ? "#7A3330" : "#2C5646" }}>
          {message.text}
        </div>
      )}
      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", flexWrap: "wrap" }}>
        <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
          <Download size={14} /> Export to Excel
        </button>
        <button onClick={() => fileInputRef.current?.click()} disabled={importing} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: importing ? "default" : "pointer", opacity: importing ? 0.6 : 1 }}>
          <Upload size={14} /> {importing ? "Importing…" : "Upload Excel"}
        </button>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileSelected} style={{ display: "none" }} />
      </div>
      <p style={{ fontSize: 11.5, color: "#859195", margin: "6px 0 0", textAlign: "right" }}>
        Offline backup: export, fill it in during an outage, then upload the same file once you're back online.
      </p>
    </div>
  );
}

function PartsPage({ parts, selectedSiteId, onRefresh }) {
  const [subTab, setSubTab] = useState("inventory"); // "inventory" | "quotes"
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");
  const [importing, setImporting] = useState(false);
  const [importMessage, setImportMessage] = useState(null);
  const fileInputRef = React.useRef(null);

  const filtered = useMemo(() => {
    if (!query.trim()) return parts;
    const q = query.toLowerCase();
    return parts.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
  }, [parts, query]);

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const dataRows = parts.map((p) => [p.part_no, p.description, p.qty_in_stock, p.minimum_qty, p.supplier, p.unit_cost]);
    const aoa = [["Parts Inventory"], [`Exported: ${timestamp}`], [], PARTS_IMPORT_HEADERS, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: PARTS_IMPORT_HEADERS.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: PARTS_IMPORT_HEADERS.length - 1 } },
    ];
    ws["!cols"] = PARTS_IMPORT_HEADERS.map((h) => ({ wch: Math.max(h.length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Parts Inventory");
    XLSX.writeFile(wb, `Parts_Inventory_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  const handleFileSelected = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImporting(true);
    setImportMessage(null);

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      // The exported file has a title + timestamp + blank row before the
      // real header row - find it by matching on the known header text
      // rather than assuming a fixed row number, so a plain simple sheet
      // (just headers + data from row 1) also works.
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headerRowIndex = raw.findIndex((row) => row[0] === "Part No");
      const dataStart = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;

      const rows = raw.slice(dataStart)
        .filter((row) => row[0] !== "" && row[0] != null)
        .map((row) => {
          const obj = {};
          PARTS_IMPORT_KEYS.forEach((key, i) => { obj[key] = row[i]; });
          obj.qty_in_stock = Number(obj.qty_in_stock) || 0;
          obj.minimum_qty = Number(obj.minimum_qty) || 0;
          obj.unit_cost = Number(obj.unit_cost) || 0;
          obj.part_no = String(obj.part_no).trim();
          obj.site_id = selectedSiteId;
          return obj;
        })
        .filter((row) => row.part_no);

      if (rows.length === 0) {
        setImportMessage({ type: "error", text: "No valid rows found - make sure the file has a 'Part No' column with data underneath it." });
        setImporting(false);
        return;
      }

      const { error } = await supabase.from("parts_inventory").upsert(rows, { onConflict: "site_id,part_no" });
      if (error) throw error;

      setImportMessage({ type: "success", text: `Imported ${rows.length} part${rows.length === 1 ? "" : "s"} - existing part numbers were updated, new ones were added.` });
      onRefresh();
    } catch (err) {
      setImportMessage({ type: "error", text: err.message || String(err) });
    } finally {
      setImporting(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 4, marginBottom: 16, borderBottom: "1px solid #E2E6E3" }}>
        {[["inventory", "Inventory"], ["quotes", "Quote Price List"]].map(([key, label]) => (
          <button
            key={key}
            onClick={() => setSubTab(key)}
            style={{
              padding: "9px 16px", fontSize: 13.5, fontWeight: 600, border: "none", background: "none", cursor: "pointer",
              color: subTab === key ? NAVY : "#859195",
              borderBottom: subTab === key ? `2px solid ${NAVY}` : "2px solid transparent",
              marginBottom: -1,
            }}
          >
            {label}
          </button>
        ))}
      </div>

      {subTab === "quotes" ? (
        <QuotePriceListPage selectedSiteId={selectedSiteId} parts={parts} />
      ) : (
      <>
      {importMessage && (
        <div style={{
          background: importMessage.type === "error" ? "#F6E2E0" : "#E2EFE9",
          border: `1px solid ${importMessage.type === "error" ? "#DDB6B2" : "#B7D89A"}`,
          color: importMessage.type === "error" ? "#7A3330" : "#2C5646",
          borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13,
        }}>
          {importMessage.text}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search parts"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
            <Download size={14} /> Export to Excel
          </button>
          <button onClick={() => fileInputRef.current?.click()} disabled={importing} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: importing ? "default" : "pointer", opacity: importing ? 0.6 : 1 }}>
            <Upload size={14} /> {importing ? "Importing…" : "Upload Excel"}
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileSelected} style={{ display: "none" }} />
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Add Part
          </button>
        </div>
      </div>
      <p style={{ fontSize: 11.5, color: "#859195", margin: "-6px 0 16px" }}>
        Export the current list, edit it in Excel (add rows, change quantities), then Upload the same file - existing Part Nos get updated, new ones get added. Nothing gets deleted by an upload.
      </p>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Part Number", "Part", "Qty in Stock", "Minimum Qty", "Status"].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr key={row.part_no ?? i} onClick={() => { setEditing(row); setShowForm(true); }}
                style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none", cursor: "pointer" }}>
                <td style={{ padding: "9px 12px" }}>{row.part_no}</td>
                <td style={{ padding: "9px 12px" }}>{row.description}</td>
                <td style={{ padding: "9px 12px" }}>{row.qty_in_stock}</td>
                <td style={{ padding: "9px 12px" }}>{row.minimum_qty}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={row.reorder_status} /></td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={5} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {parts.length === 0 ? "No parts added yet." : "No parts match your search."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <PartForm existing={editing} selectedSiteId={selectedSiteId} onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); onRefresh(); }} />
      )}
      </>
      )}
    </div>
  );
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      // reader.result is "data:<mediatype>;base64,<data>" - keep just the data part
      const result = reader.result;
      const comma = result.indexOf(",");
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function QuotePriceListPage({ selectedSiteId, parts }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState("");
  const [partNoFilter, setPartNoFilter] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [extracting, setExtracting] = useState(false);
  const [extractError, setExtractError] = useState("");
  const [review, setReview] = useState(null); // { file, items: [{supplier, part_no, part_description, price}] }
  const [saving, setSaving] = useState(false);
  const fileInputRef = React.useRef(null);

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from("quote_prices")
      .select("*")
      .eq("site_id", selectedSiteId)
      .order("price_date", { ascending: false });
    if (!error) setRows(data || []);
    setLoading(false);
  }, [selectedSiteId]);

  useEffect(() => { load(); }, [load]);

  // Keyed by part number (case/space-insensitive) so a quote line can be
  // matched against the site's current stock regardless of minor
  // formatting differences in how each was typed/read.
  const stockByPartNo = useMemo(() => {
    const map = new Map();
    parts.forEach((p) => {
      if (p.part_no) map.set(String(p.part_no).trim().toLowerCase(), p);
    });
    return map;
  }, [parts]);
  const stockFor = (partNo) => (partNo ? stockByPartNo.get(String(partNo).trim().toLowerCase()) : null);

  const supplierOptions = useMemo(
    () => [...new Set(rows.map((r) => r.supplier).filter(Boolean))].sort(),
    [rows]
  );

  const filtered = useMemo(() => {
    let out = rows;
    if (partNoFilter.trim()) {
      const p = partNoFilter.trim().toLowerCase();
      out = out.filter((r) => String(r.part_no ?? "").toLowerCase().includes(p));
    }
    if (supplierFilter) out = out.filter((r) => r.supplier === supplierFilter);
    if (query.trim()) {
      const q = query.toLowerCase();
      out = out.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return out;
  }, [rows, query, partNoFilter, supplierFilter]);

  const processFile = async (file) => {
    if (!file) return;
    setExtractError("");
    setExtracting(true);
    try {
      const base64 = await fileToBase64(file);
      const mediaType = file.type || (file.name.toLowerCase().endsWith(".pdf") ? "application/pdf" : "image/jpeg");
      const resp = await fetch("/api/extract-quote", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ imageBase64: base64, mediaType }),
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data?.error || "Couldn't read this quote.");
      const items = (data.items || []).map((it) => ({
        supplier: it.supplier || "",
        part_no: it.part_no || "",
        part_description: it.part_description || "",
        price: it.price ?? "",
      }));
      if (items.length === 0) {
        setExtractError("No part/price lines were found on this document - you can still add rows manually below, or try a clearer scan.");
      }
      setReview({ file, items: items.length ? items : [{ supplier: "", part_no: "", part_description: "", price: "" }], detectedLanguage: data.detectedLanguage || null });
    } catch (err) {
      setExtractError(err.message || String(err));
    } finally {
      setExtracting(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    processFile(file);
  };

  const updateReviewRow = (i, field, value) => {
    setReview((r) => ({ ...r, items: r.items.map((row, idx) => (idx === i ? { ...row, [field]: value } : row)) }));
  };
  const removeReviewRow = (i) => {
    setReview((r) => ({ ...r, items: r.items.filter((_, idx) => idx !== i) }));
  };
  const addReviewRow = () => {
    setReview((r) => ({ ...r, items: [...r.items, { supplier: "", part_no: "", part_description: "", price: "" }] }));
  };

  const saveReview = async () => {
    if (!review || review.items.length === 0) return;
    setSaving(true);
    try {
      let quoteDocumentPath = null;
      if (review.file) {
        const path = `quotes/${selectedSiteId}/${Date.now()}-${review.file.name}`;
        const { error: uploadErr } = await supabase.storage.from("job-card-scans").upload(path, review.file);
        if (uploadErr) throw uploadErr;
        quoteDocumentPath = path;
      }
      const today = todayForInput();
      const payload = review.items
        .filter((it) => it.supplier || it.part_no || it.part_description)
        .map((it) => ({
          site_id: selectedSiteId,
          supplier: it.supplier || "",
          part_no: it.part_no || null,
          part_description: it.part_description || null,
          price: it.price === "" ? null : Number(it.price),
          price_date: today,
          quote_document_path: quoteDocumentPath,
        }));
      if (payload.length === 0) { setSaving(false); return; }
      const { error: insertErr } = await supabase.from("quote_prices").insert(payload);
      if (insertErr) throw insertErr;
      setReview(null);
      load();
    } catch (err) {
      setExtractError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const columns = [
    ["supplier", "Supplier"], ["part_description", "Part"], ["part_no", "Part Number"],
    ["price", "Price"], ["price_date", "Last updated"],
  ];

  return (
    <div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
        style={{
          border: `2px dashed ${dragOver ? NAVY : "#E2E6E3"}`, borderRadius: 12, padding: 28, textAlign: "center",
          background: dragOver ? "#F0F3F8" : "#FAFAF7", cursor: "pointer", marginBottom: 20, transition: "background 0.15s, border-color 0.15s",
        }}
      >
        <Upload size={22} style={{ color: NAVY, marginBottom: 8 }} />
        <p style={{ fontSize: 13.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" }}>
          {extracting ? "Reading quote…" : "Drag a scanned quote here, or click to browse"}
        </p>
        <p style={{ fontSize: 12, color: "#859195", margin: 0 }}>
          Image or PDF, any language - Claude will read it, translate part descriptions into English automatically, and pull out the Supplier, Part, Part Number and Price for you to review before saving.
        </p>
        <input
          ref={fileInputRef}
          type="file"
          accept="image/*,.pdf"
          onChange={(e) => processFile(e.target.files?.[0])}
          style={{ display: "none" }}
        />
      </div>

      {extractError && (
        <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", color: "#7A3330", borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13 }}>
          {extractError}
        </div>
      )}

      {review && (
        <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, padding: 16, marginBottom: 20, background: "#fff" }}>
          <p style={{ fontSize: 13.5, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Review before saving</p>
          <p style={{ fontSize: 12, color: "#859195", margin: "0 0 14px" }}>Check what Claude read off the quote - fix anything that's wrong, remove lines that shouldn't be there, then save. Lines flagged "In stock" already have some on hand - worth checking before ordering more.</p>
          {review.detectedLanguage && !/^english$/i.test(review.detectedLanguage) && (
            <div style={{ display: "flex", alignItems: "center", gap: 6, background: "#DCEFED", border: "1px solid #A9D6D2", borderRadius: 8, padding: "7px 11px", marginBottom: 14, fontSize: 12, color: "#183642" }}>
              <span>🌐</span>
              <span>Detected language: <strong>{review.detectedLanguage}</strong> - part descriptions below were translated into English. Supplier names and part numbers are kept as printed on the original.</span>
            </div>
          )}
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#F7F8F6" }}>
                  {["Supplier", "Part", "Part Number", "Price", "Stock", ""].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: "7px 8px", fontWeight: 600, color: "#4B5659", fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {review.items.map((row, i) => {
                  const stock = stockFor(row.part_no);
                  return (
                  <tr key={i}>
                    {["supplier", "part_description", "part_no", "price"].map((field) => (
                      <td key={field} style={{ padding: "4px 8px" }}>
                        <input
                          type={field === "price" ? "number" : "text"}
                          value={row[field]}
                          onChange={(e) => updateReviewRow(i, field, e.target.value)}
                          style={{ width: "100%", padding: "6px 8px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6, boxSizing: "border-box" }}
                        />
                      </td>
                    ))}
                    <td style={{ padding: "4px 8px", whiteSpace: "nowrap" }}>
                      {stock ? (
                        <span title={`Already have ${stock.qty_in_stock ?? 0} of ${stock.part_no} in stock`} style={{ fontSize: 10.5, fontWeight: 700, color: "#C58A32", background: "#F3E4C8", padding: "3px 8px", borderRadius: 6 }}>
                          In stock: {stock.qty_in_stock ?? 0}
                        </span>
                      ) : row.part_no ? (
                        <span style={{ fontSize: 11, color: "#B4B2A9" }}>Not in inventory</span>
                      ) : null}
                    </td>
                    <td style={{ padding: "4px 8px" }}>
                      <button type="button" onClick={() => removeReviewRow(i)} style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "flex" }}>
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12 }}>
            <button type="button" onClick={addReviewRow} style={{ background: "none", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "6px 12px", borderRadius: 6, cursor: "pointer" }}>
              + Add row
            </button>
            <div style={{ display: "flex", gap: 10 }}>
              <button type="button" onClick={() => setReview(null)} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "8px 16px", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>
                Discard
              </button>
              <button type="button" onClick={saveReview} disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
                {saving ? "Saving…" : "Save to Price List"}
              </button>
            </div>
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <div style={{ position: "relative", flex: 1, minWidth: 200, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search price list"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", boxSizing: "border-box" }} />
        </div>
        <input
          value={partNoFilter}
          onChange={(e) => setPartNoFilter(e.target.value)}
          placeholder="Filter by part number"
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", minWidth: 180 }}
        />
        <select value={supplierFilter} onChange={(e) => setSupplierFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
          <option value="">All suppliers</option>
          {supplierOptions.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map(([key, label]) => (
                <th key={key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{label}</th>
              ))}
              <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>Current stock</th>
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan={columns.length + 2} style={{ padding: 20, textAlign: "center", color: "#859195" }}>Loading…</td></tr>
            ) : filtered.length === 0 ? (
              <tr><td colSpan={columns.length + 2} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {rows.length === 0 ? "No quotes added yet - drag one in above to get started." : "No entries match your filters."}
              </td></tr>
            ) : filtered.map((row, i) => {
              const stock = stockFor(row.part_no);
              return (
              <tr key={row.id ?? i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px" }}>{row.supplier || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.part_description || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.part_no || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.price != null ? Number(row.price).toFixed(2) : <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.price_date}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {stock ? (
                    <span style={{ fontSize: 10.5, fontWeight: 700, color: "#C58A32", background: "#F3E4C8", padding: "3px 8px", borderRadius: 6 }}>
                      {stock.qty_in_stock ?? 0} in stock
                    </span>
                  ) : <span style={{ color: "#B4B2A9" }}>-</span>}
                </td>
                <td style={{ padding: "9px 12px" }}>
                  {row.quote_document_path && <QuoteDocLink path={row.quote_document_path} />}
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function QuoteDocLink({ path }) {
  const [url, setUrl] = useState(null);
  useEffect(() => {
    let cancelled = false;
    supabase.storage.from("job-card-scans").createSignedUrl(path, 3600).then(({ data }) => {
      if (!cancelled) setUrl(data?.signedUrl || null);
    });
    return () => { cancelled = true; };
  }, [path]);
  if (!url) return null;
  return <a href={url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: NAVY, whiteSpace: "nowrap" }}>View quote</a>;
}

function ChangeoutForm({ component, assets, workOrders, isFirstChangeout, currentHours, onClose, onSaved }) {
  const [workOrderId, setWorkOrderId] = useState("");
  const [changeoutDate, setChangeoutDate] = useState(todayForInput());
  const [changeoutHours, setChangeoutHours] = useState(currentHours ?? "");
  const [reason, setReason] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const woOptions = useMemo(() => workOrders.filter((w) => w.asset_id === component.asset_id), [workOrders, component.asset_id]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    if (!isFirstChangeout && !reason.trim()) { setError("Enter why the previous component was changed."); return; }
    setSaving(true);
    try {
      const wo = woOptions.find((w) => String(w.id) === String(workOrderId));
      const payload = {
        component_id: component.component_id,
        asset_id: component.asset_id,
        changeout_date: changeoutDate,
        changeout_hours: changeoutHours === "" ? null : Number(changeoutHours),
        work_order_id: workOrderId || null,
        event_id: wo?.event_id || null,
        reason_for_change: isFirstChangeout ? null : reason.trim(),
      };
      const { error: dbError } = await supabase.from("component_changeout_log").insert(payload);
      if (dbError) throw dbError;
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  const fieldStyle = { width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif' };
  const labelStyle = { display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 440, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Log Changeout</h3>
        <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 16px" }}>{component.component_id} on {component.asset_id}</p>

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Linked Work Order (optional)</label>
          <select value={workOrderId} onChange={(e) => setWorkOrderId(e.target.value)} style={fieldStyle}>
            <option value="">- None -</option>
            {woOptions.map((w) => <option key={w.id} value={w.id}>{w.wo_no}{w.problem_scope ? ` - ${w.problem_scope}` : ""}</option>)}
          </select>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Changeout Date</label>
            <input type="date" value={changeoutDate} onChange={(e) => setChangeoutDate(e.target.value)} required max={todayForInput()} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Hour Meter at Changeout</label>
            <input type="number" step="0.1" value={changeoutHours} onChange={(e) => setChangeoutHours(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        {isFirstChangeout ? (
          <p style={{ fontSize: 12, color: "#859195", background: "#F7F8F6", border: "1px solid #E2E6E3", borderRadius: 8, padding: "8px 10px", marginBottom: 18 }}>
            This is the first changeout on record for this component - no previous one to explain yet.
          </p>
        ) : (
          <div style={{ marginBottom: 18 }}>
            <label style={labelStyle}>Why was the previous component changed?</label>
            <textarea value={reason} onChange={(e) => setReason(e.target.value)} rows={2} required style={{ ...fieldStyle, resize: "vertical" }} />
          </div>
        )}

        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : "Log Changeout"}
          </button>
        </div>
      </form>
    </div>
  );
}

function ComponentsPage({ assets, components, breakdowns, workOrders, dailyHours, userEmail, onRefresh }) {
  const [targets, setTargets] = useState([]);
  const [changeouts, setChangeouts] = useState([]);
  const [loadingExtra, setLoadingExtra] = useState(true);
  const [query, setQuery] = useState("");
  const [editingTarget, setEditingTarget] = useState(null); // component_id currently being edited
  const [targetInput, setTargetInput] = useState("");
  const [changeoutFor, setChangeoutFor] = useState(null); // component row

  const loadExtra = useCallback(async () => {
    setLoadingExtra(true);
    const [t, c] = await Promise.all([
      supabase.from("component_targets").select("*"),
      supabase.from("component_changeout_log").select("*").order("changeout_date", { ascending: false }),
    ]);
    setTargets(t.data || []);
    setChangeouts(c.data || []);
    setLoadingExtra(false);
  }, []);

  useEffect(() => { loadExtra(); }, [loadExtra]);

  // Latest closing hours per asset from Daily Hours - this is what
  // stands in for "the machine's current hours" throughout this page.
  const currentHoursByAsset = useMemo(() => {
    const map = new Map();
    dailyHours.forEach((h) => {
      const prev = map.get(h.asset_id);
      if (!prev || h.log_date > prev.log_date) map.set(h.asset_id, h);
    });
    const out = new Map();
    map.forEach((h, assetId) => out.set(assetId, h.closing_hours ?? h.opening_hours ?? null));
    return out;
  }, [dailyHours]);

  const targetByComponent = useMemo(() => new Map(targets.map((t) => [t.component_id, t.target_hours])), [targets]);

  const changeoutsByComponent = useMemo(() => {
    const map = new Map();
    changeouts.forEach((c) => {
      if (!map.has(c.component_id)) map.set(c.component_id, []);
      map.get(c.component_id).push(c);
    });
    return map;
  }, [changeouts]);

  const woNoById = useMemo(() => Object.fromEntries(workOrders.map((w) => [w.id, w.wo_no])), [workOrders]);

  const rows = useMemo(() => components
    .filter((c) => Object.values(c).some((v) => String(v ?? "").toLowerCase().includes(query.toLowerCase())))
    .map((c) => {
      const asset = assets.find((a) => a.asset_id === c.asset_id);
      const history = changeoutsByComponent.get(c.component_id) || [];
      const latest = history[0] || null;
      const target = targetByComponent.get(c.component_id) ?? null;
      const currentHours = currentHoursByAsset.get(c.asset_id) ?? null;
      const baselineHours = latest?.changeout_hours ?? asset?.opening_hours ?? null;
      const hoursSinceChangeout = (currentHours != null && baselineHours != null) ? currentHours - baselineHours : null;
      const pctLifeUsed = (target && hoursSinceChangeout != null) ? (hoursSinceChangeout / target) * 100 : null;
      return { ...c, target, currentHours, hoursSinceChangeout, pctLifeUsed, latest, changeoutCount: history.length };
    }), [components, assets, changeoutsByComponent, targetByComponent, currentHoursByAsset, query]);

  const saveTarget = async (componentId) => {
    const value = targetInput === "" ? null : Number(targetInput);
    await supabase.from("component_targets").upsert({ component_id: componentId, target_hours: value }, { onConflict: "component_id" });
    setEditingTarget(null);
    loadExtra();
  };

  const pctColor = (pct) => pct == null ? "#859195" : pct >= 100 ? "#7A3330" : pct >= 80 ? "#7A5320" : "#2C5646";

  return (
    <div>
      <div style={{ position: "relative", maxWidth: 280, marginBottom: 12 }}>
        <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search this table"
          style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", boxSizing: "border-box" }} />
      </div>
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Component #", "Equipment #", "Type", "Target Hrs", "Current Hrs", "Hrs Since Changeout", "% Life Used", "Status", "Linked WO", "Reason (previous component)"].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {loadingExtra ? (
              <tr><td colSpan={11} style={{ padding: 20, textAlign: "center", color: "#859195" }}>Loading…</td></tr>
            ) : rows.length === 0 ? (
              <tr><td colSpan={11} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No rows match your search.</td></tr>
            ) : rows.map((r) => (
              <tr key={r.component_id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.component_id}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.asset_id}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.component_type}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {editingTarget === r.component_id ? (
                    <span style={{ display: "flex", gap: 4 }}>
                      <input type="number" autoFocus value={targetInput} onChange={(e) => setTargetInput(e.target.value)}
                        style={{ width: 70, padding: "3px 6px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6 }} />
                      <button onClick={() => saveTarget(r.component_id)} style={{ fontSize: 11.5, color: NAVY, background: "none", border: "none", cursor: "pointer", fontWeight: 600 }}>Save</button>
                    </span>
                  ) : (
                    <span onClick={() => { setEditingTarget(r.component_id); setTargetInput(r.target ?? ""); }} style={{ cursor: "pointer", textDecoration: "underline", textDecorationStyle: "dotted", color: r.target ? "#183642" : "#859195" }}>
                      {r.target ?? "Set target"}
                    </span>
                  )}
                </td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.currentHours ?? <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.hoursSinceChangeout != null ? Math.round(r.hoursSinceChangeout) : <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {r.pctLifeUsed != null ? (
                    <span style={{ fontWeight: 700, color: pctColor(r.pctLifeUsed) }}>{r.pctLifeUsed.toFixed(0)}%</span>
                  ) : <span style={{ color: "#B4B2A9" }}>-</span>}
                </td>
                <td style={{ padding: "9px 12px" }}><Badge value={r.status} /></td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {r.latest?.work_order_id ? (woNoById[r.latest.work_order_id] || "-") : <span style={{ color: "#B4B2A9" }}>-</span>}
                </td>
                <td style={{ padding: "9px 12px", maxWidth: 220 }}>
                  {r.latest?.reason_for_change || <span style={{ color: "#B4B2A9" }}>-</span>}
                </td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  <button onClick={() => setChangeoutFor(r)} style={{ fontSize: 12, color: NAVY, background: "none", border: `1px solid ${NAVY}`, borderRadius: 6, padding: "4px 9px", cursor: "pointer", fontWeight: 600 }}>
                    Log Changeout
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {changeoutFor && (
        <ChangeoutForm
          component={changeoutFor}
          assets={assets}
          workOrders={workOrders}
          isFirstChangeout={changeoutFor.changeoutCount === 0}
          currentHours={changeoutFor.currentHours}
          onClose={() => setChangeoutFor(null)}
          onSaved={() => { setChangeoutFor(null); loadExtra(); onRefresh(); }}
        />
      )}
    </div>
  );
}

function AssetsPage({ assets, selectedSiteId, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");

  const columns = [
    ["asset_id", "Equipment #"], ["asset_name", "Name"], ["make", "Make"], ["model", "Model"],
    ["fleet", "Fleet"], ["serial_number", "Serial number"], ["status", "Status"], ["current_hours", "Current hours"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    if (!query.trim()) return assets;
    const q = query.toLowerCase();
    return assets.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
  }, [assets, query]);

  const handleSaved = () => {
    setShowForm(false);
    setEditing(null);
    onRefresh();
  };

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = columns.map((c) => c.label);
    const dataRows = filtered.map((row) => columns.map((c) => row[c.key] ?? ""));
    const aoa = [["Assets"], [`Exported: ${timestamp}`], [], headerRow, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: columns.length - 1 } },
    ];
    ws["!cols"] = columns.map((c) => ({ wch: Math.max(c.label.length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Assets");
    XLSX.writeFile(wb, `Assets_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search assets"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }}
          />
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button
            onClick={exportToExcel}
            style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}
          >
            <Download size={14} /> Export to Excel
          </button>
          <button
            onClick={() => { setEditing(null); setShowForm(true); }}
            style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}
          >
            + Add Asset
          </button>
        </div>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {columns.map((c) => (
                <th key={c.key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{c.label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((row, i) => (
              <tr
                key={row.asset_id ?? i}
                onClick={() => { setEditing(row); setShowForm(true); }}
                style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none", cursor: "pointer" }}
              >
                {columns.map((c) => (
                  <td key={c.key} style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                    {c.key === "status" ? <Badge value={row[c.key]} /> : (row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>)}
                  </td>
                ))}
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={columns.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {assets.length === 0 ? "No assets added yet." : "No assets match your search."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && (
        <AssetForm
          existing={editing}
          selectedSiteId={selectedSiteId}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={handleSaved}
        />
      )}
    </div>
  );
}


function KpiRow({ metrics }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 12, marginBottom: 20 }}>
      {METRIC_DEFS.map((m) => (
        <MetricCard key={m.key} label={m.label} value={metrics[m.key] != null ? m.fmt(metrics[m.key]) : "-"} />
      ))}
    </div>
  );
}

function aggregateMetrics(kpiRows, assetIds) {
  const rows = kpiRows.filter((r) => assetIds.includes(r.asset_id));
  if (rows.length === 0) return {};
  const numRows = (key) => rows.map((r) => Number(r[key])).filter((v) => !isNaN(v));
  const sum = (key) => numRows(key).reduce((a, b) => a + b, 0);
  const avg = (key) => {
    const vals = numRows(key);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };
  const totalWorked = sum("worked_hours");
  const totalBreakdowns = sum("num_unplanned_events");
  const avgMtbf = avg("mtbf");
  const avgMttr = avg("mttr");
  return {
    hours_worked: totalWorked,
    availability: avg("availability"),
    utilisation: avg("utilisation"),
    breakdown_count: totalBreakdowns,
    mtbf: avgMtbf,
    mttr: avgMttr,
    availability_index: avgMtbf != null && avgMttr != null && (avgMtbf + avgMttr) > 0
      ? (avgMtbf / (avgMtbf + avgMttr)) * 100
      : avg("availability_index"),
  };
}

// Every delete goes through here: logs the reason to deletion_log, then
// performs the actual delete. The existing audit trigger already captures
// what was deleted automatically - this captures why, as a required step
// rather than an afterthought.
// Best-effort - never blocks or fails the user's actual save because a
// log write hiccupped. Every insert/update/delete flows through here
// with a plain-English summary already built at the point it happens,
// rather than trying to reconstruct one later from a raw row diff.
async function logActivity(tableName, recordId, action, summary) {
  try {
    await supabase.from("activity_log").insert({ table_name: tableName, record_id: recordId != null ? String(recordId) : null, action, summary });
  } catch (err) {
    console.error("Activity log write failed:", err);
  }
}

async function deleteWithReason(tableName, recordId, idColumn, reason, userEmail) {
  const { error: logError } = await supabase.from("deletion_log").insert({
    table_name: tableName, record_id: String(recordId), reason, deleted_by_email: userEmail || null,
  });
  if (logError) throw logError;
  const { error: deleteError } = await supabase.from(tableName).delete().eq(idColumn, recordId);
  if (deleteError) throw deleteError;
  logActivity(TABLE_NAME_LABELS[tableName] || tableName, recordId, "deleted", reason);
}

function DeleteConfirmModal({ itemLabel, userEmail, onCancel, onConfirm }) {
  const [reason, setReason] = useState("");
  const [error, setError] = useState("");
  const [deleting, setDeleting] = useState(false);

  const handleConfirm = async () => {
    if (!reason.trim()) { setError("A reason is required before deleting."); return; }
    setError("");
    setDeleting(true);
    try {
      await onConfirm(reason.trim());
    } catch (err) {
      setError(err.message || String(err));
      setDeleting(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 55, padding: 16 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 400, maxWidth: "100%" }}>
        <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>Delete {itemLabel}?</p>
        <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 14px" }}>
          This removes it from the list, but the record and your reason both stay in the Audit Trail for accountability.
        </p>
        <label style={{ display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" }}>Reason for deletion (required)</label>
        <textarea
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          rows={2}
          autoFocus
          style={{ width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", resize: "vertical", fontFamily: "inherit" }}
        />
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "8px 0 0" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
          <button onClick={onCancel} disabled={deleting} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button onClick={handleConfirm} disabled={deleting} style={{ background: "#B85450", border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: deleting ? "default" : "pointer", opacity: deleting ? 0.7 : 1 }}>
            {deleting ? "Deleting…" : "Delete"}
          </button>
        </div>
      </div>
    </div>
  );
}

function FleetEquipmentFilter({ assets, selectedFleet, setSelectedFleet, selectedAsset, setSelectedAsset }) {
  const fleets = useMemo(() => [...new Set(assets.map((a) => a.fleet))].filter(Boolean), [assets]);
  const filteredAssets = selectedFleet ? assets.filter((a) => a.fleet === selectedFleet) : assets;
  const selectStyle = { padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', background: "#fff", minWidth: 160 };

  return (
    <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
      <select
        value={selectedFleet}
        onChange={(e) => { setSelectedFleet(e.target.value); setSelectedAsset(""); }}
        style={selectStyle}
      >
        <option value="">All fleets</option>
        {fleets.map((f) => <option key={f} value={f}>{f}</option>)}
      </select>
      <select
        value={selectedAsset}
        onChange={(e) => setSelectedAsset(e.target.value)}
        style={selectStyle}
      >
        <option value="">All equipment</option>
        {filteredAssets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
      </select>
    </div>
  );
}

function DateRangePicker({ fromDateTime, toDateTime, setFromDateTime, setToDateTime }) {
  const invalid = new Date(fromDateTime) > new Date(toDateTime);
  const inputStyle = {
    fontSize: 13, padding: "7px 10px", border: `1px solid ${invalid ? "#B85450" : "#E2E6E3"}`,
    borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', color: "#183642",
  };
  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <span style={{ fontSize: 12.5, color: "#4B5659", fontWeight: 600 }}>Show data from:</span>
        <input
          type="datetime-local"
          value={fromDateTime}
          onChange={(e) => setFromDateTime(e.target.value)}
          style={inputStyle}
        />
        <span style={{ fontSize: 12.5, color: "#4B5659", fontWeight: 600 }}>to:</span>
        <input
          type="datetime-local"
          value={toDateTime}
          onChange={(e) => setToDateTime(e.target.value)}
          style={inputStyle}
        />
      </div>
      {invalid && (
        <p style={{ fontSize: 12, color: "#B85450", margin: "6px 0 0" }}>The "from" date/time must be before the "to" date/time.</p>
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
  const [kpiData, setKpiData] = useState([]);
  const [kpiLoading, setKpiLoading] = useState(true);
  const [kpiError, setKpiError] = useState(null);

  // Note: the date range is intentionally NOT reset by any of the
  // drill-down navigation below - it's set once and stays applied at
  // every level (fleet list, fleet detail, equipment detail, compare)
  // until the user changes it themselves.

  useEffect(() => {
    let cancelled = false;
    async function loadKpi() {
      setKpiLoading(true);
      setKpiError(null);
      const { data, error } = await supabase.rpc("plant_performance_kpi", {
        period_start: new Date(fromDateTime).toISOString(),
        period_end: new Date(toDateTime).toISOString(),
      });
      if (cancelled) return;
      if (error) setKpiError(error.message);
      else setKpiData(data || []);
      setKpiLoading(false);
    }
    loadKpi();
    return () => { cancelled = true; };
  }, [fromDateTime, toDateTime]);

  const kpiByAsset = useMemo(() => Object.fromEntries(kpiData.map((r) => [r.asset_id, r])), [kpiData]);

  const fleets = useMemo(() => {
    const names = [...new Set(assets.map((a) => a.fleet))];
    return names.map((fleet) => {
      const fleetAssetIds = assets.filter((a) => a.fleet === fleet).map((a) => a.asset_id);
      return { fleet, count: fleetAssetIds.length, metrics: aggregateMetrics(kpiData, fleetAssetIds) };
    });
  }, [assets, kpiData]);

  const fleetAssets = selectedFleet ? assets.filter((a) => a.fleet === selectedFleet) : [];

  const resetToFleetList = () => { setSelectedFleet(null); setSelectedEquipment(null); setCompareMode(false); setCompareIds([]); };
  const backToFleet = () => { setSelectedEquipment(null); setCompareMode(false); setCompareIds([]); };

  const toggleCompare = (id) => {
    setCompareIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  };

  const crumbs = (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "#4B5659", marginBottom: 16, flexWrap: "wrap" }}>
      <button onClick={resetToFleetList} style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: selectedFleet ? NAVY : "#4B5659", fontWeight: selectedFleet ? 400 : 700, fontSize: 13 }}>All fleets</button>
      {selectedFleet && (
        <>
          <ChevronRight size={13} />
          <button onClick={backToFleet} style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: selectedEquipment ? NAVY : "#4B5659", fontWeight: selectedEquipment ? 400 : 700, fontSize: 13 }}>{selectedFleet}</button>
        </>
      )}
      {selectedEquipment && (
        <>
          <ChevronRight size={13} />
          <span style={{ fontWeight: 700, color: "#4B5659" }}>{selectedEquipment}</span>
        </>
      )}
    </div>
  );

  const rangeLabel = formatRangeForDisplay(fromDateTime, toDateTime);

  if (kpiError) {
    return (
      <div>
        <DateRangePicker fromDateTime={fromDateTime} toDateTime={toDateTime} setFromDateTime={setFromDateTime} setToDateTime={setToDateTime} />
        <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 8, padding: "10px 14px", fontSize: 13, color: "#7A3330" }}>
          Couldn't load performance data: {kpiError}
        </div>
      </div>
    );
  }

  if (kpiLoading) {
    return (
      <div>
        <DateRangePicker fromDateTime={fromDateTime} toDateTime={toDateTime} setFromDateTime={setFromDateTime} setToDateTime={setToDateTime} />
        <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
      </div>
    );
  }

  // ---------------------------------------------------------------
  // Build the content for whichever level is active - one render tree,
  // with the time frame selector rendered exactly once, above it, so it
  // never disappears or resets while drilling down or backing out.
  // ---------------------------------------------------------------
  let content;

  if (!selectedFleet) {
    // Level 1: fleet cards
    content = (
      <div>
        <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 16px" }}>Select a fleet to see its performance, then drill into an individual machine. Figures below are for {rangeLabel}.</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 14 }}>
          {fleets.map((f) => (
            <button
              key={f.fleet}
              onClick={() => setSelectedFleet(f.fleet)}
              style={{ textAlign: "left", background: "#F7F8F6", border: "1px solid #E2E6E3", borderRadius: 10, padding: 16, cursor: "pointer" }}
            >
              <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>{f.fleet}</p>
              <p style={{ fontSize: 12, color: "#4B5659", margin: "0 0 10px" }}>{f.count} equipment</p>
              <div style={{ display: "flex", gap: 14 }}>
                <span style={{ fontSize: 13 }}><b>{f.metrics.availability != null ? Math.round(f.metrics.availability * 100) : "-"}%</b> avail.</span>
                <span style={{ fontSize: 13 }}><b>{f.metrics.utilisation != null ? Math.round(f.metrics.utilisation * 100) : "-"}%</b> util.</span>
                <span style={{ fontSize: 13 }}><b>{f.metrics.breakdown_count ?? "-"}</b> breakdowns</span>
              </div>
            </button>
          ))}
        </div>
      </div>
    );
  } else if (selectedEquipment) {
    // Level 3: single equipment detail
    const asset = assets.find((a) => a.asset_id === selectedEquipment);
    const metrics = kpiByAsset[selectedEquipment] || {};
    const eqBreakdowns = breakdownsInRange(breakdowns, selectedEquipment, fromDateTime, toDateTime);
    content = (
      <div>
        {crumbs}
        <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 4px", color: NAVY }}>{asset.asset_name}</h3>
        <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 16px" }}>{asset.make} {asset.model} · Serial {asset.serial_number} · <Badge value={asset.status} /></p>
        <KpiRow metrics={{ ...metrics, hours_worked: metrics.worked_hours, breakdown_count: metrics.num_unplanned_events }} />
        <h4 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px", color: NAVY }}>Events - {rangeLabel}</h4>
        {eqBreakdowns.length === 0 ? (
          <p style={{ fontSize: 13, color: "#859195" }}>No events for this machine in the selected date range.</p>
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
    const fleetMetrics = aggregateMetrics(kpiData, fleetAssets.map((a) => a.asset_id));
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

        <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "hidden", marginBottom: compareMode && compareIds.length >= 2 ? 24 : 0 }}>
          {fleetAssets.map((a, i) => {
            const m = kpiByAsset[a.asset_id] || {};
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
                  <p style={{ fontSize: 13.5, fontWeight: 600, margin: 0, color: "#183642" }}>{a.asset_id} - {a.asset_name}</p>
                  <p style={{ fontSize: 12, color: "#859195", margin: "2px 0 0" }}>{a.serial_number}</p>
                </div>
                <span style={{ fontSize: 13 }}>{m.availability != null ? Math.round(m.availability * 100) : "-"}% avail.</span>
                <span style={{ fontSize: 13 }}>{m.utilisation != null ? Math.round(m.utilisation * 100) : "-"}% util.</span>
                <span style={{ fontSize: 13 }}>{m.num_unplanned_events ?? "-"} breakdowns</span>
                <Badge value={a.status} />
                {!compareMode && <ChevronRight size={15} style={{ color: "#B4B2A9" }} />}
              </div>
            );
          })}
        </div>

        {compareMode && compareIds.length >= 2 && (
          <div>
            <h4 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px", color: NAVY }}>Comparison - {rangeLabel}</h4>
            <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: "#F7F8F6" }}>
                    <th style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3" }}>Metric</th>
                    {compareIds.map((id) => (
                      <th key={id} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" }}>{id}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {METRIC_DEFS.map((met, i) => (
                    <tr key={met.key} style={{ borderTop: i > 0 ? "1px solid #EFEEE7" : "none" }}>
                      <td style={{ padding: "9px 12px", fontWeight: 600 }}>{met.label}</td>
                      {compareIds.map((id) => {
                        const row = kpiByAsset[id];
                        const fieldKey = met.key === "hours_worked" ? "worked_hours" : met.key === "breakdown_count" ? "num_unplanned_events" : met.key;
                        const val = row ? row[fieldKey] : null;
                        return <td key={id} style={{ padding: "9px 12px" }}>{val != null ? met.fmt(val) : "-"}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
        {compareMode && compareIds.length === 1 && (
          <p style={{ fontSize: 13, color: "#859195" }}>Select at least one more machine to compare.</p>
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
// Dark, glossy-bar KPI chart matching the plant's existing Excel report
// style (dark background, gradient bars, value labels on top, a solid
// target line, legend below) rather than the flatter light chart style
// used elsewhere in the app - these four are deliberately styled to
// match that reference rather than the rest of the app's charts.
// Rounds a domain max up to a "nice" round number (1/2/5 x a power of
// ten) instead of an arbitrary decimal like 34.04 - purely cosmetic, but
// a jagged axis top reads as broken on a phone where there's little else
// to anchor the eye.
function niceDomainMax(value) {
  if (!value || value <= 0) return 10;
  const magnitude = Math.pow(10, Math.floor(Math.log10(value)));
  const normalized = value / magnitude;
  const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return niceNormalized * magnitude;
}

function KpiBarChart({ title, data, xKey, dataKey, target, domainMax, valueFormatter, meetsTarget, unitSuffix, onClick }) {
  const gradGreen = `grad-green-${dataKey}`;
  const gradRed = `grad-red-${dataKey}`;
  const niceMax = domainMax === 100 ? 100 : niceDomainMax(domainMax);

  // On a phone-width screen, side-by-side columns with rotated labels
  // just don't have room for names like "LIEBHERR EXCAVATOR" - they
  // collide into unreadable text. Stacking equipment one per row instead
  // (each with its own line for its name) fixes that at any label length,
  // same fix already used for the Event Timeline further up this page.
  const [narrow, setNarrow] = useState(() => typeof window !== "undefined" && window.innerWidth < 560);
  useEffect(() => {
    const onResize = () => setNarrow(window.innerWidth < 560);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  const chartHeight = narrow ? Math.max(180, data.length * 40 + 30) : 240;

  return (
    <div>
      {title && <p style={{ fontSize: 12.5, fontWeight: 600, margin: "0 0 6px", color: "#4B5659" }}>{title}</p>}
      <div onClick={onClick} style={{ height: chartHeight, background: "#292929", border: "1px solid #1C1C1C", borderRadius: 10, padding: narrow ? "10px 10px 10px 4px" : "18px 8px 4px", cursor: onClick ? "pointer" : "default" }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={data} layout={narrow ? "vertical" : "horizontal"} margin={narrow ? { top: 4, right: 30, left: 4, bottom: 4 } : { top: 20, right: 8, left: 0, bottom: 20 }}>
            <defs>
              <linearGradient id={gradGreen} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#8FE8B4" />
                <stop offset="100%" stopColor="#2F9E63" />
              </linearGradient>
              <linearGradient id={gradRed} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#F3998A" />
                <stop offset="100%" stopColor="#A62A1E" />
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#3D3D3D" vertical={narrow} horizontal={!narrow} />
            {narrow ? (
              <>
                <XAxis type="number" domain={[0, niceMax]} tick={{ fontSize: 10, fill: "#CFCFCF" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}${unitSuffix || ""}`} />
                <YAxis type="category" dataKey={xKey} width={82} tick={{ fontSize: 10, fill: "#CFCFCF" }} axisLine={{ stroke: "#4A4A4A" }} tickLine={false} />
              </>
            ) : (
              <>
                <XAxis dataKey={xKey} interval={0} angle={-30} textAnchor="end" height={44} tick={{ fontSize: 10, fill: "#CFCFCF" }} axisLine={{ stroke: "#4A4A4A" }} tickLine={false} />
                <YAxis domain={[0, niceMax]} tick={{ fontSize: 11, fill: "#CFCFCF" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}${unitSuffix || ""}`} />
              </>
            )}
            <Tooltip formatter={(v) => `${Number(v).toFixed(1)}${unitSuffix || ""}`} contentStyle={{ background: "#1F1F1F", border: "1px solid #444", borderRadius: 8 }} labelStyle={{ color: "#fff" }} itemStyle={{ color: "#fff" }} />
            <Bar dataKey={dataKey} radius={narrow ? [0, 3, 3, 0] : [3, 3, 0, 0]} cursor={onClick ? "pointer" : "default"}>
              {data.map((row, i) => <Cell key={i} fill={meetsTarget(row) ? `url(#${gradGreen})` : `url(#${gradRed})`} />)}
              <LabelList dataKey={dataKey} position={narrow ? "right" : "top"} formatter={valueFormatter} fill="#fff" fontSize={11} fontWeight={700} />
            </Bar>
            <ReferenceLine {...(narrow ? { x: target } : { y: target })} stroke="#F5C518" strokeWidth={2.5} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function KpiLegend({ targetLabel }) {
  return (
    <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 11.5, color: "#4B5659" }}>
      <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#2F9E63", borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Target met</span>
      <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#A62A1E", borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Target not met</span>
      <span><span style={{ display: "inline-block", width: 12, height: 3, background: "#F5C518", borderRadius: 2, marginRight: 6, verticalAlign: "middle" }} />{targetLabel || "Target"}</span>
    </div>
  );
}

function UserMenu({ myFullName, isAdmin, myRole }) {
  const [open, setOpen] = useState(false);
  const ref = React.useRef(null);

  useEffect(() => {
    const onClick = (e) => { if (ref.current && !ref.current.contains(e.target)) { setOpen(false); } };
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, []);

  const initials = (myFullName || "?")
    .split(" ").filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join("") || "?";
  const roleLabel = isAdmin ? "Administrator" : myRole === "manager" ? "Manager" : "Operator";

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((v) => !v)}
        style={{ display: "flex", alignItems: "center", gap: 9, background: "none", border: "none", cursor: "pointer", padding: "4px 6px", borderRadius: 8 }}
      >
        <span style={{ width: 32, height: 32, borderRadius: "50%", background: NAVY, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12.5, fontWeight: 700, flexShrink: 0 }}>
          {initials}
        </span>
        <span style={{ textAlign: "left", display: window.innerWidth < 480 ? "none" : "block" }}>
          <span style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#183642", lineHeight: 1.2 }}>{myFullName || "Name not set"}</span>
          <span style={{ display: "block", fontSize: 11, color: "#859195", lineHeight: 1.2 }}>{roleLabel}</span>
        </span>
        <ChevronDown size={14} style={{ color: "#859195", transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
      </button>

      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 6px)", right: 0, background: "#fff", border: "1px solid #E2E6E3", borderRadius: 10, boxShadow: "0 4px 16px rgba(0,0,0,0.1)", padding: 10, width: 200, zIndex: 30 }}>
          {/* Names are set by an admin only, from the User Access page -
              deliberately no self-service edit here. */}
          <button onClick={() => supabase.auth.signOut()} style={{ width: "100%", textAlign: "left", background: "none", border: "none", padding: "8px 8px", fontSize: 13, color: "#B85450", cursor: "pointer", borderRadius: 6 }}>
            Sign out
          </button>
        </div>
      )}
    </div>
  );
}

function AboutPage() {
  const capabilities = [
    ["Equipment & Events", "A live register of every machine, with breakdowns and planned jobs tracked as Events - each with its own linked Work Orders, parts used, and full downtime history."],
    ["Daily Operations", "Daily Hours, Fuel Log, Oil Consumption and Daily Service compliance, all tied back to the equipment they belong to."],
    ["Maintenance & Backlogs", "Work Orders, Planned Maintenance schedules, and a Backlog Report that tracks outstanding items back to where they came from - a daily service, a monthly inspection, or a scanned service card."],
    ["Parts & Pricing", "Parts Inventory with reorder alerts, and a Quote Price List that reads scanned supplier quotes and keeps a running history of what's been quoted, by whom, and when."],
    ["Reporting", "Availability, MTBF, MTTR and Utilisation, calculated from the same underlying data as everything else in the system - no separate spreadsheet to keep in sync."],
    ["Accountability", "Every insert, update and deletion is captured in an Audit Trail, with required reasons for anything removed."],
  ];

  return (
    <div style={{ maxWidth: 760 }}>
      <div style={{ background: NAVY, borderRadius: 12, padding: "32px 28px", marginBottom: 28, color: "#fff" }}>
        <h1 style={{ fontSize: 24, fontWeight: 700, margin: "0 0 8px" }}>Fleet Tracker</h1>
        <p style={{ fontSize: 14.5, lineHeight: 1.6, margin: 0, color: "rgba(255,255,255,0.88)" }}>
          A single system for tracking mining fleet equipment from the moment it's mobilised on site: maintenance, availability, parts, and the reporting that ties it all together - replacing a spreadsheet workbook with one live, multi-site application.
        </p>
      </div>

      <h2 style={{ fontSize: 15, fontWeight: 700, color: NAVY, margin: "0 0 14px" }}>What it does</h2>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: 14, marginBottom: 28 }}>
        {capabilities.map(([title, desc]) => (
          <div key={title} style={{ border: "1px solid #E2E6E3", borderRadius: 10, padding: "14px 16px", background: "#fff" }}>
            <p style={{ fontSize: 13.5, fontWeight: 700, color: NAVY, margin: "0 0 6px" }}>{title}</p>
            <p style={{ fontSize: 13, color: "#4B5659", margin: 0, lineHeight: 1.5 }}>{desc}</p>
          </div>
        ))}
      </div>

      <div style={{ borderTop: "1px solid #E2E6E3", paddingTop: 16 }}>
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>Built by Datavera Analytics.</p>
      </div>
    </div>
  );
}

function Dashboard({ assets, breakdowns, workOrders, plannedMaintenance, components, parts, inspections, onNavigate }) {
  const [monthKpi, setMonthKpi] = useState([]);
  const [monthKpiLoading, setMonthKpiLoading] = useState(true);
  const [selectedFleet, setSelectedFleet] = useState("");

  const fleets = useMemo(() => [...new Set(assets.map((a) => a.fleet))].filter(Boolean).sort(), [assets]);
  const filteredAssets = useMemo(
    () => (selectedFleet ? assets.filter((a) => a.fleet === selectedFleet) : assets),
    [assets, selectedFleet]
  );
  const filteredAssetIds = useMemo(() => new Set(filteredAssets.map((a) => a.asset_id)), [filteredAssets]);

  const filteredBreakdowns = useMemo(() => breakdowns.filter((b) => filteredAssetIds.has(b.asset_id)), [breakdowns, filteredAssetIds]);
  const filteredWorkOrders = useMemo(() => workOrders.filter((w) => filteredAssetIds.has(w.asset_id)), [workOrders, filteredAssetIds]);
  const filteredPlannedMaintenance = useMemo(() => plannedMaintenance.filter((p) => filteredAssetIds.has(p.asset_id)), [plannedMaintenance, filteredAssetIds]);
  const filteredComponents = useMemo(() => components.filter((c) => filteredAssetIds.has(c.asset_id)), [components, filteredAssetIds]);
  const filteredInspections = useMemo(() => inspections.filter((i) => filteredAssetIds.has(i.asset_id)), [inspections, filteredAssetIds]);
  const filteredMonthKpi = useMemo(() => monthKpi.filter((r) => filteredAssetIds.has(r.asset_id)), [monthKpi, filteredAssetIds]);
  // Parts Inventory isn't tied to a specific asset/fleet in this schema -
  // stays unfiltered regardless of fleet selection.

  // Ticks forward every 60 seconds so any downtime shown for an in-progress
  // breakdown or work order stays accurate to the actual current moment -
  // not frozen at whichever instant the page data happened to load. A full
  // page refresh also naturally re-syncs this (nowTick re-initializes to
  // the real current time on mount).
  const [nowTick, setNowTick] = useState(() => new Date());
  useEffect(() => {
    const interval = setInterval(() => setNowTick(new Date()), 60000);
    return () => clearInterval(interval);
  }, []);

  const liveDowntimeHours = (startIso) => {
    if (!startIso) return null;
    return (nowTick.getTime() - new Date(startIso).getTime()) / 3600000;
  };

  const inProgressBreakdowns = useMemo(
    () => filteredBreakdowns
      .filter((b) => b.repair_status !== "Closed" && b.downtime_start)
      .map((b) => ({ ...b, liveHours: liveDowntimeHours(b.downtime_start) }))
      .sort((a, b) => (b.liveHours ?? 0) - (a.liveHours ?? 0)),
    [filteredBreakdowns, nowTick]
  );

  const inProgressMaintenance = useMemo(
    () => filteredWorkOrders
      .filter((w) => w.status !== "Closed" && w.actual_start)
      .map((w) => ({ ...w, liveHours: liveDowntimeHours(w.actual_start) }))
      .sort((a, b) => (b.liveHours ?? 0) - (a.liveHours ?? 0)),
    [filteredWorkOrders, nowTick]
  );

  const formatLiveHours = (hrs) => {
    if (hrs == null) return "-";
    if (hrs < 1) return `${Math.round(hrs * 60)} min`;
    return `${hrs.toFixed(1)} hrs`;
  };

  useEffect(() => {
    let cancelled = false;
    async function loadMonthKpi() {
      setMonthKpiLoading(true);
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const { data, error } = await supabase.rpc("plant_performance_kpi", {
        period_start: monthStart.toISOString(),
        period_end: now.toISOString(),
      });
      if (cancelled) return;
      if (!error) setMonthKpi(data || []);
      setMonthKpiLoading(false);
    }
    loadMonthKpi();
    return () => { cancelled = true; };
  }, []);

  const fleetAvailability = useMemo(() => {
    const byFleet = {};
    filteredMonthKpi.forEach((row) => {
      if (row.availability == null) return;
      if (!byFleet[row.fleet]) byFleet[row.fleet] = [];
      byFleet[row.fleet].push(Number(row.availability));
    });
    return Object.entries(byFleet).map(([fleet, vals]) => ({
      fleet, availability: Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100),
    }));
  }, [filteredMonthKpi]);

  const repeatFailureCount = useMemo(() => {
    const groups = new Set(
      filteredBreakdowns
        .filter((b) => (b.repeat_count ?? 0) >= 2)
        .map((b) => `${b.asset_id}|${b.component_affected}`)
    );
    return groups.size;
  }, [filteredBreakdowns]);

  const avgAvailability = fleetAvailability.length
    ? Math.round(fleetAvailability.reduce((a, f) => a + f.availability, 0) / fleetAvailability.length)
    : null;

  const avgUtilisation = useMemo(() => {
    const vals = filteredMonthKpi.filter((r) => r.utilisation != null).map((r) => Number(r.utilisation));
    return vals.length ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100) : null;
  }, [filteredMonthKpi]);

  const avgMtbf = useMemo(() => {
    const vals = filteredMonthKpi.filter((r) => r.mtbf != null).map((r) => Number(r.mtbf));
    return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) : null;
  }, [filteredMonthKpi]);

  const avgMttr = useMemo(() => {
    const vals = filteredMonthKpi.filter((r) => r.mttr != null).map((r) => Number(r.mttr));
    return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) : null;
  }, [filteredMonthKpi]);

  const openBreakdownCount = useMemo(
    () => filteredBreakdowns.filter((b) => b.repair_status !== "Closed").length,
    [filteredBreakdowns]
  );

  const statusDistribution = useMemo(() => {
    const counts = {};
    filteredAssets.forEach((a) => { counts[a.status || "Unknown"] = (counts[a.status || "Unknown"] || 0) + 1; });
    return Object.entries(counts).map(([status, count]) => ({ status, count }));
  }, [filteredAssets]);

  const STATUS_COLORS = {
    Operating: "#3F7D65", "Under Maintenance": "#E8A33D", Breakdown: "#B85450",
    Standby: "#8CA0BF", Disposed: "#B4B2A9", Unknown: "#E2E6E3",
  };

  // Each problem asset broken down by individual breakdown type
  // (component/cause), showing how many times and total hours lost to
  // each one - not just a single combined count.
  const topProblemEquipment = useMemo(() => {
    const byAsset = {};
    filteredBreakdowns.forEach((b) => {
      if (!byAsset[b.asset_id]) byAsset[b.asset_id] = {};
      const key = b.component_affected || b.cause_code || "Unspecified";
      if (!byAsset[b.asset_id][key]) byAsset[b.asset_id][key] = { count: 0, hours: 0 };
      byAsset[b.asset_id][key].count += 1;
      byAsset[b.asset_id][key].hours += Number(b.downtime_hours) || 0;
    });
    return Object.entries(byAsset)
      .map(([asset_id, types]) => {
        const asset = assets.find((a) => a.asset_id === asset_id);
        const entries = Object.entries(types).map(([type, v]) => ({ type, ...v })).sort((a, b) => b.hours - a.hours);
        return {
          asset_id, name: asset?.asset_name || asset_id,
          totalCount: entries.reduce((a, e) => a + e.count, 0),
          totalHours: entries.reduce((a, e) => a + e.hours, 0),
          entries,
        };
      })
      .sort((a, b) => b.totalCount - a.totalCount)
      .slice(0, 5);
  }, [filteredBreakdowns, assets]);
  const maxProblemCount = Math.max(1, ...topProblemEquipment.map((p) => p.totalCount));

  const alertRows = [
    ["Services overdue / due soon", filteredPlannedMaintenance.filter((p) => p.status !== "OK").length, AlertTriangle, "#E8A33D", "planned_maintenance"],
    ["Repeat failures (2+ times)", repeatFailureCount, AlertTriangle, "#B85450", "breakdowns"],
    ["Components requiring action", filteredComponents.filter((c) => c.status !== "OK").length, CircleDot, "#E8A33D", "components"],
    ["Parts below reorder point", parts.filter((p) => p.reorder_status === "REORDER").length, Package, "#E8A33D", "parts"],
    ["Failed inspections", filteredInspections.filter((i) => i.result === "Fail").length, ShieldCheck, "#B85450", "inspections"],
  ];

  return (
    <div style={{ position: "relative" }}>
      <svg viewBox="0 0 400 250" style={{ position: "absolute", right: -10, top: 60, width: 340, height: 213, opacity: 0.045, pointerEvents: "none", zIndex: 0 }}>
        <rect x="20" y="190" width="140" height="30" rx="15" fill={NAVY} />
        <circle cx="35" cy="205" r="15" fill={NAVY} />
        <circle cx="145" cy="205" r="15" fill={NAVY} />
        <rect x="40" y="140" width="100" height="55" rx="6" fill={NAVY} />
        <rect x="45" y="105" width="45" height="45" rx="4" fill={NAVY} />
        <path d="M 100 130 L 200 80 L 260 100" stroke={NAVY} strokeWidth="14" fill="none" strokeLinecap="round" />
        <path d="M 260 100 L 300 150" stroke={NAVY} strokeWidth="12" fill="none" strokeLinecap="round" />
        <path d="M 300 150 L 330 145 L 335 170 L 305 178 Z" fill={NAVY} />
      </svg>

      <div style={{ position: "relative", zIndex: 1 }}>
        <p style={{ fontSize: 13, fontStyle: "italic", color: "#4B5659", margin: "-12px 0 16px" }}>
          "If you can measure it, you can manage it."
        </p>

        <div style={{ marginBottom: 16 }}>
          <select
            value={selectedFleet}
            onChange={(e) => setSelectedFleet(e.target.value)}
            style={{ padding: "8px 12px", fontSize: 13, fontWeight: 600, border: "1px solid #E2E6E3", borderRadius: 8, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', background: "#fff", color: NAVY, minWidth: 180 }}
          >
            <option value="">All fleets</option>
            {fleets.map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
          {selectedFleet && <span style={{ fontSize: 12, color: "#859195", marginLeft: 10 }}>Everything below is filtered to {selectedFleet}</span>}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 14, marginBottom: 24 }}>
          <MetricCard label="Total assets" value={filteredAssets.length} icon={Truck} accentColor={NAVY} onClick={() => onNavigate?.("assets")} />
          <MetricCard label="Open work orders" value={filteredWorkOrders.filter((w) => w.status !== "Closed").length} icon={ClipboardList} accentColor="#8B5CF6" onClick={() => onNavigate?.("work_orders")} />
          <MetricCard label="Open breakdowns" value={openBreakdownCount} icon={AlertTriangle} accentColor="#B85450" onClick={() => onNavigate?.("breakdowns")} />
          <MetricCard label="Services due soon / overdue" value={filteredPlannedMaintenance.filter((p) => p.status !== "OK").length} icon={CalendarClock} accentColor="#E8A33D" onClick={() => onNavigate?.("planned_maintenance")} />
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 12 }}>
          <h3 style={{ fontSize: 15, fontWeight: 700, margin: 0, color: NAVY }}>Currently in progress</h3>
          <span style={{ fontSize: 11.5, color: "#859195" }}>Live as of {nowTick.toLocaleTimeString("en-ZA", { hour: "2-digit", minute: "2-digit" })} - updates automatically</span>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 20, marginBottom: 24 }}>
          <div style={{ border: "1px solid #E2E6E3", borderRadius: 12, background: "#fff", overflow: "hidden" }}>
            <div onClick={() => onNavigate?.("breakdowns")} style={{ padding: "10px 14px", background: "#F6E2E0", fontSize: 12.5, fontWeight: 700, color: "#7A3330", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>Events - {inProgressBreakdowns.length} in progress</span>
              <ChevronRight size={14} />
            </div>
            {inProgressBreakdowns.length === 0 ? (
              <p style={{ padding: 14, fontSize: 13, color: "#859195", margin: 0 }}>No open events right now.</p>
            ) : (
              inProgressBreakdowns.map((b, i) => {
                const barColor = b.liveHours > 8 ? "#B85450" : b.liveHours > 2 ? "#E8A33D" : "#3F7D65";
                const barPct = Math.min(100, ((b.liveHours || 0) / 24) * 100);
                return (
                  <div key={b.id ?? i} onClick={() => onNavigate?.("breakdowns")} style={{ padding: "10px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none", cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                      <div style={{ minWidth: 0 }}>
                        <p style={{ fontSize: 13, fontWeight: 600, margin: 0, color: "#183642" }}>{b.asset_id} - {b.component_affected || "-"}</p>
                        <p style={{ fontSize: 11.5, color: "#859195", margin: "2px 0 0" }}>{b.cause_code || "No cause recorded"}</p>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: barColor, whiteSpace: "nowrap", marginLeft: 10 }}>{formatLiveHours(b.liveHours)}</span>
                    </div>
                    <div style={{ background: "#F2F1EA", borderRadius: 6, height: 6 }}>
                      <div style={{ width: `${barPct}%`, background: barColor, height: 6, borderRadius: 6, transition: "width 0.4s" }} />
                    </div>
                  </div>
                );
              })
            )}
          </div>

          <div style={{ border: "1px solid #E2E6E3", borderRadius: 12, background: "#fff", overflow: "hidden" }}>
            <div onClick={() => onNavigate?.("work_orders")} style={{ padding: "10px 14px", background: "#F5E9D8", fontSize: 12.5, fontWeight: 700, color: "#7A5320", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>Maintenance - {inProgressMaintenance.length} in progress</span>
              <ChevronRight size={14} />
            </div>
            {inProgressMaintenance.length === 0 ? (
              <p style={{ padding: 14, fontSize: 13, color: "#859195", margin: 0 }}>No maintenance work in progress right now.</p>
            ) : (
              inProgressMaintenance.map((w, i) => {
                const barPct = Math.min(100, ((w.liveHours || 0) / 24) * 100);
                return (
                  <div key={w.id ?? i} onClick={() => onNavigate?.("work_orders")} style={{ padding: "10px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none", cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                      <div style={{ minWidth: 0 }}>
                        <p style={{ fontSize: 13, fontWeight: 600, margin: 0, color: "#183642" }}>{w.asset_id} - {w.problem_scope || "-"}</p>
                        <p style={{ fontSize: 11.5, color: "#859195", margin: "2px 0 0" }}>{w.technician_vendor || "Unassigned"} · {w.work_type}</p>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#7A5320", whiteSpace: "nowrap", marginLeft: 10 }}>{formatLiveHours(w.liveHours)}</span>
                    </div>
                    <div style={{ background: "#F2F1EA", borderRadius: 6, height: 6 }}>
                      <div style={{ width: `${barPct}%`, background: "#E8A33D", height: 6, borderRadius: 6, transition: "width 0.4s" }} />
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>

        <EventTimeline breakdowns={filteredBreakdowns} workOrders={filteredWorkOrders} />

        {!monthKpiLoading && filteredMonthKpi.length > 0 && (() => {
          const mtbfTarget = 40, mttrTarget = 4, utilTarget = 85;
          const mtbfMax = Math.max(mtbfTarget, ...filteredMonthKpi.map((r) => r.mtbf || 0)) * 1.15;
          const mttrMax = Math.max(mttrTarget, ...filteredMonthKpi.map((r) => r.mttr || 0)) * 1.15;
          const utilData = filteredMonthKpi.map((r) => ({ ...r, utilisationPct: r.utilisation != null ? Math.round(r.utilisation * 100) : null }));
          return (
          <div style={{ marginBottom: 8 }}>
            <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>MTBF, MTTR & Utilisation, this month</h3>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 16, marginBottom: 8 }}>
              <KpiBarChart
                title="MTBF by equipment (hrs)" data={filteredMonthKpi} xKey="asset_id" dataKey="mtbf"
                target={mtbfTarget} domainMax={mtbfMax} unitSuffix="h"
                valueFormatter={(v) => Number(v).toFixed(1)}
                meetsTarget={(r) => (r.mtbf || 0) >= mtbfTarget}
                onClick={() => onNavigate?.("mtbf_mttr")}
              />
              <KpiBarChart
                title="MTTR by equipment (hrs)" data={filteredMonthKpi} xKey="asset_id" dataKey="mttr"
                target={mttrTarget} domainMax={mttrMax} unitSuffix="h"
                valueFormatter={(v) => Number(v).toFixed(1)}
                // Lower is better for MTTR, so "meets target" means at or below it.
                meetsTarget={(r) => (r.mttr || 0) <= mttrTarget}
                onClick={() => onNavigate?.("mtbf_mttr")}
              />
              <KpiBarChart
                title="Utilisation by equipment" data={utilData} xKey="asset_id" dataKey="utilisationPct"
                target={utilTarget} domainMax={100} unitSuffix="%"
                valueFormatter={(v) => `${v}%`}
                meetsTarget={(r) => (r.utilisationPct || 0) >= utilTarget}
                onClick={() => onNavigate?.("mtbf_mttr")}
              />
            </div>
            <KpiLegend targetLabel="Target" />
          </div>
          );
        })()}

        <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 20, marginBottom: 24, alignItems: "start" }}>
          <div>
            <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Availability by fleet, this month</h3>
            {monthKpiLoading ? (
              <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
            ) : fleetAvailability.length === 0 ? (
              <p style={{ fontSize: 13, color: "#859195" }}>No hours logged yet this month.</p>
            ) : (
              <>
                <KpiBarChart
                  title="" data={fleetAvailability} xKey="fleet" dataKey="availability"
                  target={85} domainMax={100} unitSuffix="%"
                  valueFormatter={(v) => `${v}%`}
                  meetsTarget={(r) => (r.availability || 0) >= 85}
                  onClick={() => onNavigate?.("fleet_performance")}
                />
                <KpiLegend targetLabel="Avail Target" />
              </>
            )}
          </div>

          <div>
            <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Fleet status</h3>
            <div onClick={() => onNavigate?.("assets")} style={{ height: 240, background: "#fff", border: "1px solid #E2E6E3", borderRadius: 12, padding: "8px 12px", cursor: "pointer" }}>
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={statusDistribution} dataKey="count" nameKey="status" cx="50%" cy="50%" innerRadius={45} outerRadius={75} paddingAngle={2}>
                    {statusDistribution.map((entry, i) => (
                      <Cell key={entry.status} fill={STATUS_COLORS[entry.status] || "#E2E6E3"} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 8 }}>
          <div>
            <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Alerts</h3>
            <div style={{ border: "1px solid #E2E6E3", borderRadius: 12, overflow: "hidden", background: "#fff" }}>
              {alertRows.map(([label, count, Icon, color, target], i) => (
                <div key={label} onClick={() => onNavigate?.(target)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "11px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none", cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <Icon size={15} style={{ color: count > 0 ? color : "#B4B2A9", flexShrink: 0 }} />
                    <span style={{ fontSize: 13.5, color: "#183642" }}>{label}</span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <Badge value={count > 0 ? String(count) : "0"} />
                    <ChevronRight size={14} style={{ color: "#B4B2A9" }} />
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div>
            <h3 style={{ fontSize: 15, fontWeight: 700, margin: "0 0 12px", color: NAVY }}>Top problem equipment</h3>
            <div style={{ border: "1px solid #E2E6E3", borderRadius: 12, background: "#fff", padding: "14px 16px" }}>
              {topProblemEquipment.length === 0 ? (
                <p style={{ fontSize: 13, color: "#859195", margin: 0 }}>No breakdowns logged yet.</p>
              ) : (
                topProblemEquipment.map((p, pi) => (
                  <div key={p.asset_id} onClick={() => onNavigate?.("breakdowns")} style={{ marginBottom: pi < topProblemEquipment.length - 1 ? 16 : 0, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                      <span style={{ color: "#183642", fontWeight: 700 }}>{p.asset_id} - {p.name}</span>
                      <span style={{ color: "#4B5659" }}>{p.totalCount} breakdown{p.totalCount === 1 ? "" : "s"} · {p.totalHours.toFixed(1)} hrs total</span>
                    </div>
                    <div style={{ background: "#F2F1EA", borderRadius: 6, height: 8, marginBottom: 8 }}>
                      <div style={{ width: `${(p.totalCount / maxProblemCount) * 100}%`, background: "#B85450", height: 8, borderRadius: 6 }} />
                    </div>
                    {p.entries.map((e) => (
                      <div key={e.type} style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: "#4B5659", padding: "2px 0 2px 10px" }}>
                        <span>· {e.type} ({e.count}×)</span>
                        <span>{e.hours.toFixed(1)} hrs</span>
                      </div>
                    ))}
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
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
    // Cost Ledger stays on mock data - hidden from the nav per standing request, no point wiring it live yet
    cost_ledger: { title: "Cost Ledger", cols: [["cost_date","Date"],["asset_id","Equipment #"],["cost_type","Cost type"],["description","Description"],["total_cost","Total cost"]], data: costLedger.map(c => ({...c, total_cost: `R${c.total_cost.toLocaleString()}`})) },
    audit: { title: "Audit Trail", cols: [["changed_at","Timestamp"],["changed_by_email","User"],["action","Action"],["table_name","Table"],["record_id","Record #"]], data: auditLog },
  };
}

export default function App({ userEmail, isAdmin, myRole = "manager", mySites = [], myPageAccess = [], myFullName, onNameSaved }) {
  const [active, setActive] = useState(() => (myRole === "operator" ? "daily_hours" : "dashboard"));
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [engineeringExpanded, setEngineeringExpanded] = useState(true);
  const [backlogsExpanded, setBacklogsExpanded] = useState(true);
  const isMobile = useIsMobile();
  const [selectedSiteId, setSelectedSiteId] = useState(mySites[0]?.id);

  // Live data from Supabase - Assets, Daily Hours, Breakdowns (the
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
  const [activityLog, setActivityLog] = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [backlogs, setBacklogs] = useState([]);
  const [dailyServiceChecklist, setDailyServiceChecklist] = useState([]);
  const [componentCodes, setComponentCodes] = useState([]);
  const [restLoading, setRestLoading] = useState(true);
  const [restError, setRestError] = useState(null);

  const loadCoreData = useCallback(async () => {
    if (!selectedSiteId) return;
    setCoreLoading(true);
    setCoreError(null);
    try {
      const [assetsRes, hoursRes, breakdownsRes, currentHoursRes] = await Promise.all([
        supabase.from("assets").select("*").eq("site_id", selectedSiteId),
        supabase.from("daily_hours_calc").select("*"),
        supabase.from("breakdown_log_calc").select("*"),
        supabase.from("current_hours").select("*"),
      ]);
      if (assetsRes.error) throw assetsRes.error;
      if (hoursRes.error) throw hoursRes.error;
      if (breakdownsRes.error) throw breakdownsRes.error;
      if (currentHoursRes.error) throw currentHoursRes.error;

      // assets table has no current_hours column of its own - it comes
      // from the current_hours view, merged in here by asset_id
      const currentHoursByAsset = Object.fromEntries(
        (currentHoursRes.data || []).map((r) => [r.asset_id, r.current_hours])
      );
      const mergedAssets = (assetsRes.data || []).map((a) => ({
        ...a,
        current_hours: currentHoursByAsset[a.asset_id] ?? null,
      }));

      // daily_hours_calc / breakdown_log_calc aren't filtered by site at
      // the query level (RLS only guarantees "some site this user can
      // access", which could be several for a multi-site manager) - so
      // narrow to just the currently selected site's assets here.
      const siteAssetIds = new Set(mergedAssets.map((a) => a.asset_id));

      setAssets(mergedAssets);
      setDailyHours((hoursRes.data || []).filter((r) => siteAssetIds.has(r.asset_id)));
      setBreakdowns((breakdownsRes.data || []).filter((r) => siteAssetIds.has(r.asset_id)));
    } catch (err) {
      setCoreError(err.message || String(err));
    } finally {
      setCoreLoading(false);
    }
  }, [selectedSiteId]);

  const loadRestOfData = useCallback(async () => {
    if (!selectedSiteId) return;
    setRestLoading(true);
    setRestError(null);
    try {
      const [
        siteAssetsRes, fuelRes, oilRes, woRes, serviceRes, inspRes,
        compRes, tyreRes, partsRes, warrRes, docRes, backlogsRes, dailyServiceRes, componentCodesRes,
      ] = await Promise.all([
        supabase.from("assets").select("asset_id").eq("site_id", selectedSiteId),
        supabase.from("fuel_log_calc").select("*"),
        supabase.from("oil_consumption_calc").select("*"),
        supabase.from("work_orders_calc").select("*"),
        supabase.from("service_schedule_calc").select("*"),
        supabase.from("inspections").select("*"),
        supabase.from("component_status").select("*"),
        supabase.from("tyre_tracking_calc").select("*"),
        supabase.from("parts_inventory_calc").select("*").eq("site_id", selectedSiteId),
        supabase.from("warranty_register_calc").select("*"),
        supabase.from("document_register_calc").select("*"),
        supabase.from("backlogs").select("*"),
        supabase.from("daily_service_checklist").select("*"),
        supabase.from("component_codes").select("*").order("name"),
      ]);
      const results = { siteAssetsRes, fuelRes, oilRes, woRes, serviceRes, inspRes, compRes, tyreRes, partsRes, warrRes, docRes, backlogsRes, dailyServiceRes, componentCodesRes };
      for (const [name, res] of Object.entries(results)) {
        if (res.error) throw new Error(`${name}: ${res.error.message}`);
      }

      // Fetched fresh here (rather than reused from the assets state set
      // by loadCoreData) to avoid any race between the two loads - this
      // is the same site-membership filter, just self-contained.
      const siteAssetIds = new Set((siteAssetsRes.data || []).map((a) => a.asset_id));
      const bySite = (rows) => (rows || []).filter((r) => siteAssetIds.has(r.asset_id));

      // Field names below are reshaped to match what the existing table
      // columns already expect, so DataTable/export don't need changes.
      setFuelLog(bySite(fuelRes.data).map((r) => ({ ...r, consumption_rate: r.consumption_rate_l_per_hr })));
      setOilConsumption(bySite(oilRes.data).map((r) => ({ ...r, consumption_rate: r.consumption_rate_l_per_hr })));
      setWorkOrders(bySite(woRes.data));
      setPlannedMaintenance(bySite(serviceRes.data).map((r) => ({
        ...r,
        remaining: r.current_hours != null ? r.next_service_due - r.current_hours : null,
      })));
      setInspections(bySite(inspRes.data));
      setComponents(bySite(compRes.data).map((r) => ({
        ...r,
        component_id: r.component_code,
        component_type: r.component_name,
        life_used_pct: r.expected_life_hours
          ? `${Math.round((r.current_comp_hrs / r.expected_life_hours) * 100)}%`
          : "-",
      })));
      setTyres(bySite(tyreRes.data));
      setParts(partsRes.data || []);

      const warrantyRows = bySite(warrRes.data).map((r) => ({
        asset_id: r.asset_id, type: `Warranty - ${r.component || ""}`.trim(),
        reference: r.serial_number, expiry: r.warranty_end, status: r.status,
      }));
      const docRows = bySite(docRes.data).map((r) => ({
        asset_id: r.asset_id, type: r.document_type,
        reference: r.reference, expiry: r.expiry_date, status: r.status,
      }));
      setWarrantyDocs([...warrantyRows, ...docRows]);
      setBacklogs(bySite(backlogsRes.data));
      setDailyServiceChecklist(bySite(dailyServiceRes.data));
      setComponentCodes(componentCodesRes.data || []);
    } catch (err) {
      setRestError(err.message || String(err));
    } finally {
      setRestLoading(false);
    }

    // Fetched separately and deliberately not allowed to affect the rest
    // of the app - a non-admin is correctly denied this by RLS, and that
    // denial should be silent from their point of view, not an error
    // banner blocking every other tab's data. Audit Trail intentionally
    // stays un-scoped by site - as the admin, seeing everything across
    // every site is the point of it.
    const { data: auditData, error: auditError } = await supabase.from("audit_report").select("*").limit(200);
    if (!auditError) setAuditLog(auditData || []);

    const { data: activityData, error: activityError } = await supabase.from("activity_log").select("*").order("created_at", { ascending: false }).limit(500);
    if (!activityError) setActivityLog(activityData || []);

    // Everyone's names, not just the current user's - needed to show who
    // did what throughout the Audit Trail regardless of who's looking.
    const { data: profilesData, error: profilesError } = await supabase.from("profiles").select("*");
    if (!profilesError) setProfiles(profilesData || []);
  }, [selectedSiteId]);

  useEffect(() => {
    loadCoreData();
    loadRestOfData();
  }, [loadCoreData, loadRestOfData]);

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
  // Single source of truth for the page title - NAV already has a label
  // for every key, so deriving it from there means any new standalone
  // page (like Downtime Summary or MTBF/MTTR) automatically gets a
  // correct title without needing a matching entry in a separate
  // hardcoded ternary chain. That mismatch is exactly what caused
  // activeConfig to be undefined for those two pages and crash the
  // whole view with no error boundary to catch it.
  const pageTitle = active === "about" ? "About" : NAV.find((n) => n.key === active)?.label || activeConfig?.title || "";

  // Defense in depth: if an operator's active tab is ever something
  // they're not supposed to see (role changed mid-session, a stale
  // value, anything), redirect them rather than silently rendering a
  // manager-level page. The nav already hides these, this is the
  // backstop for if it's reached some other way.
  useEffect(() => {
    if (myRole !== "operator") return;
    const navItem = NAV.find((n) => n.key === active);
    if (navItem && !navItem.operatorVisible) {
      setActive("daily_hours");
    }
  }, [myRole, active]);

  const handleNavClick = (key) => {
    setActive(key);
    if (isMobile) setSidebarOpen(false);
  };

  // Per-user page visibility, layered on top of the role system rather
  // than replacing it - an empty myPageAccess means this person has no
  // explicit restrictions, so their normal role-based access applies
  // unchanged. The moment an admin has ticked/unticked anything for
  // them, that list becomes their full, explicit allow-list. Admins are
  // never restricted by this, however their own access list looks.
  const canSeePage = (key) => isAdmin || myPageAccess.length === 0 || myPageAccess.includes(key);

  const sidebarWidth = isMobile ? 240 : (sidebarOpen ? 220 : 56);

  return (
    <div style={{ position: "relative", display: "flex", minHeight: 600, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif', background: "#fff", overflow: "hidden" }}>
      <style>{`
        html, body { overflow-x: hidden; max-width: 100%; }
        .nav-item:hover { background: rgba(255,255,255,0.08) !important; color: #fff !important; }
        tbody tr:hover td { background: #F7F8F6; }
        button:not(:disabled):hover { filter: brightness(0.96); }
        button:disabled { cursor: default; }
        input:focus, select:focus, textarea:focus { outline: none; border-color: #3DA35D !important; box-shadow: 0 0 0 3px rgba(61,163,93,0.15); }
        /* Momentum scrolling for the horizontally-scrollable tables/grids
           throughout the app, so dragging a wide table sideways on an
           iPhone/iPad feels native rather than sluggish. */
        div[style*="overflow-x:auto"], div[style*="overflow-x: auto"] { -webkit-overflow-scrolling: touch; }
        /* iOS Safari auto-zooms the whole page when a focused input's
           font-size is under 16px - this is the single most common
           "why does this feel broken on my phone" complaint on forms.
           Bumping every text input/select/textarea to 16px on phone-sized
           screens avoids that without changing how anything looks on a
           laptop. */
        @media (max-width: 480px) {
          input, select, textarea { font-size: 16px !important; }
        }
      `}</style>
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
          background: SIDEBAR,
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
          {sidebarOpen && (
            <button
              onClick={() => handleNavClick("about")}
              title="About Fleet Tracker"
              style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: GREEN, fontWeight: 700, fontSize: 15, textAlign: "left" }}
            >
              Fleet Tracker
            </button>
          )}
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

        {sidebarOpen && mySites.length > 1 && (
          <div style={{ padding: "10px 14px", borderBottom: "1px solid rgba(255,255,255,0.12)" }}>
            <label style={{ display: "block", fontSize: 10.5, fontWeight: 700, letterSpacing: 0.5, color: "rgba(255,255,255,0.5)", textTransform: "uppercase", marginBottom: 5 }}>Site</label>
            <select
              value={selectedSiteId}
              onChange={(e) => setSelectedSiteId(Number(e.target.value))}
              style={{ width: "100%", padding: "7px 8px", fontSize: 12.5, fontWeight: 600, border: "1px solid rgba(255,255,255,0.25)", borderRadius: 6, background: "rgba(255,255,255,0.08)", color: "#fff" }}
            >
              {mySites.map((s) => <option key={s.id} value={s.id} style={{ color: "#000" }}>{s.site_name}</option>)}
            </select>
          </div>
        )}

        <div style={{ overflowY: "auto", flex: 1 }}>
          {sidebarOpen && (
            <button
              onClick={() => setEngineeringExpanded((v) => !v)}
              style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", background: "none", border: "none", cursor: "pointer", padding: "14px 16px 6px" }}
            >
              <span style={{ color: "rgba(255,255,255,0.5)", fontSize: 11, fontWeight: 700, letterSpacing: 0.5, textTransform: "uppercase" }}>Engineering</span>
              <ChevronDown size={13} style={{ color: "rgba(255,255,255,0.5)", transform: engineeringExpanded ? "rotate(0deg)" : "rotate(-90deg)", transition: "transform 0.15s" }} />
            </button>
          )}
          {(engineeringExpanded || !sidebarOpen) && NAV.filter((n) =>
            !n.group &&
            ((n.key !== "audit" && n.key !== "site_management" && n.key !== "component_codes") || isAdmin) &&
            (myRole !== "operator" || n.operatorVisible) &&
            canSeePage(n.key)
          ).map(({ key, label, icon: Icon }) => (
            <button
              key={key}
              className="nav-item"
              onClick={() => handleNavClick(key)}
              title={label}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: isMobile ? "12px 16px" : "9px 16px", margin: "1px 8px", borderRadius: 8,
                background: active === key ? ACCENT : "transparent",
                border: "none", color: active === key ? "#fff" : "rgba(241,245,244,0.75)",
                cursor: "pointer", fontSize: 14, textAlign: "left", justifyContent: sidebarOpen ? "flex-start" : "center",
                width: sidebarOpen ? "calc(100% - 16px)" : "auto", minHeight: 40,
              }}
            >
              <Icon size={16} style={{ flexShrink: 0 }} />
              {sidebarOpen && <span>{label}</span>}
            </button>
          ))}

          {/* Backlogs - its own collapsible group inside Engineering, same
              expand/collapse pattern as the Engineering heading itself. */}
          {(engineeringExpanded || !sidebarOpen) && (
            <>
              {sidebarOpen ? (
                <button
                  onClick={() => setBacklogsExpanded((v) => !v)}
                  style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, width: "calc(100% - 16px)", background: "transparent", border: "none", cursor: "pointer", padding: isMobile ? "12px 16px" : "9px 16px", margin: "1px 8px", borderRadius: 8, color: "rgba(241,245,244,0.75)" }}
                >
                  <span style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 14 }}>
                    <AlertTriangle size={16} style={{ flexShrink: 0 }} />
                    Backlogs
                  </span>
                  <ChevronDown size={13} style={{ transform: backlogsExpanded ? "rotate(0deg)" : "rotate(-90deg)", transition: "transform 0.15s" }} />
                </button>
              ) : null}
              {(backlogsExpanded || !sidebarOpen) && NAV.filter((n) =>
                n.group === "backlogs" && (myRole !== "operator" || n.operatorVisible) && canSeePage(n.key)
              ).map(({ key, label, icon: Icon }) => (
                <button
                  key={key}
                  className="nav-item"
                  onClick={() => handleNavClick(key)}
                  title={label}
                  style={{
                    display: "flex", alignItems: "center", gap: 10,
                    padding: isMobile ? "12px 16px" : "9px 16px",
                    paddingLeft: sidebarOpen ? 30 : (isMobile ? 16 : 9),
                    margin: "1px 8px", borderRadius: 8,
                    background: active === key ? ACCENT : "transparent",
                    border: "none", color: active === key ? "#fff" : "rgba(241,245,244,0.75)",
                    cursor: "pointer", fontSize: 13.5, textAlign: "left", justifyContent: sidebarOpen ? "flex-start" : "center",
                    width: sidebarOpen ? "calc(100% - 16px)" : "auto", minHeight: 38,
                  }}
                >
                  <Icon size={15} style={{ flexShrink: 0 }} />
                  {sidebarOpen && <span>{label}</span>}
                </button>
              ))}
            </>
          )}

          {sidebarOpen && <p style={{ color: "rgba(255,255,255,0.5)", fontSize: 11, fontWeight: 700, letterSpacing: 0.5, margin: "18px 16px 6px", textTransform: "uppercase" }}>Production</p>}
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: isMobile ? "12px 16px" : "9px 16px", margin: "1px 8px", borderRadius: 8, color: "rgba(255,255,255,0.32)", fontSize: 14, cursor: "not-allowed", minHeight: 40 }}>
            <Factory size={16} style={{ flexShrink: 0 }} />
            {sidebarOpen && <span>Coming soon</span>}
          </div>

          <button
            onClick={() => supabase.auth.signOut()}
            style={{ display: "flex", alignItems: "center", gap: 10, padding: isMobile ? "12px 16px" : "9px 16px", margin: "12px 8px 8px", borderRadius: 8, background: "transparent", border: "none", color: "rgba(241,245,244,0.75)", fontSize: 14, cursor: "pointer", minHeight: 40, justifyContent: sidebarOpen ? "flex-start" : "center" }}
          >
            <LogOut size={16} style={{ flexShrink: 0 }} />
            {sidebarOpen && <span>Sign out</span>}
          </button>
        </div>
      </div>

      <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        {/* Mobile top bar */}
        {isMobile && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "12px 16px", borderBottom: "1px solid #E2E6E3", position: "sticky", top: 0, background: "#fff", zIndex: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <button onClick={() => setSidebarOpen(true)} aria-label="Open menu" style={{ background: "none", border: "none", color: NAVY, cursor: "pointer", padding: 6, display: "flex" }}>
                <Menu size={20} />
              </button>
              <span style={{ fontSize: 16, fontWeight: 700, color: NAVY }}>
                {pageTitle}
              </span>
            </div>
            <UserMenu myFullName={myFullName} isAdmin={isAdmin} myRole={myRole} />
          </div>
        )}

        <div style={{ padding: isMobile ? "16px" : "24px 28px", overflow: "auto", flex: 1 }}>
          {!isMobile && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
              <h2 style={{ fontSize: 19, fontWeight: 700, margin: 0, color: NAVY }}>
                {pageTitle}
              </h2>
              <UserMenu myFullName={myFullName} isAdmin={isAdmin} myRole={myRole} />
            </div>
          )}

          {(coreError || restError) && (
            <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 8, padding: "10px 14px", marginBottom: 16, fontSize: 13, color: "#7A3330" }}>
              Couldn't load live data from Supabase: {coreError || restError}. Check your .env values and that the SQL files ran successfully.
            </div>
          )}

          <PageErrorBoundary key={active}>
            {active === "about" ? (
              <AboutPage />
            ) : !canSeePage(active) ? (
              <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 10, padding: 28, textAlign: "center" }}>
                <p style={{ fontWeight: 700, fontSize: 14, color: "#7A3330", margin: "0 0 6px" }}>Access restricted</p>
                <p style={{ fontSize: 13, color: "#7A3330", margin: 0 }}>You don't have access to this page. Ask your administrator if you need it.</p>
              </div>
            ) : !selectedSiteId ? (
              <div style={{ background: "#F5E9D8", border: "1px solid #E3C79B", borderRadius: 10, padding: 24, textAlign: "center" }}>
                <p style={{ fontWeight: 700, fontSize: 14, color: "#7A5320", margin: "0 0 6px" }}>No site selected</p>
                <p style={{ fontSize: 13, color: "#7A5320", margin: 0 }}>
                  Either your account has no site access yet, or App.jsx and EngineeringApp.jsx are from different versions - make sure both files were updated together.
                </p>
              </div>
            ) : (coreLoading && ["dashboard", "fleet_performance", "assets", "daily_hours", "breakdowns", "downtime_summary"].includes(active)) ||
             (restLoading && ["dashboard", "fuel_log", "oil_consumption", "work_orders", "downtime_summary", "planned_maintenance", "inspections", "components", "tyres", "parts", "warranty_docs", "audit", "backlog_report", "daily_service", "breakdowns", "component_codes"].includes(active)) ? (
              <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
            ) : active === "dashboard" ? (
              <Dashboard assets={assets} breakdowns={breakdowns} workOrders={workOrders} plannedMaintenance={plannedMaintenance} components={components} parts={parts} inspections={inspections} onNavigate={setActive} />
            ) : active === "fleet_performance" ? (
              <FleetPerformance assets={assets} breakdowns={breakdowns} />
            ) : active === "breakdowns" ? (
              <BreakdownsPage assets={assets} breakdowns={breakdowns} workOrders={workOrders} parts={parts} componentCodes={componentCodes} onRefresh={() => { loadCoreData(); loadRestOfData(); }} userEmail={userEmail} myFullName={myFullName} />
            ) : active === "assets" ? (
              <AssetsPage assets={assets} selectedSiteId={selectedSiteId} onRefresh={loadCoreData} />
            ) : active === "daily_hours" ? (
              <DailyHoursPage assets={assets} dailyHours={dailyHours} userEmail={userEmail} onRefresh={loadCoreData} />
            ) : active === "planned_maintenance" ? (
              <PlannedMaintenancePage assets={assets} plannedMaintenance={plannedMaintenance} workOrders={workOrders} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "work_orders" ? (
              <WorkOrdersPage assets={assets} workOrders={workOrders} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "downtime_summary" ? (
              <DowntimeSummaryPage assets={assets} breakdowns={breakdowns} workOrders={workOrders} />
            ) : active === "mtbf_mttr" ? (
              <MtbfMttrReportPage assets={assets} />
            ) : active === "fuel_log" ? (
              <FuelLogPage assets={assets} fuelLog={fuelLog} userEmail={userEmail} dailyHours={dailyHours} onRefresh={loadRestOfData} />
            ) : active === "oil_consumption" ? (
              <OilConsumptionPage assets={assets} oilConsumption={oilConsumption} userEmail={userEmail} dailyHours={dailyHours} onRefresh={loadRestOfData} />
            ) : active === "inspections" ? (
              <InspectionsPage assets={assets} inspections={inspections} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "backlog_report" ? (
              <BacklogsPage assets={assets} backlogs={backlogs} workOrders={workOrders} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "daily_service" ? (
              <DailyServicePage assets={assets} dailyServiceChecklist={dailyServiceChecklist} breakdowns={breakdowns} dailyHours={dailyHours} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "components" ? (
              <ComponentsPage assets={assets} components={components} breakdowns={breakdowns} workOrders={workOrders} dailyHours={dailyHours} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "parts" ? (
              <PartsPage parts={parts} selectedSiteId={selectedSiteId} onRefresh={loadRestOfData} />
            ) : active === "audit" ? (
              <AuditTrailPage activityLog={activityLog} profiles={profiles} isAdmin={isAdmin} />
            ) : active === "site_management" ? (
              <SiteManagementPage isAdmin={isAdmin} onSitesChanged={() => window.location.reload()} onNameSaved={onNameSaved} />
            ) : active === "component_codes" ? (
              <ComponentCodesAdminPage componentCodes={componentCodes} isAdmin={isAdmin} onRefresh={loadRestOfData} />
            ) : (
              <DataTable
                columns={activeConfig.cols.map(([key, label]) => ({ key, label }))}
                rows={activeConfig.data}
                exportName={activeConfig.title}
              />
            )}
          </PageErrorBoundary>
        </div>
      </div>
    </div>
  );
}

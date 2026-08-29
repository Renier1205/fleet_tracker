import React, { useState, useMemo, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Truck, Clock, AlertTriangle, ClipboardList,
  CalendarClock, ShieldCheck, CircleDot, Package, FileText,
  DollarSign, History, Factory, Download, Search, Menu, X,
  Layers, ChevronRight, GitCompare, Fuel, Droplet, LogOut, Upload, Trash2, Printer, FileBarChart, ChevronDown, Activity, MapPin,
  Wrench, BarChart3, ShoppingCart, Users, TrendingUp, Boxes, FilePlus, ShieldAlert
} from "lucide-react";
import { supabase } from "./lib/supabaseClient";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend, ReferenceLine, LabelList,
  LineChart, Line,
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
  { key: "work_requests", label: "Work Requests", icon: FilePlus, operatorVisible: true },
  { key: "defects", label: "Defect Register", icon: ShieldAlert, operatorVisible: true },
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
  { key: "component_master", label: "Component Master", icon: Package },
  { key: "component_status", label: "Component Status", icon: Activity },
  { key: "fleet_health", label: "Fleet Health", icon: Activity },
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
  if (["OVERDUE", "CRITICAL", "CHANGE OUT", "REORDER", "REPLACE NOW", "EXPIRED", "FAIL", "OPEN", "OVER SCHEDULED HOURS", "REJECTED"].includes(s))
    return { bg: "#F6E2E0", text: "#7A3330" };
  if (["DUE SOON", "PLAN CHANGE", "EXPIRING SOON", "MEDIUM", "PLANNED", "APPROACHING LIMIT", "WORK ORDER CREATED", "IN PROGRESS", "AWAITING PARTS"].includes(s))
    return { bg: "#F5E9D8", text: "#7A5320" };
  if (["OK", "CLOSED", "PASS", "ACTIVE", "VALID", "APPROVED", "CONVERTED", "MERGED", "COMPLETED", "VERIFIED"].includes(s))
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
    const dataRows = visibleRows.map((row) => columns.map((c) => row[c.key] ?? ""));

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
                    : c.key === "consumption_rate" && (val == null)
                      ? <span title="Can't calculate a rate - this hour meter reading matches (or is lower than) the previous fill for this machine. Check the hour meter value on this entry." style={{ display: "inline-flex", alignItems: "center", gap: 4, color: "#A65A1E", fontWeight: 600 }}>
                          <AlertTriangle size={13} /> Check reading
                        </span>
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
        {/* No nowrap/ellipsis here on purpose - on a narrow phone card,
            two-line wrapped text ("Services due soon /\noverdue") is
            readable; a single truncated line ("Servi...") isn't. */}
        <p style={{ fontSize: 12, color: "#4B5659", margin: "0 0 2px", lineHeight: 1.3 }}>{label}</p>
        <p style={{ fontSize: 21, fontWeight: 700, margin: 0, color: "#183642" }}>{value}</p>
      </div>
      {onClick && <ChevronRight size={16} style={{ color: "#B4B2A9", flexShrink: 0 }} />}
    </div>
  );
}

const CAUSE_CODES = ["Mechanical Failure", "Electrical Fault", "Hydraulic Failure", "Tyre/Track",
  "Operator Error", "Wear & Tear", "Structural", "Overheating", "Lubrication/Fluid", "Weather", "Unknown", "Other"];
// Downtime Responsibility - WHOSE account the lost hours go to. This is
// separate from Cause, which records WHAT broke: a hydraulic failure can
// be Plant's or the OEM's depending on circumstance.
// Engineering Availability / Engineering MTBF count PLANT and STORES
// only; the rest are lost time outside engineering's control.
const ACTION_TAKEN_OPTIONS = ["Repaired in place", "Adjusted", "Component replaced", "No fault found"];
const CHANGEOUT_DOC_TYPES = ["Commission report", "Component certificate", "Oil sample", "Failure photos", "OEM inspection", "Return note", "Invoice", "Other"];
const COMPONENT_CONDITIONS = ["New", "Reman", "Repaired"];
const CLAIM_STATUSES = ["Not lodged", "Lodged", "Under assessment", "Approved", "Rejected", "Not applicable"];

const RESPONSIBILITY_CODES = ["Plant", "Stores", "Production", "OEM", "Client", "Uncontrollable Time", "Non-Shift"];
const ENGINEERING_RESPONSIBILITY = ["Plant", "Stores"];

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

// Default column sets live at module level so their identity is stable
// across renders - useTablePrefs memoises against them.
const DAILY_HOURS_COLUMNS = [
  ["asset_id", "Equipment #"], ["fleet", "Fleet"], ["log_date", "Date"], ["shift", "Shift"],
  ["opening_hours", "Opening hours"], ["closing_hours", "Closing hours"], ["hours_run", "Hours run"],
].map(([key, label]) => ({ key, label }));

const FUEL_LOG_COLUMNS = [
  ["fill_date", "Date"], ["fill_time", "Time"], ["asset_id", "Equipment #"], ["hour_meter", "Hour meter"],
  ["litres", "Litres"], ["consumption_rate", "Rate (L/hr)"], ["recorded_by", "Recorded by"],
].map(([key, label]) => ({ key, label }));

const OIL_CONSUMPTION_COLUMNS = [
  ["fill_date", "Date"], ["asset_id", "Equipment #"], ["hour_meter", "Hour meter"], ["oil_type", "Oil type"],
  ["litres", "Litres"], ["fill_reason", "Reason"], ["consumption_rate", "Rate (L/hr)"], ["recorded_by", "Recorded by"],
].map(([key, label]) => ({ key, label }));

// ---------------------------------------------------------------------
// Per-user table preferences - column order, hidden columns and sort.
//
// Saved to Supabase (user_table_prefs) keyed on user_id + table_key, so
// a person's layout follows them to any device. Every read and write is
// wrapped so that a missing table, an RLS refusal or an offline moment
// silently falls back to the built-in defaults. A table must never fail
// to render because a preference couldn't be loaded.
// ---------------------------------------------------------------------
function useTablePrefs(tableKey, defaultColumns, initialSort) {
  const [order, setOrder] = useState(null);
  const [hidden, setHidden] = useState([]);
  // Per-column pixel widths, dragged by the user and saved with the rest
  // of their layout. Absent keys fall back to the table's own default.
  const [widths, setWidths] = useState({});
  const [sortKey, setSortKey] = useState(initialSort?.key ?? null);
  const [sortDir, setSortDir] = useState(initialSort?.dir ?? "desc");
  const userIdRef = useRef(null);
  const loadedRef = useRef(false);
  const saveTimer = useRef(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { data: auth } = await supabase.auth.getUser();
        const user = auth?.user;
        if (!user || cancelled) return;
        userIdRef.current = user.id;
        const { data, error } = await supabase
          .from("user_table_prefs")
          .select("column_order, hidden_columns, sort_key, sort_dir, column_widths")
          .eq("user_id", user.id)
          .eq("table_key", tableKey)
          .maybeSingle();
        if (error || cancelled || !data) return;
        if (Array.isArray(data.column_order) && data.column_order.length) setOrder(data.column_order);
        if (Array.isArray(data.hidden_columns)) setHidden(data.hidden_columns);
        if (data.column_widths && typeof data.column_widths === "object") setWidths(data.column_widths);
        if (data.sort_key) setSortKey(data.sort_key);
        if (data.sort_dir) setSortDir(data.sort_dir);
      } catch (err) {
        // Defaults are a perfectly good outcome - nothing to surface.
      } finally {
        if (!cancelled) loadedRef.current = true;
      }
    })();
    return () => { cancelled = true; if (saveTimer.current) clearTimeout(saveTimer.current); };
  }, [tableKey]);

  // Debounced so dragging three columns in a row is one write, not three.
  const persist = useCallback((next) => {
    if (!userIdRef.current || !loadedRef.current) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      try {
        supabase.from("user_table_prefs").upsert({
          user_id: userIdRef.current,
          table_key: tableKey,
          column_order: next.order,
          hidden_columns: next.hidden,
          column_widths: next.widths ?? {},
          sort_key: next.sortKey,
          sort_dir: next.sortDir,
          updated_at: new Date().toISOString(),
        }, { onConflict: "user_id,table_key" }).then(
          () => {},
          () => {}
        );
      } catch (err) {
        // A preference that didn't save is not worth interrupting anyone over.
      }
    }, 600);
  }, [tableKey]);

  const allKeys = useMemo(() => defaultColumns.map((c) => c.key), [defaultColumns]);

  // Stored order can go stale when columns are added or removed in a
  // release - anything unknown is dropped, anything new lands at the end.
  const effectiveOrder = useMemo(() => {
    const base = (order || []).filter((k) => allKeys.includes(k));
    return [...base, ...allKeys.filter((k) => !base.includes(k))];
  }, [order, allKeys]);

  const columns = useMemo(() => {
    const byKey = new Map(defaultColumns.map((c) => [c.key, c]));
    return effectiveOrder.filter((k) => !hidden.includes(k)).map((k) => byKey.get(k)).filter(Boolean);
  }, [defaultColumns, effectiveOrder, hidden]);

  const allColumns = useMemo(() => {
    const byKey = new Map(defaultColumns.map((c) => [c.key, c]));
    return effectiveOrder.map((k) => byKey.get(k)).filter(Boolean);
  }, [defaultColumns, effectiveOrder]);

  const setColumnWidth = useCallback((key, px) => {
    setWidths((prev) => {
      const next = { ...prev, [key]: Math.max(48, Math.round(px)) };
      persist({ order: effectiveOrder, hidden, widths: next, sortKey, sortDir });
      return next;
    });
  }, [effectiveOrder, hidden, sortKey, sortDir, persist]);

  const moveColumn = useCallback((fromKey, toKey) => {
    if (!fromKey || !toKey || fromKey === toKey) return;
    setOrder((prev) => {
      const base = (prev && prev.length ? prev : allKeys).filter((k) => allKeys.includes(k));
      const merged = [...base, ...allKeys.filter((k) => !base.includes(k))];
      const from = merged.indexOf(fromKey), to = merged.indexOf(toKey);
      if (from < 0 || to < 0) return prev;
      merged.splice(to, 0, merged.splice(from, 1)[0]);
      persist({ order: merged, hidden, widths, sortKey, sortDir });
      return merged;
    });
  }, [allKeys, hidden, sortKey, sortDir, persist]);

  const nudgeColumn = useCallback((key, delta) => {
    const idx = effectiveOrder.indexOf(key);
    const target = effectiveOrder[idx + delta];
    if (target) moveColumn(key, target);
  }, [effectiveOrder, moveColumn]);

  const toggleHidden = useCallback((key) => {
    setHidden((prev) => {
      const next = prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key];
      // Hiding the last visible column would leave an empty grid.
      if (next.length >= allKeys.length) return prev;
      persist({ order: effectiveOrder, hidden: next, widths, sortKey, sortDir });
      return next;
    });
  }, [allKeys, effectiveOrder, sortKey, sortDir, persist]);

  const toggleSort = useCallback((key) => {
    const nextDir = key === sortKey ? (sortDir === "asc" ? "desc" : "asc") : "asc";
    setSortKey(key);
    setSortDir(nextDir);
    persist({ order: effectiveOrder, hidden, widths, sortKey: key, sortDir: nextDir });
  }, [sortKey, sortDir, effectiveOrder, hidden, persist]);

  const resetPrefs = useCallback(() => {
    setOrder(null);
    setHidden([]);
    setWidths({});
    setSortKey(initialSort?.key ?? null);
    setSortDir(initialSort?.dir ?? "desc");
    persist({ order: allKeys, hidden: [], widths: {}, sortKey: initialSort?.key ?? null, sortDir: initialSort?.dir ?? "desc" });
  }, [allKeys, initialSort, persist]);

  // Sorts by the chosen column while leaving the caller's own ordering
  // untouched when no sort column is set.
  const sortRows = useCallback((rows) => {
    if (!sortKey) return rows;
    const val = (r) => {
      const v = r?.[sortKey];
      if (v == null || v === "") return null;
      const n = Number(v);
      return Number.isNaN(n) || typeof v === "boolean" ? String(v).toLowerCase() : n;
    };
    return [...rows].sort((a, b) => {
      const va = val(a), vb = val(b);
      if (va == null && vb == null) return 0;
      if (va == null) return 1;   // blanks always sink, either direction
      if (vb == null) return -1;
      if (va < vb) return sortDir === "asc" ? -1 : 1;
      if (va > vb) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
  }, [sortKey, sortDir]);

  return { columns, allColumns, hidden, widths, setColumnWidth, sortKey, sortDir, toggleSort, moveColumn, nudgeColumn, toggleHidden, resetPrefs, sortRows };
}

// Header row: click to sort, drag to reorder. Drag uses the browser's
// native HTML5 events, which do nothing on a touchscreen - that's what
// the Columns panel below is for.
function SmartTableHead({ prefs, trailingCell = true, defaultWidth }) {
  const [dragKey, setDragKey] = useState(null);
  const [overKey, setOverKey] = useState(null);

  // Column resizing. Pointer events rather than HTML5 drag, so it works
  // on a touchscreen as well as with a mouse - the same reason the
  // Columns panel exists alongside header dragging.
  const startResize = (e, key, th) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = th.getBoundingClientRect().width;
    const move = (ev) => {
      const w = Math.max(48, startW + (ev.clientX - startX));
      th.style.width = `${w}px`;
      th.style.minWidth = `${w}px`;
    };
    const up = (ev) => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      prefs.setColumnWidth(key, Math.max(48, startW + (ev.clientX - startX)));
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  return (
    <thead>
      <tr style={{ background: "#F7F8F6" }}>
        {prefs.columns.map((c) => {
          const isOver = overKey === c.key && dragKey && dragKey !== c.key;
          return (
            <th
              key={c.key}
              draggable
              onDragStart={(e) => { setDragKey(c.key); e.dataTransfer.effectAllowed = "move"; }}
              onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; setOverKey(c.key); }}
              onDragLeave={() => setOverKey((k) => (k === c.key ? null : k))}
              onDrop={(e) => { e.preventDefault(); prefs.moveColumn(dragKey, c.key); setDragKey(null); setOverKey(null); }}
              onDragEnd={() => { setDragKey(null); setOverKey(null); }}
              onClick={() => prefs.toggleSort(c.key)}
              title="Click to sort, drag to move, drag the edge to resize"
              style={{
                position: "relative",
                textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659",
                whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3", cursor: "pointer",
                userSelect: "none", opacity: dragKey === c.key ? 0.4 : 1,
                boxShadow: isOver ? `inset 3px 0 0 ${NAVY}` : "none",
                width: prefs.widths?.[c.key] ?? defaultWidth?.(c.key) ?? undefined,
                minWidth: prefs.widths?.[c.key] ?? undefined,
              }}
            >
              {c.label}{" "}
              <span style={{ display: "inline-block", width: 12, textAlign: "center", color: prefs.sortKey === c.key ? NAVY : "#C7C5BB" }}>
                {prefs.sortKey === c.key ? (prefs.sortDir === "asc" ? "▲" : "▼") : "↕"}
              </span>
              <span
                onPointerDown={(e) => startResize(e, c.key, e.currentTarget.parentElement)}
                onClick={(e) => e.stopPropagation()}
                onDragStart={(e) => e.preventDefault()}
                title="Drag to resize"
                style={{ position: "absolute", top: 0, right: 0, height: "100%", width: 8, cursor: "col-resize", touchAction: "none" }}
              />
            </th>
          );
        })}
        {trailingCell && <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>}
      </tr>
    </thead>
  );
}

// The touch-friendly half of the feature: reorder with arrows, show and
// hide with checkboxes. Works identically on desktop, tablet and phone.
function ColumnsButton({ prefs }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}
      >
        <Layers size={14} /> Columns
      </button>
      {open && (
        <>
          <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 40 }} />
          <div style={{ position: "absolute", right: 0, top: "calc(100% + 6px)", zIndex: 41, background: "#fff", border: "1px solid #E2E6E3", borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,0.12)", padding: 10, width: 268, maxHeight: 340, overflowY: "auto" }}>
            <p style={{ margin: "0 0 8px", fontSize: 12, color: "#859195" }}>Tick to show. Use the arrows to reorder.</p>
            {prefs.allColumns.map((c, i) => (
              <div key={c.key} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 2px" }}>
                <input
                  type="checkbox"
                  checked={!prefs.hidden.includes(c.key)}
                  onChange={() => prefs.toggleHidden(c.key)}
                  style={{ cursor: "pointer" }}
                />
                <span style={{ flex: 1, fontSize: 13, color: "#183642" }}>{c.label}</span>
                <button type="button" disabled={i === 0} onClick={() => prefs.nudgeColumn(c.key, -1)}
                  style={{ background: "none", border: "1px solid #E2E6E3", borderRadius: 6, padding: "2px 7px", cursor: i === 0 ? "default" : "pointer", opacity: i === 0 ? 0.3 : 1 }}>↑</button>
                <button type="button" disabled={i === prefs.allColumns.length - 1} onClick={() => prefs.nudgeColumn(c.key, 1)}
                  style={{ background: "none", border: "1px solid #E2E6E3", borderRadius: 6, padding: "2px 7px", cursor: i === prefs.allColumns.length - 1 ? "default" : "pointer", opacity: i === prefs.allColumns.length - 1 ? 0.3 : 1 }}>↓</button>
              </div>
            ))}
            <button type="button" onClick={prefs.resetPrefs}
              style={{ marginTop: 8, width: "100%", background: "#F7F8F6", border: "1px solid #E2E6E3", color: "#4B5659", fontSize: 12.5, fontWeight: 600, padding: "7px 10px", borderRadius: 8, cursor: "pointer" }}>
              Reset to default
            </button>
          </div>
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------
// Machine grouping, shared by the tabs that show per-machine sequences.
//
// buildMachineGroups always returns an array of groups. With grouping
// switched off it returns a single unbannered group holding every row,
// so the body renders a plain flat list - which is what you want before
// an export.
// ---------------------------------------------------------------------
function buildMachineGroups(rows, assets, { enabled, dateKey, hoursKey, sortKey, sortDir, sortRows, limitPerMachine }) {
  if (!enabled) {
    return [{ key: "__all__", banner: null, rows: sortKey ? sortRows(rows) : rows }];
  }

  const byAsset = new Map();
  for (const r of rows) {
    if (!byAsset.has(r.asset_id)) byAsset.set(r.asset_id, []);
    byAsset.get(r.asset_id).push(r);
  }

  const ids = [...byAsset.keys()].sort((a, b) =>
    sortKey === "asset_id" && sortDir === "desc"
      ? String(b).localeCompare(String(a))
      : String(a).localeCompare(String(b))
  );

  return ids.map((id) => {
    // Window from the newest end first, then sort - so changing the sort
    // never changes which rows are on screen, only their order.
    const newestFirst = [...byAsset.get(id)].sort((a, b) =>
      String(b?.[dateKey] || "").localeCompare(String(a?.[dateKey] || ""))
    );
    const windowed = limitPerMachine ? newestFirst.slice(0, limitPerMachine) : newestFirst;
    const finalRows = (sortKey && sortKey !== "asset_id") ? sortRows(windowed) : windowed;
    const asset = assets.find((a) => a.asset_id === id);
    return {
      key: id,
      assetId: id,
      assetName: asset?.asset_name || "",
      fleet: asset?.fleet || "",
      total: hoursKey ? finalRows.reduce((sum, r) => sum + (Number(r?.[hoursKey]) || 0), 0) : null,
      rows: finalRows,
      banner: true,
    };
  }).filter((g) => g.rows.length > 0);
}

function MachineTableBody({ groups, columns, onRowClick, onDelete, emptyMessage, totalLabel = "hrs", countNoun = "entry", countNounPlural = "entries" }) {
  const span = columns.length + 1;

  if (!groups.some((g) => g.rows.length > 0)) {
    return (
      <tbody>
        <tr><td colSpan={span} style={{ padding: 20, textAlign: "center", color: "#859195" }}>{emptyMessage}</td></tr>
      </tbody>
    );
  }

  return (
    <tbody>
      {groups.map((g, gi) => (
        <React.Fragment key={g.key}>
          {g.banner && (
            <tr>
              <td colSpan={span} style={{ padding: "8px 12px", background: "#F2F1EA", fontWeight: 700, fontSize: 12.5, color: NAVY, borderTop: gi > 0 ? "1px solid #E2E6E3" : "none" }}>
                {g.assetId}{g.assetName ? ` - ${g.assetName}` : ""}
                <span style={{ fontWeight: 500, color: "#4B5659" }}>
                  {g.fleet ? `  ·  ${g.fleet}` : ""}
                  {`  ·  ${g.rows.length} ${g.rows.length === 1 ? countNoun : countNounPlural}`}
                  {g.total != null ? `  ·  ${g.total.toFixed(1)} ${totalLabel}` : ""}
                </span>
              </td>
            </tr>
          )}
          {g.rows.map((row, i) => (
            <tr key={row.id ?? `${g.key}-${i}`} style={{ borderBottom: "1px solid #EFEEE7", background: row.exceeds_shift_limit ? "#F6E2E0" : "transparent" }}>
              {columns.map((c) => (
                <td key={c.key} onClick={() => onRowClick && onRowClick(row)} style={{ padding: "9px 12px", whiteSpace: "nowrap", cursor: onRowClick ? "pointer" : "default" }}>
                  {row[c.key] ?? <span style={{ color: "#B4B2A9" }}>-</span>}
                </td>
              ))}
              <td style={{ padding: "9px 12px" }}>
                {onDelete && (
                  <button onClick={() => onDelete(row)} title="Delete" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "inline-flex" }}>
                    <Trash2 size={15} />
                  </button>
                )}
              </td>
            </tr>
          ))}
        </React.Fragment>
      ))}
    </tbody>
  );
}

// For tables whose rows are hand-written (badges, action buttons), the
// generic body above doesn't fit. These two let a page keep its own row
// markup and just interleave machine banners into it.
function flattenGroups(groups) {
  const out = [];
  for (const g of groups) {
    if (g.banner) out.push({ kind: "banner", group: g });
    for (const row of g.rows) out.push({ kind: "row", row, group: g });
  }
  return out;
}

function MachineBannerRow({ group, colSpan, first, totalLabel = "hrs", countNoun = "entry", countNounPlural = "entries" }) {
  return (
    <tr>
      <td colSpan={colSpan} style={{ padding: "8px 12px", background: "#F2F1EA", fontWeight: 700, fontSize: 12.5, color: NAVY, borderTop: first ? "none" : "1px solid #E2E6E3" }}>
        {group.assetId}{group.assetName ? ` - ${group.assetName}` : ""}
        <span style={{ fontWeight: 500, color: "#4B5659" }}>
          {group.fleet ? `  ·  ${group.fleet}` : ""}
          {`  ·  ${group.rows.length} ${group.rows.length === 1 ? countNoun : countNounPlural}`}
          {group.total != null ? `  ·  ${group.total.toFixed(1)} ${totalLabel}` : ""}
        </span>
      </td>
    </tr>
  );
}

function GroupByMachineToggle({ value, onChange }) {
  return (
    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#4B5659", cursor: "pointer", whiteSpace: "nowrap" }}>
      <input type="checkbox" checked={value} onChange={(e) => onChange(e.target.checked)} style={{ cursor: "pointer" }} />
      Group by machine
    </label>
  );
}

function todayForInput() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// Shifts a YYYY-MM-DD string by whole days. Built at midday so a daylight
// saving jump can't roll the result onto the wrong date.
function firstOfMonthForInput() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-01`;
}

function addDaysToInput(dateStr, days) {
  if (!dateStr) return dateStr;
  const d = new Date(`${dateStr}T12:00:00`);
  d.setDate(d.getDate() + days);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// Renders a date as e.g. "Tue 13 Aug" for the shift confirmation line.
function shortDateLabel(dateStr) {
  if (!dateStr) return "";
  return new Date(`${dateStr}T12:00:00`).toLocaleDateString("en-ZA", { weekday: "short", day: "2-digit", month: "short" });
}

// Formats a price with a space as the thousands separator and 2 decimal
// places (e.g. 16938 -> "16 938.00") - easier to read at a glance than a
// long run of digits, without relying on a locale's own grouping/decimal
// conventions (which vary and can be inconsistent across browsers).
function formatMoney(value) {
  if (value == null || value === "") return null;
  const num = Number(value);
  if (Number.isNaN(num)) return null;
  const [whole, decimals] = num.toFixed(2).split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return `${grouped}.${decimals}`;
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

function DailyHoursForm({ assets, dailyHours, existing, selectedSiteId, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [captureDate, setCaptureDate] = useState(existing?.log_date || todayForInput());
  const [closingHours, setClosingHours] = useState(existing?.closing_hours ?? "");
  const [status, setStatus] = useState(existing?.status || "Operating");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  // Equipment picker filters - a long fleet list is painful to scroll on
  // a tablet at the workshop door.
  const [fleetFilter, setFleetFilter] = useState("");
  const [assetQuery, setAssetQuery] = useState("");

  // Shift changeover times come from the site, so a site running 06:00 /
  // 18:00 gets its own options rather than the 07:00 / 19:00 default.
  const [shiftTimes, setShiftTimes] = useState({ day: "07:00", night: "19:00" });
  const [captureTime, setCaptureTime] = useState("");

  useEffect(() => {
    let cancelled = false;
    (async () => {
      if (!selectedSiteId) return;
      try {
        const { data, error: siteErr } = await supabase
          .from("sites").select("day_shift_start, night_shift_start").eq("id", selectedSiteId).maybeSingle();
        if (siteErr || cancelled || !data) return;
        setShiftTimes({
          day: (data.day_shift_start || "07:00").slice(0, 5),
          night: (data.night_shift_start || "19:00").slice(0, 5),
        });
      } catch (err) {
        // 07:00 / 19:00 remain in place - never block logging over this.
      }
    })();
    return () => { cancelled = true; };
  }, [selectedSiteId]);

  // Default to whichever changeover has most recently passed.
  useEffect(() => {
    if (isEdit || captureTime) return;
    const nowHM = new Date().toTimeString().slice(0, 5);
    setCaptureTime(nowHM >= shiftTimes.night || nowHM < shiftTimes.day ? shiftTimes.night : shiftTimes.day);
  }, [shiftTimes, isEdit, captureTime]);

  // The rule: a reading at the morning changeover closes the NIGHT shift
  // that began the evening before, so it files against the previous day.
  // A reading at the evening changeover closes that same day's DAY shift.
  // Operators only ever see the date they're standing in.
  const derived = useMemo(() => {
    if (isEdit) return { logDate: existing.log_date, shift: existing.shift };
    if (captureTime === shiftTimes.day) return { logDate: addDaysToInput(captureDate, -1), shift: "Night" };
    return { logDate: captureDate, shift: "Day" };
  }, [isEdit, existing, captureTime, captureDate, shiftTimes]);

  const logDate = derived.logDate;
  const shift = derived.shift;

  const visibleAssets = useMemo(() => {
    const q = assetQuery.trim().toLowerCase();
    return assets.filter((a) => {
      if (fleetFilter && a.fleet !== fleetFilter) return false;
      if (!q) return true;
      return `${a.asset_id} ${a.asset_name || ""} ${a.fleet || ""}`.toLowerCase().includes(q);
    });
  }, [assets, fleetFilter, assetQuery]);

  const fleetOptions = useMemo(
    () => [...new Set(assets.map((a) => a.fleet).filter(Boolean))].sort(),
    [assets]
  );

  // If filtering hides the selected machine, move the selection to
  // something visible so the form can't save a machine you can't see.
  useEffect(() => {
    if (isEdit) return;
    if (visibleAssets.length && !visibleAssets.some((a) => a.asset_id === assetId)) {
      setAssetId(visibleAssets[0].asset_id);
    }
  }, [visibleAssets, assetId, isEdit]);

  const openingHours = isEdit && existing?.opening_hours != null
    ? existing.opening_hours
    : findExpectedOpeningHours(dailyHours, assetId, logDate, shift);

  const hoursRun = openingHours != null && closingHours !== ""
    ? Number(closingHours) - Number(openingHours)
    : null;
  const exceedsShiftLimit = hoursRun != null && hoursRun > 12;

  useEffect(() => {
    if (existing?.id) {
      logActivity("Daily Hours", existing.id, "viewed", `Opened ${existing.log_date || ""} ${existing.shift || ""} shift entry`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
      const assetName = assets.find((a) => a.asset_id === assetId)?.asset_name || assetId;
      const summary = `${assetName} - ${shift} shift, ${logDate}, closing hours ${closingHours}${status !== "Operating" ? ` (${status})` : ""}`;
      if (isEdit) {
        const { error: dbError } = await supabase.from("daily_hours").update(payload).eq("id", existing.id);
        if (dbError) throw dbError;
        logActivity("Daily Hours", existing.id, "updated", summary, assetId);
      } else {
        const { data, error: dbError } = await supabase.from("daily_hours").insert(payload).select().single();
        if (dbError) throw dbError;
        logActivity("Daily Hours", data?.id, "created", summary, assetId);
      }
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

        {!isEdit && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 8 }}>
            <div>
              <label style={labelStyle}>Filter by fleet</label>
              <select value={fleetFilter} onChange={(e) => setFleetFilter(e.target.value)} style={fieldStyle}>
                <option value="">All fleets</option>
                {fleetOptions.map((f) => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
            <div>
              <label style={labelStyle}>Search equipment</label>
              <input value={assetQuery} onChange={(e) => setAssetQuery(e.target.value)} placeholder="Type a number or name" style={fieldStyle} />
            </div>
          </div>
        )}

        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required disabled={isEdit} style={{ ...fieldStyle, ...(isEdit ? { background: "#F2F1EA", color: "#4B5659" } : {}) }}>
            {(isEdit ? assets : visibleAssets).map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
          {!isEdit && visibleAssets.length === 0 && (
            <p style={{ fontSize: 12, color: "#B85450", margin: "4px 0 0" }}>No equipment matches that fleet or search.</p>
          )}
        </div>

        {isEdit ? (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
            <div>
              <label style={labelStyle}>Date</label>
              <input type="date" value={logDate} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
            </div>
            <div>
              <label style={labelStyle}>Shift</label>
              <input type="text" value={shift} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
            </div>
          </div>
        ) : (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 8 }}>
              <div>
                <label style={labelStyle}>Reading date</label>
                <input type="date" value={captureDate} onChange={(e) => setCaptureDate(e.target.value)} required style={fieldStyle} />
              </div>
              <div>
                <label style={labelStyle}>Reading time</label>
                <select value={captureTime} onChange={(e) => setCaptureTime(e.target.value)} style={fieldStyle}>
                  <option value={shiftTimes.day}>{shiftTimes.day}</option>
                  <option value={shiftTimes.night}>{shiftTimes.night}</option>
                </select>
              </div>
            </div>
            <div style={{ marginBottom: 12, padding: "8px 10px", background: "#E2EFE9", border: "1px solid #B7D89A", borderRadius: 8, fontSize: 12.5, color: "#2C5646" }}>
              {captureTime} closes the <strong>{shift} shift</strong> of {shortDateLabel(logDate)}
            </div>
          </>
        )}

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

function ScheduledHoursForm({ assets, fleets, existing, onClose, onSaved }) {
  const [fleetFilter, setFleetFilter] = useState(existing?.fleet || fleets[0] || "");
  const assetsInFleet = useMemo(() => assets.filter((a) => a.fleet === fleetFilter), [assets, fleetFilter]);
  const [assetId, setAssetId] = useState(existing?.asset_id || assetsInFleet[0]?.asset_id || "");
  const [scheduledHours, setScheduledHours] = useState(existing?.scheduled_hours ?? "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!assetsInFleet.some((a) => a.asset_id === assetId)) setAssetId(assetsInFleet[0]?.asset_id || "");
  }, [fleetFilter]); // eslint-disable-line react-hooks/exhaustive-deps

  const monthStart = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}-01`;

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    if (!assetId) { setError("Choose a machine."); return; }
    if (scheduledHours === "") { setError("Enter a scheduled hours value."); return; }
    setSaving(true);
    try {
      const { error: dbError } = await supabase.from("asset_scheduled_hours")
        .upsert({ asset_id: assetId, month: monthStart, scheduled_hours: Number(scheduledHours) }, { onConflict: "asset_id,month" });
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
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Set Monthly Scheduled Hours</h3>
        <p style={{ fontSize: 12, color: "#859195", margin: "0 0 16px" }}>Each machine is scheduled individually - a fleet's total is the sum of its machines.</p>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Fleet</label>
          <select value={fleetFilter} onChange={(e) => setFleetFilter(e.target.value)} style={fieldStyle}>
            {fleets.map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Machine</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} style={fieldStyle}>
            {assetsInFleet.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
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

function DailyHoursPage({ assets, dailyHours, userEmail, selectedSiteId, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [showBudgetForm, setShowBudgetForm] = useState(false);
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [budgetLoading, setBudgetLoading] = useState(true);
  const [deleting, setDeleting] = useState(null);
  const [drillFleet, setDrillFleet] = useState(null);
  const [templateMonth, setTemplateMonth] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  });
  const [bulkImporting, setBulkImporting] = useState(false);
  const [bulkMessage, setBulkMessage] = useState(null);
  const bulkFileInputRef = React.useRef(null);

  const fleets = useMemo(() => [...new Set(assets.map((a) => a.fleet))], [assets]);

  const [assetScheduled, setAssetScheduled] = useState({}); // asset_id -> scheduled_hours, this month

  const loadBudget = React.useCallback(async () => {
    setBudgetLoading(true);
    const monthStart = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}-01`;
    const { data, error } = await supabase.from("asset_scheduled_hours").select("asset_id, scheduled_hours").eq("month", monthStart);
    if (!error) {
      setAssetScheduled(Object.fromEntries((data || []).map((r) => [r.asset_id, Number(r.scheduled_hours)])));
    }
    setBudgetLoading(false);
  }, []);

  useEffect(() => { loadBudget(); }, [loadBudget]);

  // Per-equipment hours for the current month, for the drill-down under
  // each fleet's total - so machines that are barely running (versus
  // ones running flat out) are visible for budgeting, not just hidden
  // inside a single fleet-wide average.
  const assetMonthHours = useMemo(() => {
    const now = new Date();
    const y = now.getFullYear(), m = now.getMonth();
    const totals = {};
    dailyHours.forEach((r) => {
      if (!r.log_date) return;
      const d = new Date(r.log_date + "T00:00:00");
      if (d.getFullYear() !== y || d.getMonth() !== m) return;
      totals[r.asset_id] = (totals[r.asset_id] || 0) + (Number(r.hours_run) || 0);
    });
    return totals;
  }, [dailyHours]);

  // Fleet-level budget rows, built by summing each fleet's own machines -
  // there's no separate "fleet total" set anywhere; it's always derived
  // from what's been scheduled per machine.
  const budget = useMemo(() => {
    return fleets.map((fleet) => {
      const fleetAssets = assets.filter((a) => a.fleet === fleet);
      const actual_hours = fleetAssets.reduce((sum, a) => sum + (assetMonthHours[a.asset_id] || 0), 0);
      const anyScheduled = fleetAssets.some((a) => assetScheduled[a.asset_id] != null);
      const scheduled_hours = anyScheduled ? fleetAssets.reduce((sum, a) => sum + (assetScheduled[a.asset_id] || 0), 0) : null;
      const status = scheduled_hours == null ? null : (actual_hours > scheduled_hours ? "Over Scheduled Hours" : "OK");
      return { fleet, actual_hours, scheduled_hours, status };
    });
  }, [fleets, assets, assetMonthHours, assetScheduled]);

  const drillFleetAssets = useMemo(() => {
    if (!drillFleet) return [];
    return assets
      .filter((a) => a.fleet === drillFleet)
      .map((a) => ({
        asset_id: a.asset_id,
        asset_name: a.asset_name,
        hours: assetMonthHours[a.asset_id] || 0,
        scheduled: assetScheduled[a.asset_id] ?? null,
      }))
      .sort((a, b) => b.hours - a.hours);
  }, [drillFleet, assets, assetMonthHours, assetScheduled]);

  const prefs = useTablePrefs("daily_hours", DAILY_HOURS_COLUMNS);
  const columns = prefs.columns;

  const fleetByAsset = useMemo(() => new Map(assets.map((a) => [a.asset_id, a.fleet])), [assets]);

  const RECENT_READINGS_PER_MACHINE = 2;
  const [groupByMachine, setGroupByMachine] = useState(true);

  // Rows still get filtered flat here; the grouping into machines happens
  // below, because the "latest N readings" window is per machine, not
  // across the whole fleet. A machine that wasn't logged yesterday should
  // still show its own last two readings rather than vanishing.
  const filtered = useMemo(() => {
    let rows = dailyHours;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => fleetByAsset.get(r.asset_id) === selectedFleet);
    if (dateFrom) rows = rows.filter((r) => (r.log_date || "") >= dateFrom);
    if (dateTo) rows = rows.filter((r) => (r.log_date || "") <= dateTo);
    return rows.map((r) => ({ ...r, fleet: fleetByAsset.get(r.asset_id) || "-" }));
  }, [dailyHours, selectedFleet, selectedAsset, dateFrom, dateTo, fleetByAsset]);

  const grouped = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "log_date", hoursKey: "hours_run",
    sortKey: prefs.sortKey, sortDir: prefs.sortDir, sortRows: prefs.sortRows,
    limitPerMachine: (!dateFrom && !dateTo && groupByMachine) ? RECENT_READINGS_PER_MACHINE : null,
  }), [filtered, assets, groupByMachine, dateFrom, dateTo, prefs.sortKey, prefs.sortDir, prefs.sortRows]); // eslint-disable-line react-hooks/exhaustive-deps

  // Flat list of what's on screen, for Export to Excel.
  const visibleRows = useMemo(() => grouped.flatMap((g) => g.rows), [grouped]);

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
    await deleteWithReason("daily_hours", deleting.id, "id", reason, userEmail, deleting.asset_id);
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
            <React.Fragment key={row.fleet}>
              <div
                onClick={() => setDrillFleet(drillFleet === row.fleet ? null : row.fleet)}
                style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", borderTop: i > 0 ? "1px solid #EFEEE7" : "none", cursor: "pointer", background: drillFleet === row.fleet ? "#F7F8F6" : "transparent" }}
              >
                <ChevronRight size={14} style={{ transform: drillFleet === row.fleet ? "rotate(90deg)" : "none", transition: "transform 0.15s", color: "#859195", flexShrink: 0 }} />
                <span style={{ flex: 1, fontSize: 13.5, fontWeight: 600 }}>{row.fleet}</span>
                <span style={{ fontSize: 13 }}>{Number(row.actual_hours).toFixed(1)} / {row.scheduled_hours != null ? Number(row.scheduled_hours).toFixed(0) : "-"} hrs</span>
                <Badge value={row.status} />
              </div>
              {drillFleet === row.fleet && (
                <div style={{ background: "#FBFBF9", borderTop: "1px solid #EFEEE7" }}>
                  {drillFleetAssets.length === 0 ? (
                    <p style={{ padding: "10px 14px 10px 40px", fontSize: 12.5, color: "#859195", margin: 0 }}>No equipment in this fleet.</p>
                  ) : (
                    drillFleetAssets.map((a) => (
                      <div key={a.asset_id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "7px 14px 7px 40px", fontSize: 12.5 }}>
                        <span style={{ flex: 1, color: "#4B5659" }}>{a.asset_id} - {a.asset_name}</span>
                        <span style={{ fontWeight: 600, color: "#183642" }}>{a.hours.toFixed(1)} / {a.scheduled != null ? a.scheduled.toFixed(0) : "-"} hrs</span>
                      </div>
                    ))
                  )}
                </div>
              )}
            </React.Fragment>
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

      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexWrap: "wrap", fontSize: 13, color: "#4B5659" }}>
        <span style={{ fontWeight: 600 }}>Dates</span>
        <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)}
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        <span>to</span>
        <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)}
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        {(dateFrom || dateTo) ? (
          <button onClick={() => { setDateFrom(""); setDateTo(""); }}
            style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer" }}>
            Back to latest {RECENT_READINGS_PER_MACHINE} readings
          </button>
        ) : (
          <span style={{ color: "#859195" }}>Showing each machine's latest {RECENT_READINGS_PER_MACHINE} readings — pick a date range to see more.</span>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
          <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
            <Download size={14} /> Export to Excel
          </button>
          <ColumnsButton prefs={prefs} />
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Log Daily Hours
          </button>
        </div>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <SmartTableHead prefs={prefs} />
          <MachineTableBody
            groups={grouped} columns={columns}
            onRowClick={(row) => { setEditing(row); setShowForm(true); }}
            onDelete={(row) => setDeleting(row)}
            totalLabel="hrs" countNoun="reading" countNounPlural="readings"
            emptyMessage={dailyHours.length === 0 ? "No hours logged yet." : "No entries match the current filters."}
          />
        </table>
      </div>

      {showForm && (
        <DailyHoursForm assets={assets} dailyHours={dailyHours} existing={editing} selectedSiteId={selectedSiteId}
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
        <ScheduledHoursForm assets={assets} fleets={fleets} onClose={() => setShowBudgetForm(false)} onSaved={handleBudgetSaved} />
      )}
    </div>
  );
}


function BreakdownForm({ assets, existing, activatingWorkOrder, onClose, onSaved, userEmail, myFullName, workOrders, parts, componentCodes, onRefresh }) {
  const [assetId, setAssetId] = useState(existing?.asset_id || activatingWorkOrder?.asset_id || assets[0]?.asset_id || "");
  // Planned vs Breakdown is now recorded on the event itself rather than
  // inferred from which table the row lives in.
  const [eventType, setEventType] = useState(existing?.event_type || "Breakdown");
  const [causeCode, setCauseCode] = useState(existing?.cause_code || CAUSE_CODES[0]);
  // "Other" on its own tells you nothing six months later, so it has to
  // be spelled out before the event can be saved.
  const [causeDetail, setCauseDetail] = useState(existing?.cause_detail || "");
  // What was actually DONE. Kept separate from cause: "engine overheating"
  // and "engine replaced" are different facts and both matter to a claim.
  const [actionTaken, setActionTaken] = useState(existing?.action_taken || "");
  const [severity, setSeverity] = useState(existing?.severity || "Medium");
  // Blank by default and blank on historical events - never guessed.
  const [responsibility, setResponsibility] = useState(existing?.responsibility || "");
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

  // Records that this event was opened, not just changed - so the audit
  // trail shows who looked at it, not only who edited it. Only fires for
  // an event that already existed when this form opened, once per open.
  useEffect(() => {
    if (existing?.id) {
      logActivity("Events", existing.id, "viewed", `Opened event ${existing.description ? `"${existing.description}"` : ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
    if (causeCode === "Other" && !causeDetail.trim()) {
      setError("Please specify the cause when Cause is set to Other.");
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
      event_type: eventType,
      cause_code: causeCode,
      cause_detail: causeCode === "Other" ? causeDetail.trim() : null,
      action_taken: actionTaken || null,
      responsibility: responsibility || null,
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
        logActivity("Events", currentId, status === "Closed" ? "closed" : "updated", eventSummary, assetId);
      } else {
        const { data, error: dbError } = await supabase.from("breakdown_log").insert(payload).select().single();
        if (dbError) throw dbError;
        // Stay open on the just-created record instead of closing - this
        // is what lets a Work Order (and parts against it) be added right
        // away, in the same session, instead of having to reopen it.
        setSavedRecord(data);
        logActivity("Events", data.id, "created", eventSummary, assetId);

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
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 980, maxWidth: "96vw", maxHeight: "90vh", overflowY: "auto" }}>
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

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 6 }}>
          <div>
            <label style={labelStyle}>Downtime</label>
            <DateTimeField value={downtimeStart} onChange={setDowntimeStart} max={nowForInput()} required />
          </div>
          <div>
            <label style={labelStyle}>Uptime {status !== "Closed" && <span style={{ fontWeight: 400, color: "#859195" }}>(Closed first)</span>}</label>
            <DateTimeField value={downtimeEnd} onChange={setDowntimeEnd} max={nowForInput()} disabled={status !== "Closed"} />
          </div>
          <div>
            <label style={labelStyle}>Expected Up Time {status === "Closed" && <span style={{ fontWeight: 400, color: "#859195" }}>(not needed)</span>}</label>
            <DateTimeField value={expectedUpTime} onChange={setExpectedUpTime} disabled={status === "Closed"} />
          </div>
        </div>
        {/* Collapsed by default - the guidance is long and pushed the form
            past the bottom of the screen every time it opened. */}
        <details style={{ marginBottom: 12 }}>
          <summary style={{ fontSize: 11.5, color: "#859195", cursor: "pointer" }}>How these three dates work</summary>
          <p style={{ fontSize: 11.5, color: "#859195", margin: "6px 0 0" }}>
            Downtime defaults to right now, but you can set it earlier if you're logging something that happened previously (e.g. catching up after time away from the system) - future dates and times aren't allowed. Uptime only opens up once Status is Closed - an event that's still active doesn't have an end yet, and having both set at once would throw off MTBF, MTTR, Availability and Utilisation. Expected Up Time is your best estimate of when the machine will be back - unlike the other two, future dates are fine here, since that's the whole point of it.
          </p>
        </details>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Event</label>
            <select value={eventType} onChange={(e) => setEventType(e.target.value)} style={fieldStyle}>
              <option value="Breakdown">Breakdown</option>
              <option value="Planned">Planned</option>
            </select>
            <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
              Drives the Planned / Unplanned split on the KPI report.
            </p>
          </div>
          <div>
            <label style={labelStyle}>Cause</label>
            <select value={causeCode} onChange={(e) => setCauseCode(e.target.value)} style={fieldStyle}>
              {CAUSE_CODES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Responsibility</label>
            <select value={responsibility} onChange={(e) => setResponsibility(e.target.value)} style={fieldStyle}>
              <option value="">- not set -</option>
              {RESPONSIBILITY_CODES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
              Whose account the downtime goes to. Plant and Stores count as engineering.
            </p>
          </div>
          <div>
            <label style={labelStyle}>Severity</label>
            <select value={severity} onChange={(e) => setSeverity(e.target.value)} style={fieldStyle}>
              {SEVERITIES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        </div>

        {causeCode === "Other" && (
          <div style={{ marginBottom: 12 }}>
            <label style={labelStyle}>Specify the cause *</label>
            <input
              type="text"
              value={causeDetail}
              onChange={(e) => setCauseDetail(e.target.value)}
              required
              maxLength={120}
              placeholder="e.g. Operator drove over a berm and damaged the sump guard"
              style={fieldStyle}
            />
            <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
              Required when Cause is Other. This shows in the Reason column and on reports.
            </p>
          </div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Component Code</label>
            <select
              value={componentCode}
              onChange={(e) => { setComponentCode(e.target.value); if (e.target.value) setComponentAffected(e.target.value); }}
              style={fieldStyle}
            >
              <option value="">- Select a component -</option>
              {componentCodes.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}
            </select>
            <p style={{ fontSize: 11, color: "#859195", margin: "4px 0 0" }}>Fills in Component / System Affected - still editable after.</p>
          </div>
          <div>
            <label style={labelStyle}>Component / System Affected</label>
            <input type="text" value={componentAffected} onChange={(e) => setComponentAffected(e.target.value)} placeholder="e.g. Starter Motor" style={fieldStyle} />
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Action Taken</label>
            <select value={actionTaken} onChange={(e) => setActionTaken(e.target.value)} style={fieldStyle}>
              <option value="">- not set -</option>
              {ACTION_TAKEN_OPTIONS.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
            {actionTaken === "Component replaced" && (
              <p style={{ fontSize: 11.5, color: "#2C5646", margin: "4px 0 0" }}>
                {hasSavedRecord
                  ? "Fill in the Component Change Out panel below."
                  : "Save the event first - the change out panel opens once it exists."}
              </p>
            )}
          </div>
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
          <>
            {actionTaken === "Component replaced" && (
            <EventChangeoutPanel event={savedRecord} componentCodes={componentCodes} myFullName={myFullName} />
          )}
          <EventPulledPartsPanel event={savedRecord} parts={parts} onRefresh={onRefresh} />
            <EventWorkOrdersPanel event={savedRecord} assets={assets} parts={parts} onRefresh={onRefresh} />
          </>
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
  const [groupByMachine, setGroupByMachine] = useState(true);
  const [sortKey, setSortKey] = useState("downtime_start");
  const [sortDir, setSortDir] = useState("desc");

  const columns = [
    ["asset_id", "Equipment #"], ["event_type", "Event"], ["component_affected", "Component"], ["cause_code", "Reason"],
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
      event_type: "Planned",
      component_affected: w.component || null,
      cause_code: null,
      description: w.problem_scope || "",
      wo_reference: w.wo_no,
      downtime_start: w.planned_start,
      downtime_end: null,
      downtime_hours: null,
      repair_status: "Planned",
    })), [workOrders]);

  // Anything in breakdown_log is by definition unplanned; the
  // placeholders above are the planned side. Event says WHICH KIND of
  // event it was - the component that failed has its own column.
  const allRows = useMemo(
    () => [
      ...breakdowns.map((b) => ({
        ...b,
        event_type: b.event_type || "Breakdown",
        cause_code: b.cause_code === "Other" && b.cause_detail ? `Other - ${b.cause_detail}` : b.cause_code,
      })),
      ...scheduledRows,
    ],
    [breakdowns, scheduledRows]
  );

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

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "downtime_start", hoursKey: "downtime_hours",
    sortKey: null, sortDir: null, sortRows: (r) => r,
  }), [filtered, assets, groupByMachine]);

  const handleSaved = () => {
    setShowForm(false);
    setEditing(null);
    setActivating(null);
    onRefresh();
  };

  const handleDelete = async (reason) => {
    await deleteWithReason("breakdown_log", deleting.id, "id", reason, userEmail, deleting.asset_id);
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
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
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
            {flattenGroups(groups).map((item, i) => {
              if (item.kind === "banner") {
                return <MachineBannerRow key={`b-${item.group.key}`} group={item.group} colSpan={columns.length + 1} first={i === 0} countNoun="event" countNounPlural="events" />;
              }
              const row = item.row;
              const isRepeat = (row.repeat_count ?? 0) >= 2;
              const openRow = () => {
                if (row.isScheduledPlaceholder) { setActivating(row.sourceWorkOrder); setShowForm(true); }
                else { setEditing(row); setShowForm(true); }
              };
              return (
                <tr
                  key={row.id ?? i}
                  style={{ borderBottom: "1px solid #EFEEE7", background: row.isScheduledPlaceholder ? "#FFF8EB" : isRepeat ? "#F5E9D8" : "transparent" }}
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

// Generic "type a reason, then confirm" modal - same shape as
// DeleteConfirmModal, reused for Reject/Merge actions elsewhere so
// every place that needs a reason looks and behaves the same way.
function ReasonPromptModal({ title, message, confirmLabel, confirmColor, reasonRequired = true, onCancel, onConfirm }) {
  const [reason, setReason] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const handleConfirm = async () => {
    if (reasonRequired && !reason.trim()) { setError("A reason is required."); return; }
    setError("");
    setBusy(true);
    try {
      await onConfirm(reason.trim());
    } catch (err) {
      setError(err.message || String(err));
      setBusy(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 55, padding: 16 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 400, maxWidth: "100%" }}>
        <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>{title}</p>
        {message && <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 14px" }}>{message}</p>}
        <textarea
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          rows={2}
          autoFocus
          placeholder={reasonRequired ? "Reason (required)" : "Notes (optional)"}
          style={{ width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", resize: "vertical", fontFamily: "inherit" }}
        />
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "8px 0 0" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
          <button onClick={onCancel} disabled={busy} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button onClick={handleConfirm} disabled={busy} style={{ background: confirmColor || NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: busy ? "default" : "pointer", opacity: busy ? 0.7 : 1 }}>
            {busy ? "Saving…" : confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

const WORK_REQUEST_SOURCES = ["Operator Report", "Pre-Start Inspection", "Breakdown", "Maintenance Inspection", "Planner", "Supervisor", "Condition Monitoring", "Scheduled Maintenance", "Reliability Finding"];

function WorkRequestForm({ assets, onClose, onSaved, userEmail, myFullName }) {
  const [assetId, setAssetId] = useState(assets[0]?.asset_id || "");
  const [description, setDescription] = useState("");
  const [priority, setPriority] = useState("Medium");
  const [source, setSource] = useState(WORK_REQUEST_SOURCES[0]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const reportedBy = myFullName || userEmail || "";

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
    if (!description.trim()) { setError("Describe the problem."); return; }
    setSaving(true);
    const payload = {
      asset_id: assetId, description: description.trim(), priority, source,
      reported_by: reportedBy || null, status: "Open",
    };
    try {
      const assetName = assets.find((a) => a.asset_id === assetId)?.asset_name || assetId;
      const { data, error: dbError } = await supabase.from("work_requests").insert(payload).select().single();
      if (dbError) throw dbError;
      logActivity("Work Request", data?.id, "created", `${assetName} - ${description.trim()}`, assetId);
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
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 440, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Report a Problem</h3>
        <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 16px" }}>Quick report - a planner or supervisor will review it and either create a Work Order or let you know why not.</p>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required style={fieldStyle}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>What's wrong?</label>
          <textarea value={description} onChange={(e) => setDescription(e.target.value)} required rows={3} autoFocus style={{ ...fieldStyle, resize: "vertical" }} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Priority</label>
            <select value={priority} onChange={(e) => setPriority(e.target.value)} style={fieldStyle}>
              {SEVERITIES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Where did this come from?</label>
            <select value={source} onChange={(e) => setSource(e.target.value)} style={fieldStyle}>
              {WORK_REQUEST_SOURCES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <div style={{ marginBottom: 18 }}>
          <label style={labelStyle}>Reported By</label>
          <input type="text" value={reportedBy} disabled style={{ ...fieldStyle, background: "#F2F1EA", color: "#4B5659" }} />
        </div>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Submitting…" : "Submit"}
          </button>
        </div>
      </form>
    </div>
  );
}

function WorkRequestsPage({ assets, workRequests, workOrders, userEmail, myFullName, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [statusFilter, setStatusFilter] = useState("Open");
  const [rejecting, setRejecting] = useState(null);
  const [merging, setMerging] = useState(null);
  const [busyId, setBusyId] = useState(null);
  const [error, setError] = useState("");

  const [groupByMachine, setGroupByMachine] = useState(true);

  const filtered = useMemo(() => {
    let rows = workRequests;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (statusFilter !== "All") rows = rows.filter((r) => r.status === statusFilter);
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    return [...rows].sort((a, b) => (b.reported_at || "").localeCompare(a.reported_at || ""));
  }, [workRequests, assets, query, selectedFleet, selectedAsset, statusFilter]);

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "reported_at", hoursKey: null,
    sortKey: null, sortDir: null, sortRows: (r) => r,
  }), [filtered, assets, groupByMachine]);

  const formatDateTime = (iso) => iso
    ? new Date(iso).toLocaleString("en-ZA", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", hour12: false })
    : "-";

  const handleApprove = async (req) => {
    setBusyId(req.id);
    setError("");
    try {
      const assetName = assets.find((a) => a.asset_id === req.asset_id)?.asset_name || req.asset_id;
      const { data: wo, error: woErr } = await supabase.from("work_orders").insert({
        asset_id: req.asset_id, work_type: "Corrective", priority: req.priority,
        problem_scope: req.description, status: "Open", request_date: (req.reported_at || "").slice(0, 10),
        work_request_id: req.id,
      }).select().single();
      if (woErr) throw woErr;
      const { error: reqErr } = await supabase.from("work_requests").update({
        status: "Converted", converted_work_order_id: wo.id, reviewed_by: myFullName || userEmail || null, reviewed_at: new Date().toISOString(),
      }).eq("id", req.id);
      if (reqErr) throw reqErr;
      logActivity("Work Request", req.id, "updated", `Approved and converted to Work Order ${wo.wo_no || wo.id} - ${assetName}`, req.asset_id);
      onRefresh();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusyId(null);
    }
  };

  const handleReject = async (reason) => {
    const req = rejecting;
    const { error: dbError } = await supabase.from("work_requests").update({
      status: "Rejected", rejection_reason: reason, reviewed_by: myFullName || userEmail || null, reviewed_at: new Date().toISOString(),
    }).eq("id", req.id);
    if (dbError) throw dbError;
    logActivity("Work Request", req.id, "updated", `Rejected: ${reason}`, req.asset_id);
    setRejecting(null);
    onRefresh();
  };

  const openWorkOrdersForAsset = (assetId) => workOrders.filter((w) => w.asset_id === assetId && w.status !== "Closed");

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", flex: 1 }}>
          <div style={{ position: "relative", maxWidth: 260, flex: 1 }}>
            <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search work requests" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", boxSizing: "border-box" }} />
          </div>
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
            {["Open", "Converted", "Rejected", "Merged", "All"].map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
        <button onClick={() => setShowForm(true)} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
          + Report a Problem
          </button>
        </div>
      </div>
      {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "10px 0 0" }}>{error}</p>}
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10, marginTop: 12, WebkitOverflowScrolling: "touch" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Reported", "Equipment #", "Description", "Priority", "Source", "Reported By", "Status", ""].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {flattenGroups(groups).map((item, i) => item.kind === "banner" ? (
              <MachineBannerRow key={`b-${item.group.key}`} group={item.group} colSpan={8} first={i === 0} countNoun="request" countNounPlural="requests" />
            ) : (
              (({ row: r }) => (
              <tr key={r.id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{formatDateTime(r.reported_at)}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.asset_id}</td>
                <td style={{ padding: "9px 12px", maxWidth: 260 }}>{r.description}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={r.priority} /></td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.source || "-"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.reported_by || "-"}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={r.status} /></td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {r.status === "Open" && (
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => handleApprove(r)} disabled={busyId === r.id} title="Approve & Create Work Order" style={{ background: "#2C5646", border: "none", color: "#fff", padding: "5px 9px", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: busyId === r.id ? "default" : "pointer" }}>
                        {busyId === r.id ? "…" : "Approve"}
                      </button>
                      {openWorkOrdersForAsset(r.asset_id).length > 0 && (
                        <button onClick={() => setMerging(r)} title="Merge into an existing Work Order" style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, padding: "5px 9px", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                          Merge
                        </button>
                      )}
                      <button onClick={() => setRejecting(r)} title="Reject" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4 }}>
                        <X size={16} />
                      </button>
                    </div>
                  )}
                </td>
              </tr>
              ))(item)
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={8} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {workRequests.length === 0 ? "No work requests yet." : "Nothing matches your filters."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      {showForm && (
        <WorkRequestForm assets={assets} userEmail={userEmail} myFullName={myFullName}
          onClose={() => setShowForm(false)}
          onSaved={() => { setShowForm(false); onRefresh(); }} />
      )}
      {rejecting && (
        <ReasonPromptModal
          title={`Reject request for ${rejecting.asset_id}?`}
          message="This closes the request without creating a Work Order. The reason is kept for the record."
          confirmLabel="Reject" confirmColor="#B85450"
          onCancel={() => setRejecting(null)}
          onConfirm={handleReject}
        />
      )}
      {merging && (
        <MergeWorkRequestModal
          request={merging}
          candidateWorkOrders={openWorkOrdersForAsset(merging.asset_id)}
          userEmail={userEmail} myFullName={myFullName}
          onCancel={() => setMerging(null)}
          onMerged={() => { setMerging(null); onRefresh(); }}
        />
      )}
    </div>
  );
}

function MergeWorkRequestModal({ request, candidateWorkOrders, userEmail, myFullName, onCancel, onMerged }) {
  const [woId, setWoId] = useState(candidateWorkOrders[0]?.id || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const handleConfirm = async () => {
    if (!woId) return;
    setSaving(true);
    setError("");
    try {
      const { error: dbError } = await supabase.from("work_requests").update({
        status: "Merged", merged_into_work_order_id: Number(woId), reviewed_by: myFullName || userEmail || null, reviewed_at: new Date().toISOString(),
      }).eq("id", request.id);
      if (dbError) throw dbError;
      const wo = candidateWorkOrders.find((w) => String(w.id) === String(woId));
      logActivity("Work Request", request.id, "updated", `Merged into existing Work Order ${wo?.wo_no || woId}`, request.asset_id);
      onMerged();
    } catch (err) {
      setError(err.message || String(err));
      setSaving(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 55, padding: 16 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 420, maxWidth: "100%" }}>
        <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: "0 0 8px" }}>Merge into an existing Work Order</p>
        <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 14px" }}>
          Use this when the problem is already being handled by another job on {request.asset_id}, so this doesn't create a duplicate.
        </p>
        <select value={woId} onChange={(e) => setWoId(e.target.value)} style={{ width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box" }}>
          {candidateWorkOrders.map((w) => <option key={w.id} value={w.id}>{w.wo_no || `WO-${w.id}`} - {w.problem_scope || w.work_type} ({w.status})</option>)}
        </select>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "8px 0 0" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
          <button onClick={onCancel} disabled={saving} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button onClick={handleConfirm} disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Merging…" : "Merge"}
          </button>
        </div>
      </div>
    </div>
  );
}

function DefectForm({ assets, onClose, onSaved, userEmail, myFullName }) {
  const [assetId, setAssetId] = useState(assets[0]?.asset_id || "");
  const [description, setDescription] = useState("");
  const [riskRating, setRiskRating] = useState("Medium");
  const [responsiblePerson, setResponsiblePerson] = useState("");
  const [targetCompletion, setTargetCompletion] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const reportedBy = myFullName || userEmail || "";

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
    if (!description.trim()) { setError("Describe the defect."); return; }
    setSaving(true);
    const payload = {
      asset_id: assetId, description: description.trim(), risk_rating: riskRating,
      is_critical: riskRating === "Critical" || riskRating === "High",
      reported_by: reportedBy || null, status: "Open",
      responsible_person: responsiblePerson || null, target_completion: targetCompletion || null,
    };
    try {
      const assetName = assets.find((a) => a.asset_id === assetId)?.asset_name || assetId;
      const { data, error: dbError } = await supabase.from("defects").insert(payload).select().single();
      if (dbError) throw dbError;
      logActivity("Defect", data?.id, "created", `${assetName} - ${description.trim()}`, assetId);
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
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 440, maxWidth: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>Report Defect</h3>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Equipment</label>
          <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required style={fieldStyle}>
            {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Defect Description</label>
          <textarea value={description} onChange={(e) => setDescription(e.target.value)} required rows={3} autoFocus style={{ ...fieldStyle, resize: "vertical" }} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Risk Rating</label>
          <select value={riskRating} onChange={(e) => setRiskRating(e.target.value)} style={fieldStyle}>
            {SEVERITIES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          {(riskRating === "High" || riskRating === "Critical") && (
            <p style={{ fontSize: 12, color: "#7A3330", margin: "6px 0 0", fontWeight: 600 }}>
              ⚠ This will be flagged as critical - equipment should not return to service until it's cleared.
            </p>
          )}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 18 }}>
          <div>
            <label style={labelStyle}>Responsible Person (optional)</label>
            <input type="text" value={responsiblePerson} onChange={(e) => setResponsiblePerson(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Target Completion (optional)</label>
            <input type="date" value={targetCompletion} onChange={(e) => setTargetCompletion(e.target.value)} style={fieldStyle} />
          </div>
        </div>
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
            {saving ? "Saving…" : "Report Defect"}
          </button>
        </div>
      </form>
    </div>
  );
}

function DefectsPage({ assets, defects, userEmail, myFullName, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [statusFilter, setStatusFilter] = useState("Open");
  const [busyId, setBusyId] = useState(null);
  const [error, setError] = useState("");

  const [groupByMachine, setGroupByMachine] = useState(true);

  const filtered = useMemo(() => {
    let rows = defects;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (statusFilter !== "All") rows = rows.filter((r) => r.status === statusFilter);
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    return [...rows].sort((a, b) => (b.reported_date || "").localeCompare(a.reported_date || ""));
  }, [defects, assets, query, selectedFleet, selectedAsset, statusFilter]);

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "reported_date", hoursKey: null,
    sortKey: null, sortDir: null, sortRows: (r) => r,
  }), [filtered, assets, groupByMachine]);

  const handleCreateWorkOrder = async (defect) => {
    setBusyId(defect.id);
    setError("");
    try {
      const assetName = assets.find((a) => a.asset_id === defect.asset_id)?.asset_name || defect.asset_id;
      const { data: wo, error: woErr } = await supabase.from("work_orders").insert({
        asset_id: defect.asset_id, work_type: "Corrective", priority: defect.risk_rating,
        problem_scope: defect.description, status: "Open", request_date: defect.reported_date,
        defect_id: defect.id,
      }).select().single();
      if (woErr) throw woErr;
      const { error: defErr } = await supabase.from("defects").update({
        status: "Work Order Created", linked_work_order_id: wo.id,
      }).eq("id", defect.id);
      if (defErr) throw defErr;
      logActivity("Defect", defect.id, "updated", `Work Order ${wo.wo_no || wo.id} created for ${assetName}`, defect.asset_id);
      onRefresh();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusyId(null);
    }
  };

  const handleVerify = async (defect) => {
    setBusyId(defect.id);
    setError("");
    try {
      const { error: dbError } = await supabase.from("defects").update({
        status: "Verified", verified_by: myFullName || userEmail || null, verified_at: new Date().toISOString(),
        actual_completion: defect.actual_completion || todayForInput(),
      }).eq("id", defect.id);
      if (dbError) throw dbError;
      logActivity("Defect", defect.id, "updated", "Verified and closed", defect.asset_id);
      onRefresh();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setBusyId(null);
    }
  };

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", flex: 1 }}>
          <div style={{ position: "relative", maxWidth: 260, flex: 1 }}>
            <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search defects" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none", boxSizing: "border-box" }} />
          </div>
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
            {["Open", "Work Order Created", "Completed", "Verified", "All"].map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
        <button onClick={() => setShowForm(true)} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
          + Report Defect
          </button>
        </div>
      </div>
      {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "10px 0 0" }}>{error}</p>}
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10, marginTop: 12, WebkitOverflowScrolling: "touch" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Date", "Equipment #", "Description", "Risk", "Responsible", "Status", ""].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {flattenGroups(groups).map((item, i) => item.kind === "banner" ? (
              <MachineBannerRow key={`b-${item.group.key}`} group={item.group} colSpan={7} first={i === 0} countNoun="defect" countNounPlural="defects" />
            ) : (
              (({ row: r }) => (
              <tr key={r.id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.reported_date || "-"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.asset_id}</td>
                <td style={{ padding: "9px 12px", maxWidth: 260 }}>
                  {r.description}
                  {r.is_critical && (
                    <div style={{ marginTop: 4, display: "inline-flex", alignItems: "center", gap: 4, background: "#F6E2E0", color: "#7A3330", fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 6 }}>
                      <AlertTriangle size={11} /> DO NOT OPERATE UNTIL CLEARED
                    </div>
                  )}
                </td>
                <td style={{ padding: "9px 12px" }}><Badge value={r.risk_rating} /></td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.responsible_person || "-"}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={r.status} /></td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>
                  {r.status === "Open" && (
                    <button onClick={() => handleCreateWorkOrder(r)} disabled={busyId === r.id} style={{ background: NAVY, border: "none", color: "#fff", padding: "5px 9px", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: busyId === r.id ? "default" : "pointer" }}>
                    {busyId === r.id ? "…" : "Create Work Order"}
                    </button>
                  )}
                  {r.status === "Completed" && (
                    <button onClick={() => handleVerify(r)} disabled={busyId === r.id} style={{ background: "#2C5646", border: "none", color: "#fff", padding: "5px 9px", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: busyId === r.id ? "default" : "pointer" }}>
                      {busyId === r.id ? "…" : "Verify & Close"}
                    </button>
                  )}
                </td>
              </tr>
              ))(item)
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={7} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {defects.length === 0 ? "No defects on record." : "Nothing matches your filters."}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      {showForm && (
        <DefectForm assets={assets} userEmail={userEmail} myFullName={myFullName}
          onClose={() => setShowForm(false)}
          onSaved={() => { setShowForm(false); onRefresh(); }} />
      )}
    </div>
  );
}

function WorkOrderForm({ assets, existing, defaultWorkType, defaultAssetId, eventId, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || defaultAssetId || assets[0]?.asset_id || "");
  const [workType, setWorkType] = useState(existing?.work_type || defaultWorkType || "Corrective");
  // Planned downtime counts toward Downtime Responsibility too, so a
  // Preventive job needs an account the same way an event does.
  const [responsibility, setResponsibility] = useState(existing?.responsibility || "");
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

  // Records that this work order was opened, not just changed - so the
  // audit trail shows who looked at it, not only who edited it.
  useEffect(() => {
    if (existing?.id) {
      logActivity("Work Orders", existing.id, "viewed", `Opened Work Order ${existing.wo_no || ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
      asset_id: assetId, work_type: workType, priority, problem_scope: problemScope, responsibility: responsibility || null,
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
      logActivity("Work Orders", (isEdit ? existing.id : data?.id), status === "Closed" ? "closed" : (isEdit ? "updated" : "created"), woSummary, assetId);
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
          <div>
            <label style={labelStyle}>Responsibility</label>
            <select value={responsibility} onChange={(e) => setResponsibility(e.target.value)} style={fieldStyle}>
              <option value="">- not set -</option>
              {RESPONSIBILITY_CODES.map((c) => <option key={c} value={c}>{c}</option>)}
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

function PartUsedForm({ parts, workOrder, event, onClose, onSaved }) {
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
        work_order_id: workOrder?.id ?? null,
        event_id: event?.id ?? null,
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

      logActivity(
        event ? "Events" : "Work Orders",
        event?.id ?? workOrder?.id ?? null,
        "updated",
        `Pulled ${qtyNum} × ${selectedPart?.part_no || "part"} - ${selectedPart?.description || ""}${workOrder ? ` against Work Order ${workOrder.wo_no}` : " directly from inventory"}`,
        event?.asset_id ?? workOrder?.asset_id ?? null
      );

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
        <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 4px" }}>Pull Part from Inventory</h3>
        <p style={{ fontSize: 12, color: "#859195", margin: "0 0 16px" }}>
          {workOrder ? `Against Work Order ${workOrder.wo_no}` : "Pulled directly against this event - only use this for parts already in stock."}
        </p>

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
            {saving ? "Saving…" : "Pull Part"}
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
    const p = partById[row.part_id];
    logActivity("Work Orders", workOrder.id, "updated", `Confirmed ${row.qty_used} × ${p?.part_no || "part"} pulled against Work Order ${workOrder.wo_no}`, workOrder.asset_id);
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

// Parts pulled straight from Inventory against this event - for parts
// already on hand. Separate from Work Orders (below), which are for
// parts that need to be ordered in first.
// ---------------------------------------------------------------------
// Component Change Out - the drill-down behind Action Taken.
//
// Lives against an event via component_changeout_log.event_id. Written
// in two stages for the same reason the parts and work order panels are:
// the attachment rows need a changeout_id, so the record is saved first
// and the documents attach to it afterwards.
//
// life_pct and premature_failure are derived in the database view rather
// than here, so the form, the reports and anything built later can't
// disagree about what counts as a premature failure.
// ---------------------------------------------------------------------
function EventChangeoutPanel({ event, componentCodes, myFullName }) {
  const [row, setRow] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [docs, setDocs] = useState([]);
  const [docType, setDocType] = useState(CHANGEOUT_DOC_TYPES[0]);
  const [uploading, setUploading] = useState(false);

  const [f, setF] = useState({
    component_type: "", serial_out: "", serial_in: "", part_no: "", supplier: "",
    component_condition: "New", machine_hours: "", comp_hours_achieved: "",
    expected_life_hours: "", failure_mode: "", fitted_by: myFullName || "",
    under_warranty: false, warranty_end: "", claim_ref: "", claim_status: "Not lodged",
    reason_for_change: "", changeout_date: todayForInput(),
  });
  const set = (k, v) => setF((prev) => ({ ...prev, [k]: v }));

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("component_changeout_calc").select("*").eq("event_id", event.id).maybeSingle();
    if (data) {
      setRow(data);
      setF({
        component_type: data.component_type || "", serial_out: data.serial_out || "",
        serial_in: data.serial_in || "", part_no: data.part_no || "", supplier: data.supplier || "",
        component_condition: data.component_condition || "New",
        machine_hours: data.machine_hours ?? "", comp_hours_achieved: data.comp_hours_achieved ?? "",
        expected_life_hours: data.expected_life_hours ?? "", failure_mode: data.failure_mode || "",
        fitted_by: data.fitted_by || myFullName || "", under_warranty: !!data.under_warranty,
        warranty_end: data.warranty_end || "", claim_ref: data.claim_ref || "",
        claim_status: data.claim_status || "Not lodged", reason_for_change: data.reason_for_change || "",
        changeout_date: data.changeout_date || todayForInput(),
      });
      const { data: scans } = await supabase.from("component_changeout_scans")
        .select("*").eq("changeout_id", data.id).order("uploaded_at", { ascending: false });
      setDocs(scans || []);
    }
    setLoading(false);
  }, [event.id, myFullName]);

  useEffect(() => { load(); }, [load]);

  // Shown live while typing so the technician sees the number that will
  // carry the claim, rather than discovering it on a report later.
  const livePct = useMemo(() => {
    const life = Number(f.expected_life_hours), got = Number(f.comp_hours_achieved);
    if (!life || !got || isNaN(life) || isNaN(got)) return null;
    return Math.round((got / life) * 1000) / 10;
  }, [f.expected_life_hours, f.comp_hours_achieved]);

  const save = async () => {
    setError("");
    if (!f.component_type.trim()) { setError("Component type is required."); return; }
    setSaving(true);
    const num = (v) => (v === "" || v === null ? null : Number(v));
    const payload = {
      asset_id: event.asset_id, event_id: event.id,
      changeout_date: f.changeout_date || todayForInput(),
      changeout_hours: num(f.machine_hours),
      component_type: f.component_type.trim(),
      serial_out: f.serial_out.trim() || null, serial_in: f.serial_in.trim() || null,
      part_no: f.part_no.trim() || null, supplier: f.supplier.trim() || null,
      component_condition: f.component_condition,
      machine_hours: num(f.machine_hours), comp_hours_achieved: num(f.comp_hours_achieved),
      expected_life_hours: num(f.expected_life_hours),
      failure_mode: f.failure_mode.trim() || null, fitted_by: f.fitted_by.trim() || null,
      under_warranty: f.under_warranty, warranty_end: f.warranty_end || null,
      claim_ref: f.claim_ref.trim() || null, claim_status: f.claim_status,
      reason_for_change: f.reason_for_change.trim() || null,
    };
    try {
      if (row) {
        const { error: e1 } = await supabase.from("component_changeout_log").update(payload).eq("id", row.id);
        if (e1) throw e1;
      } else {
        const { error: e2 } = await supabase.from("component_changeout_log").insert(payload);
        if (e2) throw e2;
      }
      // The register is what makes serial history work across machines.
      if (f.serial_in.trim()) {
        await supabase.from("component_register").insert({
          asset_id: event.asset_id, component_id: f.serial_in.trim(),
          component_type: f.component_type.trim(), serial_number: f.serial_in.trim(),
          installed_date: f.changeout_date || todayForInput(),
          installed_hours: num(f.machine_hours), expected_life_hours: num(f.expected_life_hours),
          condition: f.component_condition, supplier: f.supplier.trim() || null,
          warranty_end: f.warranty_end || null,
        });
      }
      await load();
    } catch (err) {
      setError(err.message || String(err));
    }
    setSaving(false);
  };

  const upload = async (file) => {
    if (!file || !row) return;
    setUploading(true);
    setError("");
    try {
      const path = `changeouts/${row.id}/${Date.now()}-${file.name.replace(/[^\w.\-]/g, "_")}`;
      const { error: upErr } = await supabase.storage.from("job-card-scans").upload(path, file);
      if (upErr) throw upErr;
      const { error: insErr } = await supabase.from("component_changeout_scans").insert({
        changeout_id: row.id, storage_path: path, file_name: file.name, document_type: docType,
      });
      if (insErr) throw insErr;
      await load();
    } catch (err) {
      setError(err.message || String(err));
    }
    setUploading(false);
  };

  const openDoc = async (d) => {
    const { data } = await supabase.storage.from("job-card-scans").createSignedUrl(d.storage_path, 3600);
    if (data?.signedUrl) window.open(data.signedUrl, "_blank");
  };

  const fs = { width: "100%", padding: "7px 9px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box" };
  const ls = { display: "block", fontSize: 12, fontWeight: 600, color: "#183642", margin: "0 0 4px" };
  const sec = { fontSize: 12, fontWeight: 700, color: "#4B5659", margin: "14px 0 6px" };

  if (loading) return <p style={{ fontSize: 12.5, color: "#859195" }}>Loading change out…</p>;

  return (
    <div style={{ border: `1px solid ${NAVY}`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
      <p style={{ margin: "0 0 2px", fontSize: 13.5, fontWeight: 700, color: NAVY }}>Component Change Out</p>
      <p style={{ margin: "0 0 8px", fontSize: 11.5, color: "#859195" }}>
        The event keeps its own cause and description - this records what was replaced.
      </p>

      {error && <p style={{ fontSize: 12.5, color: "#B85450", margin: "0 0 8px" }}>{error}</p>}

      <p style={sec}>Identity</p>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 4 }}>
        <div>
          <label style={ls}>Component type *</label>
          <input list="changeout-components" value={f.component_type} onChange={(e) => set("component_type", e.target.value)} placeholder="Engine" style={fs} />
          <datalist id="changeout-components">
            {(componentCodes || []).map((c) => <option key={c.id} value={c.name} />)}
          </datalist>
        </div>
        <div><label style={ls}>Serial out</label><input value={f.serial_out} onChange={(e) => set("serial_out", e.target.value)} placeholder="Removed unit" style={fs} /></div>
        <div><label style={ls}>Serial in</label><input value={f.serial_in} onChange={(e) => set("serial_in", e.target.value)} placeholder="Fitted unit" style={fs} /></div>
        <div><label style={ls}>Part no.</label><input value={f.part_no} onChange={(e) => set("part_no", e.target.value)} style={fs} /></div>
        <div><label style={ls}>Supplier / OEM</label><input value={f.supplier} onChange={(e) => set("supplier", e.target.value)} style={fs} /></div>
        <div>
          <label style={ls}>Condition</label>
          <select value={f.component_condition} onChange={(e) => set("component_condition", e.target.value)} style={fs}>
            {COMPONENT_CONDITIONS.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
      </div>

      <p style={sec}>Life</p>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
        <div><label style={ls}>Change out date</label><input type="date" value={f.changeout_date} onChange={(e) => set("changeout_date", e.target.value)} style={fs} /></div>
        <div><label style={ls}>Machine hours</label><input type="number" step="0.1" value={f.machine_hours} onChange={(e) => set("machine_hours", e.target.value)} style={fs} /></div>
        <div><label style={ls}>Hours achieved</label><input type="number" step="0.1" value={f.comp_hours_achieved} onChange={(e) => set("comp_hours_achieved", e.target.value)} style={fs} /></div>
        <div><label style={ls}>Expected life</label><input type="number" step="1" value={f.expected_life_hours} onChange={(e) => set("expected_life_hours", e.target.value)} style={fs} /></div>
      </div>
      {livePct != null && (
        <p style={{ margin: "8px 0 0", fontSize: 12.5, padding: "7px 10px", borderRadius: 8,
                    background: livePct < 70 ? "#F6E2E0" : "#E2EFE9", color: livePct < 70 ? "#8A2F28" : "#2C5646" }}>
          Reached {livePct}% of expected life{livePct < 70 ? " - flagged as a premature failure" : ""}
        </p>
      )}

      <p style={sec}>Warranty</p>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, alignItems: "end" }}>
        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#183642", paddingBottom: 8 }}>
          <input type="checkbox" checked={f.under_warranty} onChange={(e) => set("under_warranty", e.target.checked)} />
          Under warranty
        </label>
        <div><label style={ls}>Warranty end</label><input type="date" value={f.warranty_end} onChange={(e) => set("warranty_end", e.target.value)} style={fs} /></div>
        <div><label style={ls}>Claim ref</label><input value={f.claim_ref} onChange={(e) => set("claim_ref", e.target.value)} style={fs} /></div>
        <div>
          <label style={ls}>Claim status</label>
          <select value={f.claim_status} onChange={(e) => set("claim_status", e.target.value)} style={fs}>
            {CLAIM_STATUSES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
      </div>

      <p style={sec}>Context</p>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
        <div><label style={ls}>Failure mode</label><input value={f.failure_mode} onChange={(e) => set("failure_mode", e.target.value)} placeholder="e.g. Bearing seizure" style={fs} /></div>
        <div><label style={ls}>Reason for change</label><input value={f.reason_for_change} onChange={(e) => set("reason_for_change", e.target.value)} style={fs} /></div>
        <div><label style={ls}>Fitted by</label><input value={f.fitted_by} onChange={(e) => set("fitted_by", e.target.value)} style={fs} /></div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}>
        <button type="button" onClick={save} disabled={saving}
          style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: saving ? "default" : "pointer" }}>
          {saving ? "Saving…" : row ? "Save Change Out" : "Create Change Out"}
        </button>
      </div>

      <p style={sec}>Evidence</p>
      {!row ? (
        <p style={{ fontSize: 12, color: "#859195", margin: 0 }}>
          Save the change out first - documents attach to the saved record.
        </p>
      ) : (
        <>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <select value={docType} onChange={(e) => setDocType(e.target.value)} style={{ ...fs, width: 200 }}>
              {CHANGEOUT_DOC_TYPES.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
            <input type="file" onChange={(e) => { const file = e.target.files?.[0]; e.target.value = ""; upload(file); }} disabled={uploading} style={{ fontSize: 12.5 }} />
            {uploading && <span style={{ fontSize: 12, color: "#859195" }}>Uploading…</span>}
          </div>
          {docs.length === 0 ? (
            <p style={{ fontSize: 12, color: "#859195", margin: "8px 0 0" }}>No documents attached yet.</p>
          ) : (
            <ul style={{ listStyle: "none", padding: 0, margin: "8px 0 0" }}>
              {docs.map((d) => (
                <li key={d.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: "1px solid #EFEEE7", fontSize: 12.5 }}>
                  <span style={{ background: "#E2EFE9", color: "#2C5646", padding: "2px 8px", borderRadius: 6, fontSize: 11.5, fontWeight: 600 }}>{d.document_type}</span>
                  <button type="button" onClick={() => openDoc(d)} style={{ background: "none", border: "none", color: NAVY, textDecoration: "underline", cursor: "pointer", padding: 0, fontSize: 12.5 }}>
                    {d.file_name || "document"}
                  </button>
                </li>
              ))}
            </ul>
          )}
        </>
      )}
    </div>
  );
}

function EventPulledPartsPanel({ event, parts, onRefresh }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showAdd, setShowAdd] = useState(false);

  const partById = useMemo(() => Object.fromEntries(parts.map((p) => [p.id, p])), [parts]);

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("parts_used").select("*").eq("event_id", event.id).order("created_at", { ascending: false });
    if (!error) setRows(data || []);
    setLoading(false);
  }, [event.id]);

  useEffect(() => { load(); }, [load]);

  const confirmRow = async (row) => {
    await supabase.from("parts_used").update({ status: "Confirmed", confirmed_at: new Date().toISOString() }).eq("id", row.id);
    const p = partById[row.part_id];
    logActivity("Events", event.id, "updated", `Confirmed ${row.qty_used} × ${p?.part_no || "part"} pulled from inventory`, event.asset_id);
    load();
  };

  return (
    <div style={{ marginTop: 18, paddingTop: 16, borderTop: "1px solid #E2E6E3" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <p style={{ fontSize: 13, fontWeight: 700, color: NAVY, margin: 0 }}>Parts Pulled from Inventory</p>
        <button type="button" onClick={() => setShowAdd(true)} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "6px 12px", borderRadius: 8, cursor: "pointer" }}>
          + Pull Part
        </button>
      </div>
      <p style={{ fontSize: 11.5, color: "#859195", margin: "0 0 10px" }}>
        For parts you already have on hand - stock is deducted immediately. If a part needs to be ordered instead, use Order Parts (New Work Order) below.
      </p>

      {loading ? (
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>Loading…</p>
      ) : rows.length === 0 ? (
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>No parts pulled from inventory against this event yet.</p>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          {rows.map((r) => {
            const p = partById[r.part_id];
            return (
              <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5, background: "#F9F8F4", border: "1px solid #E2E6E3", borderRadius: 6, padding: "7px 10px" }}>
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
        <PartUsedForm
          parts={parts}
          event={event}
          onClose={() => setShowAdd(false)}
          onSaved={() => { setShowAdd(false); load(); onRefresh?.(); }}
        />
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
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <p style={{ fontSize: 13, fontWeight: 700, color: NAVY, margin: 0 }}>Work Orders</p>
        <button type="button" onClick={() => { setEditingWo(null); setShowWoForm(true); }} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "6px 12px", borderRadius: 8, cursor: "pointer" }}>
          + Order Parts (New Work Order)
        </button>
      </div>
      <p style={{ fontSize: 11.5, color: "#859195", margin: "0 0 10px" }}>
        For parts that need to be ordered in rather than pulled from stock - each Work Order gets its own number, and parts can still be logged against it once it arrives.
      </p>

      {loading ? (
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>Loading…</p>
      ) : linked.length === 0 ? (
        <p style={{ fontSize: 12.5, color: "#859195", margin: 0 }}>No Work Orders linked to this event yet.</p>
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
          .print-footer { display: block !important; position: fixed; bottom: 8px; left: 0; right: 0; text-align: center; font-size: 10px; color: #666; }
        }
        .print-footer { display: none; }
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

          <p className="print-footer">Printed: {new Date().toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" })}</p>
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

function FuelLogForm({ assets, existing, userEmail, myFullName, dailyHours, fuelLog, onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || assets[0]?.asset_id || "");
  const [fillDate, setFillDate] = useState(existing?.fill_date || todayForInput());
  const [fillTime, setFillTime] = useState(existing?.fill_time || nowForInput().slice(11, 16));
  const [hourMeter, setHourMeter] = useState(existing?.hour_meter ?? "");
  const [litres, setLitres] = useState(existing?.litres ?? "");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const recordedBy = existing?.recorded_by || myFullName || userEmail || "";

  const hourMeterWarning = useMemo(
    () => checkHourMeterPlausibility(assetId, hourMeter, fillDate, dailyHours),
    [assetId, hourMeter, fillDate, dailyHours]
  );

  // Same hour meter reading already logged for this machine, on a
  // different entry - almost always a duplicate/double-capture rather
  // than a genuine second fill, so it's blocked rather than just warned.
  const duplicateReading = useMemo(() => {
    if (hourMeter === "" || !assetId) return false;
    return (fuelLog || []).some((r) =>
      r.asset_id === assetId &&
      Number(r.hour_meter) === Number(hourMeter) &&
      (!isEdit || r.id !== existing.id)
    );
  }, [fuelLog, assetId, hourMeter, isEdit, existing]);

  useEffect(() => {
    if (existing?.id) {
      logActivity("Fuel Log", existing.id, "viewed", `Opened fuel log entry from ${existing.fill_date || ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
    if (duplicateReading) { setError("This hour meter reading has already been logged for this machine - check for a duplicate entry."); return; }
    setSaving(true);
    const payload = {
      asset_id: assetId, fill_date: fillDate, fill_time: fillTime || null, hour_meter: Number(hourMeter),
      litres: Number(litres), recorded_by: recordedBy || null, notes: notes || null,
    };
    try {
      const assetName = assets.find((a) => a.asset_id === assetId)?.asset_name || assetId;
      const summary = `${assetName} - ${litres}L on ${fillDate}, hour meter ${hourMeter}`;
      if (isEdit) {
        const { error: dbError } = await supabase.from("fuel_log").update(payload).eq("id", existing.id);
        if (dbError) throw dbError;
        logActivity("Fuel Log", existing.id, "updated", summary, assetId);
      } else {
        const { data, error: dbError } = await supabase.from("fuel_log").insert(payload).select().single();
        if (dbError) throw dbError;
        logActivity("Fuel Log", data?.id, "created", summary, assetId);
      }
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
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={labelStyle}>Date</label>
            <input type="date" value={fillDate} onChange={(e) => setFillDate(e.target.value)} required max={todayForInput()} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Time</label>
            <input type="time" value={fillTime} onChange={(e) => setFillTime(e.target.value)} style={fieldStyle} />
          </div>
          <div>
            <label style={labelStyle}>Hour Meter</label>
            <input type="number" step="0.1" value={hourMeter} onChange={(e) => setHourMeter(e.target.value)} required style={fieldStyle} />
          </div>
        </div>
        {duplicateReading && (
          <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", borderRadius: 8, padding: "8px 10px", marginBottom: 12, fontSize: 12, color: "#7A3330", fontWeight: 600 }}>
            This hour meter reading is already logged for this machine - looks like a duplicate entry.
          </div>
        )}
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
          <button type="submit" disabled={saving || duplicateReading} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: (saving || duplicateReading) ? "default" : "pointer", opacity: (saving || duplicateReading) ? 0.7 : 1 }}>
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

function InspectionForm({ assets, existing, userEmail, myFullName, onClose, onSaved }) {
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
  const inspector = existing?.inspector || myFullName || userEmail || "";

  useEffect(() => {
    if (existing?.id) {
      logActivity("Inspections", existing.id, "viewed", `Opened ${existing.inspection_type || ""} inspection from ${existing.log_date || ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
      const assetName = assets.find((a) => a.asset_id === assetId)?.asset_name || assetId;
      const summary = `${assetName} - ${inspectionType} inspection on ${logDate}, result: ${result}${result !== "Pass" ? ` (${riskRating} risk)` : ""}`;
      if (isEdit) {
        const { error: dbError } = await supabase.from("inspections").update(payload).eq("id", existing.id);
        if (dbError) throw dbError;
        logActivity("Inspections", existing.id, "updated", summary, assetId);
      } else {
        const { data, error: dbError } = await supabase.from("inspections").insert(payload).select().single();
        if (dbError) throw dbError;
        logActivity("Inspections", data?.id, "created", summary, assetId);
      }
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

function InspectionsPage({ assets, inspections, userEmail, myFullName, onRefresh }) {
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
    await deleteWithReason("inspections", deleting.id, "id", reason, userEmail, deleting.asset_id);
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
        <InspectionForm assets={assets} existing={editing} userEmail={userEmail} myFullName={myFullName}
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

  useEffect(() => {
    if (existing?.id) {
      logActivity("Backlogs", existing.id, "viewed", `Opened backlog item ${existing.description ? `"${existing.description}"` : ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
      logActivity("Backlogs", (isEdit ? existing.id : data?.id), status === "Closed" ? "closed" : (isEdit ? "updated" : "created"), backlogSummaryParts.join(" - "), assetId);
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
    await deleteWithReason("backlogs", deleting.id, "id", reason, userEmail, deleting.asset_id);
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

function DailyServiceForm({ assets, defaultAssetId, defaultDate, existing, userEmail, myFullName, onClose, onSaved }) {
  const [assetId, setAssetId] = useState(existing?.asset_id || defaultAssetId || assets[0]?.asset_id || "");
  const [serviceDate, setServiceDate] = useState(existing?.service_date || defaultDate || todayForInput());
  const [completedBy, setCompletedBy] = useState(existing?.completed_by || myFullName || userEmail || "");
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

function DailyServicePage({ assets, dailyServiceChecklist, breakdowns, dailyHours, userEmail, myFullName, onRefresh }) {
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
          myFullName={myFullName}
          onClose={() => setShowForm(false)}
          onSaved={() => { setShowForm(false); onRefresh(); }}
        />
      )}
    </div>
  );
}

function FuelLogPage({ assets, fuelLog, userEmail, myFullName, dailyHours, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const [groupByMachine, setGroupByMachine] = useState(true);
  const [dateFrom, setDateFrom] = useState(firstOfMonthForInput());
  const [dateTo, setDateTo] = useState(todayForInput());
  const prefs = useTablePrefs("fuel_log", FUEL_LOG_COLUMNS);
  const columns = prefs.columns;

  const filtered = useMemo(() => {
    let rows = fuelLog;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    if (dateFrom) rows = rows.filter((r) => (r.fill_date || "") >= dateFrom);
    if (dateTo) rows = rows.filter((r) => (r.fill_date || "") <= dateTo);
    return [...rows].sort((a, b) => (b.fill_date || "").localeCompare(a.fill_date || ""));
  }, [fuelLog, assets, query, selectedFleet, selectedAsset, dateFrom, dateTo]);

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "fill_date", hoursKey: "litres",
    sortKey: prefs.sortKey, sortDir: prefs.sortDir, sortRows: prefs.sortRows,
  }), [filtered, assets, groupByMachine, prefs.sortKey, prefs.sortDir, prefs.sortRows]);

  const handleDelete = async (reason) => {
    await deleteWithReason("fuel_log", deleting.id, "id", reason, userEmail, deleting.asset_id);
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
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap", fontSize: 13, color: "#4B5659" }}>
        <span style={{ fontWeight: 600 }}>Dates</span>
        <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)}
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        <span>to</span>
        <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)}
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        <button onClick={() => { setDateFrom(firstOfMonthForInput()); setDateTo(todayForInput()); }}
          style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer" }}>
          Month to date
        </button>
        <button onClick={() => { setDateFrom(""); setDateTo(""); }}
          style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", fontSize: 12.5, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer" }}>
          Show all
        </button>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search fuel log" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
          <ColumnsButton prefs={prefs} />
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Log Fuel
          </button>
        </div>
      </div>
      <ExcelSync
        data={fuelLog} assets={assets} fields={FUEL_LOG_FIELDS} tableName="fuel_log"
        sheetTitle="Fuel Log" filenamePrefix="Fuel_Log" onRefresh={onRefresh}
        extraOnSave={(obj) => ({ recorded_by: obj.recorded_by || userEmail || null })}
      />
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <SmartTableHead prefs={prefs} />
          <MachineTableBody
            groups={groups} columns={columns}
            onRowClick={(row) => { setEditing(row); setShowForm(true); }}
            onDelete={(row) => setDeleting(row)}
            totalLabel="L" countNoun="fill" countNounPlural="fills"
            emptyMessage={fuelLog.length === 0 ? "No fuel entries yet." : "No entries match your filters."}
          />
        </table>
      </div>
      {showForm && (
        <FuelLogForm assets={assets} existing={editing} userEmail={userEmail} myFullName={myFullName} dailyHours={dailyHours} fuelLog={fuelLog}
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

function OilConsumptionForm({ assets, existing, userEmail, myFullName, dailyHours, onClose, onSaved }) {
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
  const recordedBy = existing?.recorded_by || myFullName || userEmail || "";

  const hourMeterWarning = useMemo(
    () => checkHourMeterPlausibility(assetId, hourMeter, fillDate, dailyHours),
    [assetId, hourMeter, fillDate, dailyHours]
  );

  useEffect(() => {
    if (existing?.id) {
      logActivity("Oil Consumption", existing.id, "viewed", `Opened oil consumption entry from ${existing.fill_date || ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
      const assetName = assets.find((a) => a.asset_id === assetId)?.asset_name || assetId;
      const summary = `${assetName} - ${litres}L ${oilType} (${fillReason}) on ${fillDate}`;
      if (isEdit) {
        const { error: dbError } = await supabase.from("oil_consumption").update(payload).eq("id", existing.id);
        if (dbError) throw dbError;
        logActivity("Oil Consumption", existing.id, "updated", summary, assetId);
      } else {
        const { data, error: dbError } = await supabase.from("oil_consumption").insert(payload).select().single();
        if (dbError) throw dbError;
        logActivity("Oil Consumption", data?.id, "created", summary, assetId);
      }
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

function OilConsumptionPage({ assets, oilConsumption, userEmail, myFullName, dailyHours, onRefresh }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");

  const [groupByMachine, setGroupByMachine] = useState(true);
  const [dateFrom, setDateFrom] = useState(firstOfMonthForInput());
  const [dateTo, setDateTo] = useState(todayForInput());
  const prefs = useTablePrefs("oil_consumption", OIL_CONSUMPTION_COLUMNS);
  const columns = prefs.columns;

  const filtered = useMemo(() => {
    let rows = oilConsumption;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (query.trim()) { const q = query.toLowerCase(); rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))); }
    if (dateFrom) rows = rows.filter((r) => (r.fill_date || "") >= dateFrom);
    if (dateTo) rows = rows.filter((r) => (r.fill_date || "") <= dateTo);
    return [...rows].sort((a, b) => (b.fill_date || "").localeCompare(a.fill_date || ""));
  }, [oilConsumption, assets, query, selectedFleet, selectedAsset, dateFrom, dateTo]);

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "fill_date", hoursKey: "litres",
    sortKey: prefs.sortKey, sortDir: prefs.sortDir, sortRows: prefs.sortRows,
  }), [filtered, assets, groupByMachine, prefs.sortKey, prefs.sortDir, prefs.sortRows]);

  const handleDelete = async (reason) => {
    await deleteWithReason("oil_consumption", deleting.id, "id", reason, userEmail, deleting.asset_id);
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
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap", fontSize: 13, color: "#4B5659" }}>
        <span style={{ fontWeight: 600 }}>Dates</span>
        <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)}
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        <span>to</span>
        <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)}
          style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }} />
        <button onClick={() => { setDateFrom(firstOfMonthForInput()); setDateTo(todayForInput()); }}
          style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12.5, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer" }}>
          Month to date
        </button>
        <button onClick={() => { setDateFrom(""); setDateTo(""); }}
          style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", fontSize: 12.5, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer" }}>
          Show all
        </button>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search oil consumption" style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
          <ColumnsButton prefs={prefs} />
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Log Oil
          </button>
        </div>
      </div>
      <ExcelSync
        data={oilConsumption} assets={assets} fields={OIL_FIELDS} tableName="oil_consumption"
        sheetTitle="Oil Consumption" filenamePrefix="Oil_Consumption" onRefresh={onRefresh}
        extraOnSave={(obj) => ({ recorded_by: obj.recorded_by || userEmail || null })}
      />
      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <SmartTableHead prefs={prefs} />
          <MachineTableBody
            groups={groups} columns={columns}
            onRowClick={(row) => { setEditing(row); setShowForm(true); }}
            onDelete={(row) => setDeleting(row)}
            totalLabel="L" countNoun="fill" countNounPlural="fills"
            emptyMessage={oilConsumption.length === 0 ? "No oil entries yet." : "No entries match your filters."}
          />
        </table>
      </div>
      {showForm && (
        <OilConsumptionForm assets={assets} existing={editing} userEmail={userEmail} myFullName={myFullName} dailyHours={dailyHours}
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

// Full Plant Performance KPI column set: [key, label, group, format].
// Groups drive the merged banner row above the headers.
//   n1/n2 = numbers to 1 / 2 decimals, int = whole, pct = fraction shown as %
// ---------------------------------------------------------------------
// KPI row drill-down: the events behind one machine's figures.
//
// Reconciliation is the point of this screen. Individual event hours can
// sum to MORE than the row's Total Downtime, because plant_performance_kpi
// merges overlapping events so a machine is never down twice at once.
// That difference is shown explicitly rather than left to look like an
// error in the report.
//
// Events are clipped to the report period here exactly as the function
// clips them, so a breakdown spanning the period boundary contributes
// only the hours that fall inside it.
// ---------------------------------------------------------------------
function KpiEventsDrilldown({ machine, fromDateTime, toDateTime, onClose }) {
  const [events, setEvents] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const periodStart = useMemo(() => new Date(fromDateTime), [fromDateTime]);
  const periodEnd = useMemo(() => new Date(toDateTime), [toDateTime]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [bd, wo] = await Promise.all([
        supabase.from("breakdown_log_calc").select("*").eq("asset_id", machine.asset_id),
        supabase.from("work_orders_calc").select("*").eq("asset_id", machine.asset_id).eq("work_type", "Preventive"),
      ]);
      if (cancelled) return;
      if (bd.error || wo.error) { setError((bd.error || wo.error).message); setLoading(false); return; }

      const clip = (startIso, endIso) => {
        if (!startIso) return null;
        const s = new Date(startIso);
        const e = endIso ? new Date(endIso) : new Date();
        if (s >= periodEnd || e <= periodStart) return null;          // outside the period
        const cs = s < periodStart ? periodStart : s;
        const ce = e > periodEnd ? periodEnd : e;
        return { start: s, end: endIso ? e : null, clippedStart: cs, clippedEnd: ce,
                 hours: (ce - cs) / 3600000, trimmed: cs > s || ce < e };
      };

      const rows = [];
      (bd.data || []).forEach((b) => {
        const c = clip(b.downtime_start, b.downtime_end);
        if (!c) return;
        rows.push({
          id: `b${b.id}`, kind: b.event_type === "Planned" ? "Planned" : "Breakdown",
          description: b.description, cause: b.cause_code, component: b.component_affected,
          responsibility: b.responsibility, status: b.repair_status,
          open: !b.downtime_end, ...c,
        });
      });
      (wo.data || []).forEach((w) => {
        const c = clip(w.actual_start, w.actual_finish);
        if (!c) return;
        rows.push({
          id: `w${w.id}`, kind: "Planned", description: w.problem_scope,
          cause: w.work_type, component: w.component, responsibility: w.responsibility,
          status: w.status, open: !w.actual_finish, ...c,
        });
      });
      rows.sort((a, b) => b.clippedStart - a.clippedStart);
      setEvents(rows);
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, [machine.asset_id, periodStart, periodEnd]);

  // Merge overlapping intervals the same way the SQL does, so the
  // drill-down can show WHY the sum differs from the row.
  const merged = useMemo(() => {
    const iv = events.map((e) => [e.clippedStart.getTime(), e.clippedEnd.getTime()]).sort((a, b) => a[0] - b[0]);
    let total = 0, curS = null, curE = null;
    iv.forEach(([s, e]) => {
      if (curS === null) { curS = s; curE = e; return; }
      if (s <= curE) { curE = Math.max(curE, e); }
      else { total += curE - curS; curS = s; curE = e; }
    });
    if (curS !== null) total += curE - curS;
    return total / 3600000;
  }, [events]);

  const sums = useMemo(() => {
    const unplanned = events.filter((e) => e.kind === "Breakdown");
    const planned = events.filter((e) => e.kind === "Planned");
    const h = (arr) => arr.reduce((a, e) => a + e.hours, 0);
    return {
      unplannedHours: h(unplanned), plannedHours: h(planned),
      unplannedCount: unplanned.length, plannedCount: planned.length,
      rawTotal: h(events),
    };
  }, [events]);

  const overlap = sums.rawTotal - merged;

  const n = (v, d = 2) => (v == null || isNaN(v) ? "-" : Number(v).toLocaleString(undefined, { minimumFractionDigits: d, maximumFractionDigits: d }));
  const dt = (d) => (d ? new Date(d).toLocaleString("en-ZA", { dateStyle: "short", timeStyle: "short" }) : "-");

  const th = { textAlign: "left", padding: "7px 9px", fontSize: 11, fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" };
  const thR = { ...th, textAlign: "right" };
  const td = { padding: "7px 9px", fontSize: 12, verticalAlign: "top" };
  const tdR = { ...td, textAlign: "right", whiteSpace: "nowrap" };

  const recon = (label, value, expected, note) => {
    const diff = expected == null ? null : Math.abs(Number(value) - Number(expected));
    const ok = diff == null || diff < 0.05;
    return (
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, padding: "5px 0", fontSize: 12, borderBottom: "1px solid #EFEEE7" }}>
        <span style={{ color: "#4B5659" }}>{label}</span>
        <span style={{ whiteSpace: "nowrap" }}>
          <strong>{n(value)}</strong>
          {expected != null && (
            <span style={{ color: ok ? "#2C5646" : "#B85450", marginLeft: 8 }}>
              {ok ? "✓ matches report" : `report shows ${n(expected)}`}
            </span>
          )}
          {note && <span style={{ color: "#859195", marginLeft: 8 }}>{note}</span>}
        </span>
      </div>
    );
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(24,54,66,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: 24, zIndex: 100, overflowY: "auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 12, padding: 22, width: 1080, maxWidth: "97vw", maxHeight: "92vh", overflowY: "auto" }}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 12 }}>
          <div>
            <p style={{ margin: 0, fontSize: 17, fontWeight: 700, color: NAVY }}>
              {machine.asset_id}{machine.asset_name ? ` — ${machine.asset_name}` : ""}
            </p>
            <p style={{ margin: "2px 0 0", fontSize: 12, color: "#859195" }}>
              {[machine.fleet, machine.model].filter(Boolean).join(" · ")} · {formatRangeForDisplay(fromDateTime, toDateTime)}
            </p>
          </div>
          <button onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", padding: "6px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Close</button>
        </div>

        {error && <p style={{ fontSize: 12.5, color: "#B85450" }}>{error}</p>}
        {loading ? <p style={{ fontSize: 13, color: "#859195" }}>Loading events…</p> : (
          <>
            <div style={{ background: "#F7F8F6", borderRadius: 10, padding: "10px 14px", marginBottom: 14 }}>
              <p style={{ margin: "0 0 4px", fontSize: 12.5, fontWeight: 700, color: NAVY }}>Reconciliation against the report row</p>
              {recon(`Unplanned events (${sums.unplannedCount})`, sums.unplannedHours, machine.unplanned_downtime_hrs)}
              {recon(`Planned events (${sums.plannedCount})`, sums.plannedHours, machine.planned_downtime_hrs)}
              {recon("Sum of every event", sums.rawTotal, null, "before overlap is merged")}
              {recon("Total downtime", merged, machine.total_downtime_hrs, "overlap counted once")}
              {overlap > 0.05 && (
                <p style={{ margin: "8px 0 0", fontSize: 11.5, color: "#8A6320", background: "#F5E9D8", padding: "7px 10px", borderRadius: 8 }}>
                  {n(overlap)} hours of these events overlap each other. A machine cannot be down twice at the same time, so the report counts that time once — which is why the events above sum to more than Total Downtime.
                </p>
              )}
              <div style={{ display: "flex", gap: 18, marginTop: 10, fontSize: 12, flexWrap: "wrap" }}>
                <span style={{ color: "#4B5659" }}>MTBF <strong>{machine.mtbf ?? "-"}</strong></span>
                <span style={{ color: "#4B5659" }}>MTTR <strong>{machine.mttr ?? "-"}</strong></span>
                <span style={{ color: "#4B5659" }}>Availability <strong>{machine.availability != null ? `${Math.round(machine.availability * 100)}%` : "-"}</strong></span>
                <span style={{ color: "#4B5659" }}>Worked hrs <strong>{machine.worked_hours ?? "-"}</strong></span>
                <span style={{ color: "#859195" }}>MTTR = unplanned hours ÷ {sums.unplannedCount || 0} event{sums.unplannedCount === 1 ? "" : "s"}</span>
              </div>
            </div>

            <p style={{ margin: "0 0 6px", fontSize: 13, fontWeight: 700, color: NAVY }}>Events in this period ({events.length})</p>
            <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#F7F8F6" }}>
                    <th style={th}>Event</th><th style={th}>Component</th><th style={th}>Cause</th>
                    <th style={th}>Responsibility</th><th style={th}>Description</th>
                    <th style={thR}>Down</th><th style={thR}>Up</th><th style={thR}>Hours</th>
                  </tr>
                </thead>
                <tbody>
                  {events.map((e) => (
                    <tr key={e.id} style={{ borderBottom: "1px solid #EFEEE7", background: e.kind === "Planned" ? "#FBFBF8" : "transparent" }}>
                      <td style={td}>
                        <span style={{ background: e.kind === "Planned" ? "#E2EFE9" : "#F6E2E0", color: e.kind === "Planned" ? "#2C5646" : "#7A3330",
                                       padding: "2px 8px", borderRadius: 6, fontSize: 10.5, fontWeight: 600, whiteSpace: "nowrap" }}>{e.kind}</span>
                      </td>
                      <td style={td}>{e.component || "-"}</td>
                      <td style={td}>{e.cause || "-"}</td>
                      <td style={td}>{e.responsibility || <span style={{ color: "#B4B2A9" }}>not set</span>}</td>
                      <td style={{ ...td, maxWidth: 300, whiteSpace: "normal" }}>{e.description || "-"}</td>
                      <td style={tdR}>{dt(e.clippedStart)}</td>
                      <td style={tdR}>{e.open ? <span style={{ color: "#B85450" }}>still down</span> : dt(e.clippedEnd)}</td>
                      <td style={{ ...tdR, fontWeight: 600 }}>
                        {n(e.hours)}
                        {e.trimmed && <span title="Clipped to the report period" style={{ color: "#B07D2B", marginLeft: 4 }}>*</span>}
                      </td>
                    </tr>
                  ))}
                  {events.length === 0 && (
                    <tr><td colSpan={8} style={{ padding: 18, textAlign: "center", color: "#859195", fontSize: 12.5 }}>
                      No events for this machine in the selected period.
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
            {events.some((e) => e.trimmed) && (
              <p style={{ margin: "8px 2px 0", fontSize: 11, color: "#859195" }}>
                * Event runs beyond the report period — only the hours inside the period are counted, exactly as the report does.
              </p>
            )}
          </>
        )}
      </div>
    </div>
  );
}

const KPI_REPORT_COLUMNS = [
  ["asset_id", "Equipment", "", "text", "Equipment"],
  ["asset_name", "Name", "", "text", "Name"],
  ["model", "Model", "", "text", "Model"],
  ["fleet", "Fleet", "", "text", "Fleet"],
  ["uom", "UOM", "", "text", "UOM"],

  ["scheduled_hours", "Scheduled", "Hours", "n1", "Sched"],
  ["opening_hours", "Opening Usage", "Hours", "n1", "Open"],
  ["closing_hours", "Closing Usage", "Hours", "n1", "Close"],
  ["worked_hours", "Worked", "Hours", "n1", "Worked"],

  ["uncontrollable_hours", "Uncontrollable Time", "Prod.", "n2", "Uncontr."],

  ["planned_downtime_hrs", "Planned Downtime Hrs", "Downtime", "n2", "Planned"],
  ["unplanned_downtime_hrs", "Unplanned Downtime Hrs", "Downtime", "n2", "Unplan."],
  ["total_downtime_hrs", "Total Downtime Hrs", "Downtime", "n2", "Total"],

  ["resp_plant", "Plant", "Downtime Responsibility", "n2", "Plant"],
  ["resp_stores", "Stores", "Downtime Responsibility", "n2", "Stores"],
  ["resp_production", "Production", "Downtime Responsibility", "n2", "Prod."],
  ["resp_oem", "OEM", "Downtime Responsibility", "n2", "OEM"],
  ["resp_client", "Client", "Downtime Responsibility", "n2", "Client"],
  ["resp_uncontrollable", "Uncontrollable Time", "Downtime Responsibility", "n2", "Uncontr."],
  ["resp_non_shift", "Non-Shift", "Downtime Responsibility", "n2", "Non-Sh."],
  ["resp_not_set", "Not set", "Downtime Responsibility", "n2", "Not set"],

  ["num_planned_events", "Planned", "Number of Stoppages", "int", "Pl."],
  ["num_unplanned_events", "Unplanned", "Number of Stoppages", "int", "Unpl."],
  ["num_engineering_unplanned", "Engineering Unplanned", "Number of Stoppages", "int", "Eng."],

  ["mtbf", "Equipment MTBF", "Reliability", "n1", "Eq MTBF"],
  ["engineering_mtbf", "Engineering MTBF", "Reliability", "n1", "Eng MTBF"],
  ["mttr", "MTTR", "Reliability", "n1", "MTTR"],
  ["breakdown_pct", "Breakdown %", "Reliability", "pct", "Bkdn %"],

  ["availability_24h", "Last 24 Hours", "Equipment Availability", "pct", "24h"],
  ["availability", "Period", "Equipment Availability", "pct", "Period"],

  ["engineering_availability_24h", "Last 24 Hours", "Engineering Availability", "pct", "24h"],
  ["engineering_availability", "Period", "Engineering Availability", "pct", "Period"],

  ["utilisation_24h", "Last 24 Hours", "Equipment Utilisation", "pct", "24h"],
  ["utilisation", "Period", "Equipment Utilisation", "pct", "Period"],
];

// Columns that make sense to total across a fleet by SUMMING; the rest
// are averaged, and identity columns are left blank on subtotal rows.
const KPI_SUM_KEYS = new Set([
  "worked_hours", "uncontrollable_hours", "planned_downtime_hrs", "unplanned_downtime_hrs",
  "total_downtime_hrs", "resp_plant", "resp_stores", "resp_production", "resp_oem",
  "resp_client", "resp_uncontrollable", "resp_non_shift", "resp_not_set",
  "num_planned_events", "num_unplanned_events", "num_engineering_unplanned",
]);
const KPI_AVG_KEYS = new Set([
  "scheduled_hours", "mtbf", "engineering_mtbf", "mttr", "breakdown_pct",
  "availability_24h", "availability", "engineering_availability_24h",
  "engineering_availability", "utilisation_24h", "utilisation",
]);

// A machine cannot be working and broken at the same time. When worked
// hours plus downtime exceed the scheduled hours of the period, the two
// records contradict each other - the report flags the row rather than
// letting the clamped percentage hide it.
function hasHoursConflict(row) {
  const sched = Number(row?.scheduled_hours);
  const worked = Number(row?.worked_hours);
  const down = Number(row?.total_downtime_hrs);
  if (!isFinite(sched) || !isFinite(worked) || !isFinite(down) || sched <= 0) return false;
  return (worked + down) > sched * 1.001;   // 0.1% tolerance for rounding
}

function kpiAggregate(rowsIn) {
  const out = {};
  const nums = (key) => rowsIn.map((r) => r[key])
    .filter((v) => v !== null && v !== undefined && v !== "")
    .map(Number).filter((v) => !isNaN(v));
  for (const [key] of KPI_REPORT_COLUMNS) {
    const vals = nums(key);
    if (KPI_SUM_KEYS.has(key)) out[key] = vals.length ? vals.reduce((a, b) => a + b, 0) : null;
    else if (KPI_AVG_KEYS.has(key)) out[key] = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
    else out[key] = null;
  }
  return out;
}

const MTBF_MTTR_COLUMNS = KPI_REPORT_COLUMNS.map(([k, l]) => [k, l]);

function MtbfMttrReportPage({ assets }) {
  const [fromDateTime, setFromDateTime] = useState(DEFAULT_FROM);
  const [toDateTime, setToDateTime] = useState(DEFAULT_TO);
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [kpiData, setKpiData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [eventsDrill, setEventsDrill] = useState(null);
  const [sortKey, setSortKey] = useState("num_unplanned_events");
  const [sortDir, setSortDir] = useState("desc");

  useEffect(() => {
    let cancelled = false;
    async function load() {
      setLoading(true);
      setError(null);

      // "Last 24h" always trails the end of the selected date range (not
      // "now") - e.g. range ending 19 Aug 18:00 means last-24h is 18 Aug
      // 18:00 -> 19 Aug 18:00. "Month to date" is the 1st of that same
      // end date's month, through the end of the selected range.
      const periodEnd = new Date(toDateTime);
      const last24Start = new Date(periodEnd.getTime() - 24 * 60 * 60 * 1000);
      const mtdStart = new Date(periodEnd.getFullYear(), periodEnd.getMonth(), 1);

      const [mainRes, last24Res, mtdRes] = await Promise.all([
        supabase.rpc("plant_performance_kpi", {
          period_start: new Date(fromDateTime).toISOString(),
          period_end: periodEnd.toISOString(),
        }),
        supabase.rpc("plant_performance_kpi", {
          period_start: last24Start.toISOString(),
          period_end: periodEnd.toISOString(),
        }),
        supabase.rpc("plant_performance_kpi", {
          period_start: mtdStart.toISOString(),
          period_end: periodEnd.toISOString(),
        }),
      ]);
      if (cancelled) return;

      const rpcError = mainRes.error || last24Res.error || mtdRes.error;
      if (rpcError) {
        setError(rpcError.message);
        setLoading(false);
        return;
      }

      const availByAsset = (rows) => Object.fromEntries((rows || []).map((r) => [r.asset_id, r.availability]));
      const last24ByAsset = availByAsset(last24Res.data);
      const mtdByAsset = availByAsset(mtdRes.data);
      // Full last-24h row too - the charts plot it beside the selected
      // period, so mtbf/mttr/utilisation are needed, not just availability.
      const last24RowByAsset = Object.fromEntries((last24Res.data || []).map((r) => [r.asset_id, r]));

      // plant_performance_kpi() returns every asset across every site (it
      // has no site filter of its own) - scope it down to just the assets
      // belonging to the currently selected site before anything else
      // touches this data, so another site's real equipment never shows
      // up here.
      const siteAssetIds = new Set(assets.map((a) => a.asset_id));
      const merged = (mainRes.data || [])
        .filter((r) => siteAssetIds.has(r.asset_id))
        .map((r) => ({
          ...r,
          availability_24h: last24ByAsset[r.asset_id] ?? null,
          availability_mtd: mtdByAsset[r.asset_id] ?? null,
          mtbf_24h: last24RowByAsset[r.asset_id]?.mtbf ?? null,
          mttr_24h: last24RowByAsset[r.asset_id]?.mttr ?? null,
          utilisation_24h: last24RowByAsset[r.asset_id]?.utilisation ?? null,
          engineering_availability_24h: last24RowByAsset[r.asset_id]?.engineering_availability ?? null,
        }));

      setKpiData(merged);
      setLoading(false);
    }
    load();
    return () => { cancelled = true; };
  }, [fromDateTime, toDateTime, assets]);

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

  // One row per fleet, same pattern as the Dashboard's KPI charts - used
  // for the three charts below whenever no fleet/equipment is selected,
  // so they show 5 readable bars instead of 30+.
  const fleetKpiData = useMemo(() => {
    const byFleet = {};
    kpiData.forEach((r) => {
      if (!r.fleet) return;
      if (!byFleet[r.fleet]) byFleet[r.fleet] = [];
      byFleet[r.fleet].push(r);
    });
    return Object.keys(byFleet).sort().map((fleet) => ({
      fleet,
      ...aggregateMetrics(byFleet[fleet], byFleet[fleet].map((r) => r.asset_id)),
    }));
  }, [kpiData]);

  const handleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("desc"); }
  };

  const KPI_FMT = Object.fromEntries(KPI_REPORT_COLUMNS.map(([k, , , f]) => [k, f]));

  const fmt = (key, val) => {
    if (val == null || val === "") return "-";
    const n = Number(val);
    switch (KPI_FMT[key]) {
      case "pct": return `${Math.round(n * 100)}%`;
      case "n1":  return n === 0 ? "-" : n.toFixed(1);
      case "n2":  return n === 0 ? "-" : n.toFixed(2);
      case "int": return n === 0 ? "-" : String(Math.round(n));
      default:    return val;
    }
  };

  // Built with ExcelJS rather than the xlsx community build, which
  // cannot write fonts, fills or number formats. Just as important:
  // numbers go in as NUMBERS with a display format, not as pre-formatted
  // strings - the old export wrote "213.3" and "61%" as text, so nothing
  // in the file could be summed, sorted or charted by the person opening it.
  const exportToExcel = async () => {
    const ExcelJS = (await import("exceljs")).default;
    const now = new Date();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Mine2U";
    wb.created = now;
    const ws = wb.addWorksheet("Plant Performance KPI", {
      views: [{ state: "frozen", xSplit: 2, ySplit: 10 }],
      pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    const COLS = KPI_REPORT_COLUMNS;
    const N = COLS.length;
    const INK = "FF1F6668", HEAD = "FF203B46", BAND = "FF3D7379", WARN = "FFFDF3E3";
    const base = { name: "Arial", size: 9.5 };

    // Excel formats matching the on-screen ones. Percentages are written
    // as fractions so Excel's own % format applies - the file stays
    // numeric and can be summed, sorted and charted.
    const XL_FMT = {
      n1: '#,##0.0;-#,##0.0;"-"',
      n2: '#,##0.00;-#,##0.00;"-"',
      int: '0;-0;"-"',
      pct: "0%",           // 0% is a real reading, not an absent one
      text: null,
    };

    const bandRow = (label, value, r) => {
      ws.mergeCells(r, 2, r, Math.min(N, 8));
      const a = ws.getCell(r, 1), b = ws.getCell(r, 2);
      a.value = label; b.value = value;
      a.font = { ...base, bold: true, color: { argb: "FF4B5659" } };
      b.font = base;
    };

    ws.mergeCells(1, 1, 1, Math.min(N, 10));
    const title = ws.getCell(1, 1);
    title.value = "Plant Performance KPI";
    title.font = { name: "Arial", size: 15, bold: true, color: { argb: INK } };
    ws.getRow(1).height = 22;

    bandRow("Date Range", formatRangeForDisplay(fromDateTime, toDateTime), 3);
    bandRow("Execution", now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" }), 4);
    bandRow("Hierarchy", `Fleet: ${selectedFleet || "All"}    Equipment: ${selectedAsset || "All"}`, 5);
    bandRow("Group By", "Fleet", 6);
    bandRow("Equipment count", String(rows.length), 7);

    // --- two header rows: group banner, then column names ---
    const GROUP_ROW = 9, HEADER_ROW = 10;
    let ci = 0;
    while (ci < N) {
      const g = COLS[ci][2];
      let span = 1;
      while (ci + span < N && COLS[ci + span][2] === g) span += 1;
      const c = ws.getCell(GROUP_ROW, ci + 1);
      if (span > 1) ws.mergeCells(GROUP_ROW, ci + 1, GROUP_ROW, ci + span);
      c.value = g || null;
      c.font = { ...base, bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BAND } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      ci += span;
    }
    COLS.forEach(([, label, , fmtType], i) => {
      const c = ws.getCell(HEADER_ROW, i + 1);
      c.value = label;   // full names in the file; the screen uses short ones
      c.font = { ...base, bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEAD } };
      c.alignment = { horizontal: fmtType === "text" ? "left" : "center", vertical: "middle", wrapText: true };
    });
    ws.getRow(GROUP_ROW).height = 18;
    ws.getRow(HEADER_ROW).height = 30;

    const writeRow = (r, values, opts = {}) => {
      COLS.forEach(([key, , , fmtType], i) => {
        const c = ws.getCell(r, i + 1);
        const raw = values[key];
        c.value = raw === null || raw === undefined || raw === "" ? null
                : (fmtType === "text" ? raw : Number(raw));
        if (XL_FMT[fmtType]) c.numFmt = XL_FMT[fmtType];
        c.font = { ...base, bold: !!opts.bold, color: { argb: opts.bold ? INK : "FF183642" } };
        c.alignment = { horizontal: fmtType === "text" ? "left" : "right" };
        if (opts.fill) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
        if (opts.topBorder) c.border = { top: { style: "thin", color: { argb: INK } } };
      });
    };

    // --- data, grouped by fleet with a subtotal after each ---
    const byFleet = new Map();
    rows.forEach((r) => {
      const f = r.fleet || "(no fleet)";
      if (!byFleet.has(f)) byFleet.set(f, []);
      byFleet.get(f).push(r);
    });

    let r = HEADER_ROW + 1;
    let conflicts = 0;
    for (const [fleetName, fleetRows] of byFleet) {
      for (const row of fleetRows) {
        const conflict = hasHoursConflict(row);
        if (conflict) conflicts += 1;
        writeRow(r, row, { fill: conflict ? WARN : null });
        r += 1;
      }
      const sub = kpiAggregate(fleetRows);
      writeRow(r, { ...sub, asset_id: `${fleetName} (${fleetRows.length})`, asset_name: "", model: "", fleet: "", uom: "" },
               { bold: true, fill: "FFF2F1EA" });
      r += 1;
    }
    if (rows.length) {
      const tot = kpiAggregate(rows);
      writeRow(r, { ...tot, asset_id: `Total (${rows.length})`, asset_name: "", model: "", fleet: "", uom: "" },
               { bold: true, fill: "FFE2EFE9", topBorder: true });
      r += 1;
    }

    // --- notes: the things a reader would otherwise have to ask about ---
    const notes = [
      "Subtotal and Total rows SUM hours, downtime and stoppage counts; they AVERAGE rates (MTBF, MTTR, availability, utilisation, breakdown %). Machines with no data are excluded from averages rather than counted as zero.",
      "Engineering Availability and Engineering MTBF count downtime whose Responsibility is Plant or Stores. Events recorded before Responsibility was introduced appear under 'Not set' and are excluded from those two figures.",
      "Last 24 Hours columns cover the 24 hours up to the export time and are independent of the date range above.",
    ];
    if (conflicts) notes.push(`${conflicts} shaded row(s): worked hours plus downtime exceed the hours available in the period, so the Daily Hours and downtime records contradict each other. Utilisation is capped at 100% on those rows and the true figure is higher.`);
    r += 1;
    for (const n of notes) {
      ws.mergeCells(r, 1, r, Math.min(N, 12));
      const c = ws.getCell(r, 1);
      c.value = n;
      c.font = { name: "Arial", size: 8.5, italic: true, color: { argb: "FF859195" } };
      c.alignment = { wrapText: true, vertical: "top" };
      ws.getRow(r).height = 24;
      r += 1;
    }

    ws.columns = COLS.map(([key, , , fmtType, shortLabel]) => ({
      width: key === "asset_id" || key === "asset_name" ? 16
           : fmtType === "text" ? 9
           : Math.min(Math.max((shortLabel || "").length + 3, 8), 11),
    }));
    ws.autoFilter = { from: { row: HEADER_ROW, column: 1 }, to: { row: HEADER_ROW, column: N } };

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `Plant_Performance_KPI_${now.toISOString().slice(0, 10)}.xlsx`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  return (
    <div>
      <DateRangePicker fromDateTime={fromDateTime} toDateTime={toDateTime} setFromDateTime={setFromDateTime} setToDateTime={setToDateTime} />
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />

      {!loading && rows.length > 0 && (() => {
        const mtbfTarget = 40, mttrTarget = 4, utilTarget = 85, availTarget = 85;
        // No fleet/equipment picked -> chart by fleet (5 readable bars).
        // Fleet picked -> chart by individual machine within that fleet.
        // Equipment picked -> rows is just that one machine either way.
        const showByFleet = !selectedFleet && !selectedAsset;
        const chartData = showByFleet ? fleetKpiData : rows;
        const chartXKey = showByFleet ? "fleet" : "asset_id";
        const mtbfMax = Math.max(mtbfTarget, ...chartData.map((r) => r.mtbf || 0)) * 1.15;
        const mttrMax = Math.max(mttrTarget, ...chartData.map((r) => r.mttr || 0)) * 1.15;
        // A Total group on the right, aggregated across everything on
        // screen - the same shape as the Excel report.
        const totalRow = chartData.length > 1
          ? { [chartXKey]: "Total", ...aggregateMetrics(rows, rows.map((r) => r.asset_id)) }
          : null;
        const withTotal = totalRow ? [...chartData, totalRow] : chartData;

        const pct = (v) => (v != null ? Math.round(v * 100) : null);
        const chartRows = withTotal.map((r) => ({
          ...r,
          utilisationPct: pct(r.utilisation),
          utilisationPct24h: pct(r.utilisation_24h),
          availabilityPct: pct(r.availability),
          availabilityPct24h: pct(r.availability_24h),
        }));

        const periodLabel = "Selected period";
        const handleBarClick = (row) => {
          if (row.fleet === "Total" || row.asset_id === "Total") return;
          if (showByFleet) setSelectedFleet(row.fleet);
          else if (!selectedAsset) setSelectedAsset(row.asset_id);
        };
        return (
        <div style={{ marginBottom: 8 }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10, margin: "0 0 12px" }}>
            <p style={{ fontSize: 13, color: "#4B5659", margin: 0 }}>
              {showByFleet ? "Click a fleet to see individual machines." : `Showing ${selectedAsset || selectedFleet}.`}
            </p>
            {!showByFleet && (
              <button
                onClick={() => { setSelectedFleet(""); setSelectedAsset(""); }}
                style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: NAVY, fontSize: 12.5, fontWeight: 600, textDecoration: "underline" }}
              >
                ← back to all fleets
              </button>
            )}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(280px,100%),1fr))", gap: 16, marginBottom: 8 }}>
            <KpiBarChart
              title={showByFleet ? "MTBF by fleet (hrs)" : "MTBF by equipment (hrs)"} data={chartRows} xKey={chartXKey} dataKey="mtbf"
              target={mtbfTarget} domainMax={mtbfMax} unitSuffix="h"
              valueFormatter={(v) => Number(v).toFixed(1)}
              meetsTarget={(r) => (r.mtbf || 0) >= mtbfTarget}
              onBarClick={handleBarClick}
            />
            <KpiBarChart
              title={showByFleet ? "MTTR by fleet (hrs)" : "MTTR by equipment (hrs)"} data={chartRows} xKey={chartXKey} dataKey="mttr"
              target={mttrTarget} domainMax={mttrMax} unitSuffix="h"
              valueFormatter={(v) => Number(v).toFixed(1)}
              meetsTarget={(r) => (r.mttr || 0) <= mttrTarget}
              onBarClick={handleBarClick}
            />
            <KpiBarChart
              title={showByFleet ? "Utilisation by fleet" : "Utilisation by equipment"} data={chartRows} xKey={chartXKey} dataKey="utilisationPct" dataKey2="utilisationPct24h" seriesName={periodLabel} seriesName2="Last 24h" meetsTarget2={(r) => (r.utilisationPct24h || 0) >= utilTarget}
              target={utilTarget} domainMax={100} unitSuffix="%"
              valueFormatter={(v) => `${v}%`}
              meetsTarget={(r) => (r.utilisationPct || 0) >= utilTarget}
              onBarClick={handleBarClick}
            />
            <KpiBarChart
              title={showByFleet ? "Availability by fleet" : "Availability by equipment"} data={chartRows} xKey={chartXKey} dataKey="availabilityPct" dataKey2="availabilityPct24h" seriesName={periodLabel} seriesName2="Last 24h" meetsTarget2={(r) => (r.availabilityPct24h || 0) >= availTarget}
              target={availTarget} domainMax={100} unitSuffix="%"
              valueFormatter={(v) => `${v}%`}
              meetsTarget={(r) => (r.availabilityPct || 0) >= availTarget}
              onBarClick={handleBarClick}
            />
          </div>
          <KpiLegend targetLabel="Target" showSeries />
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
      <>
        <p style={{ fontSize: 11.5, color: "#859195", margin: "0 0 6px" }}>
          Scroll sideways inside the table to reach every column — the Equipment column and the headings stay in place.
        </p>
        <div style={{ overflow: "auto", maxHeight: "62vh", border: "1px solid #E2E6E3", borderRadius: 10, position: "relative" }}>
          {(() => {
            // Group banner spans: consecutive columns sharing a group name.
            const spans = [];
            KPI_REPORT_COLUMNS.forEach(([, , group]) => {
              const last = spans[spans.length - 1];
              if (last && last.group === group) last.span += 1;
              else spans.push({ group, span: 1 });
            });

            // Rows grouped by fleet, each fleet followed by a subtotal,
            // then a site total - the shape of the reference report.
            const byFleet = new Map();
            rows.forEach((r) => {
              const f = r.fleet || "(no fleet)";
              if (!byFleet.has(f)) byFleet.set(f, []);
              byFleet.get(f).push(r);
            });
            const anyConflict = rows.some(hasHoursConflict);

            // Sized to fit all 34 columns on a desktop screen: identity
            // columns get real room, numeric ones only what a figure needs.
            // Wide enough that nothing gets clipped. The table scrolls
            // sideways in its own box, so readability wins over fitting -
            // a truncated equipment number helps nobody.
            const widthFor = (key, fmtType) => {
              if (key === "asset_id") return 112;
              if (key === "asset_name") return 124;
              if (key === "fleet") return 94;
              if (key === "model") return 64;
              if (fmtType === "text") return 64;
              if (fmtType === "int") return 48;
              if (fmtType === "pct") return 54;
              return 64;   // fits a 5-digit subtotal like 12345.67
            };

            // Group banners get short forms too - "Equipment Availability"
            // across two 54px columns would clip no matter the font.
            const SHORT_GROUP = {
              "Downtime Responsibility": "Downtime Responsibility",
              "Number of Stoppages": "Stoppages",
              "Equipment Availability": "Equip. Availability",
              "Engineering Availability": "Eng. Availability",
              "Equipment Utilisation": "Equip. Utilisation",
            };
            // Equipment stays pinned while the rest scrolls sideways -
            // 34 columns will not fit a laptop screen at a readable size,
            // so the column you need for orientation never leaves view.
            const stickyFirst = (extra = {}) => ({
              position: "sticky", left: 0, zIndex: 2, ...extra,
            });
            const cellStyle = (fmtType) => ({
              padding: "6px 7px", whiteSpace: "nowrap",
              textAlign: fmtType === "text" ? "left" : "right",
            });

            return (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
                <colgroup>
                  {KPI_REPORT_COLUMNS.map(([key, , , fmtType]) => (
                    <col key={key} style={{ width: widthFor(key, fmtType) }} />
                  ))}
                </colgroup>
                <thead>
                  <tr style={{ background: "#E9ECEA" }}>
                    {spans.map((sp, i) => (
                      <th key={i} colSpan={sp.span} title={sp.group}
                        style={{ position: "sticky", top: 0, zIndex: 3, background: "#E9ECEA",
                                 textAlign: "center", padding: "5px 6px", fontWeight: 700, fontSize: 11, color: NAVY,
                                 borderBottom: "1px solid #E2E6E3", borderLeft: i > 0 ? "1px solid #DCE0DD" : "none", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                        {SHORT_GROUP[sp.group] || sp.group}
                      </th>
                    ))}
                  </tr>
                  <tr style={{ background: "#F7F8F6" }}>
                    {KPI_REPORT_COLUMNS.map(([key, label, , fmtType, shortLabel]) => (
                      <th key={key} onClick={() => handleSort(key)} title={label}
                        style={{ position: "sticky", top: 26, zIndex: key === "asset_id" ? 5 : 3, background: "#F7F8F6",
                                 ...(key === "asset_id" ? { left: 0 } : {}),
                                 textAlign: fmtType === "text" ? "left" : "right", padding: "6px 7px", fontWeight: 600, color: "#4B5659",
                                 fontSize: 11, lineHeight: 1.25, borderBottom: "1px solid #E2E6E3", cursor: "pointer", userSelect: "none",
                                 whiteSpace: "normal", overflowWrap: "normal", hyphens: "none" }}>
                        {shortLabel || label}{sortKey === key ? (sortDir === "asc" ? " ▲" : " ▼") : ""}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...byFleet.entries()].map(([fleetName, fleetRows]) => {
                    const sub = kpiAggregate(fleetRows);
                    return (
                      <React.Fragment key={fleetName}>
                        {fleetRows.map((row) => {
                          const conflict = hasHoursConflict(row);
                          return (
                            <tr key={row.asset_id}
                              onClick={() => setEventsDrill(row)}
                              title={conflict
                                ? "Worked hours plus downtime exceed the hours in this period - the Daily Hours and downtime records contradict each other. Click to see the events."
                                : "Click to see the events behind these figures"}
                              style={{ borderBottom: "1px solid #EFEEE7", background: conflict ? "#FDF3E3" : "transparent", cursor: "pointer" }}
                              onMouseEnter={(e) => { if (!conflict) e.currentTarget.style.background = "#F7F8F6"; }}
                              onMouseLeave={(e) => { if (!conflict) e.currentTarget.style.background = "transparent"; }}>
                              {KPI_REPORT_COLUMNS.map(([key, , , fmtType]) => (
                                <td key={key} title={fmtType === "text" ? String(row[key] ?? "") : undefined}
                                  style={{ ...cellStyle(fmtType), overflow: "hidden", textOverflow: "ellipsis",
                                  ...(key === "asset_id" ? stickyFirst({ background: conflict ? "#FDF3E3" : "#FFFFFF" }) : {}),
                                  fontWeight: key === "num_unplanned_events" && row[key] >= 3 ? 700 : 400,
                                  color: key === "num_unplanned_events" && row[key] >= 3 ? "#B85450" : "#183642" }}>
                                  {key === "asset_id" && conflict ? <span style={{ color: "#B07D2B", fontWeight: 700 }} title="Hours conflict">⚠ </span> : null}
                                  {fmt(key, row[key])}
                                </td>
                              ))}
                            </tr>
                          );
                        })}
                        <tr style={{ background: "#F2F1EA", borderBottom: "1px solid #E2E6E3" }}>
                          {KPI_REPORT_COLUMNS.map(([key, , , fmtType], ci) => (
                            <td key={key} style={{ ...cellStyle(fmtType), overflow: "hidden", textOverflow: "ellipsis", fontWeight: 700, color: NAVY,
                              ...(key === "asset_id" ? stickyFirst({ background: "#F2F1EA" }) : {}) }}>
                              {ci === 0 ? `${fleetName} (${fleetRows.length})` : (KPI_SUM_KEYS.has(key) || KPI_AVG_KEYS.has(key) ? fmt(key, sub[key]) : "")}
                            </td>
                          ))}
                        </tr>
                      </React.Fragment>
                    );
                  })}
                  {rows.length > 0 && (() => {
                    const tot = kpiAggregate(rows);
                    return (
                      <tr style={{ background: "#E2EFE9", borderTop: `2px solid ${NAVY}` }}>
                        {KPI_REPORT_COLUMNS.map(([key, , , fmtType], ci) => (
                          <td key={key} style={{ ...cellStyle(fmtType), overflow: "hidden", textOverflow: "ellipsis", fontWeight: 800, color: NAVY,
                            ...(key === "asset_id" ? stickyFirst({ background: "#E2EFE9" }) : {}) }}>
                            {ci === 0 ? `Total (${rows.length})` : (KPI_SUM_KEYS.has(key) || KPI_AVG_KEYS.has(key) ? fmt(key, tot[key]) : "")}
                          </td>
                        ))}
                      </tr>
                    );
                  })()}
                  {rows.length === 0 && (
                    <tr><td colSpan={KPI_REPORT_COLUMNS.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No data for this date range.</td></tr>
                  )}
                  {anyConflict && (
                    <tr><td colSpan={KPI_REPORT_COLUMNS.length} style={{ padding: "8px 10px", fontSize: 11.5, color: "#8A6320", background: "#FDF3E3" }}>
                      ⚠ Shaded rows: worked hours plus downtime exceed the hours available in this period, so the Daily Hours and downtime records contradict each other. Utilisation is capped at 100% on those rows and the true figure is higher.
                    </td></tr>
                  )}
                </tbody>
              </table>
            );
          })()}
        </div>
      </>
      )}

      {eventsDrill && (
        <KpiEventsDrilldown
          machine={eventsDrill}
          fromDateTime={fromDateTime}
          toDateTime={toDateTime}
          onClose={() => setEventsDrill(null)}
        />
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
function truncateLabel(s, max) {
  if (!s || s.length <= max) return s;
  return s.slice(0, max - 1).trimEnd() + "…";
}

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

    return [...actual, ...scheduled].sort((a, b) => a.offsetHrs - b.offsetHrs)
      // Y-axis labels render as one unbroken line of SVG text with no
      // wrapping - a long "asset - problem scope" label overflows its
      // row and overlaps the rows above/below it, worst on a narrow
      // phone-width chart. Truncate what's shown on the axis; the full
      // label is still available via the tooltip and is already shown
      // in full in the list above this chart.
      .map((e) => ({ ...e, shortLabel: truncateLabel(e.label, narrow ? 20 : 34) }));
  }, [breakdowns, workOrders, statusFilter, nowTick, windowStart, windowEnd, windowHrs, narrow]);

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
              <YAxis type="category" dataKey="shortLabel" width={narrow ? 110 : 220} tick={{ fontSize: narrow ? 9.5 : 11, fill: "#183642" }} axisLine={false} tickLine={false} />
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
      wb.creator = "Mine2U";
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
              .print-footer { display: block !important; position: fixed; bottom: 8px; left: 0; right: 0; text-align: center; font-size: 10px; color: #666; }
            }
            .print-footer { display: none; }
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

              <div style={{ overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
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
              </div>

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

              <p className="print-footer">Printed: {new Date().toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" })}</p>
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
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [groupByMachine, setGroupByMachine] = useState(true);

  const columns = [
    ["wo_no", "Work order #"], ["asset_id", "Equipment #"], ["work_type", "Type"],
    ["priority", "Priority"], ["status", "Status"], ["request_date", "Requested"],
  ].map(([key, label]) => ({ key, label }));

  const filtered = useMemo(() => {
    let rows = workOrders;
    if (selectedAsset) rows = rows.filter((r) => r.asset_id === selectedAsset);
    else if (selectedFleet) rows = rows.filter((r) => { const a = assets.find((x) => x.asset_id === r.asset_id); return a && a.fleet === selectedFleet; });
    if (query.trim()) {
      const q = query.toLowerCase();
      rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return [...rows].sort((a, b) => (b.request_date || "").localeCompare(a.request_date || ""));
  }, [workOrders, assets, query, selectedFleet, selectedAsset]);

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "request_date", hoursKey: null,
    sortKey: null, sortDir: null, sortRows: (r) => r,
  }), [filtered, assets, groupByMachine]);

  const handleSaved = () => { setShowForm(false); setEditing(null); onRefresh(); };

  const handleDelete = async (reason) => {
    await deleteWithReason("work_orders", deleting.id, "id", reason, userEmail, deleting.asset_id);
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
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, gap: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search work orders"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
          <button onClick={() => { setEditing(null); setShowForm(true); }} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
            + Add Work Order
          </button>
        </div>
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
            {flattenGroups(groups).map((item, i) => item.kind === "banner" ? (
              <MachineBannerRow key={`b-${item.group.key}`} group={item.group} colSpan={columns.length + 1} first={i === 0} countNoun="work order" countNounPlural="work orders" />
            ) : (
              (({ row }) => (
              <tr key={row.id ?? i} style={{ borderBottom: "1px solid #EFEEE7" }}>
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
              ))(item)
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
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [groupByMachine, setGroupByMachine] = useState(true);

  const plannedJobs = useMemo(
    () => [...workOrders]
      .filter((w) => w.work_type === "Preventive")
      .filter((w) => {
        if (selectedAsset) return w.asset_id === selectedAsset;
        if (selectedFleet) { const a = assets.find((x) => x.asset_id === w.asset_id); return a && a.fleet === selectedFleet; }
        return true;
      })
      .sort((a, b) => (b.request_date || "").localeCompare(a.request_date || "")),
    [workOrders, assets, selectedFleet, selectedAsset]
  );

  const jobGroups = useMemo(() => buildMachineGroups(plannedJobs, assets, {
    enabled: groupByMachine, dateKey: "request_date", hoursKey: null,
    sortKey: null, sortDir: null, sortRows: (r) => r,
  }), [plannedJobs, assets, groupByMachine]);

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
    await deleteWithReason("work_orders", deletingJob.id, "id", reason, userEmail, deletingJob.asset_id);
    onRefresh();
    setDeletingJob(null);
  };

  return (
    <div>
      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />
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
        <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
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
            {flattenGroups(jobGroups).map((item, i) => item.kind === "banner" ? (
              <MachineBannerRow key={`b-${item.group.key}`} group={item.group} colSpan={7} first={i === 0} countNoun="job" countNounPlural="jobs" />
            ) : (
              (({ row }) => (
              <tr key={row.id ?? i} style={{ borderBottom: "1px solid #EFEEE7" }}>
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
              ))(item)
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

function AssetForm({ existing, selectedSiteId, isAdmin, mySites = [], onClose, onSaved }) {
  const isEdit = !!existing;
  const [assetId, setAssetId] = useState(existing?.asset_id || "");
  const [assetName, setAssetName] = useState(existing?.asset_name || "");
  const [make, setMake] = useState(existing?.make || "");
  const [model, setModel] = useState(existing?.model || "");
  const [fleet, setFleet] = useState(existing?.fleet || "");
  const [serialNumber, setSerialNumber] = useState(existing?.serial_number || "");
  const [status, setStatus] = useState(existing?.status || "Operating");
  const [siteLocation, setSiteLocation] = useState(existing?.site_location || "");
  const [siteId, setSiteId] = useState(existing?.site_id || selectedSiteId || "");
  const [openingHours, setOpeningHours] = useState(existing?.opening_hours ?? "");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (existing?.asset_id) {
      logActivity("Assets", existing.asset_id, "viewed", `Opened equipment record ${existing.asset_id} - ${existing.asset_name || ""}`.trim(), existing.asset_id);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSaving(true);

    const payload = {
      asset_id: assetId, asset_name: assetName, make, model, fleet,
      serial_number: serialNumber || null, status, site_location: siteLocation || null,
      opening_hours: openingHours === "" ? null : Number(openingHours),
      notes: notes || null,
      site_id: isAdmin ? (siteId || null) : (isEdit ? undefined : selectedSiteId),
    };
    if (payload.site_id === undefined) delete payload.site_id;

    try {
      const { error: dbError } = isEdit
        ? await supabase.from("assets").update(payload).eq("asset_id", existing.asset_id)
        : await supabase.from("assets").insert(payload);
      if (dbError) throw dbError;
      logActivity("Assets", assetId, isEdit ? "updated" : "created", `${assetId} - ${assetName}${status !== "Operating" ? ` (${status})` : ""}`, assetId);
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
            <label style={labelStyle}>Location (e.g. Yard, Pit)</label>
            <input type="text" value={siteLocation} onChange={(e) => setSiteLocation(e.target.value)} style={fieldStyle} />
          </div>
        </div>

        {isAdmin && mySites.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <label style={labelStyle}>Site</label>
            <select value={siteId} onChange={(e) => setSiteId(e.target.value)} style={fieldStyle}>
              {mySites.map((s) => <option key={s.id} value={s.id}>{s.site_name}</option>)}
            </select>
            <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
              Which site this equipment belongs to - only that site's users will see it. Change this to move equipment between sites.
            </p>
          </div>
        )}

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
      const summary = `${partNo} - ${description} (qty: ${qtyInStock}${supplier ? `, supplier: ${supplier}` : ""})`;
      if (isEdit) {
        const { error: dbError } = await supabase.from("parts_inventory").update(payload).eq("id", existing.id);
        if (dbError) throw dbError;
        logActivity("Parts Inventory", existing.id, "updated", summary);
      } else {
        const { data, error: dbError } = await supabase.from("parts_inventory").insert(payload).select().single();
        if (dbError) throw dbError;
        logActivity("Parts Inventory", data?.id, "created", summary);
      }
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
  const [editingSite, setEditingSite] = useState(null);
  const [editSiteName, setEditSiteName] = useState("");
  const [editSiteLocation, setEditSiteLocation] = useState("");
  const [editSiteActive, setEditSiteActive] = useState(true);
  const [savingEditSite, setSavingEditSite] = useState(false);

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

  const openEditSite = (site) => {
    setEditingSite(site);
    setEditSiteName(site.site_name || "");
    setEditSiteLocation(site.location || "");
    setEditSiteActive(site.active !== false);
  };

  const handleSaveSite = async (e) => {
    e.preventDefault();
    if (!editSiteName.trim()) return;
    setSavingEditSite(true);
    const { error } = await supabase.from("sites").update({
      site_name: editSiteName.trim(),
      location: editSiteLocation.trim() || null,
      active: editSiteActive,
    }).eq("id", editingSite.id);
    setSavingEditSite(false);
    if (error) {
      setMessage({ type: "error", text: error.message });
    } else {
      setEditingSite(null);
      setMessage({ type: "success", text: `"${editSiteName.trim()}" updated.` });
      loadAll();
      onSitesChanged?.();
    }
  };

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
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {sites.map((s, i) => (
              <tr key={s.id} style={{ borderBottom: i < sites.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px", fontWeight: 600 }}>{s.site_name}</td>
                <td style={{ padding: "9px 12px" }}>{s.location || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}><Badge value={s.active ? "Active" : "Inactive"} /></td>
                <td style={{ padding: "9px 12px" }}>
                  <button type="button" onClick={() => openEditSite(s)} style={{ background: "none", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12, fontWeight: 600, padding: "5px 10px", borderRadius: 6, cursor: "pointer" }}>
                    Edit
                  </button>
                </td>
              </tr>
            ))}
            {sites.length === 0 && !loading && (
              <tr><td colSpan={4} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No sites yet - add one above.</td></tr>
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

      {editingSite && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60, padding: 16 }}>
          <form onSubmit={handleSaveSite} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 380, maxWidth: "100%" }}>
            <h3 style={{ fontSize: 16, fontWeight: 700, color: NAVY, margin: "0 0 16px" }}>Edit Site</h3>

            <label style={{ display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" }}>Site Name</label>
            <input
              type="text" value={editSiteName} onChange={(e) => setEditSiteName(e.target.value)} required
              style={{ width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", marginBottom: 12 }}
            />

            <label style={{ display: "block", fontSize: 12.5, fontWeight: 600, color: "#183642", margin: "0 0 4px" }}>Location</label>
            <input
              type="text" value={editSiteLocation} onChange={(e) => setEditSiteLocation(e.target.value)} placeholder="Optional"
              style={{ width: "100%", padding: "8px 10px", fontSize: 13.5, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", marginBottom: 12 }}
            />

            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#183642", cursor: "pointer", marginBottom: 16 }}>
              <input type="checkbox" checked={editSiteActive} onChange={(e) => setEditSiteActive(e.target.checked)} />
              Site is active
            </label>
            <p style={{ fontSize: 11.5, color: "#859195", margin: "-10px 0 16px" }}>
              Marking a site inactive doesn't delete anything or remove anyone's access - it's just a status flag so you can tell which sites are actually operating at a glance.
            </p>

            {message?.type === "error" && (
              <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 12px" }}>{message.text}</p>
            )}

            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button type="button" onClick={() => setEditingSite(null)} disabled={savingEditSite} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, cursor: "pointer" }}>
                Cancel
              </button>
              <button type="submit" disabled={savingEditSite} style={{ background: NAVY, border: "none", color: "#fff", padding: "9px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: savingEditSite ? "default" : "pointer", opacity: savingEditSite ? 0.7 : 1 }}>
                {savingEditSite ? "Saving…" : "Save Changes"}
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}

function PageAccessModal({ user, onClose }) {
  const tickablePages = NAV
    .filter((n) => n.key !== "audit" && n.key !== "site_management" && n.key !== "component_codes")
    .map((n) => n.group === "backlogs" ? { ...n, label: `Backlogs - ${n.label}` } : n);
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

const ACTIVITY_COLUMNS = [
  ["created_at", "Date & Time"], ["table_name", "Tab"], ["asset_id", "Machine #"],
  ["action", "Action"], ["summary", "Details"], ["user_name", "User"],
];

// Groups an already-sorted (newest first) list of rows into buckets keyed
// by calendar day, so the table can show a day heading instead of one long
// undifferentiated list - this is what keeps a much busier audit trail
// (now that views and part confirmations are logged too, not just
// creates/edits) readable rather than an overwhelming wall of rows.
function groupByDay(rows) {
  const groups = [];
  let current = null;
  for (const r of rows) {
    const dayKey = r.created_at ? new Date(r.created_at).toDateString() : "Unknown date";
    if (!current || current.dayKey !== dayKey) {
      current = { dayKey, label: r.created_at ? new Date(r.created_at).toLocaleDateString("en-ZA", { weekday: "long", day: "2-digit", month: "long", year: "numeric" }) : "Unknown date", rows: [] };
      groups.push(current);
    }
    current.rows.push(r);
  }
  return groups;
}

function AuditTrailPage({ activityLog, profiles, isAdmin }) {
  const [tableFilter, setTableFilter] = useState("");
  const [actionFilter, setActionFilter] = useState("");
  const [userFilter, setUserFilter] = useState("");
  const [assetFilter, setAssetFilter] = useState("");
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
  const machines = useMemo(() => [...new Set(activityLog.map((r) => r.asset_id))].filter(Boolean).sort(), [activityLog]);

  const filtered = useMemo(() => {
    return activityLog.filter((r) => {
      if (tableFilter && r.table_name !== tableFilter) return false;
      if (actionFilter && r.action !== actionFilter) return false;
      if (userFilter && nameByUser.get(r.user_id) !== userFilter) return false;
      if (assetFilter && r.asset_id !== assetFilter) return false;
      if (fromDate && new Date(r.created_at) < new Date(fromDate)) return false;
      if (toDate && new Date(r.created_at) > new Date(toDate + "T23:59:59")) return false;
      return true;
    });
  }, [activityLog, tableFilter, actionFilter, userFilter, assetFilter, fromDate, toDate, nameByUser]);

  const dayGroups = useMemo(() => groupByDay(filtered), [filtered]);

  const formatDT = (v) => v ? new Date(v).toLocaleString("en-ZA", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "-";
  const nowPrinted = () => new Date().toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });

  const exportToExcel = () => {
    const now = new Date();
    const timestamp = now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" });
    const headerRow = ACTIVITY_COLUMNS.map((c) => c[1]);
    const dataRows = filtered.map((r) => [
      formatDT(r.created_at), r.table_name, r.asset_id || "-", r.action, r.summary || "", nameByUser.get(r.user_id) || "-",
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
  const dayHeaderStyle = { padding: "8px 12px", fontWeight: 700, fontSize: 12, color: NAVY, background: "#EFF3F2", textTransform: "uppercase", letterSpacing: 0.3 };

  return (
    <div>
      <p style={{ fontSize: 12.5, color: "#859195", margin: "0 0 14px" }}>
        Covers every create, edit, close and delete across Events, Work Orders, Backlogs, Daily Hours, Fuel Log, Oil Consumption, Inspections, Assets, Parts Inventory and the Quote Price List - plus every time someone opens an existing record to view it, and every part pulled or confirmed against an Event or Work Order. Filter by tab, action, user, machine or date range below, or export the current view to Excel.
      </p>

      <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
        <select value={tableFilter} onChange={(e) => setTableFilter(e.target.value)} style={selectStyle}>
          <option value="">All tabs</option>
          {tables.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
        <select value={actionFilter} onChange={(e) => setActionFilter(e.target.value)} style={selectStyle}>
          <option value="">All actions</option>
          <option value="viewed">Viewed</option>
          <option value="created">Created</option>
          <option value="updated">Updated</option>
          <option value="closed">Closed</option>
          <option value="deleted">Deleted</option>
        </select>
        <select value={userFilter} onChange={(e) => setUserFilter(e.target.value)} style={selectStyle}>
          <option value="">All users</option>
          {users.map((u) => <option key={u} value={u}>{u}</option>)}
        </select>
        <select value={assetFilter} onChange={(e) => setAssetFilter(e.target.value)} style={selectStyle}>
          <option value="">All machines</option>
          {machines.map((m) => <option key={m} value={m}>{m}</option>)}
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

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {ACTIVITY_COLUMNS.map(([key, label]) => (
                <th key={key} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={ACTIVITY_COLUMNS.length} style={{ padding: 24, textAlign: "center", color: "#859195" }}>No audit entries match your filters.</td></tr>
            ) : dayGroups.map((group) => (
              <React.Fragment key={group.dayKey}>
                <tr>
                  <td colSpan={ACTIVITY_COLUMNS.length} style={dayHeaderStyle}>{group.label}</td>
                </tr>
                {group.rows.map((r, i) => (
                  <tr key={r.id ?? i} style={{ borderBottom: i < group.rows.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{formatDT(r.created_at)}</td>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.table_name || "-"}</td>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{r.asset_id || "-"}</td>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap", textTransform: "capitalize" }}>{r.action || "-"}</td>
                    <td style={{ padding: "9px 12px" }}>{r.summary || "-"}</td>
                    <td style={{ padding: "9px 12px", whiteSpace: "nowrap", fontWeight: 600, color: "#183642" }}>{nameByUser.get(r.user_id) || "Unknown"}</td>
                  </tr>
                ))}
              </React.Fragment>
            ))}
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
              .print-footer { display: block !important; position: fixed; bottom: 8px; left: 0; right: 0; text-align: center; font-size: 10px; color: #666; }
            }
            .print-footer { display: none; }
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
                {fromDate || toDate ? `${fromDate || "earliest"} to ${toDate || "latest"}` : "All recorded activity"} · Generated {nowPrinted()}
              </p>

              {dayGroups.map((group) => (
                <div key={group.dayKey} style={{ marginBottom: 14 }}>
                  <p style={{ fontSize: 12.5, fontWeight: 700, color: NAVY, margin: "0 0 6px", textTransform: "uppercase", letterSpacing: 0.3 }}>{group.label}</p>
                  <div style={{ overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 6, fontSize: 11.5 }}>
                      <thead>
                        <tr>
                          {ACTIVITY_COLUMNS.map(([key, label]) => (
                            <th key={key} style={{ textAlign: "left", padding: "6px 8px", border: "1px solid #ccc", background: "#F0F3F8" }}>{label}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {group.rows.map((r, i) => (
                          <tr key={r.id ?? i}>
                            <td style={{ padding: "6px 8px", border: "1px solid #ccc", whiteSpace: "nowrap" }}>{formatDT(r.created_at)}</td>
                            <td style={{ padding: "6px 8px", border: "1px solid #ccc", whiteSpace: "nowrap" }}>{r.table_name || "-"}</td>
                            <td style={{ padding: "6px 8px", border: "1px solid #ccc", whiteSpace: "nowrap" }}>{r.asset_id || "-"}</td>
                            <td style={{ padding: "6px 8px", border: "1px solid #ccc", whiteSpace: "nowrap", textTransform: "capitalize" }}>{r.action || "-"}</td>
                            <td style={{ padding: "6px 8px", border: "1px solid #ccc" }}>{r.summary || "-"}</td>
                            <td style={{ padding: "6px 8px", border: "1px solid #ccc", whiteSpace: "nowrap", fontWeight: 700 }}>{nameByUser.get(r.user_id) || "Unknown"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}

              <p className="print-footer">Printed: {nowPrinted()}</p>
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

function PartsPage({ parts, selectedSiteId, onRefresh, userEmail, isAdmin }) {
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
        <QuotePriceListPage selectedSiteId={selectedSiteId} parts={parts} userEmail={userEmail} isAdmin={isAdmin} />
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
        Export the current list, edit it in Excel (add rows, change quantities), then Upload the same file - existing Part Nr get updated, new ones get added. Nothing gets deleted by an upload.
      </p>

      <div style={{ overflowX: "auto", border: "1px solid #E2E6E3", borderRadius: 10, WebkitOverflowScrolling: "touch" }}>
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
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.part_no}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.description}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.qty_in_stock}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.minimum_qty}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}><Badge value={row.reorder_status} /></td>
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
      <p style={{ fontSize: 11.5, color: "#859195", margin: "8px 2px 0" }}>Swipe the table sideways to see every column, including Status.</p>

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

function QuotePriceListPage({ selectedSiteId, parts, userEmail, isAdmin }) {
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
  const [deleting, setDeleting] = useState(null); // row pending delete confirmation
  const [showDeletedPanel, setShowDeletedPanel] = useState(false);
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
        currency: it.currency || "",
        price: it.price ?? "",
      }));
      if (items.length === 0) {
        setExtractError("No part/price lines were found on this document - you can still add rows manually below, or try a clearer scan.");
      }
      setReview({ file, items: items.length ? items : [{ supplier: "", part_no: "", part_description: "", currency: "", price: "" }], detectedLanguage: data.detectedLanguage || null });
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
    setReview((r) => ({ ...r, items: [...r.items, { supplier: "", part_no: "", part_description: "", currency: "", price: "" }] }));
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
          currency: it.currency || null,
          price: it.price === "" ? null : Number(it.price),
          price_date: today,
          quote_document_path: quoteDocumentPath,
        }));
      if (payload.length === 0) { setSaving(false); return; }
      const { error: insertErr } = await supabase.from("quote_prices").insert(payload);
      if (insertErr) throw insertErr;
      const suppliers = [...new Set(payload.map((p) => p.supplier).filter(Boolean))].join(", ") || "Unknown supplier";
      const partList = payload.map((p) => p.part_description || p.part_no || "item").join(", ");
      logActivity("Quote Price List", `quote-${Date.now()}`, "created", `Loaded quote from ${suppliers} - ${payload.length} line item${payload.length !== 1 ? "s" : ""}: ${partList}`);
      setReview(null);
      load();
    } catch (err) {
      setExtractError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  // Backs up the full quote row (not just its ID) into quote_prices_deleted
  // before removing it - so an admin can restore it later if the user who
  // deleted it comes back asking for it. deleteWithReason still handles the
  // reason + Audit Trail entry exactly as it does everywhere else in the app.
  const handleDeleteQuote = async (reason) => {
    const row = deleting;
    if (!row) return;
    const { id, ...rest } = row;
    const { error: backupErr } = await supabase.from("quote_prices_deleted").insert({
      ...rest,
      original_id: id,
      deleted_by_email: userEmail || null,
      deletion_reason: reason,
    });
    if (backupErr) throw backupErr;
    await deleteWithReason("quote_prices", row.id, "id", reason, userEmail);
    setDeleting(null);
    load();
  };

  const columns = [
    ["supplier", "Supplier"], ["part_description", "Part"], ["part_no", "Part Number"],
    ["currency", "Currency"], ["price", "Price"], ["price_date", "Last updated"],
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
          Image or PDF, any language - Mine2U will read it, translate part descriptions into English automatically, detect the currency, and pull out the Supplier, Part, Part Number and Price for you to review before saving.
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
          <p style={{ fontSize: 12, color: "#859195", margin: "0 0 14px" }}>Check what Mine2U read off the quote - fix anything that's wrong, remove lines that shouldn't be there, then save. Lines flagged "In stock" already have some on hand - worth checking before ordering more.</p>
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
                  {["Supplier", "Part", "Part Number", "Currency", "Price", "Stock", ""].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: "7px 8px", fontWeight: 600, color: "#4B5659", fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {review.items.map((row, i) => {
                  const stock = stockFor(row.part_no);
                  return (
                  <tr key={i}>
                    {["supplier", "part_description", "part_no"].map((field) => (
                      <td key={field} style={{ padding: "4px 8px" }}>
                        <input
                          type="text"
                          value={row[field]}
                          onChange={(e) => updateReviewRow(i, field, e.target.value)}
                          style={{ width: "100%", padding: "6px 8px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6, boxSizing: "border-box" }}
                        />
                      </td>
                    ))}
                    <td style={{ padding: "4px 8px" }}>
                      <input
                        type="text"
                        value={row.currency}
                        onChange={(e) => updateReviewRow(i, "currency", e.target.value.toUpperCase())}
                        placeholder="e.g. ZAR"
                        maxLength={6}
                        style={{ width: 64, padding: "6px 8px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6, boxSizing: "border-box", textTransform: "uppercase" }}
                      />
                    </td>
                    <td style={{ padding: "4px 8px" }}>
                      <input
                        type="number"
                        value={row.price}
                        onChange={(e) => updateReviewRow(i, "price", e.target.value)}
                        style={{ width: "100%", padding: "6px 8px", fontSize: 12.5, border: "1px solid #E2E6E3", borderRadius: 6, boxSizing: "border-box" }}
                      />
                    </td>
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
              <th style={{ borderBottom: "1px solid #E2E6E3" }}></th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan={columns.length + 3} style={{ padding: 20, textAlign: "center", color: "#859195" }}>Loading…</td></tr>
            ) : filtered.length === 0 ? (
              <tr><td colSpan={columns.length + 3} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                {rows.length === 0 ? "No quotes added yet - drag one in above to get started." : "No entries match your filters."}
              </td></tr>
            ) : filtered.map((row, i) => {
              const stock = stockFor(row.part_no);
              return (
              <tr key={row.id ?? i} style={{ borderBottom: i < filtered.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px" }}>{row.supplier || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.part_description || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.part_no || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.currency || <span style={{ color: "#B4B2A9" }}>-</span>}</td>
                <td style={{ padding: "9px 12px" }}>{row.price != null ? formatMoney(row.price) : <span style={{ color: "#B4B2A9" }}>-</span>}</td>
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
                <td style={{ padding: "9px 12px" }}>
                  <button type="button" onClick={() => setDeleting(row)} title="Delete quote" style={{ background: "none", border: "none", color: "#B85450", cursor: "pointer", padding: 4, display: "flex" }}>
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {deleting && (
        <DeleteConfirmModal
          itemLabel={`the quote for ${deleting.part_description || deleting.part_no || "this part"} from ${deleting.supplier || "this supplier"}`}
          userEmail={userEmail}
          onCancel={() => setDeleting(null)}
          onConfirm={handleDeleteQuote}
        />
      )}

      {isAdmin && (
        <div style={{ marginTop: 20 }}>
          <button type="button" onClick={() => setShowDeletedPanel((v) => !v)} style={{ background: "none", border: "none", color: NAVY, fontSize: 12.5, fontWeight: 600, cursor: "pointer", padding: 0, display: "flex", alignItems: "center", gap: 5 }}>
            <ChevronDown size={14} style={{ transform: showDeletedPanel ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
            Deleted Quotes (admin only)
          </button>
          {showDeletedPanel && (
            <DeletedQuotesPanel selectedSiteId={selectedSiteId} userEmail={userEmail} onRestored={load} />
          )}
        </div>
      )}
    </div>
  );
}

// Admin-only backup/restore view for quotes someone deleted from the Price
// List above. Every delete is backed up in full (not just logged) before
// the row is removed, specifically so a user coming back later saying "I
// accidentally deleted a quote" can have it restored rather than re-scanned
// from scratch. Regular users don't see this panel - they'd need to ask an
// admin, same as any other admin-only page in the app.
function DeletedQuotesPanel({ selectedSiteId, userEmail, onRestored }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [restoringId, setRestoringId] = useState(null);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error: err } = await supabase
      .from("quote_prices_deleted")
      .select("*")
      .eq("site_id", selectedSiteId)
      .order("deleted_at", { ascending: false });
    if (!err) setRows(data || []);
    setLoading(false);
  }, [selectedSiteId]);

  useEffect(() => { load(); }, [load]);

  const restore = async (row) => {
    setError("");
    setRestoringId(row.id);
    try {
      const { id, original_id, deleted_by_email, deleted_at, deletion_reason, ...quoteFields } = row;
      const { error: insertErr } = await supabase.from("quote_prices").insert(quoteFields);
      if (insertErr) throw insertErr;
      const { error: deleteErr } = await supabase.from("quote_prices_deleted").delete().eq("id", id);
      if (deleteErr) throw deleteErr;
      logActivity("Quote Price List", original_id, "created", `Restored a previously deleted quote: ${quoteFields.part_description || quoteFields.part_no || "item"} from ${quoteFields.supplier || "unknown supplier"}`);
      onRestored?.();
      load();
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setRestoringId(null);
    }
  };

  return (
    <div style={{ marginTop: 10, border: "1px solid #E2E6E3", borderRadius: 10, overflow: "hidden" }}>
      {error && (
        <div style={{ background: "#F6E2E0", border: "1px solid #DDB6B2", color: "#7A3330", borderRadius: 8, padding: "10px 14px", margin: 10, fontSize: 13 }}>
          {error}
        </div>
      )}
      <div style={{ overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F7F8F6" }}>
              {["Supplier", "Part", "Price", "Deleted", "Deleted by", "Reason", ""].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan={7} style={{ padding: 20, textAlign: "center", color: "#859195" }}>Loading…</td></tr>
            ) : rows.length === 0 ? (
              <tr><td colSpan={7} style={{ padding: 20, textAlign: "center", color: "#859195" }}>No deleted quotes on record.</td></tr>
            ) : rows.map((row, i) => (
              <tr key={row.id} style={{ borderBottom: i < rows.length - 1 ? "1px solid #EFEEE7" : "none" }}>
                <td style={{ padding: "9px 12px" }}>{row.supplier || "-"}</td>
                <td style={{ padding: "9px 12px" }}>{row.part_description || row.part_no || "-"}</td>
                <td style={{ padding: "9px 12px" }}>{row.price != null ? formatMoney(row.price) : "-"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.deleted_at ? new Date(row.deleted_at).toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" }) : "-"}</td>
                <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{row.deleted_by_email || "-"}</td>
                <td style={{ padding: "9px 12px" }}>{row.deletion_reason || "-"}</td>
                <td style={{ padding: "9px 12px" }}>
                  <button type="button" onClick={() => restore(row)} disabled={restoringId === row.id} style={{ background: "none", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12, fontWeight: 600, padding: "5px 10px", borderRadius: 6, cursor: restoringId === row.id ? "default" : "pointer" }}>
                    {restoringId === row.id ? "Restoring…" : "Restore"}
                  </button>
                </td>
              </tr>
            ))}
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

// ---------------------------------------------------------------------
// Component Status report (§16/§17).
//
// Reads component_status_calc and displays it. Deliberately NO
// calculation happens here: component hours, life used, status, UVIC and
// RVIC are all computed in the database view, so this screen, the future
// dashboard and any export cannot drift apart. If a number looks wrong,
// there is exactly one place to fix it.
// ---------------------------------------------------------------------
// ---------------------------------------------------------------------
// Mine2U report export - ONE implementation, used by every report.
//
// Layout: report name top left, Mine2U mark top right, a rule, the
// filter block, a rule, then the data. Numbers go in as NUMBERS with a
// display format, never as pre-formatted strings, so the file can be
// summed, sorted and charted by whoever opens it.
//
// Column widths are derived from the header text, so a heading is never
// clipped - the whole point of exporting is that someone can read it.
// ---------------------------------------------------------------------
const M2U_BRAND = "Mine2U";
const M2U_OWNER = "Datavera Analytics";
const M2U_XL_FMT = {
  n1: '#,##0.0;-#,##0.0;"-"',
  n2: '#,##0.00;-#,##0.00;"-"',
  int: '0;-0;"-"',
  pct: "0%",
  pct100: '0.00"%"',
  money: '#,##0.00;-#,##0.00;"-"',
  date: "yyyy-mm-dd",
  datetime: "yyyy-mm-dd hh:mm",
  text: null,
};

async function exportMine2UReport({ title, columns, rows, filterLines = [], fileName, sheetName, notes = [] }) {
  const ExcelJS = (await import("exceljs")).default;
  const now = new Date();
  const YEAR = now.getFullYear();
  const NOTICE = `© ${YEAR} ${M2U_OWNER}. Generated by ${M2U_BRAND}.`;

  const wb = new ExcelJS.Workbook();
  wb.creator = M2U_BRAND;
  wb.company = M2U_OWNER;
  wb.created = now;
  wb.description = NOTICE;

  const N = columns.length;
  const INK = "FF1F6668", HEAD = "FF203B46", BAND = "FF3D7379";
  const base = { name: "Arial", size: 9.5 };
  const RULE = { style: "medium", color: { argb: INK } };

  const ws = wb.addWorksheet(sheetName || "Report", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 9 }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    headerFooter: { oddFooter: `&L&8${NOTICE}&R&8Page &P of &N`, evenFooter: `&L&8${NOTICE}&R&8Page &P of &N` },
  });

  // masthead: title left, brand pinned to the LAST two columns so it
  // lands top-right whether the report has 8 columns or 34
  const brandFrom = Math.max(1, N - 1);
  ws.mergeCells(1, 1, 2, Math.max(1, brandFrom - 1));
  const t = ws.getCell(1, 1);
  t.value = title;
  t.font = { name: "Arial", size: 20, bold: true, color: { argb: INK } };
  t.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  ws.getRow(1).height = 20;
  ws.getRow(2).height = 20;

  ws.mergeCells(1, brandFrom, 2, N);
  const brand = ws.getCell(1, brandFrom);
  brand.value = M2U_BRAND;
  brand.font = { name: "Arial", size: 17, bold: true, color: { argb: INK } };
  brand.alignment = { vertical: "middle", horizontal: "right", indent: 1 };

  for (let c = 1; c <= N; c++) ws.getCell(3, c).border = { bottom: RULE };
  ws.getRow(3).height = 6;

  const band = (label, value, r) => {
    ws.mergeCells(r, 1, r, 2);
    ws.mergeCells(r, 4, r, N);
    const a = ws.getCell(r, 1), colon = ws.getCell(r, 3), b = ws.getCell(r, 4);
    a.value = label; colon.value = ":"; b.value = value;
    a.font = { ...base, bold: true, color: { argb: "FF183642" } };
    colon.font = { ...base, bold: true, color: { argb: "FF183642" } };
    colon.alignment = { horizontal: "center" };
    a.alignment = { indent: 1 };
    b.font = base;
  };
  const lines = [...filterLines, ["Execution", `[${now.toLocaleString("en-ZA", { dateStyle: "medium", timeStyle: "short" })}]`], ["Records", `[${rows.length}]`]];
  lines.slice(0, 4).forEach(([l, v], i) => band(l, v, 4 + i));

  for (let c = 1; c <= N; c++) ws.getCell(8, c).border = { bottom: RULE };
  ws.getRow(8).height = 6;

  // headers
  const HEADER_ROW = 9;
  columns.forEach((col, i) => {
    const c = ws.getCell(HEADER_ROW, i + 1);
    c.value = col.label;
    c.font = { ...base, bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEAD } };
    c.alignment = { horizontal: col.fmt === "text" ? "left" : "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(HEADER_ROW).height = 30;

  // data
  rows.forEach((row, ri) => {
    columns.forEach((col, i) => {
      const c = ws.getCell(HEADER_ROW + 1 + ri, i + 1);
      const raw = row[col.key];
      if (raw === null || raw === undefined || raw === "") {
        c.value = null;
      } else if (col.fmt === "text") {
        c.value = String(raw);
      } else if (col.fmt === "date" || col.fmt === "datetime") {
        const d = raw instanceof Date ? raw : new Date(raw);
        c.value = isNaN(d.getTime()) ? String(raw) : d;
      } else {
        const n = Number(raw);
        c.value = isNaN(n) ? String(raw) : n;
      }
      if (M2U_XL_FMT[col.fmt]) c.numFmt = M2U_XL_FMT[col.fmt];
      c.font = base;
      c.alignment = { horizontal: col.fmt === "text" ? "left" : "right" };
      if (ri % 2 === 1) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F7F6" } };
    });
  });

  // notes then the copyright band
  let r = HEADER_ROW + rows.length + 2;
  for (const n of notes) {
    ws.mergeCells(r, 1, r, N);
    const c = ws.getCell(r, 1);
    c.value = n;
    c.font = { name: "Arial", size: 8.5, italic: true, color: { argb: "FF859195" } };
    c.alignment = { wrapText: true, vertical: "top" };
    ws.getRow(r).height = 22;
    r += 1;
  }
  const noticeRow = r + 1;
  ws.mergeCells(noticeRow, 1, noticeRow, N);
  const note = ws.getCell(noticeRow, 1);
  note.value = NOTICE;
  note.font = { name: "Arial", size: 8.5, italic: true, color: { argb: "FFFFFFFF" } };
  note.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BAND } };
  note.alignment = { wrapText: true, vertical: "middle", indent: 1 };
  ws.getRow(noticeRow).height = 24;

  // Width from the header text and the widest value, so nothing is cut.
  ws.columns = columns.map((col) => {
    const headerNeed = Math.max(...String(col.label).split(" ").map((w) => w.length)) + 3;
    const longest = rows.reduce((m, row) => Math.max(m, String(row[col.key] ?? "").length), 0);
    const valueNeed = col.fmt === "text" ? Math.min(longest + 2, 42) : Math.min(longest + 2, 16);
    return { width: Math.max(headerNeed, valueNeed, 10) };
  });

  ws.autoFilter = { from: { row: HEADER_ROW, column: 1 }, to: { row: HEADER_ROW, column: N } };

  // Title, brand and notice locked; data left editable so people can
  // still sort and filter. A deterrent, not encryption.
  ws.eachRow({ includeEmpty: true }, (row, rn) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.protection = { locked: rn === 1 || rn === 2 || rn === noticeRow };
    });
  });
  await ws.protect("M2U-report", {
    selectLockedCells: true, selectUnlockedCells: true,
    formatCells: true, formatColumns: true, formatRows: true,
    sort: true, autoFilter: true,
    insertRows: false, deleteRows: false, insertColumns: false, deleteColumns: false,
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileName || title.replace(/\s+/g, "_")}_${now.toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---------------------------------------------------------------------
// Component Master (§3) and the installation / changeout workflow (§5, §6).
//
// Master data is the catalogue: what a component IS, what it should
// cost, and how long it should last. Installations are transactions
// against a specific machine. A changeout closes one row and opens
// another - the old row is never edited away.
// ---------------------------------------------------------------------
const CM_CATEGORIES = ["Powertrain", "Hydraulics", "Undercarriage", "Electrical", "Cooling", "Braking", "Structural", "GET", "Tyres", "Other"];
const CM_CRITICALITY = ["Low", "Medium", "High", "Critical"];
const CI_DESTINATIONS = ["Rebuild", "Repair", "Scrap", "Returned to OEM", "Stored"];

// ---------------------------------------------------------------------
// Consumption charts (homepage) and Fleet Health report.
//
// Both read from views, so no arithmetic happens in the browser. Rates
// rather than raw litres throughout: total litres only says which
// machine worked hardest, litres per hour says which one is drinking.
// ---------------------------------------------------------------------
// ---------------------------------------------------------------------
// Consumption charts (homepage).
//
// Built on the same KpiBarChart the availability charts use, so they
// read as part of the same system. The reference line is the SITE
// AVERAGE rather than an invented target: for consumption there is no
// right number, only "worse than the rest of the fleet".
//
// Click a fleet to see the machines behind it.
// ---------------------------------------------------------------------
function ConsumptionCharts() {
  const [rows, setRows] = useState([]);
  const [machines, setMachines] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [drill, setDrill] = useState(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [f, m] = await Promise.all([
        supabase.from("consumption_by_fleet_month").select("*").order("month", { ascending: true }),
        supabase.from("consumption_by_month").select("*").order("month", { ascending: true }),
      ]);
      if (cancelled) return;
      if (f.error || m.error) setError((f.error || m.error).message);
      else { setRows(f.data || []); setMachines(m.data || []); }
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, []);

  const { chartRows, months, fuelAvg, oilAvg } = useMemo(() => {
    const ms = [...new Set(rows.map((r) => r.month))].sort();
    const cur = ms[ms.length - 1];
    const prev = ms[ms.length - 2];
    const fleets = [...new Set(rows.map((r) => r.fleet))].filter(Boolean).sort();

    const data = fleets.map((fleet) => {
      const c = rows.find((r) => r.fleet === fleet && r.month === cur);
      const p = rows.find((r) => r.fleet === fleet && r.month === prev);
      return {
        fleet,
        fuel: c?.fuel_l_per_hour ?? null,
        fuelPrev: p?.fuel_l_per_hour ?? null,
        oil: c?.oil_l_per_100_hours ?? null,
        oilPrev: p?.oil_l_per_100_hours ?? null,
      };
    }).filter((r) => r.fuel != null || r.oil != null);

    // Site average weighted by hours, not an average of averages - a
    // machine that ran two hours shouldn't count as much as one that ran
    // two hundred.
    const inCur = rows.filter((r) => r.month === cur);
    const hours = inCur.reduce((a, r) => a + Number(r.engine_hours || 0), 0);
    const fuelL = inCur.reduce((a, r) => a + Number(r.fuel_litres || 0), 0);
    const oilL = inCur.reduce((a, r) => a + Number(r.oil_litres || 0), 0);

    return {
      chartRows: data,
      months: { cur, prev },
      fuelAvg: hours > 0 ? Math.round((fuelL / hours) * 100) / 100 : 0,
      oilAvg: hours > 0 ? Math.round((oilL / hours) * 100 * 100) / 100 : 0,
    };
  }, [rows]);

  const monthLabel = (m) => (m ? new Date(`${String(m).slice(0, 7)}-01T12:00:00`).toLocaleDateString("en-ZA", { month: "short", year: "numeric" }) : "");

  if (error) {
    return (
      <p style={{ fontSize: 12.5, color: "#B85450", background: "#F6E2E0", padding: "9px 12px", borderRadius: 8 }}>
        {error.includes("does not exist")
          ? "Consumption views aren't in the database yet - run fleet_health_and_consumption.sql."
          : error}
      </p>
    );
  }
  if (loading) return <p style={{ fontSize: 13, color: "#859195" }}>Loading consumption…</p>;
  if (chartRows.length === 0) return null;

  const fuelMax = Math.max(fuelAvg, ...chartRows.map((r) => Math.max(r.fuel || 0, r.fuelPrev || 0))) * 1.2;
  const oilMax = Math.max(oilAvg, ...chartRows.map((r) => Math.max(r.oil || 0, r.oilPrev || 0))) * 1.2;

  return (
    <>
      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, margin: "0 0 8px" }}>
        <p style={{ fontSize: 14, fontWeight: 700, color: NAVY, margin: 0 }}>Consumption</p>
        <p style={{ fontSize: 11.5, color: "#859195", margin: 0 }}>
          {monthLabel(months.cur)}{months.prev ? ` against ${monthLabel(months.prev)}` : ""} · click a fleet for its machines
        </p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 8 }}>
        <KpiBarChart
          title="Fuel — litres per engine hour"
          data={chartRows} xKey="fleet"
          dataKey="fuel" dataKey2="fuelPrev"
          seriesName="This month" seriesName2="Last month"
          target={fuelAvg} domainMax={fuelMax} unitSuffix=" L/hr"
          valueFormatter={(v) => (v == null ? "" : Number(v).toFixed(1))}
          meetsTarget={(r) => (r.fuel ?? 0) <= fuelAvg}
          meetsTarget2={(r) => (r.fuelPrev ?? 0) <= fuelAvg}
          onBarClick={(r) => setDrill({ fleet: r.fleet, metric: "fuel" })}
        />
        <KpiBarChart
          title="Oil — litres per 100 engine hours"
          data={chartRows} xKey="fleet"
          dataKey="oil" dataKey2="oilPrev"
          seriesName="This month" seriesName2="Last month"
          target={oilAvg} domainMax={oilMax} unitSuffix=" L/100h"
          valueFormatter={(v) => (v == null ? "" : Number(v).toFixed(1))}
          meetsTarget={(r) => (r.oil ?? 0) <= oilAvg}
          meetsTarget2={(r) => (r.oilPrev ?? 0) <= oilAvg}
          onBarClick={(r) => setDrill({ fleet: r.fleet, metric: "oil" })}
        />
      </div>

      <div style={{ display: "flex", gap: 16, fontSize: 11.5, color: "#4B5659", flexWrap: "wrap", marginBottom: 4 }}>
        <span><span style={{ display: "inline-block", width: 10, height: 10, background: CHART_COLOURS.last24h, borderRadius: 2, marginRight: 5 }} />This month</span>
        <span><span style={{ display: "inline-block", width: 10, height: 10, background: CHART_COLOURS.period, borderRadius: 2, marginRight: 5 }} />Last month</span>
        <span><span style={{ display: "inline-block", width: 12, height: 3, background: CHART_COLOURS.target, borderRadius: 2, marginRight: 6, verticalAlign: "middle" }} />Site average — {fuelAvg} L/hr · {oilAvg} L/100h</span>
        <span style={{ color: "#859195" }}>A bar above the line is drinking more than the rest of the fleet</span>
      </div>
      <p style={{ fontSize: 11, color: "#859195", margin: "0 0 4px" }}>
        Rates, not totals — a machine that ran twice as long should not look twice as thirsty. Oil counts top-ups only; scheduled changes are a service, not consumption.
      </p>

      {drill && (
        <ConsumptionDrilldown
          fleet={drill.fleet} metric={drill.metric}
          rows={machines.filter((m) => m.fleet === drill.fleet && m.month === months.cur)}
          prevRows={machines.filter((m) => m.fleet === drill.fleet && m.month === months.prev)}
          fleetAvg={drill.metric === "fuel" ? fuelAvg : oilAvg}
          monthLabel={monthLabel(months.cur)}
          onClose={() => setDrill(null)}
        />
      )}
    </>
  );
}

function ConsumptionDrilldown({ fleet, metric, rows, prevRows, fleetAvg, monthLabel, onClose }) {
  const key = metric === "fuel" ? "fuel_l_per_hour" : "oil_l_per_100_hours";
  const litreKey = metric === "fuel" ? "fuel_litres" : "oil_litres";
  const unit = metric === "fuel" ? "L/hr" : "L/100h";
  const prevBy = useMemo(() => Object.fromEntries(prevRows.map((r) => [r.asset_id, r[key]])), [prevRows, key]);

  const sorted = useMemo(
    () => [...rows].sort((a, b) => (Number(b[key]) || 0) - (Number(a[key]) || 0)),
    [rows, key]
  );

  const num = (v, d = 1) => (v == null || v === "" ? "-" : Number(v).toLocaleString(undefined, { maximumFractionDigits: d }));
  const th = { textAlign: "left", padding: "7px 9px", fontSize: 11, fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" };
  const thR = { ...th, textAlign: "right" };
  const td = { padding: "7px 9px", fontSize: 12, whiteSpace: "nowrap" };
  const tdR = { ...td, textAlign: "right" };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(24,54,66,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: 24, zIndex: 100, overflowY: "auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 12, padding: 22, width: 860, maxWidth: "96vw", maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 4 }}>
          <div>
            <p style={{ margin: 0, fontSize: 17, fontWeight: 700, color: NAVY }}>{fleet} — {metric === "fuel" ? "fuel" : "oil"} consumption</p>
            <p style={{ margin: "2px 0 0", fontSize: 12, color: "#859195" }}>{monthLabel} · site average {fleetAvg} {unit}</p>
          </div>
          <button onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", padding: "6px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Close</button>
        </div>

        <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "auto", marginTop: 12 }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: "#F7F8F6" }}>
                <th style={th}>Equipment</th><th style={th}>Name</th>
                <th style={thR}>Engine hrs</th><th style={thR}>Litres</th>
                <th style={thR}>{unit}</th><th style={thR}>Last month</th><th style={thR}>Change</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((r) => {
                const now = Number(r[key]);
                const was = Number(prevBy[r.asset_id]);
                const delta = isFinite(now) && isFinite(was) && was > 0 ? ((now - was) / was) * 100 : null;
                const over = isFinite(now) && now > fleetAvg;
                return (
                  <tr key={r.asset_id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                    <td style={{ ...td, fontWeight: 600 }}>{r.asset_id}</td>
                    <td style={td}>{r.asset_name}</td>
                    <td style={tdR}>{num(r.engine_hours)}</td>
                    <td style={tdR}>{num(r[litreKey])}</td>
                    <td style={{ ...tdR, fontWeight: 700, color: over ? "#B85450" : "#2C5646" }}>{num(now, 2)}</td>
                    <td style={{ ...tdR, color: "#859195" }}>{num(was, 2)}</td>
                    <td style={{ ...tdR, color: delta == null ? "#B4B2A9" : delta > 10 ? "#B85450" : delta < -10 ? "#2C5646" : "#4B5659" }}>
                      {delta == null ? "-" : `${delta > 0 ? "+" : ""}${delta.toFixed(0)}%`}
                    </td>
                  </tr>
                );
              })}
              {sorted.length === 0 && (
                <tr><td colSpan={7} style={{ padding: 18, textAlign: "center", color: "#859195", fontSize: 12.5 }}>
                  No consumption recorded for this fleet in {monthLabel}.
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
        <p style={{ margin: "8px 2px 0", fontSize: 11, color: "#859195" }}>
          Sorted worst first. Change compares against the same machine last month — a machine can sit below the site average and still be climbing.
        </p>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------
// Fleet Health drill-down: one machine's components.
//
// Reads the same component_status_calc as the main report, filtered to
// one asset, so the drill-down can never disagree with the row that
// opened it. Removed components are shown too - the history is the
// point of tracking serials at all.
// ---------------------------------------------------------------------
function FleetHealthDrilldown({ machine, onClose }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [showHistory, setShowHistory] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error: err } = await supabase
        .from("component_status_calc")
        .select("*")
        .eq("asset_id", machine.asset_id)
        .order("life_used_pct", { ascending: false, nullsFirst: false });
      if (cancelled) return;
      if (err) setError(err.message); else setRows(data || []);
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, [machine.asset_id]);

  const active = useMemo(() => rows.filter((r) => !r.removed_date), [rows]);
  const history = useMemo(() => rows.filter((r) => r.removed_date), [rows]);

  const tone = (s) =>
    s === "OVERDUE" || s === "FAILED" ? { bg: "#F6E2E0", text: "#7A3330", bar: "#B85450" }
    : s === "DUE" || s === "APPROACHING LIFE" ? { bg: "#F5E9D8", text: "#7A5A22", bar: "#C79A12" }
    : s === "NORMAL" || s === "NEW" ? { bg: "#E2EFE9", text: "#2C5646", bar: "#2F9E63" }
    : { bg: "#EDEBE4", text: "#4B5659", bar: "#B4B2A9" };

  const num = (v, d = 1) => (v == null || v === "" ? "-" : Number(v).toLocaleString(undefined, { maximumFractionDigits: d }));
  const money = (v) => (v == null || v === "" ? "-" : Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }));

  const th = { textAlign: "left", padding: "7px 9px", fontSize: 11, fontWeight: 600, color: "#4B5659", borderBottom: "1px solid #E2E6E3", whiteSpace: "nowrap" };
  const thR = { ...th, textAlign: "right" };
  const td = { padding: "7px 9px", fontSize: 12, whiteSpace: "nowrap" };
  const tdR = { ...td, textAlign: "right" };

  const stat = (label, value, sub) => (
    <div style={{ background: "#F7F8F6", borderRadius: 8, padding: "9px 12px", minWidth: 120 }}>
      <p style={{ margin: 0, fontSize: 11, color: "#859195" }}>{label}</p>
      <p style={{ margin: "2px 0 0", fontSize: 17, fontWeight: 700, color: NAVY }}>{value}</p>
      {sub && <p style={{ margin: "1px 0 0", fontSize: 10.5, color: "#859195" }}>{sub}</p>}
    </div>
  );

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(24,54,66,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: 24, zIndex: 100, overflowY: "auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 12, padding: 22, width: 1180, maxWidth: "97vw", maxHeight: "92vh", overflowY: "auto" }}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 12 }}>
          <div>
            <p style={{ margin: 0, fontSize: 17, fontWeight: 700, color: NAVY }}>
              {machine.asset_id}{machine.asset_name ? ` — ${machine.asset_name}` : ""}
            </p>
            <p style={{ margin: "2px 0 0", fontSize: 12, color: "#859195" }}>
              {[machine.fleet, machine.model, machine.make].filter(Boolean).join(" · ")}
            </p>
          </div>
          <button onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", padding: "6px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Close</button>
        </div>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
          {stat("Machine hours", num(machine.current_equip_hours), machine.last_reading_date ? `read ${String(machine.last_reading_date).slice(0, 10)}` : null)}
          {stat("Health", machine.health_score == null ? "not scored" : Number(machine.health_score).toFixed(0), machine.coverage)}
          {stat("Overdue", machine.overdue_count || 0, machine.due_count ? `${machine.due_count} due` : null)}
          {stat("Hours / day", num(machine.avg_hours_per_day, 2), "last 30 days")}
          {stat("Next changeout", machine.days_to_first_changeout == null ? "-" : `${machine.days_to_first_changeout} d`, machine.first_changeout_date ? String(machine.first_changeout_date).slice(0, 10) : "insufficient data")}
        </div>

        {error && <p style={{ fontSize: 12.5, color: "#B85450" }}>{error}</p>}
        {loading ? <p style={{ fontSize: 13, color: "#859195" }}>Loading components…</p> : (
          <>
            <p style={{ margin: "0 0 6px", fontSize: 13, fontWeight: 700, color: NAVY }}>Fitted components ({active.length})</p>
            <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#F7F8F6" }}>
                    <th style={th}>Component</th><th style={th}>Position</th><th style={th}>Serial</th>
                    <th style={thR}>Fitted at</th><th style={thR}>Comp. hrs</th><th style={thR}>Target</th>
                    <th style={thR}>Remaining</th><th style={{ ...th, width: 150 }}>Life used</th>
                    <th style={th}>Status</th><th style={thR}>Next at</th><th style={thR}>Est. date</th>
                    <th style={thR}>RVIC</th>
                  </tr>
                </thead>
                <tbody>
                  {active.map((c) => {
                    const t = tone(c.status);
                    const pct = c.life_used_pct == null ? null : Number(c.life_used_pct);
                    return (
                      <tr key={c.id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                        <td style={{ ...td, fontWeight: 600 }}>{c.component_type}</td>
                        <td style={td}>{c.position_label || "-"}</td>
                        <td style={td}>{c.serial_number || "-"}</td>
                        <td style={tdR}>{num(c.installed_equip_hours)}</td>
                        <td style={tdR}>{num(c.component_hours)}</td>
                        <td style={tdR}>{num(c.target_life_hours, 0)}</td>
                        <td style={{ ...tdR, color: Number(c.remaining_hours) < 0 ? "#B85450" : "#183642", fontWeight: Number(c.remaining_hours) < 0 ? 700 : 400 }}>{num(c.remaining_hours)}</td>
                        <td style={td}>
                          {pct == null ? <span style={{ fontSize: 11, color: "#859195" }}>no target</span> : (
                            <>
                              <span style={{ display: "block", height: 6, background: "#F2F1EA", borderRadius: 4, overflow: "hidden" }}>
                                <span style={{ display: "block", height: "100%", width: `${Math.max(2, Math.min(100, pct))}%`, background: t.bar }} />
                              </span>
                              <span style={{ fontSize: 10.5, color: t.text }}>{pct.toFixed(1)}%</span>
                            </>
                          )}
                        </td>
                        <td style={td}><span style={{ background: t.bg, color: t.text, padding: "2px 8px", borderRadius: 6, fontSize: 10.5, fontWeight: 600 }}>{c.status}</span></td>
                        <td style={tdR}>{num(c.next_changeout_equip_hours)}</td>
                        <td style={{ ...tdR, color: c.estimated_changeout_date ? "#183642" : "#859195", fontSize: c.estimated_changeout_date ? 12 : 10.5 }}>
                          {c.estimated_changeout_date ? String(c.estimated_changeout_date).slice(0, 10) : "insufficient data"}
                        </td>
                        <td style={tdR}>{money(c.rvic)}</td>
                      </tr>
                    );
                  })}
                  {active.length === 0 && (
                    <tr><td colSpan={12} style={{ padding: 18, textAlign: "center", color: "#859195", fontSize: 12.5 }}>
                      No components registered against this machine yet — which is why it reads NOT SCORED.
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>

            {history.length > 0 && (
              <>
                <button onClick={() => setShowHistory((v) => !v)}
                  style={{ background: "none", border: "none", color: NAVY, fontSize: 12.5, fontWeight: 600, cursor: "pointer", padding: "12px 0 6px" }}>
                  {showHistory ? "▾" : "▸"} Change out history ({history.length})
                </button>
                {showHistory && (
                  <div style={{ border: "1px solid #E2E6E3", borderRadius: 10, overflow: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse" }}>
                      <thead>
                        <tr style={{ background: "#F7F8F6" }}>
                          <th style={th}>Component</th><th style={th}>Serial</th>
                          <th style={thR}>Fitted</th><th style={thR}>Removed</th>
                          <th style={thR}>Achieved</th><th style={thR}>Target</th><th style={thR}>Life used</th>
                          <th style={th}>Reason</th><th style={th}>Destination</th>
                        </tr>
                      </thead>
                      <tbody>
                        {history.map((c) => (
                          <tr key={c.id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                            <td style={{ ...td, fontWeight: 600 }}>{c.component_type}</td>
                            <td style={td}>{c.serial_number || "-"}</td>
                            <td style={tdR}>{c.installed_date || "-"}</td>
                            <td style={tdR}>{c.removed_date || "-"}</td>
                            <td style={tdR}>{num(c.component_hours)}</td>
                            <td style={tdR}>{num(c.target_life_hours, 0)}</td>
                            <td style={{ ...tdR, color: Number(c.life_used_pct) < 70 ? "#B85450" : "#183642", fontWeight: Number(c.life_used_pct) < 70 ? 700 : 400 }}>
                              {c.life_used_pct == null ? "-" : `${Number(c.life_used_pct).toFixed(1)}%`}
                            </td>
                            <td style={td}>{c.removal_reason || "-"}</td>
                            <td style={td}>{c.destination || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <p style={{ margin: "8px 10px", fontSize: 11, color: "#859195" }}>
                      A component removed well short of its target is a warranty case — life used below 70% is highlighted.
                    </p>
                  </div>
                )}
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

const FLEET_HEALTH_COLUMNS = [
  ["asset_id", "Equipment", "", "text"],
  ["asset_name", "Name", "", "text"],
  ["fleet", "Fleet", "", "text"],
  ["current_equip_hours", "Machine Hours", "", "n1"],
  ["health_score", "Health", "", "pct100"],
  ["health_band", "Band", "", "text"],
  ["overdue_count", "Overdue", "Components", "int"],
  ["due_count", "Due", "Components", "int"],
  ["approaching_count", "Approaching", "Components", "int"],
  ["coverage", "Coverage", "Components", "text"],
  ["components_scored", "Scored", "Components", "int"],
  ["components", "Fitted", "Components", "int"],
  ["worst_component", "Worst Component", "Limiting", "text"],
  ["worst_life_used_pct", "Life Used", "Limiting", "pct100"],
  ["days_to_first_changeout", "Days to First", "Next Changeout", "int"],
  ["first_changeout_date", "Expected Date", "Next Changeout", "date"],
];
const FLEET_HEALTH_PREF_COLUMNS = FLEET_HEALTH_COLUMNS.map(([key, label]) => ({ key, label }));

function FleetHealthPage({ assets }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [bandFilter, setBandFilter] = useState("");
  const [drill, setDrill] = useState(null);
  const prefs = useTablePrefs("fleet_health", FLEET_HEALTH_PREF_COLUMNS);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error: err } = await supabase.from("fleet_health_calc").select("*").order("health_score", { ascending: true, nullsFirst: false });
      if (cancelled) return;
      if (err) setError(err.message); else setRows(data || []);
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, []);

  const filtered = useMemo(() => {
    let r = rows;
    if (selectedAsset) r = r.filter((x) => x.asset_id === selectedAsset);
    else if (selectedFleet) r = r.filter((x) => x.fleet === selectedFleet);
    if (bandFilter) r = r.filter((x) => x.health_band === bandFilter);
    if (query.trim()) {
      const q = query.toLowerCase();
      r = r.filter((x) => Object.values(x).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return prefs.sortRows(r);
  }, [rows, selectedAsset, selectedFleet, bandFilter, query, prefs.sortRows]); // eslint-disable-line react-hooks/exhaustive-deps

  const totals = useMemo(() => ({
    machines: filtered.length,
    atRisk: filtered.filter((r) => r.health_band === "AT RISK").length,
    monitor: filtered.filter((r) => r.health_band === "MONITOR").length,
    healthy: filtered.filter((r) => r.health_band === "HEALTHY").length,
    notScored: filtered.filter((r) => r.health_band === "NOT SCORED").length,
    overdue: filtered.reduce((a, r) => a + Number(r.overdue_count || 0), 0),
  }), [filtered]);

  const bandTone = (b) =>
    b === "AT RISK" ? { bg: "#F6E2E0", text: "#7A3330", bar: "#B85450" }
    : b === "MONITOR" ? { bg: "#F5E9D8", text: "#7A5A22", bar: "#C79A12" }
    : b === "HEALTHY" ? { bg: "#E2EFE9", text: "#2C5646", bar: "#2F9E63" }
    : { bg: "#EDEBE4", text: "#4B5659", bar: "#B4B2A9" };

  const fmt = (key, val) => {
    const type = FLEET_HEALTH_COLUMNS.find((c) => c[0] === key)?.[3];
    if (val == null || val === "") return "-";
    switch (type) {
      case "n1": return Number(val).toLocaleString(undefined, { maximumFractionDigits: 1 });
      case "int": return String(Math.round(Number(val)));
      case "pct100": return `${Number(val).toFixed(1)}%`;
      case "date": return String(val).slice(0, 10);
      default: return val;
    }
  };

  const exportReport = () =>
    exportMine2UReport({
      title: "Fleet Health",
      sheetName: "Fleet Health",
      fileName: "Mine2U_Fleet_Health",
      columns: prefs.columns.map((c) => {
        const def = FLEET_HEALTH_COLUMNS.find((x) => x[0] === c.key);
        return { key: c.key, label: def ? `${def[2] ? def[2] + " - " : ""}${def[1]}` : c.label, fmt: def?.[3] || "text" };
      }),
      rows: filtered,
      filterLines: [
        ["Hierarchy", `[Fleet: ${selectedFleet || "All"}]  [Equipment: ${selectedAsset || "All"}]`],
        ["Additional Filters", `[Band: ${bandFilter || "All"}]`],
      ],
      notes: [
        "Health is the straight average of life remaining across each machine's scored components, calculated from meter readings at the moment of export.",
        "Overdue components are counted separately rather than folded into the score, so one overdue item does not mask an otherwise sound machine.",
        "Components with no life target are excluded from the average but included in the Fitted count.",
        "Coverage shows how much of the machine the score actually covers. A machine scored on one component of eight is not a whole-machine verdict - read the score against its coverage.",
      ],
    });

  const widthFor = (key, type) => {
    if (key === "asset_id" || key === "asset_name" || key === "worst_component") return 120;
    if (key === "health_score") return 190;
    if (key === "coverage") return 110;
    if (key === "health_band" || key === "fleet") return 96;
    if (type === "date") return 104;
    return 72;
  };

  const spans = [];
  FLEET_HEALTH_COLUMNS.forEach(([, , group]) => {
    const last = spans[spans.length - 1];
    if (last && last.group === group) last.span += 1;
    else spans.push({ group, span: 1 });
  });

  const chip = (label, value, tone) => (
    <div style={{ background: tone.bg, color: tone.text, borderRadius: 8, padding: "7px 12px", fontSize: 12.5 }}>
      <strong style={{ fontSize: 15 }}>{value}</strong> {label}
    </div>
  );

  return (
    <div>

      {error && (
        <p style={{ fontSize: 13, color: "#B85450", background: "#F6E2E0", padding: "9px 12px", borderRadius: 8 }}>
          {error.includes("does not exist")
            ? "The Fleet Health view isn't in the database yet - run fleet_health_and_consumption.sql."
            : error}
        </p>
      )}

      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
        <div style={{ position: "relative", width: 250, flexShrink: 0 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search equipment"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <select value={bandFilter} onChange={(e) => setBandFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
          <option value="">All bands</option>
          {["AT RISK", "MONITOR", "HEALTHY", "NOT SCORED"].map((b) => <option key={b} value={b}>{b}</option>)}
        </select>
        <ColumnsButton prefs={prefs} />
        <button onClick={() => { setQuery(""); setBandFilter(""); setSelectedFleet(""); setSelectedAsset(""); }}
          style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", fontSize: 12.5, fontWeight: 600, padding: "8px 12px", borderRadius: 8, cursor: "pointer" }}>
          Clear filters
        </button>
        <button onClick={exportReport}
          style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
          <Download size={14} /> Export to Excel
        </button>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
        {chip("machines", totals.machines, { bg: "#F2F1EA", text: "#183642" })}
        {chip("at risk", totals.atRisk, { bg: "#F6E2E0", text: "#7A3330" })}
        {chip("monitor", totals.monitor, { bg: "#F5E9D8", text: "#7A5A22" })}
        {chip("healthy", totals.healthy, { bg: "#E2EFE9", text: "#2C5646" })}
        {totals.notScored > 0 && chip("not scored", totals.notScored, { bg: "#EDEBE4", text: "#4B5659" })}
        {totals.overdue > 0 && chip("overdue components", totals.overdue, { bg: "#F6E2E0", text: "#7A3330" })}
      </div>

      {loading ? <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p> : (
        <div style={{ overflow: "auto", maxHeight: "62vh", border: "1px solid #E2E6E3", borderRadius: 10 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
            <colgroup>
              {prefs.columns.map((c) => {
                const def = FLEET_HEALTH_COLUMNS.find((x) => x[0] === c.key);
                return <col key={c.key} style={{ width: prefs.widths?.[c.key] ?? widthFor(c.key, def?.[3]) }} />;
              })}
            </colgroup>
            <thead>
              <tr style={{ background: "#E9ECEA" }}>
                {spans.map((sp, i) => (
                  <th key={i} colSpan={sp.span} title={sp.group}
                    style={{ position: "sticky", top: 0, zIndex: 3, background: "#E9ECEA", textAlign: "center", padding: "5px 6px",
                             fontWeight: 700, fontSize: 11, color: NAVY, borderBottom: "1px solid #E2E6E3",
                             borderLeft: i > 0 ? "1px solid #DCE0DD" : "none", whiteSpace: "nowrap" }}>
                    {sp.group}
                  </th>
                ))}
              </tr>
              <tr style={{ background: "#F7F8F6" }}>
                {prefs.columns.map((c) => {
                  const def = FLEET_HEALTH_COLUMNS.find((x) => x[0] === c.key);
                  const type = def?.[3];
                  return (
                    <th key={c.key} onClick={() => prefs.toggleSort(c.key)} title={c.label}
                      style={{ position: "sticky", top: 26, zIndex: c.key === "asset_id" ? 5 : 3, background: "#F7F8F6",
                               ...(c.key === "asset_id" ? { left: 0 } : {}),
                               textAlign: type === "text" ? "left" : "right", padding: "6px 7px",
                               fontWeight: 600, color: "#4B5659", fontSize: 11, lineHeight: 1.25,
                               borderBottom: "1px solid #E2E6E3", cursor: "pointer", userSelect: "none",
                               whiteSpace: "normal", overflowWrap: "normal", hyphens: "none" }}>
                      {c.label}{prefs.sortKey === c.key ? (prefs.sortDir === "asc" ? " ▲" : " ▼") : ""}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => {
                const tone = bandTone(row.health_band);
                return (
                  <tr key={row.asset_id}
                    onClick={() => setDrill(row)}
                    title="Open this machine's components"
                    style={{ borderBottom: "1px solid #EFEEE7", cursor: "pointer" }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = "#F7F8F6"; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}>
                    {prefs.columns.map((c) => {
                      const def = FLEET_HEALTH_COLUMNS.find((x) => x[0] === c.key);
                      const type = def?.[3];

                      if (c.key === "health_score") {
                        const pct = row.health_score;
                        return (
                          <td key={c.key} style={{ padding: "6px 7px" }}>
                            {pct == null ? <span style={{ fontSize: 11, color: "#859195" }}>not scored</span> : (
                              <>
                                <span style={{ display: "block", height: 7, background: "#F2F1EA", borderRadius: 4, overflow: "hidden" }}>
                                  <span style={{ display: "block", height: "100%", width: `${Math.max(2, Math.min(100, pct))}%`, background: tone.bar }} />
                                </span>
                                <span style={{ fontSize: 10.5, color: tone.text }}>{Number(pct).toFixed(0)}</span>
                                {row.thin_coverage && (
                                  <span title={`Scored on only ${row.components_scored} component(s) - not a whole-machine verdict`}
                                    style={{ fontSize: 10, color: "#B07D2B", marginLeft: 6 }}>
                                    ⚠ {row.coverage}
                                  </span>
                                )}
                              </>
                            )}
                          </td>
                        );
                      }

                      if (c.key === "overdue_count") {
                        const n = Number(row.overdue_count || 0);
                        return (
                          <td key={c.key} style={{ padding: "6px 7px", textAlign: "right" }}>
                            {n > 0
                              ? <span style={{ background: "#F6E2E0", color: "#7A3330", padding: "2px 8px", borderRadius: 6, fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>{n} overdue</span>
                              : <span style={{ color: "#B4B2A9" }}>-</span>}
                          </td>
                        );
                      }

                      if (c.key === "health_band") {
                        return (
                          <td key={c.key} style={{ padding: "6px 7px" }}>
                            <span style={{ background: tone.bg, color: tone.text, padding: "2px 8px", borderRadius: 6, fontSize: 11, fontWeight: 600, whiteSpace: "nowrap" }}>{row.health_band}</span>
                          </td>
                        );
                      }

                      if (c.key === "worst_component") {
                        return (
                          <td key={c.key} style={{ padding: "6px 7px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
                            title={`${row.worst_component || ""}${row.worst_position ? " " + row.worst_position : ""}`}>
                            {row.worst_component ? `${row.worst_component}${row.worst_position ? ` ${row.worst_position}` : ""}` : "-"}
                          </td>
                        );
                      }

                      return (
                        <td key={c.key} title={String(row[c.key] ?? "")}
                          style={{ padding: "6px 7px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                                   textAlign: type === "text" ? "left" : "right",
                                   ...(c.key === "asset_id" ? { position: "sticky", left: 0, zIndex: 2, background: "inherit", fontWeight: 600 } : {}) }}>
                          {fmt(c.key, row[c.key])}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
              {filtered.length === 0 && (
                <tr><td colSpan={prefs.columns.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                  {rows.length === 0 ? "No equipment with components fitted yet." : "No machines match the current filters."}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {drill && <FleetHealthDrilldown machine={drill} onClose={() => setDrill(null)} />}
    </div>
  );
}

function ComponentMasterPage({ assets, selectedSiteId, myFullName }) {
  const [tab, setTab] = useState("master");
  const [master, setMaster] = useState([]);
  const [installs, setInstalls] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [installing, setInstalling] = useState(null);
  const [removing, setRemoving] = useState(null);
  const [query, setQuery] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    const [m, i] = await Promise.all([
      supabase.from("component_master").select("*").order("component_type"),
      supabase.from("component_status_calc").select("*").order("asset_id"),
    ]);
    if (m.error || i.error) setError((m.error || i.error).message);
    else { setError(""); setMaster(m.data || []); setInstalls(i.data || []); }
    setLoading(false);
  }, []);
  useEffect(() => { load(); }, [load]);

  const filteredMaster = useMemo(() => {
    if (!query.trim()) return master;
    const q = query.toLowerCase();
    return master.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
  }, [master, query]);

  const activeInstalls = useMemo(() => installs.filter((r) => !r.removed_date), [installs]);

  const tabBtn = (key, label) => (
    <button onClick={() => setTab(key)}
      style={{ background: tab === key ? NAVY : "#fff", color: tab === key ? "#fff" : "#4B5659",
               border: `1px solid ${tab === key ? NAVY : "#E2E6E3"}`, padding: "7px 14px", borderRadius: 8,
               fontSize: 13, fontWeight: 600, cursor: "pointer" }}>{label}</button>
  );

  return (
    <div>

      {error && (
        <p style={{ fontSize: 13, color: "#B85450", background: "#F6E2E0", padding: "9px 12px", borderRadius: 8 }}>
          {error.includes("does not exist")
            ? "The Component Life tables aren't in the database yet - run component_life_foundation.sql first."
            : error}
        </p>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        {tabBtn("master", `Catalogue (${master.length})`)}
        {tabBtn("fitted", `Fitted Components (${activeInstalls.length})`)}
      </div>

      {loading ? <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p> : tab === "master" ? (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, marginBottom: 12, flexWrap: "wrap" }}>
            <div style={{ position: "relative", flex: 1, maxWidth: 280 }}>
              <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
              <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search catalogue"
                style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
            </div>
            <button onClick={() => { setEditing(null); setShowForm(true); }}
              style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 16px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>
              + Add Component
            </button>
          </div>

          <div style={{ overflow: "auto", maxHeight: "62vh", border: "1px solid #E2E6E3", borderRadius: 10 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ background: "#F7F8F6" }}>
                  {["Component", "Category", "OEM", "Part No", "Target Life", "Warning %", "Expected Cost", "Applies To", "Criticality", "Active", ""].map((h, i) => (
                    <th key={h + i} style={{ position: "sticky", top: 0, zIndex: 2, background: "#F7F8F6", textAlign: i >= 4 && i <= 6 ? "right" : "left",
                                             padding: "8px 10px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredMaster.map((r) => (
                  <tr key={r.id} style={{ borderBottom: "1px solid #EFEEE7", opacity: r.active ? 1 : 0.5 }}>
                    <td style={{ padding: "8px 10px", cursor: "pointer" }} onClick={() => { setEditing(r); setShowForm(true); }}>
                      <strong>{r.component_type}</strong>{r.description ? <span style={{ color: "#859195" }}> — {r.description}</span> : null}
                    </td>
                    <td style={{ padding: "8px 10px" }}>{r.category || "-"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.oem || "-"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.oem_part_no || "-"}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right" }}>
                      {r.target_life_hours
                        ? Number(r.target_life_hours).toLocaleString()
                        : <span style={{ color: "#B85450" }}>not set</span>}
                    </td>
                    <td style={{ padding: "8px 10px", textAlign: "right" }}>{r.warning_pct ?? "site default"}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right" }}>{r.expected_cost != null ? Number(r.expected_cost).toLocaleString(undefined, { minimumFractionDigits: 2 }) : "-"}</td>
                    <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{[r.applies_make, r.applies_model].filter(Boolean).join(" ") || "All"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.criticality || "-"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.active ? "Yes" : "No"}</td>
                    <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>
                      <button onClick={() => setInstalling(r)} style={{ background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 12, fontWeight: 600, padding: "5px 10px", borderRadius: 7, cursor: "pointer" }}>
                        Install
                      </button>
                    </td>
                  </tr>
                ))}
                {filteredMaster.length === 0 && (
                  <tr><td colSpan={11} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                    No components defined yet. Add one to start tracking life.
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      ) : (
        <div style={{ overflow: "auto", maxHeight: "62vh", border: "1px solid #E2E6E3", borderRadius: 10 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead>
              <tr style={{ background: "#F7F8F6" }}>
                {["Equipment", "Component", "Position", "Serial", "Fitted", "Comp. Hrs", "Target", "Life Used", "Status", ""].map((h, i) => (
                  <th key={h + i} style={{ position: "sticky", top: 0, zIndex: 2, background: "#F7F8F6", textAlign: i >= 5 && i <= 7 ? "right" : "left",
                                           padding: "8px 10px", fontWeight: 600, color: "#4B5659", whiteSpace: "nowrap", borderBottom: "1px solid #E2E6E3" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {activeInstalls.map((r) => {
                const tone = statusColor(r.status);
                return (
                  <tr key={r.id} style={{ borderBottom: "1px solid #EFEEE7" }}>
                    <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{r.asset_id}</td>
                    <td style={{ padding: "8px 10px" }}>{r.component_type}</td>
                    <td style={{ padding: "8px 10px" }}>{r.position_label || "-"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.serial_number || "-"}</td>
                    <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{r.installed_date || "-"}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right" }}>{r.component_hours != null ? Number(r.component_hours).toLocaleString(undefined, { maximumFractionDigits: 1 }) : "-"}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right" }}>{r.target_life_hours != null ? Number(r.target_life_hours).toLocaleString() : "-"}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right" }}>{r.life_used_pct != null ? `${Number(r.life_used_pct).toFixed(1)}%` : "-"}</td>
                    <td style={{ padding: "8px 10px" }}>
                      <span style={{ background: tone.bg, color: tone.text, padding: "2px 8px", borderRadius: 6, fontSize: 11, fontWeight: 600, whiteSpace: "nowrap" }}>{r.status}</span>
                    </td>
                    <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>
                      <button onClick={() => setRemoving(r)} style={{ background: "#fff", border: "1px solid #B85450", color: "#B85450", fontSize: 12, fontWeight: 600, padding: "5px 10px", borderRadius: 7, cursor: "pointer" }}>
                        Change Out
                      </button>
                    </td>
                  </tr>
                );
              })}
              {activeInstalls.length === 0 && (
                <tr><td colSpan={10} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                  Nothing fitted yet. Use Install on the catalogue tab.
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {showForm && <ComponentMasterForm existing={editing} selectedSiteId={selectedSiteId} onClose={() => setShowForm(false)} onSaved={() => { setShowForm(false); load(); }} />}
      {installing && <ComponentInstallForm master={installing} assets={assets} selectedSiteId={selectedSiteId} myFullName={myFullName} onClose={() => setInstalling(null)} onSaved={() => { setInstalling(null); load(); }} />}
      {removing && <ComponentChangeOutForm install={removing} assets={assets} selectedSiteId={selectedSiteId} myFullName={myFullName} onClose={() => setRemoving(null)} onSaved={() => { setRemoving(null); load(); }} />}
    </div>
  );
}

function ComponentMasterForm({ existing, selectedSiteId, onClose, onSaved }) {
  const [f, setF] = useState({
    component_type: existing?.component_type || "", description: existing?.description || "",
    category: existing?.category || "", subcategory: existing?.subcategory || "",
    oem: existing?.oem || "", oem_part_no: existing?.oem_part_no || "", alt_part_no: existing?.alt_part_no || "",
    supplier: existing?.supplier || "",
    target_life_hours: existing?.target_life_hours ?? "", min_expected_hours: existing?.min_expected_hours ?? "",
    max_expected_hours: existing?.max_expected_hours ?? "",
    warning_hours: existing?.warning_hours ?? "", critical_hours: existing?.critical_hours ?? "",
    uses_positions: existing?.uses_positions ?? false,
    expected_cost: existing?.expected_cost ?? "", warranty_hours: existing?.warranty_hours ?? "",
    warranty_months: existing?.warranty_months ?? "",
    applies_equip_type: existing?.applies_equip_type || "", applies_make: existing?.applies_make || "",
    applies_model: existing?.applies_model || "", default_position: existing?.default_position || "",
    criticality: existing?.criticality || "Medium", active: existing?.active ?? true,
  });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const save = async (e) => {
    e.preventDefault();
    if (!f.component_type.trim()) { setError("Component type is required."); return; }
    setSaving(true); setError("");
    const num = (v) => (v === "" || v == null ? null : Number(v));
    const payload = {
      ...f,
      component_type: f.component_type.trim(),
      site_id: selectedSiteId || null,
      target_life_hours: num(f.target_life_hours), min_expected_hours: num(f.min_expected_hours),
      max_expected_hours: num(f.max_expected_hours), warning_hours: num(f.warning_hours),
      critical_hours: num(f.critical_hours), expected_cost: num(f.expected_cost),
      warranty_hours: num(f.warranty_hours), warranty_months: num(f.warranty_months),
    };
    const { error: err } = existing
      ? await supabase.from("component_master").update(payload).eq("id", existing.id)
      : await supabase.from("component_master").insert(payload);
    setSaving(false);
    if (err) setError(err.message); else onSaved();
  };

  const fs = { width: "100%", padding: "7px 9px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box" };
  const ls = { display: "block", fontSize: 12, fontWeight: 600, color: "#183642", margin: "0 0 4px" };
  const sec = { fontSize: 12, fontWeight: 700, color: "#4B5659", margin: "14px 0 6px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(24,54,66,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: 24, zIndex: 100, overflowY: "auto" }}>
      <form onSubmit={save} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 900, maxWidth: "96vw", maxHeight: "90vh", overflowY: "auto" }}>
        <p style={{ margin: "0 0 14px", fontSize: 16, fontWeight: 700, color: NAVY }}>{existing ? "Edit Component" : "Add Component"}</p>
        {error && <p style={{ fontSize: 12.5, color: "#B85450", margin: "0 0 8px" }}>{error}</p>}

        <p style={sec}>Identity</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
          <div><label style={ls}>Component type *</label><input value={f.component_type} onChange={(e) => set("component_type", e.target.value)} required style={fs} /></div>
          <div><label style={ls}>Description</label><input value={f.description} onChange={(e) => set("description", e.target.value)} style={fs} /></div>
          <div>
            <label style={ls}>Category</label>
            <select value={f.category} onChange={(e) => set("category", e.target.value)} style={fs}>
              <option value="">-</option>
              {CM_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div><label style={ls}>Subcategory</label><input value={f.subcategory} onChange={(e) => set("subcategory", e.target.value)} style={fs} /></div>
          <div><label style={ls}>OEM</label><input value={f.oem} onChange={(e) => set("oem", e.target.value)} style={fs} /></div>
          <div><label style={ls}>OEM part no.</label><input value={f.oem_part_no} onChange={(e) => set("oem_part_no", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Alt part no.</label><input value={f.alt_part_no} onChange={(e) => set("alt_part_no", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Supplier</label><input value={f.supplier} onChange={(e) => set("supplier", e.target.value)} style={fs} /></div>
        </div>

        <p style={sec}>Life &amp; thresholds</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 8 }}>
          <div><label style={ls}>Target life (hrs)</label><input type="number" step="1" value={f.target_life_hours} onChange={(e) => set("target_life_hours", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Min expected</label><input type="number" step="1" value={f.min_expected_hours} onChange={(e) => set("min_expected_hours", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Max expected</label><input type="number" step="1" value={f.max_expected_hours} onChange={(e) => set("max_expected_hours", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Warning hrs</label><input type="number" step="1" value={f.warning_hours} onChange={(e) => set("warning_hours", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Critical hrs</label><input type="number" step="1" value={f.critical_hours} onChange={(e) => set("critical_hours", e.target.value)} style={fs} /></div>
        </div>
        <p style={{ fontSize: 11.5, color: "#859195", margin: "4px 0 0" }}>
          The warning and critical PERCENTAGES that drive status are set per site in component_life_settings, not here. Without a target life this component reports LIFE TARGET NOT CONFIGURED rather than a percentage.
        </p>

        <p style={sec}>Cost &amp; warranty</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
          <div><label style={ls}>Expected cost</label><input type="number" step="0.01" value={f.expected_cost} onChange={(e) => set("expected_cost", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Warranty hours</label><input type="number" step="1" value={f.warranty_hours} onChange={(e) => set("warranty_hours", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Warranty months</label><input type="number" step="1" value={f.warranty_months} onChange={(e) => set("warranty_months", e.target.value)} style={fs} /></div>
        </div>

        <p style={sec}>Applies to</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 8 }}>
          <div><label style={ls}>Equipment type</label><input value={f.applies_equip_type} onChange={(e) => set("applies_equip_type", e.target.value)} placeholder="Dump Truck" style={fs} /></div>
          <div><label style={ls}>Manufacturer</label><input value={f.applies_make} onChange={(e) => set("applies_make", e.target.value)} placeholder="Komatsu" style={fs} /></div>
          <div><label style={ls}>Model</label><input value={f.applies_model} onChange={(e) => set("applies_model", e.target.value)} placeholder="HD785" style={fs} /></div>
          <div><label style={ls}>Default position</label><input value={f.default_position} onChange={(e) => set("default_position", e.target.value)} placeholder="LH / RH" style={fs} /></div>
          <div>
            <label style={ls}>Criticality</label>
            <select value={f.criticality} onChange={(e) => set("criticality", e.target.value)} style={fs}>
              {CM_CRITICALITY.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>
        <div style={{ display: "flex", gap: 18, marginTop: 10, flexWrap: "wrap" }}>
          <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#183642" }}>
            <input type="checkbox" checked={f.active} onChange={(e) => set("active", e.target.checked)} /> Active
          </label>
          <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#183642" }}>
            <input type="checkbox" checked={f.uses_positions} onChange={(e) => set("uses_positions", e.target.checked)} /> Tracked by position (LH / RH)
          </label>
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 18px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer" }}>
            {saving ? "Saving…" : "Save Component"}
          </button>
        </div>
      </form>
    </div>
  );
}

function ComponentInstallForm({ master, assets, selectedSiteId, myFullName, onClose, onSaved }) {
  const [assetId, setAssetId] = useState(assets[0]?.asset_id || "");
  const [meterHours, setMeterHours] = useState(null);
  const [f, setF] = useState({
    position_label: master.default_position || "", serial_number: "", part_no: master.oem_part_no || "",
    supplier: master.supplier || "", condition: "New",
    installed_date: todayForInput(), installed_equip_hours: "", hours_at_install: "0",
    purchase_cost: master.expected_cost ?? "",
    install_work_order: "", install_reason: "",
    warranty_end: "", warranty_hours: master.warranty_hours ?? "",
    purchase_order: "", grn_ref: "", technician: myFullName || "", notes: "",
  });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  // Prefill the meter reading from the machine's own history rather than
  // asking someone to remember it.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      if (!assetId) return;
      const { data } = await supabase.from("equipment_current_hours").select("current_equip_hours").eq("asset_id", assetId).maybeSingle();
      if (cancelled) return;
      setMeterHours(data?.current_equip_hours ?? null);
      setF((p) => ({ ...p, installed_equip_hours: data?.current_equip_hours ?? p.installed_equip_hours }));
    })();
    return () => { cancelled = true; };
  }, [assetId]);

  const save = async (e) => {
    e.preventDefault();
    if (!assetId) { setError("Select the equipment."); return; }
    if (f.installed_equip_hours === "") { setError("Equipment hours at installation are required - the whole life calculation depends on it."); return; }
    setSaving(true); setError("");
    const num = (v) => (v === "" || v == null ? null : Number(v));
    const { error: err } = await supabase.from("component_installations").insert({
      asset_id: assetId,
      master_id: master.id,
      component_type: master.component_type,
      position_label: f.position_label.trim() || null,
      serial_number: f.serial_number.trim() || null,
      part_no: f.part_no.trim() || null,
      supplier: f.supplier.trim() || null,
      condition_at_install: f.condition,
      installed_date: f.installed_date,
      installed_equip_hours: num(f.installed_equip_hours),
      installed_comp_hours: num(f.hours_at_install) ?? 0,
      purchase_cost: num(f.purchase_cost),
      warranty_end: f.warranty_end || null,
      warranty_hours: num(f.warranty_hours),
      purchase_order: f.purchase_order.trim() || null,
      grn_reference: f.grn_ref.trim() || null,
      technician: f.technician.trim() || null,
      notes: f.notes.trim() || null,
    });
    setSaving(false);
    if (err) {
      setError(err.message.includes("uq_ci_active_position")
        ? "That machine already has an active component in this position. Change out the existing one first."
        : err.message);
    } else onSaved();
  };

  const fs = { width: "100%", padding: "7px 9px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box" };
  const ls = { display: "block", fontSize: 12, fontWeight: 600, color: "#183642", margin: "0 0 4px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(24,54,66,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: 24, zIndex: 100, overflowY: "auto" }}>
      <form onSubmit={save} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 900, maxWidth: "96vw", maxHeight: "90vh", overflowY: "auto" }}>
        <p style={{ margin: "0 0 4px", fontSize: 16, fontWeight: 700, color: NAVY }}>Install {master.component_type}</p>
        <p style={{ margin: "0 0 14px", fontSize: 12, color: "#859195" }}>
          Target life {master.target_life_hours ? `${Number(master.target_life_hours).toLocaleString()} hrs` : "not configured"}
        </p>
        {error && <p style={{ fontSize: 12.5, color: "#B85450", margin: "0 0 8px" }}>{error}</p>}

        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
          <div>
            <label style={ls}>Equipment *</label>
            <select value={assetId} onChange={(e) => setAssetId(e.target.value)} required style={fs}>
              {assets.map((a) => <option key={a.asset_id} value={a.asset_id}>{a.asset_id} - {a.asset_name}</option>)}
            </select>
          </div>
          <div><label style={ls}>Position</label><input value={f.position_label} onChange={(e) => set("position_label", e.target.value)} placeholder="LH / RH / blank" style={fs} /></div>
          <div><label style={ls}>Serial number</label><input value={f.serial_number} onChange={(e) => set("serial_number", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Part no.</label><input value={f.part_no} onChange={(e) => set("part_no", e.target.value)} style={fs} /></div>

          <div><label style={ls}>Install date *</label><input type="date" value={f.installed_date} onChange={(e) => set("installed_date", e.target.value)} required max={todayForInput()} style={fs} /></div>
          <div>
            <label style={ls}>Equipment hours *</label>
            <input type="number" step="0.1" value={f.installed_equip_hours} onChange={(e) => set("installed_equip_hours", e.target.value)} required style={fs} />
            {meterHours != null && <p style={{ fontSize: 11, color: "#859195", margin: "3px 0 0" }}>Meter now: {Number(meterHours).toLocaleString()}</p>}
          </div>
          <div>
            <label style={ls}>Component hrs at fit</label>
            <input type="number" step="0.1" value={f.hours_at_install} onChange={(e) => set("hours_at_install", e.target.value)} style={fs} />
            <p style={{ fontSize: 11, color: "#859195", margin: "3px 0 0" }}>0 for new; the run hours for a reman unit.</p>
          </div>
          <div>
            <label style={ls}>Condition</label>
            <select value={f.condition} onChange={(e) => set("condition", e.target.value)} style={fs}>
              {["New", "Reman", "Repaired"].map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>

          <div><label style={ls}>Purchase cost</label><input type="number" step="0.01" value={f.purchase_cost} onChange={(e) => set("purchase_cost", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Work order</label><input value={f.install_work_order} onChange={(e) => set("install_work_order", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Reason</label><input value={f.install_reason} onChange={(e) => set("install_reason", e.target.value)} placeholder="New build / Changeout" style={fs} /></div>
          <div><label style={ls}>Supplier</label><input value={f.supplier} onChange={(e) => set("supplier", e.target.value)} style={fs} /></div>

          <div><label style={ls}>Warranty end</label><input type="date" value={f.warranty_end} onChange={(e) => set("warranty_end", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Warranty hours</label><input type="number" step="1" value={f.warranty_hours} onChange={(e) => set("warranty_hours", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Purchase order</label><input value={f.purchase_order} onChange={(e) => set("purchase_order", e.target.value)} style={fs} /></div>
          <div><label style={ls}>GRN ref</label><input value={f.grn_ref} onChange={(e) => set("grn_ref", e.target.value)} style={fs} /></div>
        </div>

        <div style={{ marginTop: 8 }}>
          <label style={ls}>Notes</label>
          <input value={f.notes} onChange={(e) => set("notes", e.target.value)} style={fs} />
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 18px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer" }}>
            {saving ? "Saving…" : "Install Component"}
          </button>
        </div>
      </form>
    </div>
  );
}

function ComponentChangeOutForm({ install, selectedSiteId, myFullName, onClose, onSaved }) {
  const [f, setF] = useState({
    removed_date: todayForInput(), removed_equip_hours: install.current_equip_hours ?? "",
    removal_reason: "", failure_code: "", failure_description: "",
    removed_condition: "", removed_destination: "Rebuild",
    repair_cost: "", technician: myFullName || "", supervisor: "",
    fitReplacement: true,
    new_serial: "", new_part_no: install.part_no || "", new_condition: "New",
    new_purchase_cost: "", new_hours_at_install: "0", new_warranty_end: "",
  });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const achieved = useMemo(() => {
    const removedAt = Number(f.removed_equip_hours);
    if (isNaN(removedAt)) return null;
    return (removedAt - Number(install.installed_equip_hours)) + Number(install.installed_comp_hours || 0);
  }, [f.removed_equip_hours, install]);

  const pct = useMemo(() => {
    if (achieved == null || !install.target_life_hours) return null;
    return Math.round((achieved / Number(install.target_life_hours)) * 1000) / 10;
  }, [achieved, install.target_life_hours]);

  const save = async (e) => {
    e.preventDefault();
    if (f.removed_equip_hours === "") { setError("Equipment hours at removal are required."); return; }
    if (Number(f.removed_equip_hours) < Number(install.installed_equip_hours)) {
      setError("Removal hours can't be less than the hours the component was fitted at.");
      return;
    }
    setSaving(true); setError("");
    const num = (v) => (v === "" || v == null ? null : Number(v));
    try {
      // Close the existing row. It is never deleted or overwritten - the
      // history of that physical component stays exactly as it was.
      const { error: e1 } = await supabase.from("component_installations").update({
        removed_date: f.removed_date,
        removed_equip_hours: num(f.removed_equip_hours),
        removal_reason: f.removal_reason.trim() || null,
        failure_code: f.failure_code.trim() || null,
        failure_description: f.failure_description.trim() || null,
        condition_at_removal: f.removed_condition.trim() || null,
        destination: f.removed_destination || null,
        repair_cost: num(f.repair_cost),
        technician: f.technician.trim() || null,
        supervisor: f.supervisor.trim() || null,
      }).eq("id", install.id);
      if (e1) throw e1;

      await supabase.from("component_audit").insert({
        installation_id: install.id, asset_id: install.asset_id, action: "CHANGE OUT",
        field_name: "removed_date", old_value: null, new_value: f.removed_date,
        reason: f.removal_reason.trim() || null,
      });

      if (f.fitReplacement) {
        const { error: e2 } = await supabase.from("component_installations").insert({
          asset_id: install.asset_id,
          master_id: install.master_id,
          component_type: install.component_type,
          position_label: install.position_label,
          serial_number: f.new_serial.trim() || null,
          part_no: f.new_part_no.trim() || null,
          supplier: install.supplier,
          condition_at_install: f.new_condition,
          installed_date: f.removed_date,
          installed_equip_hours: num(f.removed_equip_hours),
          installed_comp_hours: num(f.new_hours_at_install) ?? 0,
          purchase_cost: num(f.new_purchase_cost),
          warranty_end: f.new_warranty_end || null,
          technician: f.technician.trim() || null,
        });
        if (e2) throw e2;
      }
      onSaved();
    } catch (err) {
      setError(err.message || String(err));
    }
    setSaving(false);
  };

  const fs = { width: "100%", padding: "7px 9px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box" };
  const ls = { display: "block", fontSize: 12, fontWeight: 600, color: "#183642", margin: "0 0 4px" };
  const sec = { fontSize: 12, fontWeight: 700, color: "#4B5659", margin: "14px 0 6px" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(24,54,66,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: 24, zIndex: 100, overflowY: "auto" }}>
      <form onSubmit={save} style={{ background: "#fff", borderRadius: 12, padding: 24, width: 900, maxWidth: "96vw", maxHeight: "90vh", overflowY: "auto" }}>
        <p style={{ margin: "0 0 4px", fontSize: 16, fontWeight: 700, color: NAVY }}>
          Change out {install.component_type}{install.position_label ? ` (${install.position_label})` : ""} on {install.asset_id}
        </p>
        <p style={{ margin: "0 0 14px", fontSize: 12, color: "#859195" }}>
          Fitted {install.installed_date} at {Number(install.installed_equip_hours).toLocaleString()} equipment hours
        </p>
        {error && <p style={{ fontSize: 12.5, color: "#B85450", margin: "0 0 8px" }}>{error}</p>}

        <p style={sec}>Removal</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
          <div><label style={ls}>Removal date *</label><input type="date" value={f.removed_date} onChange={(e) => set("removed_date", e.target.value)} required max={todayForInput()} style={fs} /></div>
          <div><label style={ls}>Equipment hours *</label><input type="number" step="0.1" value={f.removed_equip_hours} onChange={(e) => set("removed_equip_hours", e.target.value)} required style={fs} /></div>
          <div><label style={ls}>Reason for removal</label><input value={f.removal_reason} onChange={(e) => set("removal_reason", e.target.value)} placeholder="Failure / Life expired / Planned" style={fs} /></div>
          <div><label style={ls}>Failure code</label><input value={f.failure_code} onChange={(e) => set("failure_code", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Failure description</label><input value={f.failure_description} onChange={(e) => set("failure_description", e.target.value)} style={fs} /></div>
          <div><label style={ls}>Condition at removal</label><input value={f.removed_condition} onChange={(e) => set("removed_condition", e.target.value)} style={fs} /></div>
          <div>
            <label style={ls}>Destination</label>
            <select value={f.removed_destination} onChange={(e) => set("removed_destination", e.target.value)} style={fs}>
              {CI_DESTINATIONS.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
          <div><label style={ls}>Repair / rebuild cost</label><input type="number" step="0.01" value={f.repair_cost} onChange={(e) => set("repair_cost", e.target.value)} style={fs} /></div>
        </div>

        {achieved != null && (
          <p style={{ margin: "10px 0 0", fontSize: 12.5, padding: "8px 10px", borderRadius: 8,
                      background: pct != null && pct < 70 ? "#F6E2E0" : "#E2EFE9",
                      color: pct != null && pct < 70 ? "#8A2F28" : "#2C5646" }}>
            Life achieved: {achieved.toLocaleString(undefined, { maximumFractionDigits: 1 })} hrs
            {pct != null ? ` — ${pct}% of target${pct < 70 ? ", a premature failure worth a warranty claim" : ""}` : " — no target set, so no percentage"}
          </p>
        )}

        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#183642", margin: "14px 0 0" }}>
          <input type="checkbox" checked={f.fitReplacement} onChange={(e) => set("fitReplacement", e.target.checked)} />
          Fit a replacement now
        </label>

        {f.fitReplacement && (
          <>
            <p style={sec}>Replacement</p>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 8 }}>
              <div><label style={ls}>Serial in</label><input value={f.new_serial} onChange={(e) => set("new_serial", e.target.value)} style={fs} /></div>
              <div><label style={ls}>Part no.</label><input value={f.new_part_no} onChange={(e) => set("new_part_no", e.target.value)} style={fs} /></div>
              <div>
                <label style={ls}>Condition</label>
                <select value={f.new_condition} onChange={(e) => set("new_condition", e.target.value)} style={fs}>
                  {["New", "Reman", "Repaired"].map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div><label style={ls}>Hrs at fit</label><input type="number" step="0.1" value={f.new_hours_at_install} onChange={(e) => set("new_hours_at_install", e.target.value)} style={fs} /></div>
              <div><label style={ls}>Purchase cost</label><input type="number" step="0.01" value={f.new_purchase_cost} onChange={(e) => set("new_purchase_cost", e.target.value)} style={fs} /></div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginTop: 8 }}>
              <div><label style={ls}>Warranty end</label><input type="date" value={f.new_warranty_end} onChange={(e) => set("new_warranty_end", e.target.value)} style={fs} /></div>
              <div><label style={ls}>Technician</label><input value={f.technician} onChange={(e) => set("technician", e.target.value)} style={fs} /></div>
              <div><label style={ls}>Supervisor</label><input value={f.supervisor} onChange={(e) => set("supervisor", e.target.value)} style={fs} /></div>
            </div>
            <p style={{ fontSize: 11.5, color: "#859195", margin: "6px 0 0" }}>
              The replacement is fitted at the same equipment hours as the removal, so its life starts from there.
            </p>
          </>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
          <button type="button" onClick={onClose} style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
          <button type="submit" disabled={saving} style={{ background: NAVY, color: "#fff", border: "none", padding: "8px 18px", borderRadius: 8, fontSize: 13.5, fontWeight: 700, cursor: saving ? "default" : "pointer" }}>
            {saving ? "Saving…" : "Record Change Out"}
          </button>
        </div>
      </form>
    </div>
  );
}

const COMPONENT_STATUS_COLUMNS = [
  ["asset_id", "Equipment", "", "text"],
  ["asset_name", "Name", "", "text"],
  ["model", "Model", "", "text"],
  ["component_type", "Component", "", "text"],
  ["position_label", "Position", "", "text"],
  ["oem", "OEM", "", "text"],
  ["serial_number", "Serial No", "", "text"],

  ["install_work_order", "Work Order", "Last Changeout", "text"],
  ["installed_equip_hours", "Equip. Hrs", "Last Changeout", "n1"],
  ["installed_comp_hours", "Comp. Hrs at Fit", "Last Changeout", "n1"],
  ["installed_date", "Fitted", "Last Changeout", "date"],

  ["current_equip_hours", "Equip. Hrs", "Current", "n1"],
  ["target_life_hours", "Target Life", "Current", "n1"],
  ["component_hours", "Comp. Hrs", "Current", "n1"],
  ["remaining_hours", "Remaining", "Current", "n1"],
  ["life_used_pct", "Life Used", "Current", "pct100"],
  ["overdue_hours", "Overdue", "Current", "n1"],
  ["status", "Status", "Current", "status"],

  ["next_changeout_equip_hours", "Equip. Hrs", "Next Changeout", "n1"],
  ["estimated_changeout_date", "Est. Date", "Next Changeout", "dateOrText"],
  ["estimated_days_remaining", "Days Left", "Next Changeout", "int"],

  ["purchase_cost", "Purchase", "Cost", "money"],
  ["rebuild_cost", "Rebuild", "Cost", "money"],
  ["lifetime_cost", "Lifetime", "Cost", "money"],
  ["cost_per_hour", "Per Hour", "Cost", "money"],
  ["uvic", "UVIC", "Cost", "money"],
  ["rvic", "RVIC", "Cost", "money"],

  ["warranty_status", "Status", "Warranty", "text"],
  ["warranty_end", "Expiry", "Warranty", "date"],
];

// Module level so its identity is stable across renders - useTablePrefs
// memoises against this array, and rebuilding it each render would make
// every memo inside the hook recompute continuously.
const COMPONENT_STATUS_PREF_COLUMNS = COMPONENT_STATUS_COLUMNS.map(([key, label]) => ({ key, label }));

function ComponentStatusPage({ assets, selectedSiteId }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [query, setQuery] = useState("");
  const [selectedFleet, setSelectedFleet] = useState("");
  const [selectedAsset, setSelectedAsset] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [typeFilter, setTypeFilter] = useState("");
  const [hideRemoved, setHideRemoved] = useState(true);
  const [groupByMachine, setGroupByMachine] = useState(true);
  const prefs = useTablePrefs("component_status", COMPONENT_STATUS_PREF_COLUMNS);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      const { data, error: err } = await supabase
        .from("component_status_calc")
        .select("*")
        .order("asset_id", { ascending: true });
      if (cancelled) return;
      if (err) setError(err.message);
      else { setRows(data || []); setError(""); }
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, [selectedSiteId]);

  const filtered = useMemo(() => {
    let r = rows;
    if (hideRemoved) r = r.filter((x) => x.status !== "REMOVED");
    if (selectedAsset) r = r.filter((x) => x.asset_id === selectedAsset);
    else if (selectedFleet) r = r.filter((x) => x.fleet === selectedFleet);
    if (statusFilter) r = r.filter((x) => x.status === statusFilter);
    if (typeFilter) r = r.filter((x) => x.component_type === typeFilter);
    if (query.trim()) {
      const q = query.toLowerCase();
      r = r.filter((x) => Object.values(x).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }
    return prefs.sortRows(r);
  }, [rows, hideRemoved, selectedAsset, selectedFleet, statusFilter, typeFilter, query, prefs.sortRows]); // eslint-disable-line react-hooks/exhaustive-deps

  const groups = useMemo(() => buildMachineGroups(filtered, assets, {
    enabled: groupByMachine, dateKey: "installed_date", hoursKey: null,
    sortKey: prefs.sortKey, sortDir: prefs.sortDir, sortRows: prefs.sortRows,
  }), [filtered, assets, groupByMachine, prefs.sortKey, prefs.sortDir, prefs.sortRows]);

  // Totals follow the filters (§30) - a filtered report whose totals
  // still describe everything is worse than no totals.
  const totals = useMemo(() => {
    const n = (k) => filtered.map((r) => Number(r[k])).filter((v) => !isNaN(v));
    const sum = (k) => n(k).reduce((a, b) => a + b, 0);
    const count = (s) => filtered.filter((r) => r.status === s).length;
    return {
      rows: filtered.length,
      overdue: count("OVERDUE"), due: count("DUE"), approaching: count("APPROACHING LIFE"),
      unconfigured: count("LIFE TARGET NOT CONFIGURED"), noMeter: count("NO METER DATA"),
      lifetime: sum("lifetime_cost"), rvic: sum("rvic"), uvic: sum("uvic"),
    };
  }, [filtered]);

  // Exports precisely what the filters have left on screen, in the
  // column order the user arranged - not the full unfiltered table.
  const exportReport = () =>
    exportMine2UReport({
      title: "Component Status",
      sheetName: "Component Status",
      fileName: "Mine2U_Component_Status",
      columns: prefs.columns.map((c) => {
        const def = COMPONENT_STATUS_COLUMNS.find((x) => x[0] === c.key);
        return { key: c.key, label: def ? `${def[2] ? def[2] + " - " : ""}${def[1]}` : c.label, fmt: def?.[3] === "status" ? "text" : (def?.[3] || "text") };
      }),
      rows: filtered,
      filterLines: [
        ["Hierarchy", `[Fleet: ${selectedFleet || "All"}]  [Equipment: ${selectedAsset || "All"}]`],
        ["Additional Filters", `[Status: ${statusFilter || "All"}]  [Component: ${typeFilter || "All"}]  [Removed: ${hideRemoved ? "hidden" : "shown"}]`],
      ],
      notes: [
        "Component hours, life used, status, next changeout, UVIC and RVIC are calculated from equipment meter readings at the moment of export - none of these figures is stored.",
        "UVIC = parts cost x (component hours / target life), capped at the full cost once life is reached. RVIC = parts cost - UVIC.",
        "Where utilisation history is insufficient, the estimated changeout date is left blank rather than estimated.",
      ],
    });

  const statusOptions = useMemo(() => [...new Set(rows.map((r) => r.status).filter(Boolean))].sort(), [rows]);
  const typeOptions = useMemo(() => [...new Set(rows.map((r) => r.component_type).filter(Boolean))].sort(), [rows]);

  const money = (v) => (v == null || v === "" ? "-" : Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
  const fmt = (key, val, row) => {
    const type = COMPONENT_STATUS_COLUMNS.find((c) => c[0] === key)?.[3];
    if (val == null || val === "") {
      // A missing estimate is a statement, not an empty cell.
      if (key === "estimated_changeout_date" && row?.status !== "REMOVED") return "Insufficient utilisation data";
      return "-";
    }
    switch (type) {
      case "n1": return Number(val).toLocaleString(undefined, { maximumFractionDigits: 1 });
      case "int": return String(Math.round(Number(val)));
      case "pct100": return `${Number(val).toFixed(2)}%`;
      case "money": return money(val);
      case "date": case "dateOrText": return String(val).slice(0, 10);
      default: return val;
    }
  };

  const cellStyle = (type) => ({ padding: "6px 7px", whiteSpace: "nowrap", textAlign: type === "text" || type === "status" ? "left" : "right", overflow: "hidden", textOverflow: "ellipsis" });
  const widthFor = (key, type) => {
    if (key === "asset_id" || key === "asset_name" || key === "component_type") return 110;
    if (key === "estimated_changeout_date") return 130;
    if (key === "status") return 130;
    if (type === "text") return 82;
    if (type === "money") return 84;
    return 68;
  };

  const spans = [];
  COMPONENT_STATUS_COLUMNS.forEach(([, , group]) => {
    const last = spans[spans.length - 1];
    if (last && last.group === group) last.span += 1;
    else spans.push({ group, span: 1 });
  });

  const chip = (label, value, tone) => (
    <div style={{ background: tone.bg, color: tone.text, borderRadius: 8, padding: "7px 12px", fontSize: 12.5 }}>
      <strong style={{ fontSize: 15 }}>{value}</strong> {label}
    </div>
  );

  return (
    <div>

      {error && (
        <p style={{ fontSize: 13, color: "#B85450", background: "#F6E2E0", padding: "9px 12px", borderRadius: 8 }}>
          {error.includes("does not exist")
            ? "The Component Life tables aren't in the database yet - run component_life_foundation.sql first."
            : error}
        </p>
      )}

      <FleetEquipmentFilter assets={assets} selectedFleet={selectedFleet} setSelectedFleet={setSelectedFleet} selectedAsset={selectedAsset} setSelectedAsset={setSelectedAsset} />

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 260 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#859195" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search components"
            style={{ width: "100%", padding: "8px 10px 8px 32px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8, outline: "none" }} />
        </div>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
          <option value="">All statuses</option>
          {statusOptions.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} style={{ padding: "8px 10px", fontSize: 13, border: "1px solid #E2E6E3", borderRadius: 8 }}>
          <option value="">All component types</option>
          {typeOptions.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#4B5659", cursor: "pointer" }}>
          <input type="checkbox" checked={hideRemoved} onChange={(e) => setHideRemoved(e.target.checked)} /> Hide removed
        </label>
        <GroupByMachineToggle value={groupByMachine} onChange={setGroupByMachine} />
        <ColumnsButton prefs={prefs} />
        <button onClick={() => { setQuery(""); setStatusFilter(""); setTypeFilter(""); setSelectedFleet(""); setSelectedAsset(""); }}
          style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#4B5659", fontSize: 12.5, fontWeight: 600, padding: "8px 12px", borderRadius: 8, cursor: "pointer" }}>
          Clear filters
        </button>
        <button onClick={exportReport}
          style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${NAVY}`, color: NAVY, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" }}>
          <Download size={14} /> Export to Excel
        </button>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
        {chip("components", totals.rows, { bg: "#F2F1EA", text: "#183642" })}
        {chip("overdue", totals.overdue, { bg: "#F6E2E0", text: "#7A3330" })}
        {chip("due", totals.due, { bg: "#F5E9D8", text: "#7A5A22" })}
        {chip("approaching life", totals.approaching, { bg: "#F5E9D8", text: "#7A5A22" })}
        {totals.unconfigured > 0 && chip("no life target", totals.unconfigured, { bg: "#EDEBE4", text: "#4B5659" })}
        {totals.noMeter > 0 && chip("no meter data", totals.noMeter, { bg: "#EDEBE4", text: "#4B5659" })}
        {chip("remaining value", money(totals.rvic), { bg: "#E2EFE9", text: "#2C5646" })}
      </div>

      {loading ? (
        <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
      ) : (
        <>
          <p style={{ fontSize: 11.5, color: "#859195", margin: "0 0 6px" }}>
            Every figure is calculated from meter readings when this page loads - nothing here is stored. Totals follow the filters above.
          </p>
          <div style={{ overflow: "auto", maxHeight: "60vh", border: "1px solid #E2E6E3", borderRadius: 10, position: "relative" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
              <colgroup>
                {prefs.columns.map((c) => {
                  const def = COMPONENT_STATUS_COLUMNS.find((x) => x[0] === c.key);
                  return <col key={c.key} style={{ width: widthFor(c.key, def?.[3]) }} />;
                })}
              </colgroup>
              <thead>
                <tr style={{ background: "#E9ECEA" }}>
                  {spans.map((sp, i) => (
                    <th key={i} colSpan={sp.span} title={sp.group}
                      style={{ position: "sticky", top: 0, zIndex: 3, background: "#E9ECEA", textAlign: "center", padding: "5px 6px",
                               fontWeight: 700, fontSize: 11, color: NAVY, borderBottom: "1px solid #E2E6E3",
                               borderLeft: i > 0 ? "1px solid #DCE0DD" : "none", whiteSpace: "nowrap" }}>
                      {sp.group}
                    </th>
                  ))}
                </tr>
                <tr style={{ background: "#F7F8F6" }}>
                  {prefs.columns.map((c) => {
                    const def = COMPONENT_STATUS_COLUMNS.find((x) => x[0] === c.key);
                    const type = def?.[3];
                    return (
                      <th key={c.key} onClick={() => prefs.toggleSort(c.key)} title={c.label}
                        style={{ position: "sticky", top: 26, zIndex: c.key === "asset_id" ? 5 : 3, background: "#F7F8F6",
                                 ...(c.key === "asset_id" ? { left: 0 } : {}),
                                 textAlign: type === "text" || type === "status" ? "left" : "right", padding: "6px 7px",
                                 fontWeight: 600, color: "#4B5659", fontSize: 11, lineHeight: 1.25,
                                 borderBottom: "1px solid #E2E6E3", cursor: "pointer", userSelect: "none",
                                 whiteSpace: "normal", overflowWrap: "normal", hyphens: "none" }}>
                        {c.label}{prefs.sortKey === c.key ? (prefs.sortDir === "asc" ? " ▲" : " ▼") : ""}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {groups.map((g, gi) => (
                  <React.Fragment key={g.key}>
                    {g.banner && (
                      <tr>
                        <td colSpan={prefs.columns.length} style={{ padding: "7px 10px", background: "#F2F1EA", fontWeight: 700, fontSize: 12, color: NAVY, borderTop: gi > 0 ? "1px solid #E2E6E3" : "none" }}>
                          {g.assetId}{g.assetName ? ` - ${g.assetName}` : ""}
                          <span style={{ fontWeight: 500, color: "#4B5659" }}>{`  ·  ${g.rows.length} component${g.rows.length === 1 ? "" : "s"}`}</span>
                        </td>
                      </tr>
                    )}
                    {g.rows.map((row, i) => {
                      const tone = statusColor(row.status);
                      return (
                        <tr key={row.id ?? `${g.key}-${i}`} style={{ borderBottom: "1px solid #EFEEE7" }}>
                          {prefs.columns.map((c) => {
                            const def = COMPONENT_STATUS_COLUMNS.find((x) => x[0] === c.key);
                            const type = def?.[3];
                            const value = fmt(c.key, row[c.key], row);
                            return (
                              <td key={c.key} title={String(row[c.key] ?? "")}
                                style={{ ...cellStyle(type), ...(c.key === "asset_id" ? { position: "sticky", left: 0, zIndex: 2, background: "#fff" } : {}),
                                         color: c.key === "overdue_hours" && Number(row.overdue_hours) > 0 ? "#B85450" : "#183642",
                                         fontWeight: c.key === "overdue_hours" && Number(row.overdue_hours) > 0 ? 700 : 400 }}>
                                {type === "status"
                                  ? <span style={{ background: tone.bg, color: tone.text, padding: "2px 8px", borderRadius: 6, fontSize: 11, fontWeight: 600, whiteSpace: "nowrap" }}>{value}</span>
                                  : value}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </React.Fragment>
                ))}
                {filtered.length === 0 && (
                  <tr><td colSpan={prefs.columns.length} style={{ padding: 20, textAlign: "center", color: "#859195" }}>
                    {rows.length === 0 ? "No components installed yet." : "No components match the current filters."}
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}
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

function AssetsPage({ assets, selectedSiteId, onRefresh, isAdmin, mySites = [] }) {
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
          isAdmin={isAdmin}
          mySites={mySites}
          onClose={() => { setShowForm(false); setEditing(null); }}
          onSaved={handleSaved}
        />
      )}
    </div>
  );
}


function KpiRow({ metrics }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(120px,100%),1fr))", gap: 12, marginBottom: 20 }}>
      {METRIC_DEFS.map((m) => (
        <MetricCard key={m.key} label={m.label} value={metrics[m.key] != null ? m.fmt(metrics[m.key]) : "-"} />
      ))}
    </div>
  );
}

function aggregateMetrics(kpiRows, assetIds) {
  const rows = kpiRows.filter((r) => assetIds.includes(r.asset_id));
  if (rows.length === 0) return {};
  // Missing values must be EXCLUDED, not treated as zero. Number(null)
  // is 0, not NaN, so an isNaN filter alone silently counts a machine
  // with no data as a 0% machine and drags its fleet's average down.
  const numRows = (key) => rows
    .filter((r) => r[key] !== null && r[key] !== undefined && r[key] !== "")
    .map((r) => Number(r[key]))
    .filter((v) => !isNaN(v));
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
    availability_24h: avg("availability_24h"),
    mtbf_24h: avg("mtbf_24h"),
    mttr_24h: avg("mttr_24h"),
    utilisation_24h: avg("utilisation_24h"),
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
async function logActivity(tableName, recordId, action, summary, assetId) {
  try {
    const { data: { user } } = await supabase.auth.getUser();
    await supabase.from("activity_log").insert({
      table_name: tableName,
      record_id: recordId != null ? String(recordId) : null,
      action, summary,
      user_id: user?.id ?? null,
      asset_id: assetId ?? null,
    });
  } catch (err) {
    console.error("Activity log write failed:", err);
  }
}

async function deleteWithReason(tableName, recordId, idColumn, reason, userEmail, assetId) {
  const { error: logError } = await supabase.from("deletion_log").insert({
    table_name: tableName, record_id: String(recordId), reason, deleted_by_email: userEmail || null,
  });
  if (logError) throw logError;
  const { error: deleteError } = await supabase.from(tableName).delete().eq(idColumn, recordId);
  if (deleteError) throw deleteError;
  logActivity(TABLE_NAME_LABELS[tableName] || tableName, recordId, "deleted", reason, assetId);
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
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(200px,100%),1fr))", gap: 14 }}>
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

// Two-series charts use fixed colours per series (like the Excel report)
// rather than target-based colouring, which becomes unreadable once two
// bars sit side by side. The target line carries the pass/fail meaning.
const CHART_COLOURS = {
  target:      "#F5C518",  // target reference line
  met:         "#2F9E63",  // single-series bar meeting target
  missed:      "#A62A1E",  // single-series bar missing target
  // Two-series bars take the brand teal and a gold, both as gradients so
  // they read as part of the same system as the MTBF/MTTR bars. Green is
  // deliberately NOT reused here - on a two-series chart colour means
  // "which series", and green already means "target met" elsewhere.
  last24h:      "#1F6668",
  last24hLight: "#5FB3B0",
  period:       "#C79A12",
  periodLight:  "#F2D46B",
  metLight:    "#8FE8B4",  // gradient top, target met
  missedLight: "#F3998A",  // gradient top, target missed
  gridline:    "#3D3D3D",
  axisText:    "#CFCFCF",
  panel:       "#292929",
  panelBorder: "#1C1C1C",
};
const SERIES_1_COLOUR = CHART_COLOURS.last24h;
const SERIES_2_COLOUR = CHART_COLOURS.period;

function KpiBarChart({ title, data, xKey, dataKey, dataKey2, seriesName, seriesName2, target, domainMax, valueFormatter, meetsTarget, meetsTarget2, unitSuffix, onClick, onBarClick }) {
  const gradGreen = `grad-green-${dataKey}`;
  const gradRed = `grad-red-${dataKey}`;
  const gradS1 = `grad-s1-${dataKey}`;
  const gradS2 = `grad-s2-${dataKey}`;
  const niceMax = domainMax === 100 ? 100 : niceDomainMax(domainMax);

  // Bars are always upright columns now, never sideways - on a phone,
  // a chart with many machines (e.g. 17 dump trucks) scrolls
  // horizontally instead, keeping each bar wide enough to read rather
  // than switching orientation.
  const isMobile = useIsMobile();
  const minBarWidth = isMobile ? 62 : 46;
  const needsScroll = isMobile && data.length > 6;
  const chartWidth = needsScroll ? Math.max(data.length * minBarWidth, 320) : "100%";

  return (
    <div>
      {title && <p style={{ fontSize: 12.5, fontWeight: 600, margin: "0 0 6px", color: "#4B5659" }}>{title}</p>}
      <div
        onClick={onBarClick ? undefined : onClick}
        style={{ height: 240, background: CHART_COLOURS.panel, border: `1px solid ${CHART_COLOURS.panelBorder}`, borderRadius: 10, padding: "18px 8px 4px", cursor: (onClick || onBarClick) ? "pointer" : "default", overflowX: needsScroll ? "auto" : "hidden", WebkitOverflowScrolling: "touch" }}
      >
        <div style={{ width: chartWidth, height: "100%" }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={data} margin={{ top: 20, right: 8, left: 0, bottom: isMobile ? 34 : 20 }}>
              <defs>
                <linearGradient id={gradGreen} x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={CHART_COLOURS.metLight} />
                  <stop offset="100%" stopColor={CHART_COLOURS.met} />
                </linearGradient>
                <linearGradient id={gradRed} x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={CHART_COLOURS.missedLight} />
                  <stop offset="100%" stopColor={CHART_COLOURS.missed} />
                </linearGradient>
                <linearGradient id={gradS1} x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={CHART_COLOURS.last24hLight} />
                  <stop offset="100%" stopColor={CHART_COLOURS.last24h} />
                </linearGradient>
                <linearGradient id={gradS2} x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={CHART_COLOURS.periodLight} />
                  <stop offset="100%" stopColor={CHART_COLOURS.period} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={CHART_COLOURS.gridline} vertical={false} />
              <XAxis dataKey={xKey} interval={0} angle={-40} textAnchor="end" height={isMobile ? 54 : 44} tick={{ fontSize: 10, fill: CHART_COLOURS.axisText }} axisLine={{ stroke: "#4A4A4A" }} tickLine={false} tickFormatter={(v) => truncateLabel(v, isMobile ? 12 : 14)} />
              <YAxis domain={[0, niceMax]} tick={{ fontSize: 11, fill: CHART_COLOURS.axisText }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}${unitSuffix || ""}`} />
              <Tooltip formatter={(v) => `${Number(v).toFixed(1)}${unitSuffix || ""}`} contentStyle={{ background: "#1F1F1F", border: "1px solid #444", borderRadius: 8 }} labelStyle={{ color: "#fff" }} itemStyle={{ color: "#fff" }} />
              <Bar dataKey={dataKey} name={seriesName} radius={[3, 3, 0, 0]} cursor={(onClick || onBarClick) ? "pointer" : "default"} onClick={onBarClick ? (row) => onBarClick(row) : undefined}>
                {data.map((row, i) => (
                  <Cell key={i} fill={dataKey2 ? `url(#${gradS1})` : (meetsTarget(row) ? `url(#${gradGreen})` : `url(#${gradRed})`)} />
                ))}
                <LabelList dataKey={dataKey} position="top" formatter={valueFormatter} fill="#fff" fontSize={11} fontWeight={700} />
              </Bar>
              {dataKey2 && (
                <Bar dataKey={dataKey2} name={seriesName2} fill={`url(#${gradS2})`} radius={[3, 3, 0, 0]} cursor={(onClick || onBarClick) ? "pointer" : "default"} onClick={onBarClick ? (row) => onBarClick(row) : undefined}>
                  <LabelList dataKey={dataKey2} position="top" formatter={valueFormatter} fill="#fff" fontSize={10} fontWeight={700} />
                </Bar>
              )}
              <ReferenceLine y={target} stroke={CHART_COLOURS.target} strokeWidth={2.5} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
      {needsScroll && <p style={{ fontSize: 11, color: "#859195", margin: "6px 2px 0" }}>Swipe sideways to see every machine.</p>}
    </div>
  );
}

function KpiLegend({ targetLabel, showSeries }) {
  return (
    <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 11.5, color: "#4B5659", flexWrap: "wrap" }}>
      {showSeries && (
        <>
          <span><span style={{ display: "inline-block", width: 10, height: 10, background: SERIES_1_COLOUR, borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Last 24 hours (Utilisation &amp; Availability)</span>
          <span><span style={{ display: "inline-block", width: 10, height: 10, background: SERIES_2_COLOUR, borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Selected period</span>
        </>
      )}
      <span><span style={{ display: "inline-block", width: 10, height: 10, background: CHART_COLOURS.met, borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Target met</span>
      <span><span style={{ display: "inline-block", width: 10, height: 10, background: CHART_COLOURS.missed, borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />Target not met</span>
      <span><span style={{ display: "inline-block", width: 12, height: 3, background: CHART_COLOURS.target, borderRadius: 2, marginRight: 6, verticalAlign: "middle" }} />{targetLabel || "Target"}</span>
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

function AboutPage({ assets, breakdowns, workOrders, plannedMaintenance, components, parts, inspections, onNavigate }) {
  const TEAL = "#1F6668";
  const MINT = "#8FE0D6";
  const INK = "#183642";
  const BODY = "#4B5659";
  const MUTED = "#859195";
  const LINE = "#E5EBEC";
  const isMobile = useIsMobile();

  const heroRef = React.useRef(null);

  const onHeroMove = (e) => {
    const host = heroRef.current;
    if (!host) return;
    const b = host.getBoundingClientRect();
    const dx = (e.clientX - b.left) / b.width - 0.5;
    const dy = (e.clientY - b.top) / b.height - 0.5;
    host.querySelectorAll("[data-depth]").forEach((el) => {
      const d = parseFloat(el.getAttribute("data-depth")) || 0;
      el.style.transform = `translate(${(-dx * 30 * d).toFixed(2)}px,${(-dy * 20 * d).toFixed(2)}px)`;
    });
  };
  const onHeroLeave = () => {
    const host = heroRef.current;
    if (!host) return;
    host.querySelectorAll("[data-depth]").forEach((el) => { el.style.transform = ""; });
  };

  const PIT = "430,180 520,92 660,72 800,140 850,262 780,400 630,460 490,392";
  const ring = (s, o, w) => (
    <polygon
      points={PIT}
      transform={s === 1 ? undefined : `translate(640,266) scale(${s}) translate(-640,-266)`}
      strokeOpacity={o}
      strokeWidth={w}
    />
  );

  const valueProps = [
    { label: "Live data", text: "Always up to date", dot: true },
    { label: "One system", text: "Everything connected", Icon: Boxes },
    { label: "Built for mining", text: "Reliable. Scalable. Secure.", Icon: ShieldCheck },
  ];

  const features = [
    { Icon: Truck, title: "Equipment & Events", text: "Track every asset, breakdown and planned event with full linked work-order history.", nav: "assets" },
    { Icon: Clock, title: "Daily Operations", text: "Manage operating hours, fuel, oil consumption and daily service compliance.", nav: "daily_hours" },
    { Icon: Wrench, title: "Maintenance & Backlogs", text: "Control work orders, planned maintenance and outstanding maintenance backlog.", nav: "planned_maintenance" },
    { Icon: Package, title: "Parts & Pricing", text: "Monitor inventory, reorder levels and supplier pricing from one connected system.", nav: "parts" },
    { Icon: BarChart3, title: "Reporting & Reliability", text: "Live Availability, Utilisation, MTBF and MTTR reporting from the same underlying data.", nav: "mtbf_mttr" },
    { Icon: ShieldCheck, title: "Accountability", text: "Every critical action is traceable through a complete audit history.", nav: "audit" },
  ];

  const platform = [
    { Icon: Wrench, name: "Engineering & Maintenance", sub: "Equipment, work orders, reliability" },
    { Icon: BarChart3, name: "Production", sub: "Underground and open pit reporting" },
    { Icon: ShoppingCart, name: "Procurement", sub: "Requisitions, orders and suppliers" },
    { Icon: Users, name: "Workforce", sub: "Time & attendance, rosters" },
    { Icon: TrendingUp, name: "Budget & Forecasting", sub: "Cost tracking against plan" },
    { Icon: ShieldCheck, name: "Safety & HSE", sub: "Incidents, inspections, compliance" },
  ];

  const benefits = [
    { Icon: Wrench, title: "CUSTOM BUILT", text: "Developed around the client's actual operational workflows." },
    { Icon: Layers, title: "CONNECTED DATA", text: "Engineering, production, procurement and workforce information can operate from one platform." },
    { Icon: TrendingUp, title: "DIRECT DEVELOPMENT", text: "The system can evolve as operational requirements change." },
    { Icon: Factory, title: "MULTI-SITE READY", text: "Designed to support multiple mining operations with site-specific access and permissions." },
  ];

  const card = {
    background: "#fff",
    border: `1px solid ${LINE}`,
    borderRadius: 12,
    padding: "22px 22px 24px",
    boxShadow: "0 1px 2px rgba(20,40,46,0.04)",
    transition: "box-shadow .16s ease, border-color .16s ease, transform .16s ease",
  };
  const kicker = { fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 11, letterSpacing: "0.22em", textTransform: "uppercase", color: TEAL };
  const h2 = { margin: "12px 0 0", fontSize: 29, lineHeight: 1.14, letterSpacing: "-0.03em", fontWeight: 800, color: INK, maxWidth: "26ch" };
  const iconTile = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 38, height: 38, borderRadius: 10, background: "#EAF4F3", border: "1px solid #D6E7E6", flex: "0 0 auto" };

  return (
    <div style={{ maxWidth: 1240, margin: "0 auto", padding: "28px 32px 0" }}>
      <style>{`
        @keyframes ftGridDrift { from { transform: translate3d(0,0,0); } to { transform: translate3d(-60px,-30px,0); } }
        @keyframes ftRingTurn { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes ftScan { to { stroke-dashoffset: -320; } }
        @keyframes ftPulse { 0%, 100% { opacity: .30; } 50% { opacity: .85; } }
        .ft-card:hover { transform: translateY(-2px); border-color: #C9DBDC; box-shadow: 0 16px 32px -22px rgba(20,40,46,0.30), 0 1px 2px rgba(20,40,46,0.04); }
        @media (prefers-reduced-motion: reduce) { .ft-hero * { animation: none !important; } }
      `}</style>

      {/* 1-2. BRAND AREA + HERO */}
      <section
        ref={heroRef}
        className="ft-hero"
        onMouseMove={onHeroMove}
        onMouseLeave={onHeroLeave}
        style={{
          position: "relative",
          overflow: "hidden",
          borderRadius: isMobile ? 0 : 14,
          margin: isMobile ? "-28px -32px 0" : 0,
          background: "linear-gradient(122deg, #14282E 0%, #16323A 42%, #1F6668 100%)",
          boxShadow: isMobile ? "none" : "0 20px 46px -26px rgba(20,40,46,0.55)",
        }}
      >
        <svg viewBox="0 0 900 520" preserveAspectRatio="xMaxYMid slice" style={{ position: "absolute", inset: 0, width: "100%", height: "100%", opacity: 0.62, pointerEvents: "none" }} aria-hidden="true">
          <defs>
            <radialGradient id="ftFade" cx="74%" cy="46%" r="50%">
              <stop offset="42%" stopColor="#fff" stopOpacity="1" />
              <stop offset="100%" stopColor="#fff" stopOpacity="0" />
            </radialGradient>
            <mask id="ftMask"><rect x="0" y="0" width="900" height="520" fill="url(#ftFade)" /></mask>
            <pattern id="ftDots" x="0" y="0" width="26" height="26" patternUnits="userSpaceOnUse">
              <circle cx="1.2" cy="1.2" r="1.2" fill={MINT} fillOpacity="0.3" />
            </pattern>
          </defs>
          <g mask="url(#ftMask)">
            <g data-depth="0.35" style={{ transition: "transform .35s ease-out" }}>
              <g style={{ animation: "ftGridDrift 34s ease-in-out infinite alternate" }}>
                <rect x="-60" y="-40" width="1040" height="620" fill="url(#ftDots)" />
                <g stroke={MINT} strokeOpacity="0.12" strokeWidth="1">
                  <path d="M0 130H960M0 260H960M0 390H960" />
                  <path d="M520 0V560M650 0V560M780 0V560M910 0V560" />
                </g>
              </g>
            </g>
            <g data-depth="0.7" style={{ transition: "transform .35s ease-out" }}>
              <g style={{ animation: "ftRingTurn 320s linear infinite", transformOrigin: "640px 266px" }}>
                <g fill="none" stroke={MINT} strokeLinejoin="round">
                  {ring(1, 0.4, 1.5)}{ring(0.86, 0.34, 1.35)}{ring(0.72, 0.29, 1.25)}
                  {ring(0.58, 0.24, 1.15)}{ring(0.44, 0.19, 1.1)}{ring(0.3, 0.15, 1)}{ring(0.17, 0.12, 1)}
                </g>
              </g>
            </g>
            <g data-depth="1" style={{ transition: "transform .35s ease-out" }}>
              <path d="M812 150 C 744 196, 762 244, 700 268 C 640 292, 654 336, 606 350" fill="none" stroke={MINT} strokeOpacity="0.4" strokeWidth="1.5" strokeDasharray="10 10" strokeLinecap="round" style={{ animation: "ftScan 16s linear infinite" }} />
              <g fill={MINT} fillOpacity="0.5">
                <circle cx="606" cy="350" r="3" /><circle cx="566" cy="206" r="3" /><circle cx="742" cy="366" r="2.6" />
              </g>
              <circle cx="700" cy="268" r="3.8" fill={MINT} fillOpacity="0.8" />
              <circle cx="700" cy="268" r="13" fill="none" stroke={MINT} strokeOpacity="0.34" strokeWidth="1.1" style={{ animation: "ftPulse 4.5s ease-in-out infinite" }} />
              <circle cx="700" cy="268" r="24" fill="none" stroke={MINT} strokeOpacity="0.18" strokeWidth="1" style={{ animation: "ftPulse 4.5s ease-in-out infinite .6s" }} />
            </g>
            <g data-depth="0.5" style={{ transition: "transform .35s ease-out" }}>
              <g fill="none" stroke={MINT} strokeOpacity="0.22" strokeWidth="1.2" strokeLinejoin="round">
                <path d="M340 470 L430 470 L470 442 L556 442 L596 414 L688 414 L730 386 L900 386" />
                <path d="M300 500 L470 500 L512 472 L604 472 L646 444 L760 444 L800 416 L900 416" />
              </g>
            </g>
          </g>
        </svg>

        <div style={{ position: "relative", padding: isMobile ? "32px 24px" : "42px 44px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(316px,100%),1fr))", gap: 44, alignItems: "center" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                <svg width="26" height="26" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <rect x="1.6" y="1.6" width="20.8" height="20.8" rx="6" stroke={MINT} strokeOpacity="0.6" strokeWidth="1.4" />
                  <path d="M6.5 15.5 L10 9 L13.2 13.2 L15.4 10.4 L17.6 15.5" stroke={MINT} strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                <div>
                  <div style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 12.5, fontWeight: 600, letterSpacing: "0.21em", textTransform: "uppercase", color: "#fff" }}>Datavera Analytics</div>
                  <div style={{ marginTop: 3, fontSize: 12.5, fontWeight: 500, color: "rgba(143,224,214,0.88)" }}>Mining Operations Management Platform</div>
                </div>
              </div>
              <span style={{ width: 1, height: 30, background: "rgba(255,255,255,0.14)" }} />
              <span style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "6px 13px", borderRadius: 999, border: "1px solid rgba(143,224,214,0.30)", background: "rgba(143,224,214,0.07)", fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 10.5, letterSpacing: "0.16em", textTransform: "uppercase", color: MINT }}>
                <Truck size={13} strokeWidth={2} /> Connected Mine Operations
              </span>
            </div>

            <h1 style={{ margin: "30px 0 0", fontSize: 56, lineHeight: 1.02, letterSpacing: "-0.035em", fontWeight: 800, color: "#fff" }}>Mine2U</h1>
            <p style={{ margin: "15px 0 0", fontSize: 22, lineHeight: 1.32, fontWeight: 700, letterSpacing: "-0.015em", color: MINT, maxWidth: "34ch" }}>Complete visibility. Better decisions. Maximum uptime.</p>
            <p style={{ margin: "15px 0 0", fontSize: 16, lineHeight: 1.6, color: "rgba(255,255,255,0.68)", maxWidth: "54ch" }}>A connected mining operations platform designed to manage equipment, maintenance, availability, parts, production and operational reporting from one live source of truth.</p>
          </div>

          <div style={{ display: "grid", gap: 1, background: "rgba(255,255,255,0.10)", borderRadius: 12, overflow: "hidden", border: "1px solid rgba(255,255,255,0.10)" }}>
            {valueProps.map(({ label, text, dot, Icon }) => (
              <div key={label} style={{ background: "rgba(20,40,46,0.35)", padding: "15px 18px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  {dot ? <span style={{ width: 6, height: 6, borderRadius: 999, background: MINT, boxShadow: "0 0 0 3px rgba(143,224,214,0.18)" }} /> : <Icon size={14} strokeWidth={2} color={MINT} />}
                  <span style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 10.5, letterSpacing: "0.16em", textTransform: "uppercase", color: MINT }}>{label}</span>
                </div>
                <div style={{ marginTop: 7, fontSize: 13.5, color: "rgba(255,255,255,0.72)" }}>{text}</div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* 3. EXPLORE THE PLATFORM */}
      <section style={{ padding: "54px 0 0" }}>
        <div style={kicker}>Explore the platform</div>
        <h2 style={h2}>Every module works from the same record</h2>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 340px), 1fr))", gap: 14, marginTop: 26 }}>
          {features.map(({ Icon, title, text, nav }) => (
            <div
              key={title}
              className="ft-card"
              
              style={{ ...card, cursor: nav ? "pointer" : "default" }}
            >
              <span style={iconTile}><Icon size={19} strokeWidth={1.9} color={TEAL} /></span>
              <h3 style={{ margin: "18px 0 0", fontSize: 17.5, fontWeight: 800, letterSpacing: "-0.02em", color: INK }}>{title}</h3>
              <p style={{ margin: "9px 0 0", fontSize: 14.5, lineHeight: 1.6, color: BODY }}>{text}</p>
            </div>
          ))}
        </div>
      </section>

      {/* 4. PLATFORM MODULES */}
      <section style={{ padding: "54px 0 0" }}>
        <div style={{ display: "flex", flexWrap: "wrap", alignItems: "baseline", justifyContent: "space-between", gap: "8px 24px" }}>
          <div>
            <div style={kicker}>Platform modules</div>
            <h2 style={h2}>One platform, extended as you need it</h2>
          </div>
          <p style={{ margin: 0, fontSize: 13.5, lineHeight: 1.55, color: MUTED, maxWidth: "42ch" }}>Mine2U covers engineering and maintenance. Datavera can scope, develop and connect any of the operational areas below on the same platform as your requirements grow.</p>
        </div>
        <div style={{ marginTop: 24, border: `1px solid ${LINE}`, borderRadius: 12, background: "#fff", boxShadow: "0 1px 2px rgba(20,40,46,0.04)", overflow: "hidden" }}>
          {platform.map(({ Icon, name, sub }, i) => (
            <div key={name} style={{ display: "flex", alignItems: "center", gap: 14, padding: "18px 22px", borderBottom: i < platform.length - 1 ? "1px solid #EFF3F3" : "none" }}>
              <span style={{ ...iconTile, width: 34, height: 34, borderRadius: 9 }}><Icon size={17} strokeWidth={1.9} color={TEAL} /></span>
              <div style={{ flex: "1 1 auto", minWidth: 0 }}>
                <div style={{ fontSize: 15, fontWeight: 700, letterSpacing: "-0.01em", color: INK }}>{name}</div>
                <div style={{ marginTop: 2, fontSize: 12.5, color: MUTED }}>{sub}</div>
              </div>
            </div>
          ))}
        </div>
        <p style={{ margin: "12px 0 0", fontSize: 12.5, color: "#A6AFB1" }}>Each area can be scoped and delivered on request, on the same data and permissions model as Mine2U.</p>
      </section>

      {/* 5. WHY DATAVERA */}
      <section style={{ padding: "54px 0 0" }}>
        <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 28, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(216px,100%),1fr))", gap: "24px 30px" }}>
          <div style={{ gridColumn: "1 / -1", ...kicker }}>Why Datavera</div>
          {benefits.map(({ Icon, title, text }) => (
            <div key={title}>
              <Icon size={20} strokeWidth={1.8} color={TEAL} />
              <div style={{ marginTop: 12, fontSize: 14, fontWeight: 800, letterSpacing: "0.02em", color: INK }}>{title}</div>
              <p style={{ margin: "7px 0 0", fontSize: 14, lineHeight: 1.58, color: BODY }}>{text}</p>
            </div>
          ))}
        </div>
      </section>

      {/* 6. DATAVERA BRAND SECTION */}
      <section style={{ margin: "54px 0 0", background: "#EDF4F4", border: "1px solid #DCE8E8", borderRadius: 12, padding: "38px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(290px,100%),1fr))", gap: "26px 44px", alignItems: "start" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <rect x="1.6" y="1.6" width="20.8" height="20.8" rx="6" stroke={TEAL} strokeOpacity="0.45" strokeWidth="1.4" />
                <path d="M6.5 15.5 L10 9 L13.2 13.2 L15.4 10.4 L17.6 15.5" stroke={TEAL} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              <span style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 12, letterSpacing: "0.2em", textTransform: "uppercase", color: TEAL }}>Datavera Analytics</span>
            </div>
            <p style={{ margin: "20px 0 0", fontSize: 25, lineHeight: 1.28, fontWeight: 800, letterSpacing: "-0.025em", color: INK, maxWidth: "26ch" }}>Mining software built around the way your operation works.</p>
          </div>
          <div>
            <p style={{ margin: 0, fontSize: 15.5, lineHeight: 1.64, color: BODY, maxWidth: "56ch" }}>Datavera Analytics develops custom operational systems that connect mining data, maintenance workflows and management reporting into one integrated platform.</p>
            <div style={{ marginTop: 22, fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 11.5, letterSpacing: "0.18em", textTransform: "uppercase", color: TEAL }}>Measure. Monitor. Improve.</div>
          </div>
        </div>
      </section>

      {/* 7. FOOTER */}
      <footer style={{ marginTop: 22, borderTop: `1px solid ${LINE}`, padding: "18px 0 26px", display: "flex", flexWrap: "wrap", gap: "10px 24px", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 13, color: MUTED }}>Built and maintained by Datavera Analytics.</span>
        <span style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", fontSize: 11, letterSpacing: "0.14em", color: MUTED }}>© {new Date().getFullYear()}</span>
      </footer>
    </div>
  );
}

function Dashboard({ assets, breakdowns, workOrders, plannedMaintenance, components, parts, inspections, onNavigate }) {
  const [monthKpi, setMonthKpi] = useState([]);
  const [monthKpiLoading, setMonthKpiLoading] = useState(true);
  const [selectedFleet, setSelectedFleet] = useState("");
  const isMobile = useIsMobile();

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
  // One row per fleet (Excavators, Dump Trucks, etc.) instead of one row
  // per machine - this is what the MTBF/MTTR/Utilisation charts show by
  // default, since 30+ individual equipment bars in one chart isn't
  // readable. Clicking a fleet's bar drills into that fleet by reusing
  // the same selectedFleet filter the dropdown above already uses, so
  // the whole Dashboard (not just the chart) narrows to that fleet.
  const fleetMonthKpi = useMemo(() => {
    const byFleet = {};
    // Use the already site-scoped filteredMonthKpi here, not the raw
    // monthKpi - plant_performance_kpi() itself returns every asset across
    // every site, so aggregating the unfiltered data would mix in other
    // sites' real equipment.
    filteredMonthKpi.forEach((r) => {
      if (!r.fleet) return;
      if (!byFleet[r.fleet]) byFleet[r.fleet] = [];
      byFleet[r.fleet].push(r);
    });
    return Object.keys(byFleet).sort().map((fleet) => ({
      fleet,
      ...aggregateMetrics(byFleet[fleet], byFleet[fleet].map((r) => r.asset_id)),
    }));
  }, [filteredMonthKpi]);
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

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(160px,100%),1fr))", gap: 14, marginBottom: 24 }}>
          <MetricCard label="Total assets" value={filteredAssets.length} icon={Truck} accentColor={NAVY} onClick={() => onNavigate?.("assets")} />
          <MetricCard label="Open work orders" value={filteredWorkOrders.filter((w) => w.status !== "Closed").length} icon={ClipboardList} accentColor="#8B5CF6" onClick={() => onNavigate?.("work_orders")} />
          <MetricCard label="Open breakdowns" value={openBreakdownCount} icon={AlertTriangle} accentColor="#B85450" onClick={() => onNavigate?.("breakdowns")} />
          <MetricCard label="Services due soon / overdue" value={filteredPlannedMaintenance.filter((p) => p.status !== "OK").length} icon={CalendarClock} accentColor="#E8A33D" onClick={() => onNavigate?.("planned_maintenance")} />
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 12 }}>
          <h3 style={{ fontSize: 15, fontWeight: 700, margin: 0, color: NAVY }}>Currently in progress</h3>
          <span style={{ fontSize: 11.5, color: "#859195" }}>Live as of {nowTick.toLocaleTimeString("en-ZA", { hour: "2-digit", minute: "2-digit" })} - updates automatically</span>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(280px,100%),1fr))", gap: 20, marginBottom: 24 }}>
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
          // With no fleet chosen, chart by fleet (one bar per fleet - readable
          // at any fleet size). Once a fleet is selected (dropdown above, or
          // by clicking a fleet's bar below), chart by individual machine
          // within just that fleet.
          const chartData = selectedFleet ? filteredMonthKpi : fleetMonthKpi;
          const chartXKey = selectedFleet ? "asset_id" : "fleet";
          const mtbfMax = Math.max(mtbfTarget, ...chartData.map((r) => r.mtbf || 0)) * 1.15;
          const mttrMax = Math.max(mttrTarget, ...chartData.map((r) => r.mttr || 0)) * 1.15;
          const utilData = chartData.map((r) => ({ ...r, utilisationPct: r.utilisation != null ? Math.round(r.utilisation * 100) : null }));
          const handleBarClick = (row) => {
            if (!selectedFleet) setSelectedFleet(row.fleet);
            else onNavigate?.("mtbf_mttr");
          };
          return (
          <div style={{ marginBottom: 8 }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10, margin: "0 0 12px" }}>
              <h3 style={{ fontSize: 15, fontWeight: 700, margin: 0, color: NAVY }}>MTBF, MTTR & Utilisation, this month{selectedFleet ? ` - ${selectedFleet}` : ""}</h3>
              {selectedFleet ? (
                <button onClick={() => setSelectedFleet("")} style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: NAVY, fontSize: 12.5, fontWeight: 600, textDecoration: "underline" }}>
                  ← back to all fleets
                </button>
              ) : (
                <span style={{ fontSize: 12, color: "#859195" }}>click a fleet to see individual machines</span>
              )}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(260px,100%),1fr))", gap: 16, marginBottom: 8 }}>
              <KpiBarChart
                title={selectedFleet ? "MTBF by equipment (hrs)" : "MTBF by fleet (hrs)"} data={chartData} xKey={chartXKey} dataKey="mtbf"
                target={mtbfTarget} domainMax={mtbfMax} unitSuffix="h"
                valueFormatter={(v) => Number(v).toFixed(1)}
                meetsTarget={(r) => (r.mtbf || 0) >= mtbfTarget}
                onBarClick={handleBarClick}
              />
              <KpiBarChart
                title={selectedFleet ? "MTTR by equipment (hrs)" : "MTTR by fleet (hrs)"} data={chartData} xKey={chartXKey} dataKey="mttr"
                target={mttrTarget} domainMax={mttrMax} unitSuffix="h"
                valueFormatter={(v) => Number(v).toFixed(1)}
                // Lower is better for MTTR, so "meets target" means at or below it.
                meetsTarget={(r) => (r.mttr || 0) <= mttrTarget}
                onBarClick={handleBarClick}
              />
              <KpiBarChart
                title={selectedFleet ? "Utilisation by equipment" : "Utilisation by fleet"} data={utilData} xKey={chartXKey} dataKey="utilisationPct"
                target={utilTarget} domainMax={100} unitSuffix="%"
                valueFormatter={(v) => `${v}%`}
                meetsTarget={(r) => (r.utilisationPct || 0) >= utilTarget}
                onBarClick={handleBarClick}
              />
            </div>
            <KpiLegend targetLabel="Target" />
          </div>
          );
        })()}

        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1.3fr 1fr", gap: 20, marginBottom: 24, alignItems: "start" }}>
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

        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 20, marginBottom: 8 }}>
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

        <ConsumptionCharts />
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
  const [active, setActive] = useState(() => (myRole === "operator" ? "daily_hours" : "about"));
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
  const [workRequests, setWorkRequests] = useState([]);
  const [defects, setDefects] = useState([]);
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
        workRequestsRes, defectsRes,
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
        supabase.from("work_requests").select("*"),
        supabase.from("defects").select("*"),
      ]);
      const results = { siteAssetsRes, fuelRes, oilRes, woRes, serviceRes, inspRes, compRes, tyreRes, partsRes, warrRes, docRes, backlogsRes, dailyServiceRes, componentCodesRes, workRequestsRes, defectsRes };
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
      setWorkRequests(bySite(workRequestsRes.data));
      setDefects(bySite(defectsRes.data));
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

    const { data: activityData, error: activityError } = await supabase.from("activity_log").select("*").order("created_at", { ascending: false }).limit(1500);
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

  // Per-user page visibility is now the single source of truth for what
  // shows in the sidebar - it used to be layered on top of a separate,
  // hidden Operator-role restriction that silently overrode it, which is
  // exactly what caused Manage Pages to look like it wasn't working. An
  // empty myPageAccess means this person has no explicit restrictions and
  // sees everything; admins are never restricted by this, however their
  // own access list looks.
  const canSeePage = (key) => isAdmin || myPageAccess.length === 0 || myPageAccess.includes(key);

  // Defense in depth: if a non-admin's active tab is ever something
  // they're not supposed to see (their access was changed mid-session, a
  // stale value, anything), redirect them rather than silently rendering
  // a page they shouldn't have. The nav already hides these, this is the
  // backstop for if it's reached some other way.
  useEffect(() => {
    if (isAdmin || active === "about") return;
    if (!canSeePage(active)) {
      setActive(myPageAccess[0] || "about");
    }
  }, [isAdmin, active, myPageAccess]);

  const handleNavClick = (key) => {
    setActive(key);
    if (isMobile) setSidebarOpen(false);
  };

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
              title="About Mine2U"
              style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: GREEN, fontWeight: 700, fontSize: 15, textAlign: "left" }}
            >
              Mine2U
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
                n.group === "backlogs" && canSeePage(n.key)
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
              <AboutPage assets={assets} breakdowns={breakdowns} workOrders={workOrders} plannedMaintenance={plannedMaintenance} components={components} parts={parts} inspections={inspections} onNavigate={setActive} />
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
             (restLoading && ["dashboard", "fuel_log", "oil_consumption", "work_orders", "downtime_summary", "planned_maintenance", "inspections", "components", "tyres", "parts", "warranty_docs", "audit", "backlog_report", "daily_service", "breakdowns", "component_codes", "work_requests", "defects"].includes(active)) ? (
              <p style={{ fontSize: 13, color: "#859195" }}>Loading…</p>
            ) : active === "dashboard" ? (
              <Dashboard assets={assets} breakdowns={breakdowns} workOrders={workOrders} plannedMaintenance={plannedMaintenance} components={components} parts={parts} inspections={inspections} onNavigate={setActive} />
            ) : active === "fleet_performance" ? (
              <FleetPerformance assets={assets} breakdowns={breakdowns} />
            ) : active === "breakdowns" ? (
              <BreakdownsPage assets={assets} breakdowns={breakdowns} workOrders={workOrders} parts={parts} componentCodes={componentCodes} onRefresh={() => { loadCoreData(); loadRestOfData(); }} userEmail={userEmail} myFullName={myFullName} />
            ) : active === "work_requests" ? (
              <WorkRequestsPage assets={assets} workRequests={workRequests} workOrders={workOrders} userEmail={userEmail} myFullName={myFullName} onRefresh={loadRestOfData} />
            ) : active === "defects" ? (
              <DefectsPage assets={assets} defects={defects} userEmail={userEmail} myFullName={myFullName} onRefresh={loadRestOfData} />
            ) : active === "assets" ? (
              <AssetsPage assets={assets} selectedSiteId={selectedSiteId} onRefresh={loadCoreData} isAdmin={isAdmin} mySites={mySites} />
            ) : active === "daily_hours" ? (
              <DailyHoursPage assets={assets} dailyHours={dailyHours} userEmail={userEmail} selectedSiteId={selectedSiteId} onRefresh={loadCoreData} />
            ) : active === "planned_maintenance" ? (
              <PlannedMaintenancePage assets={assets} plannedMaintenance={plannedMaintenance} workOrders={workOrders} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "work_orders" ? (
              <WorkOrdersPage assets={assets} workOrders={workOrders} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "downtime_summary" ? (
              <DowntimeSummaryPage assets={assets} breakdowns={breakdowns} workOrders={workOrders} />
            ) : active === "mtbf_mttr" ? (
              <MtbfMttrReportPage assets={assets} />
            ) : active === "fuel_log" ? (
              <FuelLogPage assets={assets} fuelLog={fuelLog} userEmail={userEmail} myFullName={myFullName} dailyHours={dailyHours} onRefresh={loadRestOfData} />
            ) : active === "oil_consumption" ? (
              <OilConsumptionPage assets={assets} oilConsumption={oilConsumption} userEmail={userEmail} myFullName={myFullName} dailyHours={dailyHours} onRefresh={loadRestOfData} />
            ) : active === "inspections" ? (
              <InspectionsPage assets={assets} inspections={inspections} userEmail={userEmail} myFullName={myFullName} onRefresh={loadRestOfData} />
            ) : active === "backlog_report" ? (
              <BacklogsPage assets={assets} backlogs={backlogs} workOrders={workOrders} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "daily_service" ? (
              <DailyServicePage assets={assets} dailyServiceChecklist={dailyServiceChecklist} breakdowns={breakdowns} dailyHours={dailyHours} userEmail={userEmail} myFullName={myFullName} onRefresh={loadRestOfData} />
            ) : active === "components" ? (
              <ComponentsPage assets={assets} components={components} breakdowns={breakdowns} workOrders={workOrders} dailyHours={dailyHours} userEmail={userEmail} onRefresh={loadRestOfData} />
            ) : active === "fleet_health" ? (
              <FleetHealthPage assets={assets} />
            ) : active === "component_master" ? (
              <ComponentMasterPage assets={assets} selectedSiteId={selectedSiteId} myFullName={myFullName} />
            ) : active === "component_status" ? (
              <ComponentStatusPage assets={assets} selectedSiteId={selectedSiteId} />
            ) : active === "parts" ? (
              <PartsPage parts={parts} selectedSiteId={selectedSiteId} onRefresh={loadRestOfData} userEmail={userEmail} isAdmin={isAdmin} />
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

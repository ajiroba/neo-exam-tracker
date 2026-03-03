import { useState, useMemo, useEffect, useCallback } from "react";

// ============================================================
// PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL BELOW
// ============================================================
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxtpdCLNcaR37MzpPmGCbmRWZgwmqwgZ3KptcHO9D068C47RvT0ID2DFPrHHRnjqlU6/exec";
// ============================================================

const PASS_MARK = 70;

const DEFAULT_DEPT_COLORS: Record<string, string> = {
  Finance: "#f59e0b",
  HR: "#10b981",
  Engineering: "#6366f1",
  Sales: "#f43f5e",
  Legal: "#8b5cf6",
  Marketing: "#06b6d4",
  Operations: "#84cc16",
};

const PALETTE = [
  "#f59e0b","#10b981","#6366f1","#f43f5e","#8b5cf6",
  "#06b6d4","#84cc16","#ec4899","#14b8a6","#f97316",
  "#a855f7","#0ea5e9","#22c55e","#eab308","#ef4444",
];

type EscalationLevel = "critical" | "high" | "medium" | "low";

interface Escalation {
  level: EscalationLevel;
  label: string;
  detail: string;
}

interface Employee {
  id: number;
  name: string;
  department: string;
  startDate: string;
  examDate: string | null;
  examScore: number | null;
  attempts: number;
  status: "passed" | "failed" | "scheduled" | "unscheduled";
}

interface EnrichedEmployee extends Employee {
  escalation: Escalation | null;
  deadline: Date;
}

interface SheetsResponse {
  success: boolean;
  data?: Record<string, string | number | null>[];
  error?: string;
}

function getDeadline(startDate: string): Date {
  const d = new Date(startDate);
  d.setMonth(d.getMonth() + 6);
  return d;
}

function getEscalation(emp: Employee): Escalation | null {
  const today = new Date();
  const deadline = getDeadline(emp.startDate);
  const daysToDeadline = Math.ceil((deadline.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (emp.status === "passed") return null;
  if (deadline < today) return { level: "critical", label: "Overdue", detail: `${Math.abs(daysToDeadline)}d past deadline` };
  if (emp.status === "failed" && emp.attempts >= 3) return { level: "critical", label: "Max Attempts", detail: "Escalate to manager" };
  if (emp.status === "failed") return { level: "high", label: "Failed", detail: `Resit required · ${daysToDeadline}d left` };
  if (emp.status === "unscheduled") return { level: daysToDeadline < 60 ? "high" : "medium", label: "Not Scheduled", detail: `${daysToDeadline}d to deadline` };
  if (daysToDeadline < 30) return { level: "high", label: "Urgent", detail: `${daysToDeadline}d to deadline` };
  if (daysToDeadline < 60) return { level: "medium", label: "Watch", detail: `${daysToDeadline}d to deadline` };
  return { level: "low", label: "On Track", detail: `${daysToDeadline}d to deadline` };
}

const escalationConfig: Record<EscalationLevel, { bg: string; border: string; text: string; dot: string; pill: string }> = {
  critical: { bg: "rgba(239,68,68,0.08)", border: "rgba(239,68,68,0.3)", text: "#ef4444", dot: "#ef4444", pill: "rgba(239,68,68,0.15)" },
  high:     { bg: "rgba(249,115,22,0.08)", border: "rgba(249,115,22,0.3)", text: "#f97316", dot: "#f97316", pill: "rgba(249,115,22,0.15)" },
  medium:   { bg: "rgba(234,179,8,0.08)", border: "rgba(234,179,8,0.3)", text: "#eab308", dot: "#eab308", pill: "rgba(234,179,8,0.15)" },
  low:      { bg: "rgba(34,197,94,0.08)", border: "rgba(34,197,94,0.3)", text: "#22c55e", dot: "#22c55e", pill: "rgba(34,197,94,0.15)" },
};

function formatDate(d: string | null | undefined): string {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function daysFromNow(date: Date): number {
  return Math.ceil((date.getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24));
}

const css = `
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@400;500;600;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body { font-family: 'Plus Jakarta Sans', sans-serif; }
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); border-radius: 10px; }
  input, select, button { font-family: 'Plus Jakarta Sans', sans-serif; }
  input:focus, select:focus { outline: none; }
  input[type="date"]::-webkit-calendar-picker-indicator { filter: invert(0.5); cursor: pointer; }

  .neo-app { min-height: 100vh; background: #080b14; color: #e2e8f0; }

  /* Sidebar */
  .sidebar { position: fixed; left: 0; top: 0; bottom: 0; width: 240px; background: rgba(13,17,30,0.95); border-right: 1px solid rgba(255,255,255,0.06); backdrop-filter: blur(20px); z-index: 50; display: flex; flex-direction: column; transition: transform 0.3s cubic-bezier(0.4,0,0.2,1); }
  .sidebar-logo { padding: 24px 20px 20px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid rgba(255,255,255,0.06); }
  .logo-icon { width: 40px; height: 40px; border-radius: 12px; background: linear-gradient(135deg,#6366f1,#8b5cf6); display: flex; align-items: center; justify-content: center; font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 18px; color: #fff; flex-shrink: 0; box-shadow: 0 8px 20px rgba(99,102,241,0.4); }
  .logo-text { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 15px; color: #f1f5f9; line-height: 1.2; }
  .logo-sub { font-size: 10px; color: #64748b; font-weight: 400; letter-spacing: 0.05em; }
  .nav-section { padding: 16px 12px 8px; flex: 1; overflow-y: auto; }
  .nav-label { font-size: 10px; font-weight: 600; letter-spacing: 0.1em; color: #475569; text-transform: uppercase; padding: 0 8px; margin-bottom: 6px; margin-top: 12px; }
  .nav-item { display: flex; align-items: center; gap: 10px; padding: 10px 12px; border-radius: 10px; cursor: pointer; transition: all 0.2s; color: #64748b; font-size: 14px; font-weight: 500; border: none; background: none; width: 100%; text-align: left; }
  .nav-item:hover { background: rgba(255,255,255,0.05); color: #cbd5e1; }
  .nav-item.active { background: rgba(99,102,241,0.15); color: #a5b4fc; border: 1px solid rgba(99,102,241,0.2); }
  .nav-item .nav-icon { font-size: 16px; width: 20px; text-align: center; flex-shrink: 0; }
  .nav-badge { margin-left: auto; background: #ef4444; color: #fff; border-radius: 20px; padding: 1px 7px; font-size: 10px; font-weight: 700; }
  .sidebar-footer { padding: 16px 12px; border-top: 1px solid rgba(255,255,255,0.06); }
  .pass-mark-pill { background: rgba(99,102,241,0.1); border: 1px solid rgba(99,102,241,0.2); border-radius: 8px; padding: 10px 14px; font-size: 12px; color: #94a3b8; display: flex; justify-content: space-between; align-items: center; }
  .pass-mark-pill span { color: #a5b4fc; font-weight: 700; font-size: 15px; }

  /* Main content */
  .main { margin-left: 240px; min-height: 100vh; display: flex; flex-direction: column; }
  .topbar { position: sticky; top: 0; background: rgba(8,11,20,0.9); backdrop-filter: blur(20px); border-bottom: 1px solid rgba(255,255,255,0.06); padding: 0 28px; height: 64px; display: flex; align-items: center; justify-content: space-between; z-index: 40; gap: 16px; }
  .page-title { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 20px; color: #f1f5f9; }
  .page-sub { font-size: 12px; color: #475569; margin-top: 2px; }
  .topbar-actions { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
  .btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px; border-radius: 10px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; border: none; white-space: nowrap; }
  .btn-primary { background: linear-gradient(135deg,#6366f1,#8b5cf6); color: #fff; box-shadow: 0 4px 14px rgba(99,102,241,0.3); }
  .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(99,102,241,0.4); }
  .btn-ghost { background: rgba(255,255,255,0.05); color: #94a3b8; border: 1px solid rgba(255,255,255,0.08); }
  .btn-ghost:hover { background: rgba(255,255,255,0.1); color: #e2e8f0; }
  .btn-danger { background: rgba(239,68,68,0.1); color: #ef4444; border: 1px solid rgba(239,68,68,0.2); }
  .btn-danger:hover { background: rgba(239,68,68,0.2); }
  .btn-sm { padding: 6px 12px; font-size: 12px; }
  .btn-icon { padding: 8px; border-radius: 10px; }

  /* Content area */
  .content { padding: 28px; flex: 1; }

  /* Cards */
  .card { background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.07); border-radius: 16px; overflow: hidden; }
  .card-glass { background: rgba(255,255,255,0.03); backdrop-filter: blur(12px); border: 1px solid rgba(255,255,255,0.07); border-radius: 16px; }

  /* Stat cards */
  .stats-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 16px; margin-bottom: 24px; }
  .stat-card { border-radius: 16px; padding: 22px; position: relative; overflow: hidden; cursor: default; transition: transform 0.2s; }
  .stat-card:hover { transform: translateY(-2px); }
  .stat-card::before { content: ''; position: absolute; top: 0; right: 0; width: 120px; height: 120px; border-radius: 50%; opacity: 0.06; transform: translate(30px,-30px); }
  .stat-label { font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 12px; opacity: 0.7; }
  .stat-value { font-family: 'Space Grotesk',sans-serif; font-size: 42px; font-weight: 700; line-height: 1; margin-bottom: 6px; }
  .stat-sub { font-size: 12px; opacity: 0.6; }
  .stat-trend { position: absolute; top: 20px; right: 20px; font-size: 24px; opacity: 0.2; }

  /* Section headers */
  .section-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
  .section-title { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 16px; color: #f1f5f9; }
  .section-sub { font-size: 12px; color: #64748b; margin-top: 2px; }

  /* Search & filters */
  .filters-row { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }
  .search-box { position: relative; flex: 1; min-width: 200px; max-width: 320px; }
  .search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #475569; font-size: 14px; pointer-events: none; }
  .search-input { width: 100%; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); color: #e2e8f0; border-radius: 10px; padding: 9px 12px 9px 36px; font-size: 13px; transition: all 0.2s; }
  .search-input:focus { background: rgba(255,255,255,0.07); border-color: rgba(99,102,241,0.4); box-shadow: 0 0 0 3px rgba(99,102,241,0.1); }
  .search-input::placeholder { color: #475569; }
  .filter-select { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); color: #94a3b8; border-radius: 10px; padding: 9px 12px; font-size: 13px; cursor: pointer; transition: all 0.2s; }
  .filter-select:focus { border-color: rgba(99,102,241,0.4); }
  .filter-select option { background: #0d111e; }
  .record-count { margin-left: auto; font-size: 12px; color: #475569; white-space: nowrap; }

  /* Table */
  .table-wrap { border-radius: 16px; overflow: hidden; border: 1px solid rgba(255,255,255,0.07); }
  .table-head { display: grid; padding: 12px 20px; background: rgba(255,255,255,0.02); border-bottom: 1px solid rgba(255,255,255,0.06); }
  .table-head-cell { font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: #475569; }
  .table-row { display: grid; padding: 14px 20px; border-bottom: 1px solid rgba(255,255,255,0.04); transition: background 0.15s; align-items: center; cursor: pointer; }
  .table-row:last-child { border-bottom: none; }
  .table-row:hover { background: rgba(255,255,255,0.03); }
  .table-row.selected { background: rgba(99,102,241,0.06); }
  .emp-name { font-weight: 600; color: #f1f5f9; font-size: 14px; }
  .emp-meta { font-size: 11px; color: #475569; margin-top: 3px; }
  .dept-pill { display: inline-flex; align-items: center; gap: 5px; border-radius: 20px; padding: 3px 10px; font-size: 11px; font-weight: 600; }
  .dept-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
  .score-val { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 18px; }
  .esc-pill { display: inline-flex; align-items: center; gap: 5px; border-radius: 20px; padding: 4px 10px; font-size: 11px; font-weight: 600; }
  .esc-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
  .actions-cell { display: flex; gap: 6px; }

  /* Progress bars */
  .progress-track { height: 6px; background: rgba(255,255,255,0.06); border-radius: 10px; overflow: hidden; }
  .progress-fill { height: 100%; border-radius: 10px; transition: width 0.6s cubic-bezier(0.4,0,0.2,1); }

  /* Two column grid */
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
  .three-col { display: grid; grid-template-columns: repeat(3,1fr); gap: 16px; }

  /* Alert banner */
  .alert-banner { display: flex; align-items: center; gap: 12px; padding: 14px 18px; border-radius: 12px; margin-bottom: 20px; }
  .alert-critical { background: rgba(239,68,68,0.08); border: 1px solid rgba(239,68,68,0.2); }

  /* Modals */
  .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.7); backdrop-filter: blur(6px); display: flex; align-items: center; justify-content: center; z-index: 200; padding: 20px; }
  .modal { background: #0d111e; border: 1px solid rgba(255,255,255,0.1); border-radius: 20px; width: 100%; max-width: 460px; overflow: hidden; box-shadow: 0 25px 60px rgba(0,0,0,0.5); }
  .modal-header { padding: 24px 24px 0; }
  .modal-title { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 20px; color: #f1f5f9; }
  .modal-sub { font-size: 13px; color: #64748b; margin-top: 4px; }
  .modal-body { padding: 20px 24px; }
  .modal-footer { padding: 0 24px 24px; display: flex; gap: 10px; }
  .form-label { font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: #64748b; display: block; margin-bottom: 8px; }
  .form-input { width: 100%; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); color: #e2e8f0; border-radius: 10px; padding: 10px 14px; font-size: 14px; transition: all 0.2s; }
  .form-input:focus { border-color: rgba(99,102,241,0.5); background: rgba(255,255,255,0.07); box-shadow: 0 0 0 3px rgba(99,102,241,0.1); }
  .form-group { margin-bottom: 16px; }
  .info-row { display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid rgba(255,255,255,0.05); }
  .info-row:last-child { border-bottom: none; }

  /* Calendar cards */
  .month-card { background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.07); border-radius: 16px; padding: 20px; transition: transform 0.2s; }
  .month-card:hover { transform: translateY(-2px); border-color: rgba(99,102,241,0.2); }
  .month-name { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 15px; color: #f1f5f9; }
  .month-count { font-family: 'Space Grotesk',sans-serif; font-weight: 700; font-size: 28px; color: #6366f1; }

  /* Dept management */
  .dept-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px,1fr)); gap: 12px; }
  .dept-card { background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.07); border-radius: 12px; padding: 16px; display: flex; align-items: center; gap: 12px; transition: all 0.2s; }
  .dept-card:hover { border-color: rgba(255,255,255,0.12); background: rgba(255,255,255,0.05); }
  .dept-color-circle { width: 40px; height: 40px; border-radius: 12px; flex-shrink: 0; display: flex; align-items: center; justify-content: center; font-size: 18px; }
  .dept-info { flex: 1; min-width: 0; }
  .dept-name { font-weight: 600; font-size: 14px; color: #f1f5f9; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .dept-count { font-size: 11px; color: #64748b; margin-top: 2px; }
  .color-picker-row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
  .color-swatch { width: 28px; height: 28px; border-radius: 8px; cursor: pointer; transition: transform 0.15s; border: 2px solid transparent; }
  .color-swatch:hover { transform: scale(1.15); }
  .color-swatch.selected { border-color: #fff; transform: scale(1.1); }

  /* Hamburger */
  .hamburger { display: none; background: none; border: none; color: #94a3b8; cursor: pointer; padding: 6px; border-radius: 8px; }
  .sidebar-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 49; backdrop-filter: blur(2px); }

  /* Save indicator */
  .save-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; margin-right: 6px; }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
  .pulsing { animation: pulse 1.5s infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .spinning { animation: spin 1s linear infinite; display: inline-block; }
  @keyframes fadein { from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:translateY(0)} }
  .fadein { animation: fadein 0.35s ease forwards; }
  @keyframes slideup { from{opacity:0;transform:translateY(20px)} to{opacity:1;transform:translateY(0)} }
  .slideup { animation: slideup 0.3s ease forwards; }

  /* Empty state */
  .empty { text-align: center; padding: 60px 20px; color: #475569; }
  .empty-icon { font-size: 48px; margin-bottom: 12px; opacity: 0.4; }
  .empty-text { font-size: 15px; color: #64748b; }
  .empty-sub { font-size: 13px; color: #475569; margin-top: 6px; }

  /* Deadline urgency bars */
  .deadline-item { margin-bottom: 18px; }
  .deadline-name { font-size: 14px; font-weight: 500; color: #e2e8f0; }
  .deadline-dept { font-size: 12px; color: #64748b; }
  .deadline-days { font-size: 12px; font-weight: 600; }
  .deadline-date { font-size: 11px; color: #64748b; }

  /* Responsive */
  @media (max-width: 1024px) {
    .stats-grid { grid-template-columns: repeat(2,1fr); }
    .two-col { grid-template-columns: 1fr; }
    .three-col { grid-template-columns: repeat(2,1fr); }
  }
  @media (max-width: 768px) {
    .sidebar { transform: translateX(-100%); }
    .sidebar.open { transform: translateX(0); }
    .sidebar-overlay.open { display: block; }
    .main { margin-left: 0; }
    .hamburger { display: flex; }
    .content { padding: 16px; }
    .topbar { padding: 0 16px; }
    .stats-grid { grid-template-columns: repeat(2,1fr); gap: 12px; }
    .three-col { grid-template-columns: 1fr; }
    .filters-row { gap: 8px; }
    .search-box { max-width: 100%; }
    .table-wrap { overflow-x: auto; }
    .modal { max-width: 100%; }
    .topbar-actions .btn-label { display: none; }
  }
  @media (max-width: 480px) {
    .stats-grid { grid-template-columns: 1fr 1fr; gap: 10px; }
    .stat-value { font-size: 32px; }
    .page-title { font-size: 16px; }
  }
`;

export default function NEOTracker() {
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [deptColors, setDeptColors] = useState<Record<string, string>>(DEFAULT_DEPT_COLORS);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [lastSaved, setLastSaved] = useState<Date | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);

  const [tab, setTab] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [filterDept, setFilterDept] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterEscalation, setFilterEscalation] = useState("All");
  const [selectedRow, setSelectedRow] = useState<number | null>(null);

  const [rescheduleModal, setRescheduleModal] = useState<EnrichedEmployee | null>(null);
  const [rescheduleDate, setRescheduleDate] = useState("");
  const [scoreModal, setScoreModal] = useState<EnrichedEmployee | null>(null);
  const [scoreInput, setScoreInput] = useState("");
  const [addEmpModal, setAddEmpModal] = useState(false);
  const [newEmp, setNewEmp] = useState({ name: "", department: "Finance", startDate: "" });
  const [addDeptModal, setAddDeptModal] = useState(false);
  const [newDept, setNewDept] = useState({ name: "", color: PALETTE[0] });

  const fetchData = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(SCRIPT_URL);
      const json = await res.json() as SheetsResponse;
      if (json.success && json.data) {
        const parsed: Employee[] = json.data.map((e) => ({
          id: Number(e.id),
          name: String(e.name ?? ""),
          department: String(e.department ?? ""),
          startDate: String(e.startDate ?? ""),
          examDate: e.examDate ? String(e.examDate) : null,
          examScore: e.examScore !== "" && e.examScore !== null ? Number(e.examScore) : null,
          attempts: Number(e.attempts) || 0,
          status: (e.status as Employee["status"]) ?? "unscheduled",
        }));
        setEmployees(parsed);
      } else {
        setError("Failed to load: " + (json.error ?? "Unknown error"));
      }
    } catch {
      setError("Could not reach Google Sheets. Check your SCRIPT_URL.");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  const saveEmployee = async (updated: Employee) => {
    setSaving(true);
    try {
      await fetch(SCRIPT_URL, { method: "POST", body: JSON.stringify({ action: "update", employee: updated }) });
      setEmployees((prev) => prev.map((e) => e.id === updated.id ? updated : e));
      setLastSaved(new Date());
    } catch { setError("Save failed."); }
    finally { setSaving(false); }
  };

  const addEmployeeFn = async () => {
    if (!newEmp.name || !newEmp.startDate) return;
    const id = Math.max(...employees.map((e) => e.id), 0) + 1;
    const emp: Employee = { id, ...newEmp, examDate: null, examScore: null, attempts: 0, status: "unscheduled" };
    setSaving(true);
    try {
      await fetch(SCRIPT_URL, { method: "POST", body: JSON.stringify({ action: "add", employee: emp }) });
      setEmployees((prev) => [...prev, emp]);
      setLastSaved(new Date());
      setAddEmpModal(false);
      setNewEmp({ name: "", department: "Finance", startDate: "" });
    } catch { setError("Failed to add employee."); }
    finally { setSaving(false); }
  };

  const addDeptFn = () => {
    if (!newDept.name.trim()) return;
    const name = newDept.name.trim();
    setDeptColors((prev) => ({ ...prev, [name]: newDept.color }));
    setAddDeptModal(false);
    setNewDept({ name: "", color: PALETTE[0] });
  };

  const getDeptColor = (dept: string) => deptColors[dept] || "#94a3b8";
  const allDepts = Array.from(new Set([...Object.keys(deptColors), ...employees.map((e) => e.department)]));

  const enriched: EnrichedEmployee[] = useMemo(() =>
    employees.map((e) => ({ ...e, escalation: getEscalation(e), deadline: getDeadline(e.startDate) })),
    [employees]
  );

  const stats = useMemo(() => {
    const total = enriched.length;
    const passed = enriched.filter((e) => e.status === "passed").length;
    const failed = enriched.filter((e) => e.status === "failed").length;
    const unscheduled = enriched.filter((e) => e.status === "unscheduled").length;
    const critical = enriched.filter((e) => e.escalation?.level === "critical").length;
    const overdue = enriched.filter((e) => e.escalation?.label === "Overdue").length;
    const passRate = total ? Math.round(passed / total * 100) : 0;
    return { total, passed, failed, unscheduled, critical, overdue, passRate };
  }, [enriched]);

  const filtered: EnrichedEmployee[] = useMemo(() => enriched.filter((e) => {
    const q = search.toLowerCase();
    return (!q || e.name.toLowerCase().includes(q) || e.department.toLowerCase().includes(q))
      && (filterDept === "All" || e.department === filterDept)
      && (filterStatus === "All" || e.status === filterStatus)
      && (filterEscalation === "All" || e.escalation?.level === filterEscalation.toLowerCase());
  }), [enriched, search, filterDept, filterStatus, filterEscalation]);

  const byMonth = useMemo(() => {
    const map: Record<string, { scheduled: number; passed: number; failed: number }> = {};
    enriched.forEach((e) => {
      if (!e.examDate) return;
      const key = e.examDate.slice(0, 7);
      if (!map[key]) map[key] = { scheduled: 0, passed: 0, failed: 0 };
      map[key].scheduled++;
      if (e.status === "passed") map[key].passed++;
      if (e.status === "failed") map[key].failed++;
    });
    return Object.entries(map).sort(([a], [b]) => a.localeCompare(b)).slice(-6);
  }, [enriched]);

  const criticalList = enriched.filter((e) => e.escalation?.level === "critical" || e.escalation?.level === "high");
  const upcomingDeadlines = enriched
    .filter((e) => { const d = daysFromNow(e.deadline); return d >= 0 && d <= 90 && e.status !== "passed"; })
    .sort((a, b) => a.deadline.getTime() - b.deadline.getTime());

  async function confirmReschedule() {
    if (!rescheduleDate || !rescheduleModal) return;
    await saveEmployee({ ...rescheduleModal, examDate: rescheduleDate, status: rescheduleModal.status === "unscheduled" ? "scheduled" : rescheduleModal.status });
    setRescheduleModal(null);
  }

  async function confirmScore() {
    if (!scoreModal) return;
    const score = parseInt(scoreInput, 10);
    if (isNaN(score) || score < 0 || score > 100) return;
    await saveEmployee({ ...scoreModal, examScore: score, status: score >= PASS_MARK ? "passed" : "failed", attempts: scoreModal.examScore === null ? (scoreModal.attempts || 0) + 1 : scoreModal.attempts });
    setScoreModal(null);
    setScoreInput("");
  }

  const navItems = [
    { id: "dashboard", icon: "⬛", label: "Dashboard" },
    { id: "employees", icon: "👥", label: "Employees", badge: stats.critical > 0 ? stats.critical : null },
    { id: "calendar", icon: "📅", label: "Schedule" },
    { id: "departments", icon: "🏢", label: "Departments" },
  ];

  const pageTitles: Record<string, { title: string; sub: string }> = {
    dashboard: { title: "Dashboard", sub: "NEO exam overview & escalations" },
    employees: { title: "Employees", sub: `${stats.total} employees in the programme` },
    calendar: { title: "Schedule", sub: "Exam dates & upcoming deadlines" },
    departments: { title: "Departments", sub: "Manage departments & colours" },
  };

  if (loading) return (
    <div style={{ minHeight: "100vh", background: "#080b14", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'Plus Jakarta Sans',sans-serif", color: "#6366f1" }}>
      <style>{css}</style>
      <div style={{ width: 56, height: 56, borderRadius: 16, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26, marginBottom: 20, boxShadow: "0 12px 30px rgba(99,102,241,0.4)" }}>N</div>
      <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 22, color: "#f1f5f9", marginBottom: 8 }}>NEO Exam Tracker</div>
      <div style={{ color: "#475569", fontSize: 13, marginBottom: 28 }}>Connecting to Google Sheets…</div>
      <div style={{ width: 200, height: 3, background: "rgba(255,255,255,0.06)", borderRadius: 10, overflow: "hidden" }}>
        <div style={{ height: "100%", background: "linear-gradient(90deg,#6366f1,#8b5cf6)", borderRadius: 10, animation: "slideup 1.5s ease infinite", width: "60%" }} />
      </div>
    </div>
  );

  if (error && !employees.length) return (
    <div style={{ minHeight: "100vh", background: "#080b14", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'Plus Jakarta Sans',sans-serif", padding: 24 }}>
      <style>{css}</style>
      <div style={{ fontSize: 48, marginBottom: 16 }}>⚠️</div>
      <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 20, color: "#f1f5f9", marginBottom: 8 }}>Connection Failed</div>
      <div style={{ background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 12, padding: 16, maxWidth: 440, fontSize: 13, color: "#f87171", lineHeight: 1.7, marginBottom: 16, textAlign: "center" }}>{error}</div>
      <p style={{ color: "#64748b", fontSize: 13, textAlign: "center", maxWidth: 380, lineHeight: 1.7 }}>Replace <code style={{ color: "#a5b4fc" }}>SCRIPT_URL</code> in App.tsx with your deployed Google Apps Script Web App URL.</p>
      <button className="btn btn-primary" onClick={fetchData} style={{ marginTop: 20 }}>↻ Retry</button>
    </div>
  );

  return (
    <div className="neo-app">
      <style>{css}</style>

      {/* Sidebar overlay for mobile */}
      <div className={`sidebar-overlay ${sidebarOpen ? "open" : ""}`} onClick={() => setSidebarOpen(false)} />

      {/* Sidebar */}
      <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
        <div className="sidebar-logo">
          <div className="logo-icon">N</div>
          <div>
            <div className="logo-text">NEO Tracker</div>
            <div className="logo-sub">Exam Management</div>
          </div>
        </div>
        <nav className="nav-section">
          <div className="nav-label">Navigation</div>
          {navItems.map((item) => (
            <button key={item.id} className={`nav-item ${tab === item.id ? "active" : ""}`}
              onClick={() => { setTab(item.id); setSidebarOpen(false); }}>
              <span className="nav-icon">{item.icon}</span>
              {item.label}
              {item.badge && <span className="nav-badge">{item.badge}</span>}
            </button>
          ))}
          <div className="nav-label" style={{ marginTop: 20 }}>Actions</div>
          <button className="nav-item" onClick={() => { setAddEmpModal(true); setSidebarOpen(false); }}>
            <span className="nav-icon">➕</span> Add Employee
          </button>
          <button className="nav-item" onClick={() => { setTab("departments"); setAddDeptModal(true); setSidebarOpen(false); }}>
            <span className="nav-icon">🏢</span> New Department
          </button>
        </nav>
        <div className="sidebar-footer">
          <div className="pass-mark-pill">
            <span style={{ fontSize: 12, color: "#64748b" }}>Pass Mark</span>
            <span>{PASS_MARK}%</span>
          </div>
          {lastSaved && (
            <div style={{ marginTop: 10, fontSize: 11, color: "#22c55e", textAlign: "center" }}>
              ✓ Saved at {lastSaved.toLocaleTimeString()}
            </div>
          )}
        </div>
      </aside>

      {/* Main */}
      <main className="main">
        {/* Topbar */}
        <div className="topbar">
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button className="hamburger btn" onClick={() => setSidebarOpen(true)}>☰</button>
            <div>
              <div className="page-title">{pageTitles[tab]?.title}</div>
              <div className="page-sub">{pageTitles[tab]?.sub}</div>
            </div>
          </div>
          <div className="topbar-actions">
            {saving && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#6366f1" }}>
                <span className="spinning">⟳</span>
                <span className="btn-label">Saving…</span>
              </div>
            )}
            {stats.critical > 0 && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.25)", borderRadius: 10, padding: "6px 12px", fontSize: 12, color: "#f87171" }}>
                <span className="save-dot pulsing" style={{ background: "#ef4444" }} />
                <span>{stats.critical} Critical</span>
              </div>
            )}
            <button className="btn btn-ghost btn-sm" onClick={fetchData}>↻ <span className="btn-label">Refresh</span></button>
            <button className="btn btn-primary btn-sm" onClick={() => setAddEmpModal(true)}>+ <span className="btn-label">Add Employee</span></button>
          </div>
        </div>

        {/* Content */}
        <div className="content">

          {/* ─── DASHBOARD ─── */}
          {tab === "dashboard" && (
            <div className="fadein">
              {/* Stats */}
              <div className="stats-grid">
                {[
                  { label: "Total Employees", value: stats.total, sub: "In NEO programme", color: "#6366f1", bg: "linear-gradient(135deg,rgba(99,102,241,0.12),rgba(139,92,246,0.06))", icon: "👥", border: "rgba(99,102,241,0.2)" },
                  { label: "Pass Rate", value: `${stats.passRate}%`, sub: `${stats.passed} employees passed`, color: "#22c55e", bg: "linear-gradient(135deg,rgba(34,197,94,0.12),rgba(16,185,129,0.06))", icon: "✅", border: "rgba(34,197,94,0.2)" },
                  { label: "Require Resit", value: stats.failed, sub: "Failed exam", color: "#f97316", bg: "linear-gradient(135deg,rgba(249,115,22,0.12),rgba(234,179,8,0.06))", icon: "📋", border: "rgba(249,115,22,0.2)" },
                  { label: "Critical Flags", value: stats.critical, sub: `${stats.overdue} past deadline`, color: "#ef4444", bg: "linear-gradient(135deg,rgba(239,68,68,0.12),rgba(244,63,94,0.06))", icon: "🚨", border: "rgba(239,68,68,0.2)" },
                ].map((s) => (
                  <div key={s.label} className="stat-card" style={{ background: s.bg, border: `1px solid ${s.border}` }}>
                    <div className="stat-trend">{s.icon}</div>
                    <div className="stat-label" style={{ color: s.color }}>{s.label}</div>
                    <div className="stat-value" style={{ color: s.color }}>{s.value}</div>
                    <div className="stat-sub">{s.sub}</div>
                  </div>
                ))}
              </div>

              {/* Critical escalations */}
              {criticalList.length > 0 && (
                <div style={{ background: "rgba(239,68,68,0.05)", border: "1px solid rgba(239,68,68,0.15)", borderRadius: 16, padding: 20, marginBottom: 20 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
                    <span className="save-dot pulsing" style={{ background: "#ef4444", width: 10, height: 10, borderRadius: "50%", display: "inline-block" }} />
                    <span style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, color: "#f87171", fontSize: 15 }}>Requires Immediate Action</span>
                    <span style={{ marginLeft: "auto", background: "rgba(239,68,68,0.15)", color: "#f87171", borderRadius: 20, padding: "2px 10px", fontSize: 12, fontWeight: 600 }}>{criticalList.length} employees</span>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {criticalList.map((e) => {
                      const cfg = escalationConfig[e.escalation!.level];
                      return (
                        <div key={e.id} style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 12, padding: "12px 16px", display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                          <div style={{ width: 36, height: 36, borderRadius: 10, background: getDeptColor(e.department) + "22", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, color: getDeptColor(e.department), fontWeight: 700, flexShrink: 0 }}>{e.name[0]}</div>
                          <div style={{ flex: 1, minWidth: 120 }}>
                            <div style={{ fontWeight: 600, color: "#f1f5f9", fontSize: 14 }}>{e.name}</div>
                            <div style={{ fontSize: 11, color: "#64748b" }}>{e.department}</div>
                          </div>
                          <span className="esc-pill" style={{ background: cfg.pill, color: cfg.text }}>
                            <span className="esc-dot" style={{ background: cfg.dot }} />
                            {e.escalation!.label}
                          </span>
                          <span style={{ fontSize: 12, color: cfg.text }}>{e.escalation!.detail}</span>
                          <button className="btn btn-ghost btn-sm" onClick={() => { setRescheduleModal(e); setRescheduleDate(e.examDate || ""); }}>Reschedule</button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              <div className="two-col">
                {/* Escalation breakdown */}
                <div className="card" style={{ padding: 22 }}>
                  <div className="section-header"><div><div className="section-title">Escalation Breakdown</div><div className="section-sub">By severity level</div></div></div>
                  {(["critical","high","medium","low"] as EscalationLevel[]).map((level) => {
                    const count = enriched.filter((e) => e.escalation?.level === level).length;
                    const tot = enriched.filter((e) => e.escalation).length || 1;
                    const cfg = escalationConfig[level];
                    return (
                      <div key={level} style={{ marginBottom: 16 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, alignItems: "center" }}>
                          <span className="esc-pill" style={{ background: cfg.pill, color: cfg.text }}>
                            <span className="esc-dot" style={{ background: cfg.dot }} />
                            {level.charAt(0).toUpperCase() + level.slice(1)}
                          </span>
                          <span style={{ color: "#94a3b8", fontSize: 13, fontWeight: 600 }}>{count}</span>
                        </div>
                        <div className="progress-track">
                          <div className="progress-fill" style={{ width: `${Math.round(count/tot*100)}%`, background: cfg.dot }} />
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Department pass rates */}
                <div className="card" style={{ padding: 22 }}>
                  <div className="section-header"><div><div className="section-title">Department Pass Rates</div><div className="section-sub">Pass vs total by team</div></div></div>
                  {allDepts.filter((d) => enriched.some((e) => e.department === d)).map((dept) => {
                    const dEmps = enriched.filter((e) => e.department === dept);
                    const dPassed = dEmps.filter((e) => e.status === "passed").length;
                    const pct = Math.round(dPassed / dEmps.length * 100);
                    const color = getDeptColor(dept);
                    return (
                      <div key={dept} style={{ marginBottom: 14 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, alignItems: "center" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <div style={{ width: 8, height: 8, borderRadius: "50%", background: color }} />
                            <span style={{ fontSize: 13, color: "#94a3b8", fontWeight: 500 }}>{dept}</span>
                          </div>
                          <span style={{ fontSize: 12, color: "#64748b" }}>{dPassed}/{dEmps.length} · {pct}%</span>
                        </div>
                        <div className="progress-track">
                          <div className="progress-fill" style={{ width: `${pct}%`, background: color }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          )}

          {/* ─── EMPLOYEES ─── */}
          {tab === "employees" && (
            <div className="fadein">
              <div className="filters-row">
                <div className="search-box">
                  <span className="search-icon">🔍</span>
                  <input className="search-input" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search name or department…" />
                </div>
                {[
                  { state: filterDept, setter: setFilterDept, opts: ["All", ...allDepts], placeholder: "All Departments" },
                  { state: filterStatus, setter: setFilterStatus, opts: ["All","passed","failed","scheduled","unscheduled"], placeholder: "All Statuses" },
                  { state: filterEscalation, setter: setFilterEscalation, opts: ["All","critical","high","medium","low"], placeholder: "All Escalations" },
                ].map((f, i) => (
                  <select key={i} className="filter-select" value={f.state} onChange={(e) => f.setter(e.target.value)}>
                    {f.opts.map((o) => <option key={o} value={o}>{o === "All" ? f.placeholder : o.charAt(0).toUpperCase() + o.slice(1)}</option>)}
                  </select>
                ))}
                <div className="record-count">{filtered.length} of {enriched.length} records</div>
              </div>

              <div className="table-wrap">
                <div className="table-head" style={{ gridTemplateColumns: "2fr 1.2fr 1fr 1fr 0.8fr 1.4fr 160px" }}>
                  {["Employee","Department","Exam Date","Score","Attempts","Escalation","Actions"].map((h) => (
                    <div key={h} className="table-head-cell">{h}</div>
                  ))}
                </div>
                {filtered.length === 0 && (
                  <div className="empty">
                    <div className="empty-icon">🔍</div>
                    <div className="empty-text">No employees found</div>
                    <div className="empty-sub">Try adjusting your filters</div>
                  </div>
                )}
                {filtered.map((e, i) => {
                  const esc = e.escalation;
                  const cfg = esc ? escalationConfig[esc.level] : null;
                  const color = getDeptColor(e.department);
                  return (
                    <div key={e.id} className={`table-row ${selectedRow === e.id ? "selected" : ""}`}
                      style={{ gridTemplateColumns: "2fr 1.2fr 1fr 1fr 0.8fr 1.4fr 160px", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.01)" }}
                      onClick={() => setSelectedRow(selectedRow === e.id ? null : e.id)}>
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <div style={{ width: 34, height: 34, borderRadius: 10, background: color + "22", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, color, fontWeight: 700, flexShrink: 0 }}>{e.name[0]}</div>
                          <div>
                            <div className="emp-name">{e.name}</div>
                            <div className="emp-meta">Deadline: {formatDate(e.deadline.toISOString().split("T")[0])}</div>
                          </div>
                        </div>
                      </div>
                      <div>
                        <span className="dept-pill" style={{ background: color + "18", color }}>
                          <span className="dept-dot" style={{ background: color }} />
                          {e.department}
                        </span>
                      </div>
                      <div style={{ fontSize: 13, color: e.examDate ? "#94a3b8" : "#475569" }}>{formatDate(e.examDate)}</div>
                      <div>
                        {e.examScore !== null
                          ? <span className="score-val" style={{ color: e.examScore >= PASS_MARK ? "#22c55e" : "#ef4444" }}>{e.examScore}%</span>
                          : <span style={{ color: "#2d3344", fontSize: 18 }}>—</span>}
                      </div>
                      <div style={{ color: "#64748b", fontSize: 14, fontWeight: 500 }}>{e.attempts}</div>
                      <div>
                        {esc && cfg
                          ? <span className="esc-pill" style={{ background: cfg.pill, color: cfg.text }}><span className="esc-dot" style={{ background: cfg.dot }} />{esc.label}</span>
                          : <span style={{ color: "#22c55e", fontSize: 13, fontWeight: 600 }}>✓ Passed</span>}
                        {esc && <div style={{ fontSize: 11, color: "#64748b", marginTop: 3 }}>{esc.detail}</div>}
                      </div>
                      <div className="actions-cell" onClick={(ev) => ev.stopPropagation()}>
                        <button className="btn btn-ghost btn-sm" onClick={() => { setRescheduleModal(e); setRescheduleDate(e.examDate || ""); }} title={e.examDate ? "Reschedule" : "Schedule"}>
                          {e.examDate ? "↻" : "+"} {e.examDate ? "Resched" : "Schedule"}
                        </button>
                        {e.examDate && (
                          <button className="btn btn-sm" style={{ background: "rgba(99,102,241,0.1)", color: "#a5b4fc", border: "1px solid rgba(99,102,241,0.2)" }}
                            onClick={() => { setScoreModal(e); setScoreInput(e.examScore !== null ? String(e.examScore) : ""); }}>
                            Score
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ─── CALENDAR / SCHEDULE ─── */}
          {tab === "calendar" && (
            <div className="fadein">
              <div className="section-header" style={{ marginBottom: 20 }}>
                <div><div className="section-title">Exam Schedule by Month</div><div className="section-sub">Last 6 months with results</div></div>
              </div>
              {byMonth.length === 0 ? (
                <div className="empty"><div className="empty-icon">📅</div><div className="empty-text">No exams scheduled yet</div></div>
              ) : (
                <div className="three-col" style={{ marginBottom: 28 }}>
                  {byMonth.map(([month, data]) => {
                    const label = new Date(month + "-01").toLocaleDateString("en-GB", { month: "long", year: "numeric" });
                    const monthEmps = enriched.filter((e) => e.examDate?.slice(0, 7) === month);
                    const pending = data.scheduled - data.passed - data.failed;
                    return (
                      <div key={month} className="month-card">
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
                          <div className="month-name">{label}</div>
                          <div className="month-count">{data.scheduled}</div>
                        </div>
                        <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
                          {[{ label: "Passed", val: data.passed, color: "#22c55e" }, { label: "Failed", val: data.failed, color: "#ef4444" }, { label: "Pending", val: pending, color: "#6366f1" }].map((s) => (
                            <div key={s.label} style={{ flex: 1, textAlign: "center", background: s.color + "10", borderRadius: 10, padding: "10px 0" }}>
                              <div style={{ fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, fontSize: 20, color: s.color }}>{s.val}</div>
                              <div style={{ fontSize: 10, color: "#64748b", fontWeight: 600, letterSpacing: "0.05em", textTransform: "uppercase" }}>{s.label}</div>
                            </div>
                          ))}
                        </div>
                        <div style={{ borderTop: "1px solid rgba(255,255,255,0.06)", paddingTop: 12 }}>
                          {monthEmps.slice(0, 3).map((e) => {
                            const cfg = e.escalation ? escalationConfig[e.escalation.level] : null;
                            return (
                              <div key={e.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                                <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 500 }}>{e.name}</span>
                                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                  {e.status === "passed" && <span style={{ color: "#22c55e", fontSize: 12 }}>✓</span>}
                                  {e.status === "failed" && <span style={{ color: "#ef4444", fontSize: 12 }}>✗</span>}
                                  {cfg && e.status !== "passed" && <span className="esc-pill" style={{ background: cfg.pill, color: cfg.text, fontSize: 10, padding: "2px 7px" }}>{e.escalation!.label}</span>}
                                </div>
                              </div>
                            );
                          })}
                          {monthEmps.length > 3 && <div style={{ fontSize: 11, color: "#475569" }}>+{monthEmps.length - 3} more</div>}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* Upcoming deadlines */}
              <div className="card" style={{ padding: 22 }}>
                <div className="section-header">
                  <div><div className="section-title">Month-6 Deadlines</div><div className="section-sub">Next 90 days — non-passed employees</div></div>
                </div>
                {upcomingDeadlines.length === 0 && <div className="empty" style={{ padding: "30px 0" }}><div className="empty-text">No upcoming deadlines in the next 90 days</div></div>}
                {upcomingDeadlines.map((e) => {
                  const days = daysFromNow(e.deadline);
                  const urgentColor = days < 14 ? "#ef4444" : days < 30 ? "#f97316" : "#eab308";
                  const pct = Math.max(0, Math.min(100, 100 - (days / 90 * 100)));
                  return (
                    <div key={e.id} className="deadline-item">
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <div>
                          <span className="deadline-name">{e.name}</span>
                          <span style={{ fontSize: 12, color: "#64748b", marginLeft: 8 }}>· {e.department}</span>
                        </div>
                        <div style={{ textAlign: "right" }}>
                          <div className="deadline-days" style={{ color: urgentColor }}>{days}d left</div>
                          <div className="deadline-date">Due {formatDate(e.deadline.toISOString().split("T")[0])}</div>
                        </div>
                      </div>
                      <div className="progress-track">
                        <div className="progress-fill" style={{ width: `${pct}%`, background: urgentColor }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ─── DEPARTMENTS ─── */}
          {tab === "departments" && (
            <div className="fadein">
              <div className="section-header" style={{ marginBottom: 20 }}>
                <div><div className="section-title">Departments</div><div className="section-sub">{allDepts.length} departments configured</div></div>
                <button className="btn btn-primary" onClick={() => setAddDeptModal(true)}>+ New Department</button>
              </div>
              <div className="dept-grid">
                {allDepts.map((dept) => {
                  const count = employees.filter((e) => e.department === dept).length;
                  const passed = employees.filter((e) => e.department === dept && e.status === "passed").length;
                  const color = getDeptColor(dept);
                  return (
                    <div key={dept} className="dept-card">
                      <div className="dept-color-circle" style={{ background: color + "20", color }}>
                        {dept[0]}
                      </div>
                      <div className="dept-info">
                        <div className="dept-name">{dept}</div>
                        <div className="dept-count">{count} employees · {count ? Math.round(passed/count*100) : 0}% passed</div>
                        <div style={{ marginTop: 8 }}>
                          <div className="progress-track" style={{ height: 4 }}>
                            <div className="progress-fill" style={{ width: `${count ? Math.round(passed/count*100) : 0}%`, background: color }} />
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
                {/* Add new dept card */}
                <div className="dept-card" style={{ cursor: "pointer", border: "1px dashed rgba(255,255,255,0.1)", background: "transparent" }} onClick={() => setAddDeptModal(true)}>
                  <div className="dept-color-circle" style={{ background: "rgba(99,102,241,0.1)", color: "#6366f1", fontSize: 22 }}>+</div>
                  <div className="dept-info">
                    <div className="dept-name" style={{ color: "#6366f1" }}>Add Department</div>
                    <div className="dept-count">Click to create new</div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* ─── RESCHEDULE MODAL ─── */}
      {rescheduleModal && (
        <div className="modal-overlay" onClick={() => setRescheduleModal(null)}>
          <div className="modal slideup" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">{rescheduleModal.examDate ? "Reschedule Exam" : "Schedule Exam"}</div>
              <div className="modal-sub">{rescheduleModal.name} · {rescheduleModal.department}</div>
            </div>
            <div className="modal-body">
              <div style={{ background: "rgba(255,255,255,0.02)", borderRadius: 12, padding: 14, marginBottom: 20 }}>
                {[
                  { label: "Start Date", value: formatDate(rescheduleModal.startDate) },
                  { label: "Month-6 Deadline", value: formatDate(getDeadline(rescheduleModal.startDate).toISOString().split("T")[0]), color: "#f97316" },
                  rescheduleModal.examDate ? { label: "Current Exam Date", value: formatDate(rescheduleModal.examDate) } : null,
                ].filter(Boolean).map((row) => (
                  <div key={row!.label} className="info-row">
                    <span style={{ fontSize: 13, color: "#64748b" }}>{row!.label}</span>
                    <span style={{ fontSize: 13, fontWeight: 600, color: row!.color || "#94a3b8" }}>{row!.value}</span>
                  </div>
                ))}
              </div>
              <div className="form-group">
                <label className="form-label">New Exam Date</label>
                <input type="date" className="form-input" value={rescheduleDate} onChange={(e) => setRescheduleDate(e.target.value)}
                  max={getDeadline(rescheduleModal.startDate).toISOString().split("T")[0]} />
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" style={{ flex: 1 }} onClick={() => setRescheduleModal(null)}>Cancel</button>
              <button className="btn btn-primary" style={{ flex: 2 }} onClick={confirmReschedule} disabled={saving}>{saving ? "Saving…" : "Confirm Schedule"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ─── SCORE MODAL ─── */}
      {scoreModal && (
        <div className="modal-overlay" onClick={() => setScoreModal(null)}>
          <div className="modal slideup" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Enter Exam Score</div>
              <div className="modal-sub">{scoreModal.name} · Attempt #{(scoreModal.attempts || 0) + 1}</div>
            </div>
            <div className="modal-body">
              <div style={{ textAlign: "center", padding: "20px 0" }}>
                <input type="number" min="0" max="100" value={scoreInput} onChange={(e) => setScoreInput(e.target.value)}
                  placeholder="0 – 100"
                  style={{ width: 140, background: "rgba(255,255,255,0.05)", border: "2px solid rgba(255,255,255,0.1)", color: "#f1f5f9", borderRadius: 16, padding: "16px", fontSize: 36, textAlign: "center", fontFamily: "'Space Grotesk',sans-serif", fontWeight: 700, transition: "all 0.2s", outline: "none" }} />
                {scoreInput !== "" && (
                  <div style={{ marginTop: 16, fontSize: 15, fontWeight: 600 }}>
                    {parseInt(scoreInput) >= PASS_MARK
                      ? <span style={{ color: "#22c55e" }}>✓ PASS — above {PASS_MARK}% threshold</span>
                      : <span style={{ color: "#ef4444" }}>✗ FAIL — below {PASS_MARK}% threshold</span>}
                  </div>
                )}
                <div style={{ marginTop: 12, fontSize: 12, color: "#475569" }}>Pass mark: {PASS_MARK}%</div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" style={{ flex: 1 }} onClick={() => setScoreModal(null)}>Cancel</button>
              <button className="btn btn-primary" style={{ flex: 2 }} onClick={confirmScore} disabled={saving}>{saving ? "Saving…" : "Record Score"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ─── ADD EMPLOYEE MODAL ─── */}
      {addEmpModal && (
        <div className="modal-overlay" onClick={() => setAddEmpModal(false)}>
          <div className="modal slideup" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Add New Employee</div>
              <div className="modal-sub">They'll be added as unscheduled</div>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label className="form-label">Full Name</label>
                <input type="text" className="form-input" value={newEmp.name} onChange={(e) => setNewEmp((p) => ({ ...p, name: e.target.value }))} placeholder="e.g. Jane Smith" />
              </div>
              <div className="form-group">
                <label className="form-label">Department</label>
                <select className="form-input" value={newEmp.department} onChange={(e) => setNewEmp((p) => ({ ...p, department: e.target.value }))}>
                  {allDepts.map((d) => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>
              <div className="form-group">
                <label className="form-label">Start Date</label>
                <input type="date" className="form-input" value={newEmp.startDate} onChange={(e) => setNewEmp((p) => ({ ...p, startDate: e.target.value }))} />
              </div>
              {newEmp.startDate && (
                <div style={{ background: "rgba(99,102,241,0.08)", border: "1px solid rgba(99,102,241,0.15)", borderRadius: 10, padding: "10px 14px", fontSize: 13, color: "#a5b4fc" }}>
                  📅 Month-6 deadline: <strong>{formatDate(getDeadline(newEmp.startDate).toISOString().split("T")[0])}</strong>
                </div>
              )}
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" style={{ flex: 1 }} onClick={() => setAddEmpModal(false)}>Cancel</button>
              <button className="btn btn-primary" style={{ flex: 2 }} onClick={addEmployeeFn} disabled={saving || !newEmp.name || !newEmp.startDate}>{saving ? "Adding…" : "Add Employee"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ─── ADD DEPARTMENT MODAL ─── */}
      {addDeptModal && (
        <div className="modal-overlay" onClick={() => setAddDeptModal(false)}>
          <div className="modal slideup" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">New Department</div>
              <div className="modal-sub">Choose a name and colour</div>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label className="form-label">Department Name</label>
                <input type="text" className="form-input" value={newDept.name} onChange={(e) => setNewDept((p) => ({ ...p, name: e.target.value }))} placeholder="e.g. Customer Success" />
              </div>
              <div className="form-group">
                <label className="form-label">Colour</label>
                <div className="color-picker-row">
                  {PALETTE.map((c) => (
                    <div key={c} className={`color-swatch ${newDept.color === c ? "selected" : ""}`}
                      style={{ background: c }} onClick={() => setNewDept((p) => ({ ...p, color: c }))} />
                  ))}
                </div>
              </div>
              {newDept.name && (
                <div style={{ marginTop: 4 }}>
                  <label className="form-label">Preview</label>
                  <span className="dept-pill" style={{ background: newDept.color + "20", color: newDept.color, fontSize: 13, padding: "6px 14px" }}>
                    <span className="dept-dot" style={{ background: newDept.color }} />
                    {newDept.name}
                  </span>
                </div>
              )}
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" style={{ flex: 1 }} onClick={() => setAddDeptModal(false)}>Cancel</button>
              <button className="btn btn-primary" style={{ flex: 2 }} onClick={addDeptFn} disabled={!newDept.name.trim()}>Create Department</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

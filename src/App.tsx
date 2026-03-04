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
type EmployeeStatus  = "passed" | "failed" | "scheduled" | "unscheduled";

interface Escalation { level: EscalationLevel; label: string; detail: string; }
interface Employee {
  id: number; name: string; department: string; startDate: string;
  examDate: string | null; examScore: number | null; attempts: number;
  status: EmployeeStatus;
}
interface EnrichedEmployee extends Employee {
  escalation: Escalation | null; deadline: Date; deadlineValid: boolean;
}

// ── Safe date helpers ──────────────────────────────────────────
function safeParseDate(raw: unknown): Date | null {
  if (raw === null || raw === undefined || raw === "" || raw === "null" || raw === "undefined") return null;
  const s = String(raw).trim();
  if (!s) return null;
  const direct = new Date(s);
  if (!isNaN(direct.getTime())) return direct;
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) {
    const d = new Date(`${ddmm[3]}-${ddmm[2].padStart(2,"0")}-${ddmm[1].padStart(2,"0")}`);
    if (!isNaN(d.getTime())) return d;
  }
  const serial = Number(s);
  if (!isNaN(serial) && serial > 40000) {
    const d = new Date((serial - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}
function formatDate(raw: unknown): string {
  const d = safeParseDate(raw);
  if (!d) return "—";
  try { return d.toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" }); } catch { return "—"; }
}
function toISO(raw: unknown): string | null {
  const d = safeParseDate(raw);
  if (!d) return null;
  try { return d.toISOString().split("T")[0]; } catch { return null; }
}
function getDeadline(s: unknown): { date: Date; valid: boolean } {
  const start = safeParseDate(s);
  if (!start) return { date: new Date(), valid: false };
  const d = new Date(start); d.setMonth(d.getMonth() + 6);
  return { date: d, valid: true };
}
function daysFromNow(d: Date): number {
  return Math.ceil((d.getTime() - Date.now()) / 86400000);
}
function getEscalation(emp: Employee, deadline: Date, valid: boolean): Escalation | null {
  if (emp.status === "passed") return null;
  if (!valid) return { level:"medium", label:"Check Date", detail:"Start date may be invalid" };
  const days = daysFromNow(deadline);
  if (days < 0) return { level:"critical", label:"Overdue",      detail:`${Math.abs(days)}d past deadline` };
  if (emp.status === "failed" && emp.attempts >= 3) return { level:"critical", label:"Max Attempts", detail:"Escalate to manager" };
  if (emp.status === "failed") return { level:"high",   label:"Failed",       detail:`Resit required · ${days}d left` };
  if (emp.status === "unscheduled") return { level: days < 60 ? "high":"medium", label:"Not Scheduled", detail:`${days}d to deadline` };
  if (days < 30) return { level:"high",   label:"Urgent",       detail:`${days}d to deadline` };
  if (days < 60) return { level:"medium", label:"Watch",        detail:`${days}d to deadline` };
  return { level:"low", label:"On Track", detail:`${days}d to deadline` };
}
const ESC: Record<EscalationLevel,{pill:string;text:string;dot:string}> = {
  critical:{ pill:"rgba(239,68,68,0.15)",  text:"#ef4444", dot:"#ef4444" },
  high:    { pill:"rgba(249,115,22,0.15)", text:"#f97316", dot:"#f97316" },
  medium:  { pill:"rgba(234,179,8,0.15)",  text:"#eab308", dot:"#eab308" },
  low:     { pill:"rgba(34,197,94,0.15)",  text:"#22c55e", dot:"#22c55e" },
};
function parseEmployee(raw: Record<string,unknown>): Employee {
  const status = String(raw.status??"").toLowerCase().trim();
  const valid: EmployeeStatus[] = ["passed","failed","scheduled","unscheduled"];
  const safeStatus: EmployeeStatus = valid.includes(status as EmployeeStatus) ? (status as EmployeeStatus) : "unscheduled";
  const scoreRaw   = raw.examScore ?? raw.ExamScore ?? null;
  const examScore  = (scoreRaw !== null && scoreRaw !== "" && !isNaN(Number(scoreRaw))) ? Number(scoreRaw) : null;
  const examDateRaw  = raw.examDate  ?? raw.ExamDate  ?? null;
  const startDateRaw = raw.startDate ?? raw.StartDate ?? "";
  return {
    id: Number(raw.id ?? 0),
    name: String(raw.name ?? raw.Name ?? "").trim(),
    department: String(raw.department ?? raw.Department ?? "").trim(),
    startDate: toISO(startDateRaw) ?? String(startDateRaw),
    examDate: toISO(examDateRaw),
    examScore, attempts: Number(raw.attempts ?? 0) || 0, status: safeStatus,
  };
}

// ── Theme tokens ────────────────────────────────────────────────
const DARK = {
  bg:       "#080b14", bg2: "#0d111e", bg3: "rgba(255,255,255,0.03)",
  border:   "rgba(255,255,255,0.07)", border2:"rgba(255,255,255,0.06)",
  text:     "#e2e8f0", text2:"#94a3b8", text3:"#64748b", text4:"#475569",
  sb:       "rgba(10,13,24,0.98)", tb:"rgba(8,11,20,0.92)",
  input:    "rgba(255,255,255,0.05)", inputBd:"rgba(255,255,255,0.08)",
  row2:     "rgba(255,255,255,0.01)", rowHov:"rgba(255,255,255,0.03)",
  rowSel:   "rgba(99,102,241,0.06)",  rowHd:"rgba(255,255,255,0.02)",
  alertBg:  "rgba(239,68,68,0.06)",   alertBd:"rgba(239,68,68,0.15)",
  infoBg:   "rgba(255,255,255,0.02)", infoHd:"rgba(255,255,255,0.06)",
  pmBg:     "rgba(99,102,241,0.09)",  pmBd:"rgba(99,102,241,0.18)",
  scPass:   "rgba(34,197,94,0.08)",   scPassBd:"rgba(34,197,94,0.18)",
  scTotal:  "rgba(99,102,241,0.08)",  scTotalBd:"rgba(99,102,241,0.18)",
  scFail:   "rgba(249,115,22,0.08)",  scFailBd:"rgba(249,115,22,0.18)",
  scCrit:   "rgba(239,68,68,0.08)",   scCritBd:"rgba(239,68,68,0.18)",
  toggle:   "#1e2535",
};
const LIGHT = {
  bg:       "#f1f5f9", bg2:"#ffffff", bg3:"rgba(0,0,0,0.02)",
  border:   "rgba(0,0,0,0.08)", border2:"rgba(0,0,0,0.06)",
  text:     "#0f172a", text2:"#334155", text3:"#475569", text4:"#64748b",
  sb:       "rgba(255,255,255,0.98)", tb:"rgba(241,245,249,0.94)",
  input:    "rgba(0,0,0,0.04)", inputBd:"rgba(0,0,0,0.1)",
  row2:     "rgba(0,0,0,0.015)", rowHov:"rgba(0,0,0,0.025)",
  rowSel:   "rgba(99,102,241,0.06)", rowHd:"rgba(0,0,0,0.02)",
  alertBg:  "rgba(239,68,68,0.05)", alertBd:"rgba(239,68,68,0.2)",
  infoBg:   "rgba(0,0,0,0.02)", infoHd:"rgba(0,0,0,0.06)",
  pmBg:     "rgba(99,102,241,0.07)", pmBd:"rgba(99,102,241,0.15)",
  scPass:   "rgba(34,197,94,0.08)",  scPassBd:"rgba(34,197,94,0.2)",
  scTotal:  "rgba(99,102,241,0.08)", scTotalBd:"rgba(99,102,241,0.2)",
  scFail:   "rgba(249,115,22,0.08)", scFailBd:"rgba(249,115,22,0.2)",
  scCrit:   "rgba(239,68,68,0.08)",  scCritBd:"rgba(239,68,68,0.2)",
  toggle:   "#e2e8f0",
};

const TCOL = "minmax(0,2fr) minmax(0,1.2fr) minmax(0,1fr) minmax(0,0.7fr) minmax(0,0.6fr) minmax(0,1.2fr) minmax(0,1.3fr)";

export default function App() {
  const [dark, setDark]               = useState(true);
  const [employees, setEmployees]     = useState<Employee[]>([]);
  const [deptColors, setDeptColors]   = useState<Record<string,string>>(DEFAULT_DEPT_COLORS);
  const [loading, setLoading]         = useState(true);
  const [saving, setSaving]           = useState(false);
  const [error, setError]             = useState<string|null>(null);
  const [lastSaved, setLastSaved]     = useState<Date|null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [tab, setTab]                 = useState("dashboard");
  const [search, setSearch]           = useState("");
  const [filterDept, setFilterDept]   = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterEsc, setFilterEsc]     = useState("All");
  const [selectedRow, setSelectedRow] = useState<number|null>(null);

  const [reschedModal, setReschedModal] = useState<EnrichedEmployee|null>(null);
  const [reschedDate,  setReschedDate]  = useState("");
  const [scoreModal,   setScoreModal]   = useState<EnrichedEmployee|null>(null);
  const [scoreInput,   setScoreInput]   = useState("");
  const [addEmpModal,  setAddEmpModal]  = useState(false);
  const [newEmp, setNewEmp]             = useState({ name:"", department:"Finance", startDate:"" });
  const [addDeptModal, setAddDeptModal] = useState(false);
  const [newDept, setNewDept]           = useState({ name:"", color:PALETTE[0] });

  const T = dark ? DARK : LIGHT;

  // Persist dark/light preference
  useEffect(() => {
    const saved = localStorage.getItem("neo-theme");
    if (saved === "light") setDark(false);
  }, []);
  const toggleTheme = () => {
    setDark((d) => { localStorage.setItem("neo-theme", d ? "light" : "dark"); return !d; });
  };

  const fetchData = useCallback(async () => {
    setLoading(true); setError(null);
    try {
      const res  = await fetch(SCRIPT_URL);
      const json = await res.json() as { success:boolean; data?:Record<string,unknown>[]; error?:string };
      if (json.success && Array.isArray(json.data)) {
        const parsed = json.data
          .filter((r) => r && typeof r === "object" && Object.keys(r).length > 0)
          .map((r) => parseEmployee(r as Record<string,unknown>))
          .filter((e) => e.name.length > 0);
        setEmployees(parsed);
      } else {
        setError("Could not load: " + (json.error ?? "Empty response from Sheets"));
      }
    } catch (err) {
      console.error(err);
      setError("Cannot reach Google Sheets. Check SCRIPT_URL and deployment.");
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  const saveEmployee = async (u: Employee) => {
    setSaving(true);
    try {
      await fetch(SCRIPT_URL, { method:"POST", body: JSON.stringify({ action:"update", employee:u }) });
      setEmployees((p) => p.map((e) => e.id === u.id ? u : e));
      setLastSaved(new Date());
    } catch { setError("Save failed — check connection."); }
    finally { setSaving(false); }
  };

  const addEmployeeFn = async () => {
    if (!newEmp.name.trim() || !newEmp.startDate) return;
    const id  = employees.length ? Math.max(...employees.map((e) => e.id)) + 1 : 1;
    const emp: Employee = { id, name:newEmp.name.trim(), department:newEmp.department, startDate:newEmp.startDate, examDate:null, examScore:null, attempts:0, status:"unscheduled" };
    setSaving(true);
    try {
      await fetch(SCRIPT_URL, { method:"POST", body: JSON.stringify({ action:"add", employee:emp }) });
      setEmployees((p) => [...p, emp]);
      setLastSaved(new Date());
      setAddEmpModal(false);
      setNewEmp({ name:"", department:newEmp.department, startDate:"" });
    } catch { setError("Failed to add."); }
    finally { setSaving(false); }
  };

  const addDeptFn = () => {
    const name = newDept.name.trim();
    if (!name) return;
    setDeptColors((p) => ({ ...p, [name]: newDept.color }));
    setNewDept({ name:"", color:PALETTE[0] });
    setAddDeptModal(false);
  };

  const gc = (d: string) => deptColors[d] ?? "#94a3b8";

  const allDepts = useMemo(() =>
    Array.from(new Set([...Object.keys(deptColors), ...employees.map((e) => e.department).filter(Boolean)])).sort(),
    [deptColors, employees]);

  const enriched: EnrichedEmployee[] = useMemo(() =>
    employees.map((e) => {
      const { date, valid } = getDeadline(e.startDate);
      return { ...e, escalation: getEscalation(e, date, valid), deadline: date, deadlineValid: valid };
    }), [employees]);

  const stats = useMemo(() => {
    const total        = enriched.length;
    const passed       = enriched.filter((e) => e.status === "passed").length;
    const failed       = enriched.filter((e) => e.status === "failed").length;
    const examTaken    = enriched.filter((e) => e.examScore !== null).length;   // sat exam at least once
    const critical     = enriched.filter((e) => e.escalation?.level === "critical").length;
    const overdue      = enriched.filter((e) => e.escalation?.label === "Overdue").length;
    const totalAttempts = enriched.reduce((sum, e) => sum + (e.attempts || 0), 0);
    const avgAttempts  = examTaken > 0 ? (totalAttempts / examTaken).toFixed(1) : "0";
    // Pass rate = passed ÷ employees who have sat the exam (not total enrolled)
    const passRate     = examTaken > 0 ? Math.round(passed / examTaken * 100) : 0;
    return { total, passed, failed, examTaken, critical, overdue, totalAttempts, avgAttempts, passRate };
  }, [enriched]);

  const filtered = useMemo(() => enriched.filter((e) => {
    const q = search.toLowerCase();
    return (!q || e.name.toLowerCase().includes(q) || e.department.toLowerCase().includes(q))
      && (filterDept   === "All" || e.department === filterDept)
      && (filterStatus === "All" || e.status === filterStatus)
      && (filterEsc    === "All" || e.escalation?.level === filterEsc);
  }), [enriched, search, filterDept, filterStatus, filterEsc]);

  const byMonth = useMemo(() => {
    const map: Record<string,{scheduled:number;passed:number;failed:number}> = {};
    enriched.forEach((e) => {
      if (!e.examDate) return;
      const k = e.examDate.slice(0,7);
      if (!map[k]) map[k] = { scheduled:0, passed:0, failed:0 };
      map[k].scheduled++;
      if (e.status==="passed") map[k].passed++;
      if (e.status==="failed") map[k].failed++;
    });
    return Object.entries(map).sort(([a],[b]) => a.localeCompare(b)).slice(-6);
  }, [enriched]);

  const criticalList = useMemo(() => enriched.filter((e) => e.escalation?.level==="critical"||e.escalation?.level==="high"), [enriched]);
  const upcomingDL   = useMemo(() => enriched.filter((e) => { const d=daysFromNow(e.deadline); return e.deadlineValid&&d>=0&&d<=90&&e.status!=="passed"; }).sort((a,b)=>a.deadline.getTime()-b.deadline.getTime()), [enriched]);

  const confirmResched = async () => {
    if (!reschedDate||!reschedModal) return;
    await saveEmployee({ ...reschedModal, examDate:reschedDate, status:reschedModal.status==="unscheduled"?"scheduled":reschedModal.status });
    setReschedModal(null); setReschedDate("");
  };
  const confirmScore = async () => {
    if (!scoreModal) return;
    const score = parseInt(scoreInput, 10);
    if (isNaN(score)||score<0||score>100) return;
    await saveEmployee({ ...scoreModal, examScore:score, status:score>=PASS_MARK?"passed":"failed", attempts:scoreModal.examScore===null?(scoreModal.attempts||0)+1:scoreModal.attempts });
    setScoreModal(null); setScoreInput("");
  };

  const navGo = (id: string) => { setTab(id); setSidebarOpen(false); };
  const META: Record<string,{title:string;sub:string}> = {
    dashboard:   { title:"Dashboard",   sub:"NEO exam overview & escalations" },
    employees:   { title:"Employees",   sub:`${stats.total} employees in programme` },
    calendar:    { title:"Schedule",    sub:"Exam dates & month-6 deadlines" },
    departments: { title:"Departments", sub:`${allDepts.length} departments configured` },
  };

  // ── Dynamic CSS with theme variables ─────────────────────────
  const CSS = `
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap');
    *, *::before, *::after { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:'Plus Jakarta Sans',sans-serif; background:${T.bg}; transition:background .3s; }
    ::-webkit-scrollbar { width:4px; height:4px; }
    ::-webkit-scrollbar-thumb { background:${dark?"rgba(255,255,255,0.08)":"rgba(0,0,0,0.12)"}; border-radius:10px; }
    input,select,button { font-family:'Plus Jakarta Sans',sans-serif; }
    input[type=date]::-webkit-calendar-picker-indicator { filter:${dark?"invert(.5)":"invert(.3)"}; cursor:pointer; }

    .neo { min-height:100vh; background:${T.bg}; color:${T.text}; display:flex; transition:background .3s,color .3s; }

    /* Sidebar */
    .sb { position:fixed; left:0; top:0; bottom:0; width:224px; background:${T.sb}; border-right:1px solid ${T.border2}; z-index:50; display:flex; flex-direction:column; transition:transform .3s cubic-bezier(.4,0,.2,1),background .3s; }
    .sb-logo { padding:20px 16px 16px; display:flex; align-items:center; gap:11px; border-bottom:1px solid ${T.border2}; }
    .sb-icon { width:36px; height:36px; border-radius:10px; background:linear-gradient(135deg,#6366f1,#8b5cf6); display:flex; align-items:center; justify-content:center; font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:16px; color:#fff; flex-shrink:0; box-shadow:0 5px 14px rgba(99,102,241,.35); }
    .sb-title { font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:13px; color:${T.text}; transition:color .3s; }
    .sb-sub   { font-size:10px; color:${T.text3}; margin-top:1px; }
    .sb-nav   { padding:12px 8px; flex:1; overflow-y:auto; }
    .sb-lbl   { font-size:10px; font-weight:600; letter-spacing:.1em; color:${T.text4}; text-transform:uppercase; padding:0 8px; margin-bottom:4px; margin-top:10px; }
    .sb-item  { display:flex; align-items:center; gap:9px; padding:9px 11px; border-radius:9px; cursor:pointer; transition:all .2s; color:${T.text3}; font-size:13px; font-weight:500; border:1px solid transparent; background:none; width:100%; text-align:left; }
    .sb-item:hover { background:${dark?"rgba(255,255,255,0.05)":"rgba(0,0,0,0.04)"}; color:${T.text2}; }
    .sb-item.on { background:rgba(99,102,241,.13); color:#a5b4fc; border-color:rgba(99,102,241,.2); }
    .sb-badge { margin-left:auto; background:#ef4444; color:#fff; border-radius:20px; padding:1px 7px; font-size:10px; font-weight:700; }
    .sb-foot  { padding:12px 8px; border-top:1px solid ${T.border2}; }
    .pm       { background:${T.pmBg}; border:1px solid ${T.pmBd}; border-radius:8px; padding:9px 12px; font-size:12px; color:${T.text3}; display:flex; justify-content:space-between; align-items:center; }
    .pm strong { color:#a5b4fc; font-size:14px; }

    /* Theme toggle */
    .theme-toggle { display:flex; align-items:center; justify-content:space-between; padding:9px 12px; border-radius:8px; background:${T.toggle}; border:none; cursor:pointer; width:100%; margin-bottom:8px; color:${T.text2}; font-size:12px; font-weight:600; transition:all .2s; }
    .theme-toggle:hover { filter:brightness(${dark?"1.1":"0.95"}); }
    .toggle-track { width:36px; height:20px; border-radius:20px; background:${dark?"#6366f1":"#cbd5e1"}; position:relative; transition:background .3s; flex-shrink:0; }
    .toggle-knob  { position:absolute; top:2px; left:${dark?"18px":"2px"}; width:16px; height:16px; border-radius:50%; background:#fff; transition:left .3s; box-shadow:0 1px 4px rgba(0,0,0,.25); }

    /* Main */
    .mn { margin-left:224px; flex:1; min-height:100vh; display:flex; flex-direction:column; min-width:0; }
    .tb { position:sticky; top:0; background:${T.tb}; backdrop-filter:blur(20px); border-bottom:1px solid ${T.border2}; padding:0 22px; height:60px; display:flex; align-items:center; justify-content:space-between; z-index:40; gap:10px; transition:background .3s; }
    .tb-l { display:flex; align-items:center; gap:10px; min-width:0; }
    .tb-title { font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:17px; color:${T.text}; transition:color .3s; }
    .tb-sub   { font-size:11px; color:${T.text4}; margin-top:1px; }
    .tb-r     { display:flex; align-items:center; gap:7px; flex-shrink:0; }
    .btn  { display:inline-flex; align-items:center; gap:5px; padding:7px 13px; border-radius:9px; font-size:13px; font-weight:600; cursor:pointer; transition:all .2s; border:none; white-space:nowrap; }
    .btn:disabled { opacity:.5; cursor:not-allowed; }
    .bp   { background:linear-gradient(135deg,#6366f1,#8b5cf6); color:#fff; box-shadow:0 3px 10px rgba(99,102,241,.28); }
    .bp:hover:not(:disabled) { transform:translateY(-1px); box-shadow:0 5px 16px rgba(99,102,241,.38); }
    .bg   { background:${T.input}; color:${T.text2}; border:1px solid ${T.inputBd}; }
    .bg:hover:not(:disabled) { background:${dark?"rgba(255,255,255,0.09)":"rgba(0,0,0,0.07)"}; color:${T.text}; }
    .ba   { background:rgba(99,102,241,.1); color:#a5b4fc; border:1px solid rgba(99,102,241,.2); }
    .bsm  { padding:5px 10px; font-size:12px; }

    /* Content */
    .ct  { padding:22px; flex:1; min-width:0; }

    /* Stats */
    .sg  { display:grid; grid-template-columns:repeat(4,1fr); gap:13px; margin-bottom:20px; }
    .sc  { border-radius:13px; padding:18px; position:relative; overflow:hidden; transition:transform .2s,background .3s; }
    .sc:hover { transform:translateY(-2px); }
    .sc-ic  { position:absolute; right:14px; top:14px; font-size:20px; opacity:.15; }
    .sc-lbl { font-size:11px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; margin-bottom:9px; }
    .sc-val { font-family:'Space Grotesk',sans-serif; font-size:36px; font-weight:700; line-height:1; margin-bottom:4px; }
    .sc-sub { font-size:11px; opacity:.7; }

    /* Cards */
    .card { background:${T.bg3}; border:1px solid ${T.border}; border-radius:13px; overflow:hidden; transition:background .3s,border .3s; }
    .cp   { padding:18px; }
    .g2   { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:16px; }

    /* Insight strip */
    .insight-strip { display:grid; grid-template-columns:repeat(3,1fr); gap:13px; margin-bottom:20px; }
    .insight-card  { background:${T.bg3}; border:1px solid ${T.border}; border-radius:13px; padding:16px 18px; display:flex; align-items:center; gap:14px; transition:background .3s; }
    .insight-icon  { font-size:26px; flex-shrink:0; }
    .insight-val   { font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:22px; color:${T.text}; line-height:1; }
    .insight-lbl   { font-size:11px; color:${T.text4}; margin-top:3px; font-weight:500; }
    .insight-sub   { font-size:11px; color:${T.text3}; margin-top:2px; }

    /* Section header */
    .sh  { display:flex; align-items:center; justify-content:space-between; margin-bottom:13px; }
    .sh-t { font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:14px; color:${T.text}; transition:color .3s; }
    .sh-s { font-size:11px; color:${T.text4}; margin-top:1px; }

    /* Progress */
    .pt  { height:5px; background:${dark?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.07)"}; border-radius:10px; overflow:hidden; }
    .pf  { height:100%; border-radius:10px; transition:width .5s cubic-bezier(.4,0,.2,1); }

    /* Pills */
    .pill { display:inline-flex; align-items:center; gap:5px; border-radius:20px; padding:3px 9px; font-size:11px; font-weight:600; }
    .dot  { width:6px; height:6px; border-radius:50%; flex-shrink:0; }

    /* Filters */
    .fr   { display:flex; gap:8px; margin-bottom:16px; flex-wrap:wrap; align-items:center; }
    .si   { position:relative; flex:1; min-width:160px; max-width:280px; }
    .si-ic{ position:absolute; left:10px; top:50%; transform:translateY(-50%); color:${T.text4}; font-size:12px; pointer-events:none; }
    .si input { width:100%; background:${T.input}; border:1px solid ${T.inputBd}; color:${T.text}; border-radius:9px; padding:8px 10px 8px 30px; font-size:13px; transition:all .2s; }
    .si input:focus { background:${dark?"rgba(255,255,255,0.07)":"rgba(0,0,0,0.06)"}; border-color:rgba(99,102,241,.4); box-shadow:0 0 0 3px rgba(99,102,241,.1); outline:none; }
    .si input::placeholder { color:${T.text4}; }
    .fsel { background:${T.input}; border:1px solid ${T.inputBd}; color:${T.text2}; border-radius:9px; padding:8px 10px; font-size:13px; cursor:pointer; transition:all .2s; }
    .fsel:focus { outline:none; border-color:rgba(99,102,241,.4); }
    .fsel option { background:${T.bg2}; color:${T.text}; }
    .rc   { margin-left:auto; font-size:12px; color:${T.text4}; }

    /* Table */
    .tw  { border-radius:13px; overflow:auto; border:1px solid ${T.border}; transition:border .3s; }
    .th  { padding:10px 16px; background:${T.rowHd}; border-bottom:1px solid ${T.border2}; display:grid; gap:8px; min-width:700px; transition:background .3s; }
    .thc { font-size:10px; font-weight:600; letter-spacing:.07em; text-transform:uppercase; color:${T.text4}; }
    .tr  { padding:11px 16px; border-bottom:1px solid ${T.border2}; transition:background .15s; display:grid; gap:8px; align-items:center; cursor:pointer; min-width:700px; }
    .tr:last-child { border-bottom:none; }
    .tr:hover { background:${T.rowHov}; }
    .tr.sel   { background:${T.rowSel}; }
    .en  { font-weight:600; color:${T.text}; font-size:13px; }
    .em  { font-size:10px; color:${T.text4}; margin-top:2px; }
    .av  { width:30px; height:30px; border-radius:8px; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:700; flex-shrink:0; }
    .ac  { display:flex; gap:5px; flex-wrap:nowrap; }

    /* Month cards */
    .mc  { background:${T.bg3}; border:1px solid ${T.border}; border-radius:13px; padding:17px; transition:transform .2s,background .3s,border .3s; }
    .mc:hover { transform:translateY(-2px); border-color:rgba(99,102,241,.2); }
    .g3  { display:grid; grid-template-columns:repeat(3,1fr); gap:13px; margin-bottom:20px; }

    /* Dept grid */
    .dg  { display:grid; grid-template-columns:repeat(auto-fill,minmax(170px,1fr)); gap:11px; }
    .dc  { background:${T.bg3}; border:1px solid ${T.border}; border-radius:11px; padding:14px; display:flex; align-items:center; gap:11px; transition:all .2s; }
    .dc:hover { border-color:${dark?"rgba(255,255,255,0.14)":"rgba(0,0,0,0.14)"}; background:${dark?"rgba(255,255,255,0.05)":"rgba(0,0,0,0.04)"}; }
    .dci { width:36px; height:36px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:15px; font-weight:700; flex-shrink:0; }
    .dcn { font-weight:600; font-size:13px; color:${T.text}; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .dcs { font-size:11px; color:${T.text4}; margin-top:2px; }
    .cs  { width:24px; height:24px; border-radius:6px; cursor:pointer; transition:transform .15s; border:2px solid transparent; }
    .cs:hover { transform:scale(1.15); }
    .cs.sel { border-color:${dark?"#fff":"#000"}; transform:scale(1.1); }

    /* Modals */
    .ov  { position:fixed; inset:0; background:rgba(0,0,0,0.7); backdrop-filter:blur(6px); display:flex; align-items:center; justify-content:center; z-index:200; padding:16px; }
    .md  { background:${T.bg2}; border:1px solid ${T.border}; border-radius:16px; width:100%; max-width:440px; box-shadow:0 24px 56px rgba(0,0,0,.4); animation:slideup .22s ease; transition:background .3s; }
    .mh  { padding:20px 20px 0; }
    .mt  { font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:18px; color:${T.text}; }
    .ms  { font-size:12px; color:${T.text4}; margin-top:3px; }
    .mb  { padding:16px 20px; }
    .mf  { padding:0 20px 20px; display:flex; gap:9px; }
    .fl  { font-size:11px; font-weight:600; letter-spacing:.07em; text-transform:uppercase; color:${T.text4}; display:block; margin-bottom:6px; }
    .fi  { width:100%; background:${T.input}; border:1px solid ${T.inputBd}; color:${T.text}; border-radius:9px; padding:9px 12px; font-size:14px; transition:all .2s; }
    .fi:focus { border-color:rgba(99,102,241,.5); background:${dark?"rgba(255,255,255,0.07)":"rgba(0,0,0,0.06)"}; box-shadow:0 0 0 3px rgba(99,102,241,.1); outline:none; }
    .fg  { margin-bottom:13px; }
    .ir  { display:flex; justify-content:space-between; padding:7px 0; border-bottom:1px solid ${T.border2}; font-size:13px; }
    .ir:last-child { border-bottom:none; }

    /* Alert */
    .alert { background:${T.alertBg}; border:1px solid ${T.alertBd}; border-radius:13px; padding:16px; margin-bottom:16px; }

    /* Animations */
    @keyframes slideup { from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)} }
    @keyframes fadein  { from{opacity:0;transform:translateY(7px)}  to{opacity:1;transform:translateY(0)} }
    @keyframes pulse   { 0%,100%{opacity:1} 50%{opacity:.4} }
    @keyframes spin    { to{transform:rotate(360deg)} }
    .fadein   { animation:fadein .3s ease forwards; }
    .pulsing  { animation:pulse 1.5s infinite; }
    .spinning { animation:spin 1s linear infinite; display:inline-block; }

    /* Hamburger */
    .hb     { display:none; background:none; border:none; color:${T.text2}; cursor:pointer; padding:5px; border-radius:7px; font-size:19px; }
    .ov-sb  { display:none; position:fixed; inset:0; background:rgba(0,0,0,.5); z-index:49; }

    /* Empty */
    .empty  { text-align:center; padding:44px 20px; }
    .ei     { font-size:36px; margin-bottom:8px; opacity:.3; }

    /* Responsive */
    @media(max-width:1100px) {
      .sg { grid-template-columns:repeat(2,1fr); }
      .g2 { grid-template-columns:1fr; }
      .g3 { grid-template-columns:repeat(2,1fr); }
      .insight-strip { grid-template-columns:1fr 1fr; }
    }
    @media(max-width:768px) {
      .sb { transform:translateX(-100%); }
      .sb.open { transform:translateX(0); box-shadow:4px 0 28px rgba(0,0,0,.4); }
      .ov-sb.open { display:block; }
      .mn { margin-left:0; }
      .hb { display:flex; }
      .ct { padding:13px; }
      .tb { padding:0 13px; }
      .sg { grid-template-columns:repeat(2,1fr); gap:10px; }
      .g3 { grid-template-columns:1fr; }
      .fr { gap:6px; }
      .si { max-width:100%; }
      .hm { display:none; }
      .insight-strip { grid-template-columns:1fr; }
    }
  `;

  // ── Loading screen ───────────────────────────────────────────
  if (loading) return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", fontFamily:"'Plus Jakarta Sans',sans-serif", transition:"background .3s" }}>
      <style>{CSS}</style>
      <div style={{ width:50, height:50, borderRadius:13, background:"linear-gradient(135deg,#6366f1,#8b5cf6)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:22, color:"#fff", fontWeight:700, marginBottom:16, boxShadow:"0 8px 24px rgba(99,102,241,.38)" }}>N</div>
      <div style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, fontSize:19, color:T.text, marginBottom:5 }}>NEO Exam Tracker</div>
      <div style={{ color:T.text4, fontSize:13, marginBottom:22 }}>Connecting to Google Sheets…</div>
      <div style={{ width:160, height:3, background:dark?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.08)", borderRadius:10, overflow:"hidden" }}>
        <div style={{ width:"55%", height:"100%", background:"linear-gradient(90deg,#6366f1,#8b5cf6)", borderRadius:10 }} />
      </div>
    </div>
  );

  // ── Error screen ─────────────────────────────────────────────
  if (error && employees.length === 0) return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", fontFamily:"'Plus Jakarta Sans',sans-serif", padding:24 }}>
      <style>{CSS}</style>
      <div style={{ fontSize:42, marginBottom:12 }}>⚠️</div>
      <div style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, fontSize:19, color:T.text, marginBottom:10 }}>Connection Error</div>
      <div style={{ background:"rgba(239,68,68,0.08)", border:"1px solid rgba(239,68,68,0.2)", borderRadius:11, padding:15, maxWidth:400, fontSize:13, color:"#ef4444", lineHeight:1.7, marginBottom:13, textAlign:"center" }}>{error}</div>
      <p style={{ color:T.text4, fontSize:13, textAlign:"center", maxWidth:340, lineHeight:1.7 }}>Replace <code style={{ color:"#a5b4fc" }}>SCRIPT_URL</code> in App.tsx with your deployed Google Apps Script URL, with <strong>Who has access: Anyone</strong>.</p>
      <button className="btn bp" onClick={fetchData} style={{ marginTop:16 }}>↻ Retry</button>
    </div>
  );

  return (
    <div className="neo">
      <style>{CSS}</style>
      <div className={`ov-sb ${sidebarOpen?"open":""}`} onClick={() => setSidebarOpen(false)} />

      {/* ── Sidebar ── */}
      <aside className={`sb ${sidebarOpen?"open":""}`}>
        <div className="sb-logo">
          <div className="sb-icon">N</div>
          <div><div className="sb-title">NEO Tracker</div><div className="sb-sub">Exam Management</div></div>
        </div>
        <nav className="sb-nav">
          <div className="sb-lbl">Menu</div>
          {[
            { id:"dashboard",   icon:"📊", label:"Dashboard" },
            { id:"employees",   icon:"👥", label:"Employees", badge: stats.critical>0 ? stats.critical : null },
            { id:"calendar",    icon:"📅", label:"Schedule" },
            { id:"departments", icon:"🏢", label:"Departments" },
          ].map((n) => (
            <button key={n.id} className={`sb-item ${tab===n.id?"on":""}`} onClick={() => navGo(n.id)}>
              <span>{n.icon}</span> {n.label}
              {n.badge!==null && <span className="sb-badge">{n.badge}</span>}
            </button>
          ))}
          <div className="sb-lbl">Quick Add</div>
          <button className="sb-item" onClick={() => { setAddEmpModal(true); setSidebarOpen(false); }}><span>➕</span> Add Employee</button>
          <button className="sb-item" onClick={() => { setAddDeptModal(true); setSidebarOpen(false); }}><span>🏷️</span> New Department</button>
        </nav>
        <div className="sb-foot">
          {/* Theme toggle */}
          <button className="theme-toggle" onClick={toggleTheme}>
            <span>{dark ? "☀️ Light Mode" : "🌙 Dark Mode"}</span>
            <div className="toggle-track"><div className="toggle-knob" /></div>
          </button>
          <div className="pm"><span>Pass Mark</span><strong>{PASS_MARK}%</strong></div>
          {lastSaved && <div style={{ marginTop:7, fontSize:11, color:"#22c55e", textAlign:"center" }}>✓ Saved {lastSaved.toLocaleTimeString()}</div>}
        </div>
      </aside>

      {/* ── Main ── */}
      <main className="mn">
        <div className="tb">
          <div className="tb-l">
            <button className="hb" onClick={() => setSidebarOpen(true)}>☰</button>
            <div><div className="tb-title">{META[tab]?.title}</div><div className="tb-sub">{META[tab]?.sub}</div></div>
          </div>
          <div className="tb-r">
            {saving && <span style={{ fontSize:12, color:"#6366f1", display:"flex", alignItems:"center", gap:4 }}><span className="spinning">⟳</span><span className="hm">Saving</span></span>}
            {stats.critical > 0 && (
              <div className="pulsing" style={{ display:"flex", alignItems:"center", gap:5, background:"rgba(239,68,68,.1)", border:"1px solid rgba(239,68,68,.22)", borderRadius:9, padding:"5px 10px", fontSize:12, color:"#f87171" }}>
                <span style={{ width:6, height:6, background:"#ef4444", borderRadius:"50%", display:"inline-block" }} />
                {stats.critical}
              </div>
            )}
            {/* Topbar theme toggle (visible on mobile) */}
            <button className="btn bg bsm" onClick={toggleTheme} title="Toggle theme">{dark?"☀️":"🌙"}</button>
            <button className="btn bg bsm" onClick={fetchData}>↻<span className="hm"> Refresh</span></button>
            <button className="btn bp bsm" onClick={() => setAddEmpModal(true)}>+<span className="hm"> Employee</span></button>
          </div>
        </div>

        <div className="ct fadein">

          {/* ════ DASHBOARD ════ */}
          {tab === "dashboard" && (<>
            {/* Primary stat cards */}
            <div className="sg">
              {[
                { lbl:"Total Enrolled",  val:stats.total,         sub:"In NEO programme",                          c:"#6366f1", bg:T.scTotal,  bd:T.scTotalBd, ic:"👥" },
                { lbl:"Pass Rate",       val:`${stats.passRate}%`,sub:`Of ${stats.examTaken} who sat the exam`,    c:"#22c55e", bg:T.scPass,   bd:T.scPassBd,  ic:"✅" },
                { lbl:"Require Resit",   val:stats.failed,        sub:"Failed — resit needed",                     c:"#f97316", bg:T.scFail,   bd:T.scFailBd,  ic:"📋" },
                { lbl:"Critical Flags",  val:stats.critical,      sub:`${stats.overdue} past deadline`,            c:"#ef4444", bg:T.scCrit,   bd:T.scCritBd,  ic:"🚨" },
              ].map((s) => (
                <div key={s.lbl} className="sc" style={{ background:s.bg, border:`1px solid ${s.bd}` }}>
                  <div className="sc-ic">{s.ic}</div>
                  <div className="sc-lbl" style={{ color:s.c }}>{s.lbl}</div>
                  <div className="sc-val" style={{ color:s.c }}>{s.val}</div>
                  <div className="sc-sub" style={{ color:s.c }}>{s.sub}</div>
                </div>
              ))}
            </div>

            {/* Insight strip — exam-takers, avg attempts, not yet sat */}
            <div className="insight-strip">
              {[
                { icon:"🎓", val:stats.examTaken,    lbl:"Sat the Exam",     sub:`${stats.total - stats.examTaken} yet to sit` },
                { icon:"🔁", val:stats.avgAttempts,  lbl:"Avg Attempts",     sub:`${stats.totalAttempts} total attempts` },
                { icon:"⏳", val:stats.total - stats.examTaken, lbl:"Not Yet Examined", sub:"Scheduled or unscheduled" },
              ].map((s) => (
                <div key={s.lbl} className="insight-card">
                  <div className="insight-icon">{s.icon}</div>
                  <div>
                    <div className="insight-val">{s.val}</div>
                    <div className="insight-lbl">{s.lbl}</div>
                    <div className="insight-sub">{s.sub}</div>
                  </div>
                </div>
              ))}
            </div>

            {/* Critical escalations */}
            {criticalList.length > 0 && (
              <div className="alert">
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12 }}>
                  <span className="dot pulsing" style={{ background:"#ef4444", width:8, height:8 }} />
                  <span style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, color:"#f87171", fontSize:14 }}>Requires Immediate Action</span>
                  <span className="pill" style={{ background:"rgba(239,68,68,.15)", color:"#f87171", marginLeft:"auto" }}>{criticalList.length}</span>
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
                  {criticalList.map((e) => {
                    const cfg = ESC[e.escalation!.level];
                    const c   = gc(e.department);
                    return (
                      <div key={e.id} style={{ background:T.infoBg, border:`1px solid ${T.infoHd}`, borderRadius:9, padding:"10px 13px", display:"flex", alignItems:"center", gap:9, flexWrap:"wrap" }}>
                        <div className="av" style={{ background:c+"22", color:c }}>{e.name[0]??""}</div>
                        <div style={{ flex:1, minWidth:90 }}>
                          <div style={{ fontWeight:600, color:T.text, fontSize:13 }}>{e.name}</div>
                          <div style={{ fontSize:11, color:T.text4 }}>{e.department}</div>
                        </div>
                        <span className="pill" style={{ background:cfg.pill, color:cfg.text }}><span className="dot" style={{ background:cfg.dot }} />{e.escalation!.label}</span>
                        <span style={{ fontSize:12, color:cfg.text }}>{e.escalation!.detail}</span>
                        <button className="btn bg bsm" onClick={() => { setReschedModal(e); setReschedDate(e.examDate??""); }}>Reschedule</button>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            <div className="g2">
              {/* Escalation breakdown */}
              <div className="card cp">
                <div className="sh"><div><div className="sh-t">Escalation Breakdown</div><div className="sh-s">By severity</div></div></div>
                {(["critical","high","medium","low"] as EscalationLevel[]).map((lvl) => {
                  const count = enriched.filter((e) => e.escalation?.level===lvl).length;
                  const tot   = enriched.filter((e) => e.escalation).length || 1;
                  const cfg   = ESC[lvl];
                  return (
                    <div key={lvl} style={{ marginBottom:13 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6, alignItems:"center" }}>
                        <span className="pill" style={{ background:cfg.pill, color:cfg.text }}><span className="dot" style={{ background:cfg.dot }} />{lvl.charAt(0).toUpperCase()+lvl.slice(1)}</span>
                        <span style={{ color:T.text2, fontSize:13, fontWeight:600 }}>{count}</span>
                      </div>
                      <div className="pt"><div className="pf" style={{ width:`${Math.round(count/tot*100)}%`, background:cfg.dot }} /></div>
                    </div>
                  );
                })}
              </div>

              {/* Department pass rates */}
              <div className="card cp">
                <div className="sh"><div><div className="sh-t">Department Pass Rates</div><div className="sh-s">Passed ÷ sat exam per dept</div></div></div>
                {allDepts.filter((d) => enriched.some((e) => e.department===d)).map((dept) => {
                  const de   = enriched.filter((e) => e.department===dept);
                  const sat  = de.filter((e) => e.examScore!==null).length;
                  const pass = de.filter((e) => e.status==="passed").length;
                  const pct  = sat > 0 ? Math.round(pass/sat*100) : 0;
                  const c    = gc(dept);
                  return (
                    <div key={dept} style={{ marginBottom:12 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:5, alignItems:"center" }}>
                        <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                          <span className="dot" style={{ background:c }} />
                          <span style={{ fontSize:13, color:T.text2, fontWeight:500 }}>{dept}</span>
                        </div>
                        <span style={{ fontSize:11, color:T.text4 }}>{pass}/{sat} sat · {pct}%</span>
                      </div>
                      <div className="pt"><div className="pf" style={{ width:`${pct}%`, background:c }} /></div>
                    </div>
                  );
                })}
                {allDepts.filter((d) => enriched.some((e) => e.department===d)).length===0 && <div style={{ color:T.text4, fontSize:13 }}>No data yet</div>}
              </div>
            </div>
          </>)}

          {/* ════ EMPLOYEES ════ */}
          {tab === "employees" && (<>
            <div className="fr">
              <div className="si">
                <span className="si-ic">🔍</span>
                <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search name or department…" />
              </div>
              <select className="fsel" value={filterDept}   onChange={(e) => setFilterDept(e.target.value)}>
                <option value="All">All Departments</option>
                {allDepts.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
              <select className="fsel" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                <option value="All">All Statuses</option>
                {["passed","failed","scheduled","unscheduled"].map((s) => <option key={s} value={s}>{s.charAt(0).toUpperCase()+s.slice(1)}</option>)}
              </select>
              <select className="fsel" value={filterEsc}    onChange={(e) => setFilterEsc(e.target.value)}>
                <option value="All">All Escalations</option>
                {["critical","high","medium","low"].map((s) => <option key={s} value={s}>{s.charAt(0).toUpperCase()+s.slice(1)}</option>)}
              </select>
              <div className="rc">{filtered.length}/{enriched.length}</div>
            </div>

            <div className="tw">
              <div className="th" style={{ gridTemplateColumns:TCOL }}>
                {["Employee","Department","Exam Date","Score","Attempts","Escalation","Actions"].map((h) => (
                  <div key={h} className="thc">{h}</div>
                ))}
              </div>
              {filtered.length===0 && (
                <div className="empty"><div className="ei">🔍</div><div style={{ color:T.text4, fontSize:13 }}>No employees match your filters</div></div>
              )}
              {filtered.map((e, i) => {
                const esc = e.escalation;
                const cfg = esc ? ESC[esc.level] : null;
                const c   = gc(e.department);
                return (
                  <div key={e.id} className={`tr ${selectedRow===e.id?"sel":""}`}
                    style={{ gridTemplateColumns:TCOL, background:i%2!==0?T.row2:"transparent" }}
                    onClick={() => setSelectedRow(selectedRow===e.id?null:e.id)}>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <div className="av" style={{ background:c+"22", color:c }}>{e.name[0]??""}</div>
                      <div>
                        <div className="en">{e.name}</div>
                        <div className="em">Due: {e.deadlineValid ? formatDate(toISO(e.deadline)) : "—"}</div>
                      </div>
                    </div>
                    <div><span className="pill" style={{ background:c+"18", color:c }}><span className="dot" style={{ background:c }} />{e.department||"—"}</span></div>
                    <div style={{ fontSize:13, color:e.examDate?T.text2:T.text4 }}>{formatDate(e.examDate)}</div>
                    <div>
                      {e.examScore!==null
                        ? <span style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, fontSize:16, color:e.examScore>=PASS_MARK?"#22c55e":"#ef4444" }}>{e.examScore}%</span>
                        : <span style={{ color:T.text4 }}>—</span>}
                    </div>
                    <div style={{ color:T.text3, fontSize:13 }}>
                      {e.attempts > 0
                        ? <span style={{ fontWeight:600, color: e.attempts>=3?"#ef4444":e.attempts===2?"#f97316":T.text3 }}>{e.attempts}</span>
                        : <span style={{ color:T.text4 }}>—</span>}
                    </div>
                    <div>
                      {esc&&cfg
                        ? <><span className="pill" style={{ background:cfg.pill, color:cfg.text }}><span className="dot" style={{ background:cfg.dot }} />{esc.label}</span><div style={{ fontSize:10, color:T.text4, marginTop:3 }}>{esc.detail}</div></>
                        : <span style={{ color:"#22c55e", fontSize:12, fontWeight:600 }}>✓ Passed</span>}
                    </div>
                    <div className="ac" onClick={(ev) => ev.stopPropagation()}>
                      <button className="btn bg bsm" onClick={() => { setReschedModal(e); setReschedDate(e.examDate??""); }}>{e.examDate?"↻ Resched":"+ Schedule"}</button>
                      {e.examDate && <button className="btn ba bsm" onClick={() => { setScoreModal(e); setScoreInput(e.examScore!==null?String(e.examScore):""); }}>Score</button>}
                    </div>
                  </div>
                );
              })}
            </div>
          </>)}

          {/* ════ CALENDAR ════ */}
          {tab === "calendar" && (<>
            <div className="sh" style={{ marginBottom:16 }}>
              <div><div className="sh-t">Exam Schedule by Month</div><div className="sh-s">Last 6 months with results</div></div>
            </div>
            {byMonth.length===0
              ? <div className="empty"><div className="ei">📅</div><div style={{ color:T.text4 }}>No exams scheduled yet</div></div>
              : <div className="g3">
                  {byMonth.map(([month, data]) => {
                    let label = month;
                    try { label = new Date(month+"-02").toLocaleDateString("en-GB",{month:"long",year:"numeric"}); } catch{/**/}
                    const mEmps  = enriched.filter((e) => e.examDate?.slice(0,7)===month);
                    const pending = Math.max(0, data.scheduled-data.passed-data.failed);
                    return (
                      <div key={month} className="mc">
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:13 }}>
                          <div style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, fontSize:14, color:T.text }}>{label}</div>
                          <div style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, fontSize:24, color:"#6366f1" }}>{data.scheduled}</div>
                        </div>
                        <div style={{ display:"flex", gap:7, marginBottom:13 }}>
                          {[{lbl:"Passed",val:data.passed,c:"#22c55e"},{lbl:"Failed",val:data.failed,c:"#ef4444"},{lbl:"Pending",val:pending,c:"#6366f1"}].map((s) => (
                            <div key={s.lbl} style={{ flex:1, textAlign:"center", background:s.c+"12", borderRadius:8, padding:"8px 0" }}>
                              <div style={{ fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, fontSize:17, color:s.c }}>{s.val}</div>
                              <div style={{ fontSize:9, color:T.text4, fontWeight:600, textTransform:"uppercase" }}>{s.lbl}</div>
                            </div>
                          ))}
                        </div>
                        <div style={{ borderTop:`1px solid ${T.border2}`, paddingTop:10 }}>
                          {mEmps.slice(0,3).map((e) => {
                            const ecfg = e.escalation ? ESC[e.escalation.level] : null;
                            return (
                              <div key={e.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
                                <span style={{ fontSize:12, color:T.text2, fontWeight:500 }}>{e.name}</span>
                                <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                                  {e.status==="passed" && <span style={{ color:"#22c55e", fontSize:11 }}>✓</span>}
                                  {e.status==="failed" && <span style={{ color:"#ef4444", fontSize:11 }}>✗</span>}
                                  {ecfg&&e.status!=="passed" && <span className="pill" style={{ background:ecfg.pill, color:ecfg.text, fontSize:10, padding:"2px 6px" }}>{e.escalation!.label}</span>}
                                </div>
                              </div>
                            );
                          })}
                          {mEmps.length>3 && <div style={{ fontSize:11, color:T.text4 }}>+{mEmps.length-3} more</div>}
                        </div>
                      </div>
                    );
                  })}
                </div>}

            <div className="card cp">
              <div className="sh"><div><div className="sh-t">Month-6 Deadlines</div><div className="sh-s">Next 90 days — non-passed</div></div></div>
              {upcomingDL.length===0
                ? <div style={{ color:T.text4, fontSize:13 }}>No deadlines in the next 90 days</div>
                : upcomingDL.map((e) => {
                    const days = daysFromNow(e.deadline);
                    const uc   = days<14?"#ef4444":days<30?"#f97316":"#eab308";
                    const pct  = Math.max(4, Math.min(100, 100-(days/90*100)));
                    return (
                      <div key={e.id} style={{ marginBottom:15 }}>
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6, flexWrap:"wrap", gap:4 }}>
                          <div>
                            <span style={{ fontSize:13, fontWeight:500, color:T.text }}>{e.name}</span>
                            <span style={{ fontSize:11, color:T.text4, marginLeft:6 }}>· {e.department}</span>
                          </div>
                          <div style={{ textAlign:"right" }}>
                            <div style={{ fontSize:12, fontWeight:600, color:uc }}>{days}d left</div>
                            <div style={{ fontSize:11, color:T.text4 }}>Due {formatDate(toISO(e.deadline))}</div>
                          </div>
                        </div>
                        <div className="pt"><div className="pf" style={{ width:`${pct}%`, background:uc }} /></div>
                      </div>
                    );
                  })}
            </div>
          </>)}

          {/* ════ DEPARTMENTS ════ */}
          {tab === "departments" && (<>
            <div className="sh" style={{ marginBottom:16 }}>
              <div><div className="sh-t">All Departments</div><div className="sh-s">{allDepts.length} departments</div></div>
              <button className="btn bp" onClick={() => setAddDeptModal(true)}>+ New Department</button>
            </div>
            <div className="dg">
              {allDepts.map((dept) => {
                const cnt  = employees.filter((e) => e.department===dept).length;
                const sat  = employees.filter((e) => e.department===dept && e.examScore!==null).length;
                const pass = employees.filter((e) => e.department===dept && e.status==="passed").length;
                const c    = gc(dept);
                const pct  = sat > 0 ? Math.round(pass/sat*100) : 0;
                return (
                  <div key={dept} className="dc">
                    <div className="dci" style={{ background:c+"22", color:c }}>{dept[0]??""}</div>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div className="dcn">{dept}</div>
                      <div className="dcs">{cnt} enrolled · {pass}/{sat} passed</div>
                      <div style={{ marginTop:7 }}><div className="pt" style={{ height:4 }}><div className="pf" style={{ width:`${pct}%`, background:c }} /></div></div>
                    </div>
                  </div>
                );
              })}
              <div className="dc" style={{ cursor:"pointer", border:`1px dashed ${T.border}`, background:"transparent" }} onClick={() => setAddDeptModal(true)}>
                <div className="dci" style={{ background:"rgba(99,102,241,.1)", color:"#6366f1", fontSize:20 }}>+</div>
                <div><div className="dcn" style={{ color:"#6366f1" }}>New Department</div><div className="dcs">Click to create</div></div>
              </div>
            </div>
          </>)}
        </div>
      </main>

      {/* ── Reschedule Modal ── */}
      {reschedModal && (
        <div className="ov" onClick={() => setReschedModal(null)}>
          <div className="md" onClick={(e) => e.stopPropagation()}>
            <div className="mh">
              <div className="mt">{reschedModal.examDate?"Reschedule":"Schedule"} Exam</div>
              <div className="ms">{reschedModal.name} · {reschedModal.department}</div>
            </div>
            <div className="mb">
              <div style={{ background:T.infoBg, borderRadius:9, padding:11, marginBottom:13 }}>
                {[
                  { lbl:"Start Date",       val:formatDate(reschedModal.startDate),                              c:T.text2 },
                  { lbl:"Month-6 Deadline", val:formatDate(toISO(getDeadline(reschedModal.startDate).date)),    c:"#f97316" },
                  ...(reschedModal.examDate ? [{ lbl:"Current Date", val:formatDate(reschedModal.examDate), c:T.text2 }] : []),
                ].map((r) => <div key={r.lbl} className="ir"><span style={{ color:T.text4 }}>{r.lbl}</span><span style={{ fontWeight:600, color:r.c }}>{r.val}</span></div>)}
              </div>
              <div className="fg">
                <label className="fl">New Exam Date</label>
                <input type="date" className="fi" value={reschedDate} onChange={(e) => setReschedDate(e.target.value)}
                  max={toISO(getDeadline(reschedModal.startDate).date) ?? undefined} />
              </div>
            </div>
            <div className="mf">
              <button className="btn bg" style={{ flex:1 }} onClick={() => { setReschedModal(null); setReschedDate(""); }}>Cancel</button>
              <button className="btn bp" style={{ flex:2 }} onClick={confirmResched} disabled={saving||!reschedDate}>{saving?"Saving…":"Confirm"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Score Modal ── */}
      {scoreModal && (
        <div className="ov" onClick={() => setScoreModal(null)}>
          <div className="md" onClick={(e) => e.stopPropagation()}>
            <div className="mh">
              <div className="mt">Record Exam Score</div>
              <div className="ms">{scoreModal.name} · Attempt #{(scoreModal.attempts||0)+(scoreModal.examScore===null?1:0)}</div>
            </div>
            <div className="mb" style={{ textAlign:"center", padding:"18px 20px" }}>
              <input type="number" min="0" max="100" value={scoreInput} onChange={(e) => setScoreInput(e.target.value)} placeholder="0–100"
                style={{ width:130, background:T.input, border:`2px solid ${T.inputBd}`, color:T.text, borderRadius:13, padding:15, fontSize:32, textAlign:"center", fontFamily:"'Space Grotesk',sans-serif", fontWeight:700, outline:"none" }} />
              {scoreInput!==""&&!isNaN(parseInt(scoreInput)) && (
                <div style={{ marginTop:12, fontSize:14, fontWeight:600 }}>
                  {parseInt(scoreInput)>=PASS_MARK
                    ? <span style={{ color:"#22c55e" }}>✅ PASS — above {PASS_MARK}%</span>
                    : <span style={{ color:"#ef4444" }}>❌ FAIL — below {PASS_MARK}%</span>}
                </div>
              )}
              <div style={{ marginTop:7, fontSize:12, color:T.text4 }}>Pass mark: {PASS_MARK}%</div>
            </div>
            <div className="mf">
              <button className="btn bg" style={{ flex:1 }} onClick={() => { setScoreModal(null); setScoreInput(""); }}>Cancel</button>
              <button className="btn bp" style={{ flex:2 }} onClick={confirmScore} disabled={saving||!scoreInput}>{saving?"Saving…":"Record Score"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Add Employee Modal ── */}
      {addEmpModal && (
        <div className="ov" onClick={() => setAddEmpModal(false)}>
          <div className="md" onClick={(e) => e.stopPropagation()}>
            <div className="mh"><div className="mt">Add New Employee</div><div className="ms">Added as unscheduled</div></div>
            <div className="mb">
              <div className="fg"><label className="fl">Full Name</label><input type="text" className="fi" value={newEmp.name} onChange={(e) => setNewEmp((p) => ({...p,name:e.target.value}))} placeholder="e.g. Jane Smith" /></div>
              <div className="fg"><label className="fl">Department</label>
                <select className="fi" value={newEmp.department} onChange={(e) => setNewEmp((p) => ({...p,department:e.target.value}))}>
                  {allDepts.map((d) => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>
              <div className="fg"><label className="fl">Start Date</label><input type="date" className="fi" value={newEmp.startDate} onChange={(e) => setNewEmp((p) => ({...p,startDate:e.target.value}))} /></div>
              {newEmp.startDate && (
                <div style={{ background:"rgba(99,102,241,.08)", border:"1px solid rgba(99,102,241,.15)", borderRadius:8, padding:"8px 12px", fontSize:13, color:"#a5b4fc" }}>
                  📅 Month-6 deadline: <strong>{formatDate(toISO(getDeadline(newEmp.startDate).date))}</strong>
                </div>
              )}
            </div>
            <div className="mf">
              <button className="btn bg" style={{ flex:1 }} onClick={() => setAddEmpModal(false)}>Cancel</button>
              <button className="btn bp" style={{ flex:2 }} onClick={addEmployeeFn} disabled={saving||!newEmp.name.trim()||!newEmp.startDate}>{saving?"Adding…":"Add Employee"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Add Department Modal ── */}
      {addDeptModal && (
        <div className="ov" onClick={() => setAddDeptModal(false)}>
          <div className="md" onClick={(e) => e.stopPropagation()}>
            <div className="mh"><div className="mt">New Department</div><div className="ms">Choose a name and colour</div></div>
            <div className="mb">
              <div className="fg"><label className="fl">Department Name</label><input type="text" className="fi" value={newDept.name} onChange={(e) => setNewDept((p) => ({...p,name:e.target.value}))} placeholder="e.g. Customer Success" /></div>
              <div className="fg">
                <label className="fl">Colour</label>
                <div style={{ display:"flex", flexWrap:"wrap", gap:7, marginTop:3 }}>
                  {PALETTE.map((c) => <div key={c} className={`cs ${newDept.color===c?"sel":""}`} style={{ background:c }} onClick={() => setNewDept((p) => ({...p,color:c}))} />)}
                </div>
              </div>
              {newDept.name.trim() && (
                <div className="fg">
                  <label className="fl">Preview</label>
                  <span className="pill" style={{ background:newDept.color+"22", color:newDept.color, fontSize:13, padding:"5px 12px" }}>
                    <span className="dot" style={{ background:newDept.color }} />{newDept.name}
                  </span>
                </div>
              )}
            </div>
            <div className="mf">
              <button className="btn bg" style={{ flex:1 }} onClick={() => setAddDeptModal(false)}>Cancel</button>
              <button className="btn bp" style={{ flex:2 }} onClick={addDeptFn} disabled={!newDept.name.trim()}>Create Department</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

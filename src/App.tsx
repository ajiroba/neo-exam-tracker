import { useState, useMemo, useEffect, useCallback } from "react";

// ============================================================
// PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL BELOW
// ============================================================
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxtpdCLNcaR37MzpPmGCbmRWZgwmqwgZ3KptcHO9D068C47RvT0ID2DFPrHHRnjqlU6/exec";
// ============================================================

const PASS_MARK = 70;

const DEPT_COLORS: Record<string, string> = {
  Finance: "#f59e0b", HR: "#10b981", Engineering: "#6366f1",
  Sales: "#f43f5e", Legal: "#8b5cf6", Marketing: "#06b6d4", Operations: "#84cc16",
};

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
  if (deadline < today) return { level: "critical", label: "OVERDUE", detail: `${Math.abs(daysToDeadline)}d past deadline` };
  if (emp.status === "failed" && emp.attempts >= 3) return { level: "critical", label: "MAX ATTEMPTS", detail: "Escalate to manager" };
  if (emp.status === "failed") return { level: "high", label: "FAILED", detail: `Resit required · ${daysToDeadline}d left` };
  if (emp.status === "unscheduled") return { level: daysToDeadline < 60 ? "high" : "medium", label: "NOT SCHEDULED", detail: `${daysToDeadline}d to deadline` };
  if (daysToDeadline < 30) return { level: "high", label: "URGENT", detail: `${daysToDeadline}d to deadline` };
  if (daysToDeadline < 60) return { level: "medium", label: "WATCH", detail: `${daysToDeadline}d to deadline` };
  return { level: "low", label: "ON TRACK", detail: `${daysToDeadline}d to deadline` };
}

const escalationStyle: Record<EscalationLevel, { bg: string; border: string; text: string; badge: string }> = {
  critical: { bg: "#2d0a0a", border: "#ef4444", text: "#f87171", badge: "#ef4444" },
  high:     { bg: "#2d1a00", border: "#f97316", text: "#fb923c", badge: "#f97316" },
  medium:   { bg: "#1e1a00", border: "#eab308", text: "#facc15", badge: "#eab308" },
  low:      { bg: "#0a1f0a", border: "#22c55e", text: "#4ade80", badge: "#22c55e" },
};

function formatDate(d: string | null): string {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

export default function NEOTracker() {
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [saving, setSaving] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [lastSaved, setLastSaved] = useState<Date | null>(null);

  const [search, setSearch] = useState("");
  const [filterDept, setFilterDept] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterEscalation, setFilterEscalation] = useState("All");
  const [rescheduleModal, setRescheduleModal] = useState<EnrichedEmployee | null>(null);
  const [rescheduleDate, setRescheduleDate] = useState("");
  const [selectedRow, setSelectedRow] = useState<number | null>(null);
  const [tab, setTab] = useState("dashboard");
  const [scoreModal, setScoreModal] = useState<EnrichedEmployee | null>(null);
  const [scoreInput, setScoreInput] = useState("");
  const [addModal, setAddModal] = useState(false);
  const [newEmp, setNewEmp] = useState<{ name: string; department: string; startDate: string }>({ name: "", department: "Finance", startDate: "" });

  const today = new Date();

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
        setError("Failed to load data: " + (json.error ?? "Unknown error"));
      }
    } catch (err) {
      setError("Could not connect to Google Sheets. Check your SCRIPT_URL.");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  const saveEmployee = async (updatedEmp: Employee) => {
    setSaving(true);
    try {
      await fetch(SCRIPT_URL, {
        method: "POST",
        body: JSON.stringify({ action: "update", employee: updatedEmp }),
      });
      setEmployees((prev) => prev.map((e) => e.id === updatedEmp.id ? updatedEmp : e));
      setLastSaved(new Date());
    } catch {
      setError("Save failed. Changes may not have persisted.");
    } finally {
      setSaving(false);
    }
  };

  const addEmployee = async () => {
    if (!newEmp.name || !newEmp.startDate) return;
    const newId = Math.max(...employees.map((e) => e.id), 0) + 1;
    const emp: Employee = { id: newId, ...newEmp, examDate: null, examScore: null, attempts: 0, status: "unscheduled" };
    setSaving(true);
    try {
      await fetch(SCRIPT_URL, {
        method: "POST",
        body: JSON.stringify({ action: "add", employee: emp }),
      });
      setEmployees((prev) => [...prev, emp]);
      setLastSaved(new Date());
      setAddModal(false);
      setNewEmp({ name: "", department: "Finance", startDate: "" });
    } catch {
      setError("Failed to add employee.");
    } finally {
      setSaving(false);
    }
  };

  const enriched: EnrichedEmployee[] = useMemo(() =>
    employees.map((e) => ({ ...e, escalation: getEscalation(e), deadline: getDeadline(e.startDate) })),
    [employees]
  );

  const stats = useMemo(() => {
    const total = enriched.length;
    const passed = enriched.filter((e) => e.status === "passed").length;
    const failed = enriched.filter((e) => e.status === "failed").length;
    const critical = enriched.filter((e) => e.escalation?.level === "critical").length;
    const overdue = enriched.filter((e) => e.escalation?.label === "OVERDUE").length;
    return { total, passed, failed, critical, overdue };
  }, [enriched]);

  const departments = ["All", ...Array.from(new Set(employees.map((e) => e.department)))];

  const filtered: EnrichedEmployee[] = useMemo(() => enriched.filter((e) => {
    const q = search.toLowerCase();
    const matchSearch = !q || e.name.toLowerCase().includes(q) || e.department.toLowerCase().includes(q);
    const matchDept = filterDept === "All" || e.department === filterDept;
    const matchStatus = filterStatus === "All" || e.status === filterStatus;
    const matchEsc = filterEscalation === "All" || e.escalation?.level === filterEscalation.toLowerCase();
    return matchSearch && matchDept && matchStatus && matchEsc;
  }), [enriched, search, filterDept, filterStatus, filterEscalation]);

  function handleReschedule(emp: EnrichedEmployee) { setRescheduleModal(emp); setRescheduleDate(emp.examDate || ""); }

  async function confirmReschedule() {
    if (!rescheduleDate || !rescheduleModal) return;
    const updated: Employee = {
      ...rescheduleModal,
      examDate: rescheduleDate,
      status: rescheduleModal.status === "unscheduled" ? "scheduled" : rescheduleModal.status,
    };
    await saveEmployee(updated);
    setRescheduleModal(null);
    setRescheduleDate("");
  }

  function handleScoreEntry(emp: EnrichedEmployee) { setScoreModal(emp); setScoreInput(emp.examScore !== null ? String(emp.examScore) : ""); }

  async function confirmScore() {
    if (!scoreModal) return;
    const score = parseInt(scoreInput, 10);
    if (isNaN(score) || score < 0 || score > 100) return;
    const updated: Employee = {
      ...scoreModal,
      examScore: score,
      status: score >= PASS_MARK ? "passed" : "failed",
      attempts: scoreModal.examScore === null ? (scoreModal.attempts || 0) + 1 : scoreModal.attempts,
    };
    await saveEmployee(updated);
    setScoreModal(null);
    setScoreInput("");
  }

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

  if (loading) return (
    <div style={{ minHeight: "100vh", background: "#0d0f14", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Mono', monospace", color: "#6366f1" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Syne:wght@800&display=swap'); @keyframes load { 0%{transform:translateX(-100%)} 100%{transform:translateX(300%)} }`}</style>
      <div style={{ fontFamily: "'Syne', sans-serif", fontSize: 22, fontWeight: 800, marginBottom: 12 }}>NEO EXAM TRACKER</div>
      <div style={{ color: "#475569", fontSize: 12, letterSpacing: "0.1em" }}>CONNECTING TO GOOGLE SHEETS...</div>
      <div style={{ marginTop: 24, width: 200, height: 3, background: "#1e2230", borderRadius: 2, overflow: "hidden" }}>
        <div style={{ height: "100%", background: "#6366f1", borderRadius: 2, animation: "load 1.5s ease infinite", width: "60%" }} />
      </div>
    </div>
  );

  if (error && employees.length === 0) return (
    <div style={{ minHeight: "100vh", background: "#0d0f14", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Mono', monospace", color: "#f87171", padding: 32 }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Syne:wght@800&display=swap');`}</style>
      <div style={{ fontFamily: "'Syne', sans-serif", fontSize: 22, fontWeight: 800, marginBottom: 16, color: "#e2e8f0" }}>Connection Error</div>
      <div style={{ background: "#2d0a0a", border: "1px solid #ef4444", borderRadius: 10, padding: 20, maxWidth: 480, fontSize: 12, lineHeight: 1.8, color: "#f87171" }}>{error}</div>
      <div style={{ marginTop: 16, fontSize: 11, color: "#475569", maxWidth: 480, textAlign: "center", lineHeight: 1.8 }}>
        Make sure you have replaced <span style={{ color: "#a5b4fc" }}>SCRIPT_URL</span> with your deployed Google Apps Script Web App URL.
      </div>
      <button onClick={fetchData} style={{ marginTop: 20, background: "#6366f1", border: "none", color: "#fff", borderRadius: 8, padding: "10px 24px", cursor: "pointer", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>Retry</button>
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: "#0d0f14", color: "#e2e8f0", fontFamily: "'IBM Plex Mono', monospace", fontSize: "13px" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Syne:wght@700;800&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 4px; height: 4px; }
        ::-webkit-scrollbar-track { background: #1a1d24; }
        ::-webkit-scrollbar-thumb { background: #3a3f4d; border-radius: 2px; }
        .row-hover:hover { background: rgba(255,255,255,0.03) !important; cursor: pointer; }
        .action-btn { transition: all 0.15s ease; cursor: pointer; border: none; outline: none; }
        .action-btn:hover { filter: brightness(1.2); transform: translateY(-1px); }
        input, select { outline: none; }
        .stat-card { transition: transform 0.2s; }
        .stat-card:hover { transform: translateY(-2px); }
        .tab-btn { transition: all 0.15s; }
        @keyframes pulse-red { 0%,100%{opacity:1} 50%{opacity:0.5} }
        .pulse { animation: pulse-red 2s infinite; }
        @keyframes fadeIn { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:translateY(0)} }
        .fadein { animation: fadeIn 0.3s ease forwards; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .spin { animation: spin 1s linear infinite; display: inline-block; }
      `}</style>

      {/* Header */}
      <div style={{ borderBottom: "1px solid #1e2230", padding: "0 32px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 64, position: "sticky", top: 0, background: "#0d0f14", zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ width: 36, height: 36, borderRadius: 8, background: "linear-gradient(135deg,#6366f1,#a855f7)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 16, color: "#fff" }}>N</div>
          <div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 17, color: "#f1f5f9" }}>NEO EXAM TRACKER</div>
            <div style={{ color: "#64748b", fontSize: 11, letterSpacing: "0.08em" }}>NEW EMPLOYEE ORIENTATION · PASS MARK {PASS_MARK}%</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          {saving && <div style={{ color: "#6366f1", fontSize: 11 }}><span className="spin">⟳</span> SAVING...</div>}
          {lastSaved && !saving && <div style={{ color: "#22c55e", fontSize: 11 }}>✓ SAVED {lastSaved.toLocaleTimeString()}</div>}
          {stats.critical > 0 && (
            <div className="pulse" style={{ display: "flex", alignItems: "center", gap: 8, background: "#2d0a0a", border: "1px solid #ef4444", borderRadius: 6, padding: "6px 12px", color: "#f87171", fontSize: 12 }}>
              <span style={{ width: 7, height: 7, borderRadius: "50%", background: "#ef4444", display: "inline-block" }} />
              {stats.critical} CRITICAL
            </div>
          )}
          <button onClick={() => setAddModal(true)} style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)", border: "none", color: "#fff", borderRadius: 8, padding: "8px 16px", fontSize: 12, cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontWeight: 600 }}>+ Add Employee</button>
          <button onClick={fetchData} style={{ background: "#1e2230", border: "1px solid #2d3344", color: "#94a3b8", borderRadius: 8, padding: "8px 14px", fontSize: 12, cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace" }}>↻ Refresh</button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ borderBottom: "1px solid #1e2230", padding: "0 32px", display: "flex" }}>
        {["dashboard", "employees", "calendar"].map((t) => (
          <button key={t} className="tab-btn" onClick={() => setTab(t)} style={{ background: "none", border: "none", cursor: "pointer", padding: "14px 20px", fontSize: 12, letterSpacing: "0.1em", color: tab === t ? "#a5b4fc" : "#475569", borderBottom: tab === t ? "2px solid #6366f1" : "2px solid transparent", fontFamily: "'IBM Plex Mono',monospace", fontWeight: tab === t ? 600 : 400, textTransform: "uppercase" }}>{t}</button>
        ))}
      </div>

      <div style={{ padding: "28px 32px", maxWidth: 1400, margin: "0 auto" }}>

        {/* DASHBOARD */}
        {tab === "dashboard" && (
          <div className="fadein">
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 16, marginBottom: 28 }}>
              {[
                { label: "TOTAL EMPLOYEES", value: stats.total, sub: "In NEO programme", color: "#6366f1", bg: "#13152a" },
                { label: "PASSED", value: stats.passed, sub: `${stats.total ? Math.round(stats.passed / stats.total * 100) : 0}% pass rate`, color: "#22c55e", bg: "#0a1a0a" },
                { label: "FAILED / RESIT", value: stats.failed, sub: `${stats.failed} require resit`, color: "#f97316", bg: "#1a1000" },
                { label: "CRITICAL FLAGS", value: stats.critical, sub: `${stats.overdue} overdue`, color: "#ef4444", bg: "#1a0808" },
              ].map((s) => (
                <div key={s.label} className="stat-card" style={{ background: s.bg, border: `1px solid ${s.color}22`, borderRadius: 12, padding: "20px 24px", borderLeft: `3px solid ${s.color}` }}>
                  <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 10 }}>{s.label}</div>
                  <div style={{ fontSize: 38, fontWeight: 600, color: s.color, lineHeight: 1, marginBottom: 6, fontFamily: "'Syne',sans-serif" }}>{s.value}</div>
                  <div style={{ fontSize: 11, color: "#475569" }}>{s.sub}</div>
                </div>
              ))}
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
              <div style={{ background: "#111318", border: "1px solid #1e2230", borderRadius: 12, padding: 24 }}>
                <div style={{ fontSize: 11, letterSpacing: "0.1em", color: "#64748b", marginBottom: 20 }}>ESCALATION BREAKDOWN</div>
                {(["critical","high","medium","low"] as EscalationLevel[]).map((level) => {
                  const count = enriched.filter((e) => e.escalation?.level === level).length;
                  const tot = enriched.filter((e) => e.escalation).length;
                  const pct = tot > 0 ? Math.round(count / tot * 100) : 0;
                  const s = escalationStyle[level];
                  return (
                    <div key={level} style={{ marginBottom: 14 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                        <span style={{ color: s.text, fontSize: 11, textTransform: "uppercase" }}>{level}</span>
                        <span style={{ color: "#94a3b8", fontSize: 11 }}>{count} employees</span>
                      </div>
                      <div style={{ height: 6, background: "#1e2230", borderRadius: 3, overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${pct}%`, background: s.badge, borderRadius: 3, transition: "width 0.5s" }} />
                      </div>
                    </div>
                  );
                })}
              </div>

              <div style={{ background: "#111318", border: "1px solid #1e2230", borderRadius: 12, padding: 24 }}>
                <div style={{ fontSize: 11, letterSpacing: "0.1em", color: "#64748b", marginBottom: 20 }}>BY DEPARTMENT</div>
                {Object.keys(DEPT_COLORS).map((dept) => {
                  const dEmps = enriched.filter((e) => e.department === dept);
                  if (!dEmps.length) return null;
                  const dPassed = dEmps.filter((e) => e.status === "passed").length;
                  const pct = Math.round(dPassed / dEmps.length * 100);
                  return (
                    <div key={dept} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 11 }}>
                      <div style={{ width: 8, height: 8, borderRadius: "50%", background: DEPT_COLORS[dept], flexShrink: 0 }} />
                      <div style={{ flex: 1, fontSize: 12, color: "#94a3b8" }}>{dept}</div>
                      <div style={{ width: 100, height: 5, background: "#1e2230", borderRadius: 3 }}>
                        <div style={{ width: `${pct}%`, height: "100%", background: DEPT_COLORS[dept], borderRadius: 3 }} />
                      </div>
                      <div style={{ fontSize: 11, color: "#64748b", width: 60, textAlign: "right" }}>{dPassed}/{dEmps.length}</div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div style={{ background: "#111318", border: "1px solid #ef444433", borderRadius: 12, padding: 24 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
                <div className="pulse" style={{ width: 8, height: 8, borderRadius: "50%", background: "#ef4444" }} />
                <div style={{ fontSize: 11, letterSpacing: "0.1em", color: "#f87171" }}>REQUIRES IMMEDIATE ACTION</div>
              </div>
              {enriched.filter((e) => e.escalation?.level === "critical" || e.escalation?.level === "high").length === 0 ? (
                <div style={{ color: "#22c55e", fontSize: 13, padding: "12px 0" }}>✓ No critical escalations at this time</div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr auto", gap: "10px 16px", alignItems: "center" }}>
                  {["EMPLOYEE","DEPARTMENT","FLAG","DETAIL","ACTION"].map((h) => (
                    <div key={h} style={{ fontSize: 10, letterSpacing: "0.1em", color: "#475569", paddingBottom: 8, borderBottom: "1px solid #1e2230" }}>{h}</div>
                  ))}
                  {enriched.filter((e) => e.escalation?.level === "critical" || e.escalation?.level === "high").map((e) => {
                    const esc = e.escalation!;
                    const s = escalationStyle[esc.level];
                    return [
                      <div key={`${e.id}n`} style={{ color: "#e2e8f0", padding: "8px 0" }}>{e.name}</div>,
                      <div key={`${e.id}d`} style={{ color: "#64748b" }}>{e.department}</div>,
                      <div key={`${e.id}l`}><span style={{ background: s.bg, border: `1px solid ${s.border}`, color: s.text, borderRadius: 4, padding: "2px 8px", fontSize: 10 }}>{esc.label}</span></div>,
                      <div key={`${e.id}dt`} style={{ color: s.text, fontSize: 12 }}>{esc.detail}</div>,
                      <button key={`${e.id}a`} className="action-btn" onClick={() => handleReschedule(e)} style={{ background: "#1e2230", border: "1px solid #2d3344", color: "#94a3b8", borderRadius: 6, padding: "5px 12px", fontSize: 11, cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace" }}>Reschedule</button>,
                    ];
                  })}
                </div>
              )}
            </div>
          </div>
        )}

        {/* EMPLOYEES */}
        {tab === "employees" && (
          <div className="fadein">
            <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap", alignItems: "center" }}>
              <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search employee or dept..." style={{ background: "#111318", border: "1px solid #1e2230", color: "#e2e8f0", borderRadius: 8, padding: "8px 14px", fontSize: 12, width: 220, fontFamily: "'IBM Plex Mono',monospace" }} />
              {[
                { label: "Department", state: filterDept, setter: setFilterDept, opts: departments },
                { label: "Status", state: filterStatus, setter: setFilterStatus, opts: ["All","passed","failed","scheduled","unscheduled"] },
                { label: "Escalation", state: filterEscalation, setter: setFilterEscalation, opts: ["All","critical","high","medium","low"] },
              ].map((f) => (
                <select key={f.label} value={f.state} onChange={(e) => f.setter(e.target.value)} style={{ background: "#111318", border: "1px solid #1e2230", color: "#94a3b8", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontFamily: "'IBM Plex Mono',monospace", cursor: "pointer" }}>
                  {f.opts.map((o) => <option key={o} value={o}>{o === "All" ? `All ${f.label}s` : o}</option>)}
                </select>
              ))}
              <div style={{ marginLeft: "auto", color: "#475569", fontSize: 12 }}>{filtered.length} records</div>
            </div>

            <div style={{ background: "#111318", border: "1px solid #1e2230", borderRadius: 12, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1fr auto", padding: "12px 20px", background: "#0d0f14", borderBottom: "1px solid #1e2230", gap: 8 }}>
                {["EMPLOYEE","DEPT","EXAM DATE","SCORE","ATTEMPTS","ESCALATION","ACTIONS"].map((h) => (
                  <div key={h} style={{ fontSize: 10, letterSpacing: "0.1em", color: "#475569" }}>{h}</div>
                ))}
              </div>
              {filtered.map((e, i) => {
                const esc = e.escalation;
                const escStyle = esc ? escalationStyle[esc.level] : null;
                return (
                  <div key={e.id} className="row-hover" onClick={() => setSelectedRow(selectedRow === e.id ? null : e.id)}
                    style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1fr auto", padding: "13px 20px", gap: 8, alignItems: "center", background: selectedRow === e.id ? "rgba(99,102,241,0.08)" : i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.01)", borderBottom: "1px solid #1a1d24", transition: "background 0.15s" }}>
                    <div>
                      <div style={{ color: "#e2e8f0", fontWeight: 500 }}>{e.name}</div>
                      <div style={{ color: "#475569", fontSize: 11, marginTop: 2 }}>Started {formatDate(e.startDate)} · Deadline {formatDate(e.deadline.toISOString().split("T")[0])}</div>
                    </div>
                    <div><span style={{ background: DEPT_COLORS[e.department] + "18", color: DEPT_COLORS[e.department], border: `1px solid ${DEPT_COLORS[e.department]}44`, borderRadius: 4, padding: "2px 8px", fontSize: 11 }}>{e.department}</span></div>
                    <div style={{ color: e.examDate ? "#94a3b8" : "#475569", fontSize: 12 }}>{formatDate(e.examDate)}</div>
                    <div>{e.examScore !== null ? <span style={{ color: e.examScore >= PASS_MARK ? "#4ade80" : "#f87171", fontWeight: 600, fontSize: 15 }}>{e.examScore}<span style={{ color: "#475569", fontSize: 11 }}>%</span></span> : <span style={{ color: "#2d3344" }}>—</span>}</div>
                    <div style={{ color: "#64748b" }}>{e.attempts}</div>
                    <div>
                      {esc && escStyle ? (
                        <div>
                          <span style={{ background: escStyle.bg, border: `1px solid ${escStyle.border}`, color: escStyle.text, borderRadius: 4, padding: "2px 8px", fontSize: 10, display: "inline-block" }}>{esc.label}</span>
                          {e.status !== "passed" && <div style={{ fontSize: 10, color: "#475569", marginTop: 3 }}>{esc.detail}</div>}
                        </div>
                      ) : <span style={{ color: "#22c55e", fontSize: 11 }}>✓ Passed</span>}
                    </div>
                    <div style={{ display: "flex", gap: 8 }} onClick={(ev) => ev.stopPropagation()}>
                      <button className="action-btn" onClick={() => handleReschedule(e)} style={{ background: "#1a1d24", border: "1px solid #2d3344", color: "#94a3b8", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace" }}>{e.examDate ? "↻ Resched" : "+ Schedule"}</button>
                      {e.examDate && <button className="action-btn" onClick={() => handleScoreEntry(e)} style={{ background: "#0a1428", border: "1px solid #1e3a5f", color: "#60a5fa", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace" }}>Score</button>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* CALENDAR */}
        {tab === "calendar" && (
          <div className="fadein">
            <div style={{ marginBottom: 24 }}>
              <div style={{ fontSize: 11, letterSpacing: "0.1em", color: "#64748b", marginBottom: 16 }}>EXAM SCHEDULE BY MONTH</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 16 }}>
                {byMonth.map(([month, data]) => {
                  const monthEmps = enriched.filter((e) => e.examDate?.slice(0, 7) === month);
                  const label = new Date(month + "-01").toLocaleDateString("en-GB", { month: "long", year: "numeric" });
                  return (
                    <div key={month} style={{ background: "#111318", border: "1px solid #1e2230", borderRadius: 12, padding: 20 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                        <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15, color: "#e2e8f0" }}>{label}</div>
                        <div style={{ fontSize: 22, fontWeight: 700, color: "#6366f1", fontFamily: "'Syne',sans-serif" }}>{data.scheduled}</div>
                      </div>
                      <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
                        {[{ label: "Passed", val: data.passed, color: "#22c55e" }, { label: "Failed", val: data.failed, color: "#ef4444" }, { label: "Pending", val: data.scheduled - data.passed - data.failed, color: "#6366f1" }].map((s) => (
                          <div key={s.label} style={{ flex: 1, textAlign: "center" }}>
                            <div style={{ fontSize: 18, fontWeight: 700, color: s.color, fontFamily: "'Syne',sans-serif" }}>{s.val}</div>
                            <div style={{ fontSize: 10, color: "#475569" }}>{s.label.toUpperCase()}</div>
                          </div>
                        ))}
                      </div>
                      <div style={{ borderTop: "1px solid #1e2230", paddingTop: 12 }}>
                        {monthEmps.slice(0, 4).map((e) => {
                          const esc = e.escalation;
                          const escStyle = esc ? escalationStyle[esc.level] : null;
                          return (
                            <div key={e.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 7 }}>
                              <div style={{ fontSize: 12, color: "#94a3b8" }}>{e.name}</div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                <span style={{ fontSize: 11, color: "#64748b" }}>{formatDate(e.examDate)}</span>
                                {e.status === "passed" && <span style={{ color: "#22c55e", fontSize: 11 }}>✓</span>}
                                {e.status === "failed" && <span style={{ color: "#ef4444", fontSize: 11 }}>✗</span>}
                                {esc && escStyle && e.status !== "passed" && <span style={{ background: escStyle.bg, border: `1px solid ${escStyle.border}`, color: escStyle.text, borderRadius: 3, padding: "1px 6px", fontSize: 9 }}>{esc.label}</span>}
                              </div>
                            </div>
                          );
                        })}
                        {monthEmps.length > 4 && <div style={{ color: "#475569", fontSize: 11 }}>+{monthEmps.length - 4} more</div>}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div style={{ background: "#111318", border: "1px solid #1e2230", borderRadius: 12, padding: 24 }}>
              <div style={{ fontSize: 11, letterSpacing: "0.1em", color: "#64748b", marginBottom: 16 }}>MONTH-6 DEADLINES (NEXT 90 DAYS)</div>
              {enriched
                .filter((e) => { const diff = Math.ceil((e.deadline.getTime() - today.getTime()) / (1000 * 60 * 60 * 24)); return diff >= 0 && diff <= 90 && e.status !== "passed"; })
                .sort((a, b) => a.deadline.getTime() - b.deadline.getTime())
                .map((e) => {
                  const daysLeft = Math.ceil((e.deadline.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
                  const urgentColor = daysLeft < 14 ? "#ef4444" : daysLeft < 30 ? "#f97316" : "#eab308";
                  return (
                    <div key={e.id} style={{ marginBottom: 16 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                        <span style={{ color: "#e2e8f0" }}>{e.name} <span style={{ color: "#475569" }}>· {e.department}</span></span>
                        <span style={{ color: urgentColor, fontSize: 12 }}>{daysLeft}d remaining · Deadline {formatDate(e.deadline.toISOString().split("T")[0])}</span>
                      </div>
                      <div style={{ height: 5, background: "#1e2230", borderRadius: 3 }}>
                        <div style={{ height: "100%", width: `${Math.max(0, Math.min(100, 100 - (daysLeft / 90 * 100)))}%`, background: urgentColor, borderRadius: 3 }} />
                      </div>
                    </div>
                  );
                })}
            </div>
          </div>
        )}
      </div>

      {/* Reschedule Modal */}
      {rescheduleModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }} onClick={() => setRescheduleModal(null)}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: "#151820", border: "1px solid #2d3344", borderRadius: 16, padding: 32, width: 420, fontFamily: "'IBM Plex Mono',monospace" }}>
            <div style={{ fontFamily: "'Syne',sans-serif", fontSize: 18, fontWeight: 800, color: "#e2e8f0", marginBottom: 6 }}>{rescheduleModal.examDate ? "Reschedule Exam" : "Schedule Exam"}</div>
            <div style={{ color: "#64748b", fontSize: 12, marginBottom: 24 }}>{rescheduleModal.name} · {rescheduleModal.department}</div>
            <div style={{ background: "#0d0f14", borderRadius: 8, padding: 16, marginBottom: 20, fontSize: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}><span style={{ color: "#64748b" }}>Start date</span><span style={{ color: "#94a3b8" }}>{formatDate(rescheduleModal.startDate)}</span></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: "#64748b" }}>Month-6 deadline</span><span style={{ color: "#f97316" }}>{formatDate(getDeadline(rescheduleModal.startDate).toISOString().split("T")[0])}</span></div>
            </div>
            <label style={{ fontSize: 11, color: "#64748b", letterSpacing: "0.08em", display: "block", marginBottom: 8 }}>NEW EXAM DATE</label>
            <input type="date" value={rescheduleDate} onChange={(e) => setRescheduleDate(e.target.value)} max={getDeadline(rescheduleModal.startDate).toISOString().split("T")[0]} style={{ width: "100%", background: "#0d0f14", border: "1px solid #2d3344", color: "#e2e8f0", borderRadius: 8, padding: "10px 14px", fontSize: 13, marginBottom: 20, fontFamily: "'IBM Plex Mono',monospace" }} />
            <div style={{ display: "flex", gap: 12 }}>
              <button onClick={() => setRescheduleModal(null)} style={{ flex: 1, background: "#1e2230", border: "1px solid #2d3344", color: "#64748b", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12 }}>Cancel</button>
              <button onClick={confirmReschedule} style={{ flex: 2, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", border: "none", color: "#fff", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 600 }}>{saving ? "Saving..." : "Confirm Schedule"}</button>
            </div>
          </div>
        </div>
      )}

      {/* Score Modal */}
      {scoreModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }} onClick={() => setScoreModal(null)}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: "#151820", border: "1px solid #2d3344", borderRadius: 16, padding: 32, width: 380, fontFamily: "'IBM Plex Mono',monospace" }}>
            <div style={{ fontFamily: "'Syne',sans-serif", fontSize: 18, fontWeight: 800, color: "#e2e8f0", marginBottom: 6 }}>Enter Exam Score</div>
            <div style={{ color: "#64748b", fontSize: 12, marginBottom: 24 }}>{scoreModal.name} · Attempt #{(scoreModal.attempts || 0) + 1}</div>
            <div style={{ textAlign: "center", marginBottom: 24 }}>
              <input type="number" min="0" max="100" value={scoreInput} onChange={(e) => setScoreInput(e.target.value)} placeholder="0–100" style={{ width: 120, background: "#0d0f14", border: "1px solid #2d3344", color: "#e2e8f0", borderRadius: 10, padding: "14px", fontSize: 28, textAlign: "center", fontFamily: "'Syne',sans-serif", fontWeight: 700 }} />
              <div style={{ marginTop: 12, fontSize: 12 }}>
                {scoreInput !== "" && (parseInt(scoreInput) >= PASS_MARK ? <span style={{ color: "#22c55e" }}>✓ PASS — above {PASS_MARK}% threshold</span> : <span style={{ color: "#f87171" }}>✗ FAIL — below {PASS_MARK}% threshold</span>)}
              </div>
            </div>
            <div style={{ display: "flex", gap: 12 }}>
              <button onClick={() => setScoreModal(null)} style={{ flex: 1, background: "#1e2230", border: "1px solid #2d3344", color: "#64748b", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12 }}>Cancel</button>
              <button onClick={confirmScore} style={{ flex: 2, background: "linear-gradient(135deg,#0369a1,#6366f1)", border: "none", color: "#fff", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 600 }}>{saving ? "Saving..." : "Record Score"}</button>
            </div>
          </div>
        </div>
      )}

      {/* Add Employee Modal */}
      {addModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200 }} onClick={() => setAddModal(false)}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: "#151820", border: "1px solid #2d3344", borderRadius: 16, padding: 32, width: 420, fontFamily: "'IBM Plex Mono',monospace" }}>
            <div style={{ fontFamily: "'Syne',sans-serif", fontSize: 18, fontWeight: 800, color: "#e2e8f0", marginBottom: 24 }}>Add New Employee</div>
            {([{ label: "FULL NAME", key: "name", type: "text", placeholder: "e.g. Jane Smith" }, { label: "START DATE", key: "startDate", type: "date", placeholder: "" }] as { label: string; key: keyof typeof newEmp; type: string; placeholder: string }[]).map((f) => (
              <div key={f.key} style={{ marginBottom: 16 }}>
                <label style={{ fontSize: 11, color: "#64748b", letterSpacing: "0.08em", display: "block", marginBottom: 8 }}>{f.label}</label>
                <input type={f.type} value={newEmp[f.key]} onChange={(e) => setNewEmp((prev) => ({ ...prev, [f.key]: e.target.value }))} placeholder={f.placeholder} style={{ width: "100%", background: "#0d0f14", border: "1px solid #2d3344", color: "#e2e8f0", borderRadius: 8, padding: "10px 14px", fontSize: 13, fontFamily: "'IBM Plex Mono',monospace" }} />
              </div>
            ))}
            <div style={{ marginBottom: 20 }}>
              <label style={{ fontSize: 11, color: "#64748b", letterSpacing: "0.08em", display: "block", marginBottom: 8 }}>DEPARTMENT</label>
              <select value={newEmp.department} onChange={(e) => setNewEmp((prev) => ({ ...prev, department: e.target.value }))} style={{ width: "100%", background: "#0d0f14", border: "1px solid #2d3344", color: "#e2e8f0", borderRadius: 8, padding: "10px 14px", fontSize: 13, fontFamily: "'IBM Plex Mono',monospace" }}>
                {Object.keys(DEPT_COLORS).map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", gap: 12 }}>
              <button onClick={() => setAddModal(false)} style={{ flex: 1, background: "#1e2230", border: "1px solid #2d3344", color: "#64748b", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12 }}>Cancel</button>
              <button onClick={addEmployee} style={{ flex: 2, background: "linear-gradient(135deg,#6366f1,#8b5cf6)", border: "none", color: "#fff", borderRadius: 8, padding: "10px", cursor: "pointer", fontFamily: "'IBM Plex Mono',monospace", fontSize: 12, fontWeight: 600 }}>{saving ? "Adding..." : "Add Employee"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

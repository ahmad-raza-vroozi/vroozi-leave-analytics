import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Cell, Legend,
  LineChart, Line, AreaChart, Area, ReferenceLine, PieChart, Pie, ResponsiveContainer
} from "recharts";
import {
  Upload, Flag, Search, AlertTriangle, ChevronRight, ChevronLeft, RefreshCw,
  Users, CalendarDays, Home, Stethoscope, PlaneTakeoff, ShieldAlert, CheckCircle2,
  LayoutGrid, ChevronDown, X
} from "lucide-react";

/* ---------- tokens: corporate SaaS ops-console look, matching reference ---------- */
const BG = "#F3F5F9";
const CARD = "#FFFFFF";
const BORDER = "#E5E9F0";
const TEXT = "#1B2430";
const MUTED = "#7C879A";
const NAVY = "#152A4E";
const NAVY_2 = "#1F3F72";
const BLUE = "#3D6FEA";
const GREEN = "#1D9A6C";
const AMBER = "#E08A1E";
const RED = "#DE3B3B";
const PURPLE = "#6D4CD1";
const TEAL = "#1E9E93";
const SANS = "'Inter', -apple-system, system-ui, sans-serif";

const DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const MONTH_ORDER = ["January","February","March","April","May","June","July",
  "August","September","October","November","December"];
const KNOWN_LABELS = {
  AL: "Annual Leave", H: "Holiday", S: "Sick Leave", "H.S": "Half-day Sick",
  WFH: "Work From Home", "H.WFH": "Half-day WFH", BL: "Bereavement Leave"
};
const TYPE_COLORS = { AL: BLUE, S: AMBER, WFH: TEAL, H: "#8A93A6", "H.S": "#F0B65C", "H.WFH": "#7FCFC7", BL: RED };

function pctile(sorted, p) {
  if (!sorted.length) return 0;
  const idx = (sorted.length - 1) * p;
  const lo = Math.floor(idx), hi = Math.ceil(idx);
  return lo === hi ? sorted[lo] : sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}
function daysInMonth(year, month1) { return new Date(year, month1, 0).getDate(); }
function weekdayName(year, month0, day) { return new Date(year, month0, day).toLocaleDateString("en-US", { weekday: "long" }); }

/* ---------- workbook parser (raw multi-tab xlsx -> long records + allocations) ---------- */
function parseWorkbook(workbook, year) {
  const sheetNames = workbook.SheetNames;
  const annualSheetName = sheetNames.find(n => n.toLowerCase().includes("annual"));
  if (!annualSheetName) throw new Error('Could not find an "Annual Summary" tab in this workbook.');

  const annualAoa = XLSX.utils.sheet_to_json(workbook.Sheets[annualSheetName], { header: 1, raw: true, defval: null });
  const topRow = annualAoa[0] || [];
  const subRow = annualAoa[1] || [];
  let lastTop = null;
  const filledTop = topRow.map(v => { if (v != null && String(v).trim() !== "") lastTop = String(v).trim().toLowerCase(); return lastTop; });
  const colKeys = subRow.map((sub, i) => {
    const top = filledTop[i] || "";
    const s = sub != null ? String(sub).trim() : "";
    if (top.startsWith("allocated")) return `Allocated_${s}`;
    if (top.startsWith("remaining")) return `Remaining_${s}`;
    if (s === "Name" || s === "Status") return s;
    return null;
  });

  const employees = [];
  const annualByName = {};
  for (let r = 2; r < annualAoa.length; r++) {
    const row = annualAoa[r] || [];
    const name = row[0] != null ? String(row[0]).trim() : "";
    if (!name) { employees.push(null); continue; }
    employees.push(name);
    const rec = {};
    colKeys.forEach((k, i) => { if (k) rec[k] = row[i]; });
    annualByName[name] = rec;
  }

  const monthSheets = MONTH_ORDER.map(m => sheetNames.find(n => n.trim().toLowerCase() === m.toLowerCase())).filter(Boolean);
  const longRows = [];
  let skippedJunk = 0;
  const unrecognizedCodes = {};

  monthSheets.forEach(sheetName => {
    const monthIndex = MONTH_ORDER.findIndex(m => m.toLowerCase() === sheetName.trim().toLowerCase());
    const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
    const headerRow = aoa[0] || [];
    const dayCols = {};
    headerRow.forEach((v, c) => { if (typeof v === "number" && Number.isInteger(v)) dayCols[c] = v; });
    const maxDay = daysInMonth(year, monthIndex + 1);

    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      if (!emp) continue;
      const row = aoa[2 + i] || [];
      for (const [cStr, day] of Object.entries(dayCols)) {
        if (day > maxDay) continue;
        const c = parseInt(cStr, 10);
        const raw = row[c];
        if (raw == null) continue;
        const cleaned = String(raw).trim();
        if (cleaned === "") continue;
        const lettersOnly = cleaned.replace(/[^A-Za-z.]/g, "");
        if (lettersOnly.length === 0) { skippedJunk++; continue; }
        const code = cleaned.toUpperCase();
        if (!KNOWN_LABELS[code]) unrecognizedCodes[code] = (unrecognizedCodes[code] || 0) + 1;
        longRows.push({
          employee: emp,
          date: `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`,
          leave_type: code,
          day_of_week: weekdayName(year, monthIndex, day),
          month: MONTH_ORDER[monthIndex],
        });
      }
    }
  });

  return { longRows, annualByName, employeeOrder: employees.filter(Boolean), skippedJunk, unrecognizedCodes };
}

export default function HRLeaveDashboard() {
  const [parsed, setParsed] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all"); // all | flagged
  const [sortMode, setSortMode] = useState("flags"); // flags | name
  const [selectedEmployee, setSelectedEmployee] = useState(null);
  const [activeTab, setActiveTab] = useState("dashboard"); // dashboard | employees | policies
  const [directorySearch, setDirectorySearch] = useState("");
  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);

  const handleUpload = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setError("");
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const result = parseWorkbook(workbook, 2026);
        if (result.longRows.length === 0) {
          setError("Parsed the file, but found zero leave/WFH records. Check the monthly tabs are named January–December.");
        } else {
          setParsed(result);
          setPage(1);
        }
      } catch (err) {
        setError(err.message || "Could not read this file. Confirm it's the original multi-tab Leave Sheet .xlsx export.");
      }
      setLoading(false);
    };
    reader.onerror = () => { setError("Failed to read the file."); setLoading(false); };
    reader.readAsArrayBuffer(file);
  }, []);

  const analysis = useMemo(() => {
    if (!parsed) return null;
    const { longRows, annualByName, employeeOrder } = parsed;
    const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : null; };

    const byEmp = {};
    for (const r of longRows) (byEmp[r.employee] ||= []).push(r);

    // Seed from every named employee (Annual Summary tab), not just ones with
    // at least one leave record — otherwise someone with 0 days used simply
    // vanishes from the dashboard instead of showing up with zero counts.
    const allNames = [...new Set([...employeeOrder, ...Object.keys(byEmp)])];

    const employeeStats = allNames.map((name) => {
      const records = byEmp[name] || [];
      const typeCounts = {}, dowCounts = {}, monthCounts = {}, monthTypeCounts = {}, dowTypeCounts = {}, weekOfMonthTypeCounts = {};
      for (const r of records) {
        typeCounts[r.leave_type] = (typeCounts[r.leave_type] || 0) + 1;
        dowCounts[r.day_of_week] = (dowCounts[r.day_of_week] || 0) + 1;
        monthCounts[r.month] = (monthCounts[r.month] || 0) + 1;
        monthTypeCounts[r.month] = monthTypeCounts[r.month] || {};
        monthTypeCounts[r.month][r.leave_type] = (monthTypeCounts[r.month][r.leave_type] || 0) + 1;
        dowTypeCounts[r.day_of_week] = dowTypeCounts[r.day_of_week] || {};
        dowTypeCounts[r.day_of_week][r.leave_type] = (dowTypeCounts[r.day_of_week][r.leave_type] || 0) + 1;
        const dayOfMonth = parseInt(r.date.slice(8, 10), 10);
        const wom = Math.min(5, Math.ceil(dayOfMonth / 7));
        weekOfMonthTypeCounts[wom] = weekOfMonthTypeCounts[wom] || {};
        weekOfMonthTypeCounts[wom][r.leave_type] = (weekOfMonthTypeCounts[wom][r.leave_type] || 0) + 1;
      }
      const al = records.filter(r => r.leave_type === "AL");
      const wfh = records.filter(r => r.leave_type === "WFH");
      const sick = records.filter(r => r.leave_type === "S" || r.leave_type === "H.S");
      const alMonFriPct = al.length >= 4 ? al.filter(r => r.day_of_week === "Monday" || r.day_of_week === "Friday").length / al.length : null;
      const wfhThuFriPct = wfh.length >= 4 ? wfh.filter(r => r.day_of_week === "Thursday" || r.day_of_week === "Friday").length / wfh.length : null;
      const ann = annualByName[name];
      const allocated = ann ? {
        annual: num(ann.Allocated_Annual), sick: num(ann.Allocated_Sick), wfh: num(ann.Allocated_WFH),
        remainingAnnual: num(ann.Remaining_Annual), remainingSick: num(ann.Remaining_Sick), remainingWFH: num(ann.Remaining_WFH)
      } : null;
      return {
        name, total: records.length, typeCounts, dowCounts, monthCounts, monthTypeCounts, dowTypeCounts, weekOfMonthTypeCounts,
        alCount: al.length, wfhCount: wfh.length, sickCount: sick.length, alMonFriPct, wfhThuFriPct, allocated
      };
    });

    const sickSorted = employeeStats.map(e => e.sickCount).sort((a, b) => a - b);
    const sickThreshold = pctile(sickSorted, 0.9);

    for (const e of employeeStats) {
      const flags = [];
      if (e.allocated?.remainingAnnual != null && e.allocated.remainingAnnual < 0) flags.push({ code: "OVER_ANNUAL", text: "Used more Annual Leave than allocated" });
      if (e.allocated?.remainingWFH != null && e.allocated.remainingWFH < 0) flags.push({ code: "OVER_WFH", text: "Used more WFH than allocated" });
      if (e.alMonFriPct != null && e.alMonFriPct >= 0.6) flags.push({ code: "LONG_WEEKEND", text: `${Math.round(e.alMonFriPct * 100)}% of Annual Leave on Mon/Fri` });
      if (sickThreshold > 0 && e.sickCount >= sickThreshold) flags.push({ code: "HIGH_SICK", text: `${e.sickCount} sick days (top 10% company-wide)` });
      if (e.wfhThuFriPct != null && e.wfhThuFriPct < 0.4) flags.push({ code: "WFH_OFFNORM", text: "WFH mostly outside Thu/Fri norm" });
      e.flags = flags;
    }

    const typeTotals = {}, monthTypeTotals = {};
    for (const r of longRows) {
      typeTotals[r.leave_type] = (typeTotals[r.leave_type] || 0) + 1;
      monthTypeTotals[r.month] ||= {};
      monthTypeTotals[r.month][r.leave_type] = (monthTypeTotals[r.month][r.leave_type] || 0) + 1;
    }

    const dowChart = DAY_ORDER.map(d => ({ day: d.slice(0, 3) }));
    Object.keys(typeTotals).forEach(t => dowChart.forEach(row => row[t] = 0));
    for (const r of longRows) {
      const row = dowChart.find(x => x.day === r.day_of_week.slice(0, 3));
      if (row) row[r.leave_type] = (row[r.leave_type] || 0) + 1;
    }

    const presentMonths = MONTH_ORDER.filter(m => monthTypeTotals[m]);
    const monthlyChart = presentMonths.map(m => ({
      month: m.slice(0, 3),
      WFH: monthTypeTotals[m].WFH || 0, AL: monthTypeTotals[m].AL || 0,
      Sick: (monthTypeTotals[m].S || 0) + (monthTypeTotals[m]["H.S"] || 0),
    }));

    const heatmap = presentMonths.map(m => {
      const row = { month: m.slice(0, 3) };
      DAY_ORDER.forEach(d => row[d] = 0);
      return row;
    });
    for (const r of longRows) {
      if (r.leave_type !== "WFH") continue;
      const mi = presentMonths.indexOf(r.month);
      if (mi === -1) continue;
      heatmap[mi][r.day_of_week] = (heatmap[mi][r.day_of_week] || 0) + 1;
    }
    let maxHeat = 0;
    heatmap.forEach(row => DAY_ORDER.forEach(d => { if (row[d] > maxHeat) maxHeat = row[d]; }));

    const topByTotal = [...employeeStats].sort((a, b) => b.total - a.total).slice(0, 8).map(e => ({ name: e.name, total: e.total }));
    const totalWFH = typeTotals.WFH || 0;
    const wfhThuFriOnly = longRows.filter(r => r.leave_type === "WFH" && (r.day_of_week === "Thursday" || r.day_of_week === "Friday")).length;

    const kpis = {
      employees: employeeStats.length,
      totalRecords: longRows.length,
      totalWFH,
      totalAL: typeTotals.AL || 0,
      totalSick: (typeTotals.S || 0) + (typeTotals["H.S"] || 0),
      flaggedCount: employeeStats.filter(e => e.flags.length > 0).length,
      wfhThuFriPct: totalWFH > 0 ? Math.round((wfhThuFriOnly / totalWFH) * 100) : 0,
    };

    return { employeeStats, typeTotals, dowChart, monthlyChart, heatmap, maxHeat, topByTotal, kpis, sickThreshold };
  }, [parsed]);

  const filtered = useMemo(() => {
    if (!analysis) return [];
    let list = analysis.employeeStats;
    if (statusFilter === "flagged") list = list.filter(e => e.flags.length > 0);
    if (search.trim()) { const q = search.toLowerCase(); list = list.filter(e => e.name.toLowerCase().includes(q)); }
    return [...list].sort((a, b) => sortMode === "flags"
      ? (b.flags.length - a.flags.length || a.name.localeCompare(b.name, undefined, { numeric: true }))
      : a.name.localeCompare(b.name, undefined, { numeric: true }));
  }, [analysis, search, statusFilter, sortMode]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / rowsPerPage));
  const pageRows = filtered.slice((page - 1) * rowsPerPage, page * rowsPerPage);

  return (
    <div style={{ fontFamily: SANS, background: BG, minHeight: "100vh", color: TEXT }}>
      <TopNav onUpload={handleUpload} fileName={fileName} hasData={!!analysis} activeTab={activeTab} onNavigate={setActiveTab} />

      {error && (
        <div style={{ margin: "16px 28px 0", padding: "11px 16px", background: "#FDEDED", border: `1px solid ${RED}`, borderRadius: 8, fontSize: 13, color: RED, display: "flex", gap: 8, alignItems: "flex-start" }}>
          <AlertTriangle size={16} style={{ flexShrink: 0, marginTop: 1 }} />{error}
        </div>
      )}

      {loading && <div style={{ padding: "16px 28px", fontSize: 13, color: MUTED }}>Reading and reshaping workbook…</div>}

      {activeTab === "dashboard" && (!analysis ? (
        <EmptyState onUpload={handleUpload} />
      ) : (
        <div style={{ padding: "24px 32px 48px" }}>
          <KpiRow kpis={analysis.kpis} />

          <SectionBar
            count={filtered.length}
            search={search} setSearch={setSearch}
            statusFilter={statusFilter} setStatusFilter={setStatusFilter}
            sortMode={sortMode} setSortMode={setSortMode}
          />
          <EmployeeTable
            rows={pageRows}
            page={page} setPage={setPage} totalPages={totalPages}
            rowsPerPage={rowsPerPage} setRowsPerPage={setRowsPerPage}
            total={filtered.length}
            sortMode={sortMode} setSortMode={setSortMode}
            onSelect={setSelectedEmployee}
          />

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginTop: 28 }}>
            <ChartCard title="Leave type breakdown — company-wide">
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={Object.entries(analysis.typeTotals).sort((a,b)=>b[1]-a[1]).map(([type, count]) => ({ type, label: KNOWN_LABELS[type] || type, count }))} margin={{ top: 4, right: 8, left: -16, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                  <XAxis dataKey="label" tick={{ fontSize: 10.5, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} interval={0} angle={-14} textAnchor="end" height={44} />
                  <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v, n, o) => [v, o.payload.label]} contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                    {Object.keys(analysis.typeTotals).sort((a,b)=>analysis.typeTotals[b]-analysis.typeTotals[a]).map((t, i) => <Cell key={i} fill={TYPE_COLORS[t] || MUTED} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </ChartCard>

            <ChartCard title="Day-of-week pattern — all leave types">
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={analysis.dowChart} margin={{ top: 4, right: 8, left: -16, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                  <XAxis dataKey="day" tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Legend wrapperStyle={{ fontFamily: SANS, fontSize: 11 }} />
                  {Object.keys(analysis.typeTotals).map(t => <Bar key={t} dataKey={t} name={KNOWN_LABELS[t] || t} stackId="a" fill={TYPE_COLORS[t] || MUTED} radius={[0,0,0,0]} />)}
                </BarChart>
              </ResponsiveContainer>
            </ChartCard>
          </div>

          <div style={{ marginTop: 24 }}>
            <ChartCard title="Monthly trend — WFH, Annual Leave, Sick" wide>
              <ResponsiveContainer width="100%" height={210}>
                <LineChart data={analysis.monthlyChart} margin={{ top: 4, right: 20, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Legend wrapperStyle={{ fontFamily: SANS, fontSize: 11 }} />
                  <Line type="monotone" dataKey="WFH" name="Work From Home" stroke={TEAL} strokeWidth={2.5} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="AL" name="Annual Leave" stroke={BLUE} strokeWidth={2.5} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="Sick" name="Sick Leave" stroke={AMBER} strokeWidth={2.5} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </ChartCard>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginTop: 24 }}>
            <BannerCard
              title="Flag Summary"
              gradient={analysis.kpis.flaggedCount === 0 ? `linear-gradient(115deg, ${GREEN}, #1BB37E)` : `linear-gradient(115deg, ${RED}, #F0705A)`}
              pill={`${analysis.kpis.flaggedCount} flagged`}
            >
              {analysis.kpis.flaggedCount === 0 ? (
                <div style={{ padding: "26px 0", textAlign: "center" }}>
                  <CheckCircle2 size={30} color={GREEN} style={{ marginBottom: 8 }} />
                  <div style={{ color: GREEN, fontWeight: 700, fontSize: 14 }}>All clear — no employees flagged!</div>
                </div>
              ) : (
                <div style={{ maxHeight: 210, overflowY: "auto" }}>
                  {analysis.employeeStats.filter(e => e.flags.length > 0).sort((a,b)=>b.flags.length-a.flags.length).map(e => (
                    <div key={e.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 4px", borderBottom: `1px solid ${BORDER}`, fontSize: 13 }}>
                      <span style={{ fontWeight: 600 }}>{e.name}</span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5, color: RED, fontSize: 12, fontWeight: 700 }}>
                        <Flag size={11} fill={RED} /> {e.flags.length} flag{e.flags.length > 1 ? "s" : ""}
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </BannerCard>

            <BannerCard title="Top Employees by Total Records" gradient={`linear-gradient(115deg, ${PURPLE}, #9B7EF0)`}>
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={analysis.topByTotal} layout="vertical" margin={{ top: 4, right: 16, left: 8, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fontSize: 10.5, fontFamily: SANS }} stroke={MUTED} width={68} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Bar dataKey="total" fill={PURPLE} radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </BannerCard>
          </div>

          <div style={{ marginTop: 24 }}>
            <ChartCard title="WFH intensity — day of week x month" wide>
              <Heatmap data={analysis.heatmap} maxVal={analysis.maxHeat} />
            </ChartCard>
          </div>
        </div>
      ))}

      {activeTab === "employees" && (!analysis ? (
        <EmptyState onUpload={handleUpload} title="No workbook loaded" message={<>Import the Leave Sheet <strong>.xlsx</strong> to browse the employee directory.</>} />
      ) : (
        <EmployeesView
          employeeStats={analysis.employeeStats}
          search={directorySearch}
          setSearch={setDirectorySearch}
          onSelect={setSelectedEmployee}
        />
      ))}

      {activeTab === "policies" && <PoliciesView analysis={analysis} />}

      {selectedEmployee && (
        <EmployeeDetailModal
          employee={analysis.employeeStats.find(e => e.name === selectedEmployee)}
          onClose={() => setSelectedEmployee(null)}
        />
      )}
    </div>
  );
}

function TopNav({ onUpload, fileName, hasData, activeTab, onNavigate }) {
  const navItems = [
    { key: "dashboard", label: "Dashboard", icon: LayoutGrid },
    { key: "employees", label: "Employees", icon: Users },
    { key: "policies", label: "Policies", icon: ShieldAlert },
  ];
  return (
    <div style={{ background: CARD, borderBottom: `1px solid ${BORDER}`, padding: "0 28px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 26 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 7 }}>
          <span style={{ fontWeight: 900, fontSize: 17, letterSpacing: "-0.01em", color: TEXT }}>LEAVE<span style={{ color: BLUE }}>OPS</span></span>
          <span style={{ fontSize: 10.5, fontWeight: 800, color: MUTED, letterSpacing: "0.03em" }}>HR</span>
        </div>
        <nav style={{ display: "flex", gap: 20, fontSize: 13.5, height: 60, alignItems: "center" }}>
          {navItems.map(({ key, label, icon: Icon }) => {
            const active = activeTab === key;
            return (
              <span key={key} onClick={() => onNavigate?.(key)} style={{
                display: "flex", alignItems: "center", gap: 6, height: "100%",
                color: active ? BLUE : MUTED, fontWeight: active ? 700 : 500,
                borderBottom: active ? `2px solid ${BLUE}` : "2px solid transparent", cursor: "pointer"
              }}>
                <Icon size={14} /> {label}
              </span>
            );
          })}
        </nav>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <label style={{ display: "flex", alignItems: "center", gap: 7, padding: "8px 14px", background: NAVY, color: "#fff", borderRadius: 7, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
          <Upload size={14} /> {hasData ? "Re-import workbook" : "Import workbook"}
          <input type="file" accept=".xlsx,.xls" onChange={onUpload} style={{ display: "none" }} />
        </label>
        <button style={{ width: 34, height: 34, borderRadius: 8, border: `1px solid ${BORDER}`, background: CARD, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
          <RefreshCw size={14} color={MUTED} />
        </button>
        <div style={{ width: 32, height: 32, borderRadius: "50%", background: PURPLE, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13 }}>H</div>
      </div>
    </div>
  );
}

function EmptyState({ onUpload, title, message }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: "90px 24px" }}>
      <div style={{ maxWidth: 480, textAlign: "center", background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: 36 }}>
        <div style={{ width: 46, height: 46, borderRadius: 10, background: "#EAF0FF", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
          <Upload size={20} color={BLUE} />
        </div>
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 6 }}>{title || "No workbook loaded"}</div>
        <div style={{ fontSize: 13.5, color: MUTED, lineHeight: 1.6, marginBottom: 20 }}>
          {message || <>Import the original Leave Sheet <strong>.xlsx</strong> — the one with an "Annual Summary" tab plus
          one tab per month. Everything below (KPIs, charts, employee table) populates automatically.</>}
        </div>
        <label style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 18px", background: NAVY, color: "#fff", borderRadius: 8, cursor: "pointer", fontSize: 13.5, fontWeight: 600 }}>
          <Upload size={15} /> Choose .xlsx file
          <input type="file" accept=".xlsx,.xls" onChange={onUpload} style={{ display: "none" }} />
        </label>
      </div>
    </div>
  );
}

function KpiRow({ kpis }) {
  const flaggedPct = kpis.employees ? Math.round((kpis.flaggedCount / kpis.employees) * 100) : 0;
  const items = [
    { label: "Employees", value: kpis.employees, icon: Users, color: BLUE, sub: "on record" },
    { label: "Total Leave Records", value: kpis.totalRecords, icon: CalendarDays, color: NAVY_2, sub: "all types" },
    { label: "WFH Days", value: kpis.totalWFH, icon: Home, color: TEAL, sub: `${kpis.wfhThuFriPct}% on Thu/Fri`, barPct: kpis.wfhThuFriPct, barColor: TEAL },
    { label: "Annual Leave Days", value: kpis.totalAL, icon: PlaneTakeoff, color: AMBER, sub: "used company-wide" },
    { label: "Sick Days", value: kpis.totalSick, icon: Stethoscope, color: PURPLE, sub: "used company-wide" },
    {
      label: "Flagged for Review", value: kpis.flaggedCount, icon: ShieldAlert, color: RED,
      sub: kpis.flaggedCount ? `${flaggedPct}% of headcount` : "all clear", alert: kpis.flaggedCount > 0,
      barPct: kpis.flaggedCount ? flaggedPct : 100, barColor: kpis.flaggedCount ? RED : GREEN,
    },
  ];
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 16 }}>
      {items.map((it, i) => (
        <div key={i} style={{
          background: CARD, border: `1px solid ${BORDER}`, borderTop: `3px solid ${it.color}`,
          borderRadius: 10, padding: "18px 18px", boxShadow: "0 1px 2px rgba(16,24,40,0.04)"
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
            <span style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", color: MUTED, fontWeight: 700 }}>{it.label}</span>
            <div style={{ width: 26, height: 26, borderRadius: "50%", background: it.color, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <it.icon size={13} color="#fff" />
            </div>
          </div>
          <div style={{ fontSize: 28, fontWeight: 800, color: it.alert ? RED : TEXT }}>{it.value}</div>
          <div style={{ fontSize: 11.5, color: it.alert ? RED : MUTED, marginTop: 3, fontWeight: it.alert ? 700 : 500 }}>{it.sub}</div>
          {it.barPct != null && (
            <div style={{ background: "#EEF1F6", borderRadius: 6, height: 4, width: "100%", marginTop: 10 }}>
              <div style={{ width: `${Math.min(100, it.barPct)}%`, height: "100%", borderRadius: 6, background: it.barColor }} />
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

function SectionBar({ count, search, setSearch, statusFilter, setStatusFilter, sortMode, setSortMode }) {
  return (
    <div style={{
      marginTop: 22, background: `linear-gradient(115deg, ${NAVY}, ${NAVY_2})`, borderRadius: "10px 10px 0 0",
      padding: "14px 18px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <span style={{ color: "#fff", fontWeight: 700, fontSize: 14.5 }}>All Employees</span>
        <span style={{ background: "rgba(255,255,255,0.15)", color: "#fff", fontSize: 11.5, padding: "2px 9px", borderRadius: 20 }}>{count} employees</span>
      </div>
      <div style={{ display: "flex", gap: 10 }}>
        <div style={{ position: "relative" }}>
          <Search size={13} style={{ position: "absolute", left: 9, top: 9, color: MUTED }} />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search employees..."
            style={{ padding: "7px 10px 7px 28px", fontSize: 12.5, borderRadius: 6, border: "none", width: 190, background: "rgba(255,255,255,0.9)", color: TEXT }} />
        </div>
        <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)}
          style={{ padding: "7px 10px", fontSize: 12.5, borderRadius: 6, border: "none", background: "rgba(255,255,255,0.9)", color: TEXT, fontWeight: 600 }}>
          <option value="all">All Status</option>
          <option value="flagged">Flagged Only</option>
        </select>
        <button
          onClick={() => setSortMode(m => m === "flags" ? "name" : "flags")}
          style={{
            display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", fontSize: 12.5, fontWeight: 600,
            borderRadius: 6, border: "none", background: "rgba(255,255,255,0.9)", color: TEXT, cursor: "pointer", whiteSpace: "nowrap"
          }}>
          <ChevronDown size={13} style={{ transform: sortMode === "flags" ? "none" : "rotate(180deg)" }} />
          {sortMode === "flags" ? "Most flagged first" : "Name A–Z"}
        </button>
      </div>
    </div>
  );
}

function EmployeeTable({ rows, page, setPage, totalPages, rowsPerPage, setRowsPerPage, total, sortMode, setSortMode, onSelect }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderTop: "none", borderRadius: "0 0 10px 10px", overflow: "hidden" }}>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8F9FC" }}>
              {["EMPLOYEE", "ANNUAL LEAVE", "SICK", "WFH", "TOTAL", "STATUS", ""].map(h => (
                <th key={h}
                  onClick={h === "EMPLOYEE" && setSortMode ? () => setSortMode(m => m === "flags" ? "name" : "flags") : undefined}
                  style={{
                    padding: "10px 14px", textAlign: "left", fontSize: 11, color: MUTED, fontWeight: 700, letterSpacing: "0.03em",
                    borderBottom: `1px solid ${BORDER}`, whiteSpace: "nowrap", cursor: h === "EMPLOYEE" && setSortMode ? "pointer" : "default"
                  }}>
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                    {h}
                    {h === "EMPLOYEE" && <ChevronDown size={11} style={{ transform: sortMode === "flags" ? "none" : "rotate(180deg)" }} />}
                  </span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(e => {
              const usedPct = e.allocated?.annual ? Math.min(100, Math.round((e.alCount / e.allocated.annual) * 100)) : null;
              const over = e.allocated?.remainingAnnual != null && e.allocated.remainingAnnual < 0;
              const accent = e.flags.length > 0 ? RED : GREEN;
              return (
                <tr key={e.name} onClick={() => onSelect?.(e.name)}
                  style={{ borderBottom: `1px solid ${BORDER}`, borderLeft: `3px solid ${accent}`, cursor: onSelect ? "pointer" : "default" }}
                  onMouseEnter={ev => ev.currentTarget.style.background = "#F8F9FC"}
                  onMouseLeave={ev => ev.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "11px 14px", fontWeight: 700, color: BLUE }}>{e.name}</td>
                  <td style={{ padding: "11px 14px", minWidth: 140 }}>
                    <div style={{ fontSize: 12, marginBottom: 4 }}>{e.alCount}{e.allocated?.annual != null ? ` / ${e.allocated.annual}` : ""}</div>
                    {usedPct != null && (
                      <div style={{ background: "#EEF1F6", borderRadius: 6, height: 5, width: 100 }}>
                        <div style={{ width: `${usedPct}%`, height: "100%", borderRadius: 6, background: over ? RED : GREEN }} />
                      </div>
                    )}
                  </td>
                  <td style={{ padding: "11px 14px" }}>{e.sickCount}</td>
                  <td style={{ padding: "11px 14px" }}>{e.wfhCount}</td>
                  <td style={{ padding: "11px 14px", fontWeight: 700 }}>{e.total}</td>
                  <td style={{ padding: "11px 14px" }}>
                    {e.flags.length > 0 ? (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: "#FDEDED", color: RED, padding: "3px 9px", borderRadius: 20, fontSize: 11.5, fontWeight: 700 }}>
                        <Flag size={10} fill={RED} /> {e.flags.length} flag{e.flags.length > 1 ? "s" : ""}
                      </span>
                    ) : (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: "#E9F7F0", color: GREEN, padding: "3px 9px", borderRadius: 20, fontSize: 11.5, fontWeight: 700 }}>
                        <CheckCircle2 size={10} /> Clear
                      </span>
                    )}
                  </td>
                  <td style={{ padding: "11px 14px" }}>
                    <span onClick={ev => { ev.stopPropagation(); onSelect?.(e.name); }}
                      style={{ display: "inline-flex", alignItems: "center", gap: 3, color: BLUE, fontWeight: 700, fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>
                      View <ChevronRight size={12} />
                    </span>
                  </td>
                </tr>
              );
            })}
            {rows.length === 0 && (
              <tr><td colSpan={7} style={{ padding: 20, color: MUTED, textAlign: "center" }}>No employees match.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 16px", fontSize: 12.5, color: MUTED }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          Rows:
          <select value={rowsPerPage} onChange={e => { setRowsPerPage(Number(e.target.value)); setPage(1); }} style={{ border: `1px solid ${BORDER}`, borderRadius: 5, padding: "3px 6px", fontSize: 12 }}>
            {[10, 20, 50].map(n => <option key={n} value={n}>{n}</option>)}
          </select>
        </div>
        <div>Showing {total === 0 ? 0 : (page - 1) * rowsPerPage + 1}–{Math.min(page * rowsPerPage, total)} of {total} employees</div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <button onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page === 1} style={{ border: `1px solid ${BORDER}`, background: CARD, borderRadius: 5, width: 26, height: 26, cursor: page === 1 ? "default" : "pointer", opacity: page === 1 ? 0.4 : 1 }}>
            <ChevronLeft size={13} />
          </button>
          <span>Page {page} of {totalPages}</span>
          <button onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={page === totalPages} style={{ border: `1px solid ${BORDER}`, background: CARD, borderRadius: 5, width: 26, height: 26, cursor: page === totalPages ? "default" : "pointer", opacity: page === totalPages ? 0.4 : 1 }}>
            <ChevronRight size={13} />
          </button>
        </div>
      </div>
    </div>
  );
}

function ChartCard({ title, children, wide }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "20px 22px", boxShadow: "0 1px 2px rgba(16,24,40,0.04)", gridColumn: wide ? "1 / -1" : "auto" }}>
      <div style={{ fontSize: 13.5, fontWeight: 700, color: TEXT, marginBottom: 14 }}>{title}</div>
      {children}
    </div>
  );
}

function BannerCard({ title, gradient, pill, children }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 2px rgba(16,24,40,0.04)" }}>
      <div style={{ background: gradient, padding: "12px 16px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ color: "#fff", fontWeight: 700, fontSize: 13.5 }}>{title}</span>
        {pill && <span style={{ background: "rgba(255,255,255,0.2)", color: "#fff", fontSize: 11, padding: "2px 9px", borderRadius: 20 }}>{pill}</span>}
      </div>
      <div style={{ padding: "18px 20px" }}>{children}</div>
    </div>
  );
}

function Heatmap({ data, maxVal }) {
  if (!data.length) return <div style={{ fontSize: 12.5, color: MUTED }}>Not enough data.</div>;
  const cellColor = (v) => {
    if (maxVal === 0) return "#F3F5F9";
    const t = v / maxVal;
    const c1 = [235, 240, 250], c2 = [61, 111, 234], c3 = [21, 42, 78];
    const [a, b] = t < 0.5 ? [c1, c2] : [c2, c3];
    const tt = t < 0.5 ? t * 2 : (t - 0.5) * 2;
    const rgb = a.map((v0, i) => Math.round(v0 + (b[i] - v0) * tt));
    return `rgb(${rgb.join(",")})`;
  };
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: `70px repeat(${DAY_ORDER.length}, 1fr)`, gap: 4, marginBottom: 4 }}>
        <div />
        {DAY_ORDER.map(d => <div key={d} style={{ fontSize: 11, color: MUTED, textAlign: "center", fontWeight: 600 }}>{d.slice(0, 3)}</div>)}
      </div>
      {data.map(row => (
        <div key={row.month} style={{ display: "grid", gridTemplateColumns: `70px repeat(${DAY_ORDER.length}, 1fr)`, gap: 4, marginBottom: 4 }}>
          <div style={{ fontSize: 12, color: MUTED, display: "flex", alignItems: "center", fontWeight: 600 }}>{row.month}</div>
          {DAY_ORDER.map(d => (
            <div key={d} title={`${row[d]}`} style={{
              background: cellColor(row[d]), borderRadius: 5, height: 28,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 11, color: row[d] / (maxVal || 1) > 0.5 ? "#fff" : TEXT, fontWeight: 700
            }}>{row[d] || ""}</div>
          ))}
        </div>
      ))}
    </div>
  );
}

/* ---------- per-employee detail modal ---------- */
const FLAG_ICON_BG = { OVER_ANNUAL: RED, OVER_WFH: RED, LONG_WEEKEND: AMBER, HIGH_SICK: AMBER, WFH_OFFNORM: AMBER };

function MiniScorecard({ label, used, allocated, remaining, color }) {
  const pct = allocated ? Math.min(100, Math.round((used / allocated) * 100)) : null;
  const over = remaining != null && remaining < 0;
  return (
    <div style={{ background: "#F8F9FC", border: `1px solid ${BORDER}`, borderRadius: 8, padding: "16px 18px", flex: 1 }}>
      <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", color: MUTED, fontWeight: 700, marginBottom: 8 }}>{label}</div>
      <div style={{ fontSize: 20, fontWeight: 800, color: over ? RED : TEXT }}>
        {used}{allocated != null ? <span style={{ fontSize: 13, color: MUTED, fontWeight: 600 }}> / {allocated}</span> : null}
      </div>
      <div style={{ fontSize: 11.5, color: over ? RED : MUTED, marginTop: 2, fontWeight: over ? 700 : 500 }}>
        {remaining != null ? `${over ? "Over by " + Math.abs(remaining) : "Remaining: " + remaining}` : "No allocation on file"}
      </div>
      {pct != null && (
        <div style={{ background: "#EEF1F6", borderRadius: 6, height: 5, width: "100%", marginTop: 8 }}>
          <div style={{ width: `${pct}%`, height: "100%", borderRadius: 6, background: over ? RED : color }} />
        </div>
      )}
    </div>
  );
}

function EmployeeDetailModal({ employee, onClose }) {
  if (!employee) return null;
  const e = employee;

  const typeData = Object.entries(e.typeCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([type, count]) => ({ type, label: KNOWN_LABELS[type] || type, count, pct: e.total ? Math.round((count / e.total) * 100) : 0 }));

  // 1. entire week — which weekday this employee WFHs the most
  const wfhByDow = DAY_ORDER.map(d => ({ day: d.slice(0, 3), full: d, count: e.dowTypeCounts[d]?.WFH || 0 }));
  const busiestWfhDay = wfhByDow.reduce((max, d) => (d.count > max.count ? d : max), wfhByDow[0]);

  // 2. entire month — which week-of-month (1st, 2nd, 3rd... week) sees the most WFH / Annual / Sick usage
  const weekOfMonthData = [1, 2, 3, 4, 5].map(w => {
    const wom = e.weekOfMonthTypeCounts[w] || {};
    return {
      week: w === 5 ? "Week 5" : `Week ${w}`,
      WFH: wom.WFH || 0,
      "Annual Leave": wom.AL || 0,
      Sick: (wom.S || 0) + (wom["H.S"] || 0),
    };
  });

  // 3. entire year — WFH, Annual Leave, and Sick trend lines, month by month
  const yearTrendData = MONTH_ORDER.map(m => {
    const mt = e.monthTypeCounts[m] || {};
    return {
      month: m.slice(0, 3),
      WFH: mt.WFH || 0,
      "Annual Leave": mt.AL || 0,
      Sick: (mt.S || 0) + (mt["H.S"] || 0),
    };
  });

  // 4. leave composition — % of total leave that's WFH vs each other category
  const donutColors = typeData.map(d => TYPE_COLORS[d.type] || MUTED);

  // 5. quota pacing — cumulative Annual Leave usage vs the allocated cap
  let alRunning = 0;
  const alPaceData = MONTH_ORDER.map(m => {
    alRunning += e.monthTypeCounts[m]?.AL || 0;
    return { month: m.slice(0, 3), cumulative: alRunning };
  });
  const alAllocated = e.allocated?.annual;
  const alMonthsWithData = Object.keys(e.monthCounts).length;
  const alOverPace = alAllocated != null && alMonthsWithData > 0 &&
    (alRunning / Math.max(1, alMonthsWithData)) * 12 > alAllocated;

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed", inset: 0, background: "rgba(15,22,38,0.55)", zIndex: 100,
        display: "flex", alignItems: "center", justifyContent: "center", padding: 20
      }}>
      <div
        onClick={ev => ev.stopPropagation()}
        style={{
          background: CARD, borderRadius: 12, width: "100%", maxWidth: 1060, maxHeight: "90vh",
          overflow: "hidden", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(15,22,38,0.35)"
        }}>
        {/* header */}
        <div style={{
          background: `linear-gradient(115deg, ${NAVY}, ${NAVY_2})`, padding: "16px 22px",
          display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0
        }}>
          <div>
            <div style={{ fontSize: 10.5, color: "rgba(255,255,255,0.65)", fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 2 }}>
              Employee Detail
            </div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>{e.name}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {e.flags.length > 0 ? (
              <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: "rgba(255,255,255,0.18)", color: "#fff", padding: "4px 11px", borderRadius: 20, fontSize: 12, fontWeight: 700 }}>
                <Flag size={11} fill="#fff" /> {e.flags.length} flag{e.flags.length > 1 ? "s" : ""}
              </span>
            ) : (
              <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: "rgba(255,255,255,0.18)", color: "#fff", padding: "4px 11px", borderRadius: 20, fontSize: 12, fontWeight: 700 }}>
                <CheckCircle2 size={11} /> Clear
              </span>
            )}
            <button onClick={onClose} style={{
              width: 30, height: 30, borderRadius: 7, border: "none", background: "rgba(255,255,255,0.15)",
              display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer"
            }}>
              <X size={15} color="#fff" />
            </button>
          </div>
        </div>

        {/* body */}
        <div style={{ padding: "28px 30px 32px", overflowY: "auto" }}>
          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: MUTED, fontWeight: 700, marginBottom: 12 }}>
            Leave Balances
          </div>
          <div style={{ display: "flex", gap: 16, marginBottom: 32 }}>
            <MiniScorecard label="Annual Leave" used={e.alCount} allocated={e.allocated?.annual} remaining={e.allocated?.remainingAnnual} color={BLUE} />
            <MiniScorecard label="Sick" used={e.sickCount} allocated={e.allocated?.sick} remaining={e.allocated?.remainingSick} color={AMBER} />
            <MiniScorecard label="WFH" used={e.wfhCount} allocated={e.allocated?.wfh} remaining={e.allocated?.remainingWFH} color={TEAL} />
          </div>

          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: MUTED, fontWeight: 700, marginBottom: 12 }}>
            Flag Reasons
          </div>
          <div style={{ marginBottom: 32 }}>
            {e.flags.length === 0 ? (
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: GREEN, fontWeight: 700, fontSize: 13.5 }}>
                <CheckCircle2 size={18} /> No HR flags for this employee
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {e.flags.map((f, i) => (
                  <div key={i} style={{
                    display: "flex", alignItems: "flex-start", gap: 8, padding: "8px 10px",
                    background: "#FDEDED", borderRadius: 8, fontSize: 12.5, color: TEXT
                  }}>
                    <Flag size={13} color={FLAG_ICON_BG[f.code] || RED} style={{ flexShrink: 0, marginTop: 1 }} />
                    <span>{f.text}</span>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: MUTED, fontWeight: 700, marginBottom: 12 }}>
            Leave Composition
          </div>
          <div style={{ marginBottom: 28 }}>
            <ChartCard title="Share of total leave by category">
              {e.total === 0 ? (
                <div style={{ padding: "24px 0", textAlign: "center", color: MUTED, fontSize: 12.5 }}>No leave records for this employee yet.</div>
              ) : (
                <div style={{ display: "flex", alignItems: "center", gap: 20, flexWrap: "wrap" }}>
                  <ResponsiveContainer width={220} height={220}>
                    <PieChart>
                      <Pie data={typeData} dataKey="count" nameKey="label" innerRadius={58} outerRadius={95} paddingAngle={2}>
                        {typeData.map((d, i) => <Cell key={i} fill={donutColors[i]} />)}
                      </Pie>
                      <Tooltip formatter={(v, n, o) => [`${v} days (${o.payload.pct}%)`, o.payload.label]} contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8, flex: 1, minWidth: 180 }}>
                    {typeData.map((d, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", fontSize: 12.5 }}>
                        <span style={{ display: "flex", alignItems: "center", gap: 7 }}>
                          <span style={{ width: 10, height: 10, borderRadius: 3, background: donutColors[i], flexShrink: 0 }} />
                          {d.label}
                        </span>
                        <span style={{ fontWeight: 700, color: TEXT }}>{d.pct}% <span style={{ color: MUTED, fontWeight: 500 }}>({d.count})</span></span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </ChartCard>
          </div>

          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: MUTED, fontWeight: 700, marginBottom: 12 }}>
            Weekly &amp; Monthly Patterns
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24, rowGap: 28, marginBottom: 28 }}>
            <ChartCard title="Which weekday they WFH the most">
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={wfhByDow} margin={{ top: 4, right: 8, left: -16, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                  <XAxis dataKey="day" tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Bar dataKey="count" name="WFH days" radius={[4, 4, 0, 0]}>
                    {wfhByDow.map((d, i) => <Cell key={i} fill={d.full === busiestWfhDay.full && d.count > 0 ? TEAL : "#C9E7E4"} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{ fontSize: 11.5, color: MUTED, marginTop: 8, lineHeight: 1.5 }}>
                {busiestWfhDay.count > 0
                  ? <>Most WFH days fall on <strong style={{ color: TEXT }}>{busiestWfhDay.full}</strong> ({busiestWfhDay.count} days).</>
                  : "No WFH days recorded for this employee."}
              </div>
            </ChartCard>

            <ChartCard title="Which week of the month, by leave type">
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={weekOfMonthData} margin={{ top: 4, right: 8, left: -16, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                  <XAxis dataKey="week" tick={{ fontSize: 10.5, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Legend wrapperStyle={{ fontFamily: SANS, fontSize: 11 }} />
                  <Bar dataKey="WFH" name="Work From Home" fill={TEAL} radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Annual Leave" name="Annual Leave" fill={BLUE} radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Sick" name="Sick Leave" fill={AMBER} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
              <div style={{ fontSize: 11.5, color: MUTED, marginTop: 8, lineHeight: 1.5 }}>
                Week 5 only exists in months with a 5th Mon–Fri span, so it will naturally look smaller.
              </div>
            </ChartCard>
          </div>

          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: MUTED, fontWeight: 700, marginBottom: 12 }}>
            Yearly Trend
          </div>
          <div style={{ marginBottom: 28 }}>
            <ChartCard title="WFH, Annual Leave &amp; Sick Leave across the year">
              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={yearTrendData} margin={{ top: 4, right: 20, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                  <Legend wrapperStyle={{ fontFamily: SANS, fontSize: 11 }} />
                  <Line type="monotone" dataKey="WFH" name="Work From Home" stroke={TEAL} strokeWidth={2.5} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="Annual Leave" name="Annual Leave" stroke={BLUE} strokeWidth={2.5} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="Sick" name="Sick Leave" stroke={AMBER} strokeWidth={2.5} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </ChartCard>
          </div>

          <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: MUTED, fontWeight: 700, marginBottom: 12 }}>
            Quota Pacing
          </div>
          <ChartCard title="Annual Leave pace vs. allocation">
            {alAllocated != null ? (
              <>
                <ResponsiveContainer width="100%" height={200}>
                  <AreaChart data={alPaceData} margin={{ top: 4, right: 12, left: -16, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={BORDER} vertical={false} />
                    <XAxis dataKey="month" tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} axisLine={{ stroke: BORDER }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fontFamily: SANS }} stroke={MUTED} allowDecimals={false} axisLine={false} tickLine={false} domain={[0, Math.max(alAllocated, alRunning) + 2]} />
                    <Tooltip contentStyle={{ fontFamily: SANS, fontSize: 12, borderRadius: 8, border: `1px solid ${BORDER}` }} />
                    <ReferenceLine y={alAllocated} stroke={RED} strokeDasharray="5 4" label={{ value: `Allocated: ${alAllocated}`, position: "insideTopRight", fill: RED, fontSize: 11, fontFamily: SANS, fontWeight: 700 }} />
                    <Area type="monotone" dataKey="cumulative" name="Cumulative Annual Leave" stroke={BLUE} fill={BLUE} fillOpacity={0.15} strokeWidth={2.5} dot={{ r: 3, fill: BLUE }} />
                  </AreaChart>
                </ResponsiveContainer>
                <div style={{ fontSize: 11.5, color: alOverPace ? RED : MUTED, marginTop: 8, lineHeight: 1.5, fontWeight: alOverPace ? 700 : 500 }}>
                  {alOverPace
                    ? `At this pace, projected usage would exceed the ${alAllocated}-day allocation before year end.`
                    : "Running total of Annual Leave days taken, checked against the allocated cap for the year."}
                </div>
              </>
            ) : (
              <div style={{ padding: "24px 0", textAlign: "center", color: MUTED, fontSize: 12.5 }}>
                No Annual Leave allocation on file for this employee — add one to the Annual Summary tab to see pacing here.
              </div>
            )}
          </ChartCard>
        </div>
      </div>
    </div>
  );
}

/* ---------- employees tab: directory ---------- */
function EmployeesView({ employeeStats, search, setSearch, onSelect }) {
  const filtered = employeeStats
    .filter(e => e.name.toLowerCase().includes(search.trim().toLowerCase()))
    .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

  return (
    <div style={{ padding: "24px 32px 48px" }}>
      <div style={{
        background: `linear-gradient(115deg, ${NAVY}, ${NAVY_2})`, borderRadius: 10, padding: "16px 20px",
        display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 20
      }}>
        <div>
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 16 }}>Employee Directory</div>
          <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 12.5, marginTop: 2 }}>
            {filtered.length} of {employeeStats.length} employees — click any card for the full breakdown
          </div>
        </div>
        <div style={{ position: "relative" }}>
          <Search size={13} style={{ position: "absolute", left: 9, top: 9, color: MUTED }} />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search employees..."
            style={{ padding: "7px 10px 7px 28px", fontSize: 12.5, borderRadius: 6, border: "none", width: 220, background: "rgba(255,255,255,0.9)", color: TEXT }} />
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 16 }}>
        {filtered.map(e => <EmployeeCard key={e.name} e={e} onSelect={onSelect} />)}
        {filtered.length === 0 && (
          <div style={{ gridColumn: "1 / -1", textAlign: "center", color: MUTED, padding: 40 }}>
            No employees match "{search}".
          </div>
        )}
      </div>
    </div>
  );
}

function EmployeeCard({ e, onSelect }) {
  const initials = e.name.split(" ").map(p => p[0]).filter(Boolean).slice(0, 2).join("").toUpperCase();
  const accent = e.flags.length > 0 ? RED : GREEN;
  return (
    <div onClick={() => onSelect(e.name)} style={{
      background: CARD, border: `1px solid ${BORDER}`, borderRadius: 10, padding: 16,
      cursor: "pointer", boxShadow: "0 1px 2px rgba(16,24,40,0.04)"
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <div style={{
          width: 36, height: 36, borderRadius: "50%", background: accent, color: "#fff",
          display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, flexShrink: 0
        }}>{initials}</div>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{e.name}</div>
          <div style={{ fontSize: 11, color: MUTED }}>{e.total} leave records</div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
        <MiniStat label="Annual" value={e.alCount} allocated={e.allocated?.annual} />
        <MiniStat label="Sick" value={e.sickCount} allocated={e.allocated?.sick} />
        <MiniStat label="WFH" value={e.wfhCount} allocated={e.allocated?.wfh} />
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        {e.flags.length > 0 ? (
          <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: "#FDEDED", color: RED, padding: "3px 9px", borderRadius: 20, fontSize: 11, fontWeight: 700 }}>
            <Flag size={9} fill={RED} /> {e.flags.length} flag{e.flags.length > 1 ? "s" : ""}
          </span>
        ) : (
          <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: "#E9F7F0", color: GREEN, padding: "3px 9px", borderRadius: 20, fontSize: 11, fontWeight: 700 }}>
            <CheckCircle2 size={9} /> Clear
          </span>
        )}
        <span style={{ display: "inline-flex", alignItems: "center", gap: 3, color: BLUE, fontWeight: 700, fontSize: 11.5 }}>
          View <ChevronRight size={11} />
        </span>
      </div>
    </div>
  );
}

function MiniStat({ label, value, allocated }) {
  return (
    <div style={{ background: "#F8F9FC", borderRadius: 6, padding: "6px 8px", textAlign: "center" }}>
      <div style={{ fontSize: 10, color: MUTED, fontWeight: 700, marginBottom: 2 }}>{label}</div>
      <div style={{ fontSize: 13, fontWeight: 800, color: TEXT }}>
        {value}{allocated != null ? <span style={{ fontSize: 10, color: MUTED, fontWeight: 600 }}>/{allocated}</span> : null}
      </div>
    </div>
  );
}

/* ---------- policies tab ---------- */
const POLICY_RULES = [
  {
    code: "OVER_ANNUAL", icon: PlaneTakeoff, color: RED, title: "Over-allocated Annual Leave",
    description: "Flags anyone whose recorded Annual Leave days exceed their allocated balance for the year — their remaining balance goes negative."
  },
  {
    code: "OVER_WFH", icon: Home, color: RED, title: "Over-allocated WFH",
    description: "Flags anyone whose recorded Work From Home days exceed their allocated WFH balance for the year."
  },
  {
    code: "LONG_WEEKEND", icon: CalendarDays, color: AMBER, title: "Long-weekend pattern",
    description: "Flags anyone who takes 60% or more of their Annual Leave on Mondays or Fridays, once they have at least 4 Annual Leave days on record — a common long-weekend extension pattern."
  },
  {
    code: "HIGH_SICK", icon: Stethoscope, color: AMBER, title: "High sick-day usage",
    description: "Flags anyone whose sick-day count sits in the top 10% company-wide for the period covered by the workbook."
  },
  {
    code: "WFH_OFFNORM", icon: ShieldAlert, color: AMBER, title: "Off-norm WFH pattern",
    description: "Flags anyone whose Work From Home days fall mostly outside the expected Thursday/Friday norm — under 40% on Thu/Fri, once they have at least 4 WFH days on record."
  },
];

function PoliciesView({ analysis }) {
  return (
    <div style={{ padding: "24px 32px 48px" }}>
      <div style={{ marginBottom: 22 }}>
        <div style={{ fontSize: 16, fontWeight: 700, color: TEXT }}>HR Flagging Policies</div>
        <div style={{ fontSize: 13, color: MUTED, marginTop: 4, maxWidth: 640, lineHeight: 1.6 }}>
          These are the automatic rules the dashboard uses to flag employees for review. They run the same way against every workbook you import — nothing here is edited by hand.
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
        {POLICY_RULES.map(rule => (
          <div key={rule.code} style={{
            background: CARD, border: `1px solid ${BORDER}`, borderTop: `3px solid ${rule.color}`,
            borderRadius: 10, padding: "18px 20px", boxShadow: "0 1px 2px rgba(16,24,40,0.04)"
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
              <div style={{ width: 28, height: 28, borderRadius: "50%", background: rule.color, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <rule.icon size={14} color="#fff" />
              </div>
              <div style={{ fontWeight: 700, fontSize: 13.5, color: TEXT }}>{rule.title}</div>
            </div>
            <div style={{ fontSize: 12.5, color: MUTED, lineHeight: 1.6 }}>{rule.description}</div>
            {rule.code === "HIGH_SICK" && analysis?.sickThreshold != null && (
              <div style={{ marginTop: 12, display: "inline-block", background: "#FFF6E9", color: AMBER, fontSize: 11.5, fontWeight: 700, padding: "3px 9px", borderRadius: 20 }}>
                Current threshold from your data: {Math.ceil(analysis.sickThreshold)}+ sick days
              </div>
            )}
          </div>
        ))}
      </div>

      <ChartCard title="Leave type reference">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 14 }}>
          {Object.entries(KNOWN_LABELS).map(([code, label]) => (
            <div key={code} style={{ display: "flex", alignItems: "center", gap: 9 }}>
              <div style={{ width: 12, height: 12, borderRadius: 4, background: TYPE_COLORS[code] || MUTED, flexShrink: 0 }} />
              <div>
                <div style={{ fontSize: 12.5, fontWeight: 700, color: TEXT }}>{code}</div>
                <div style={{ fontSize: 11, color: MUTED }}>{label}</div>
              </div>
            </div>
          ))}
        </div>
      </ChartCard>
    </div>
  );
}

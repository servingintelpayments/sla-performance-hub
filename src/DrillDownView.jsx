import { useState, useMemo } from "react";

/* ═══════════════════════════════════════════════════════════════════
   DRILL-DOWN VIEW — Embedded Component
   ═══════════════════════════════════════════════════════════════════ */
const DC = {
  bg: "#0B1120", card: "#111827", surface: "#1E293B",
  border: "#1E3A5F", borderLight: "#2d4a6f",
  text: "#E2E8F0", muted: "#94A3B8", dim: "#64748B",
  blue: "#3B82F6", blueGlow: "rgba(59,130,246,0.15)",
  green: "#10B981", greenGlow: "rgba(16,185,129,0.12)",
  red: "#EF4444", redGlow: "rgba(239,68,68,0.12)",
  orange: "#F59E0B", orangeGlow: "rgba(245,158,11,0.12)",
  purple: "#8B5CF6", cyan: "#06B6D4", gold: "#FBBF24", white: "#FFFFFF",
};

const DD_TIERS = ["All Tiers", "Tier 1", "Tier 2", "Tier 3"];
const DD_TIER_COLORS = { "Tier 1": DC.blue, "Tier 2": DC.purple, "Tier 3": DC.orange };

const DD_AGENTS = [
  {
    name: "Steve Rogers", avatar: "SR", tier: "Tier 1",
    metrics: { casesCreated: 8, casesResolved: 7, slaCompliance: 87, fcr: 71, avgResponseTime: "1h 34m", csatScore: 4.6, phoneCalls: 12, phoneAnswered: 10, phoneAbandoned: 2, avgTalkTime: "4m 12s" },
    cases: [
      { id: "CAS-4821", title: "Outlook sync issue", created: "02/28 9:14 AM", resolved: "02/28 10:36 AM", sla: "met", responseTime: "1h 22m", fcr: true, priority: "Normal" },
      { id: "CAS-4823", title: "VPN not connecting", created: "02/28 10:02 AM", resolved: "02/28 3:12 PM", sla: "breached", responseTime: "5h 10m", fcr: false, priority: "High" },
      { id: "CAS-4825", title: "Password reset request", created: "02/28 10:45 AM", resolved: "02/28 11:01 AM", sla: "met", responseTime: "16m", fcr: true, priority: "Normal" },
      { id: "CAS-4829", title: "Printer not responding", created: "02/28 11:30 AM", resolved: "02/28 12:48 PM", sla: "met", responseTime: "1h 18m", fcr: true, priority: "Normal" },
      { id: "CAS-4831", title: "SharePoint access denied", created: "02/28 1:15 PM", resolved: "02/28 2:20 PM", sla: "met", responseTime: "1h 05m", fcr: false, priority: "Normal" },
      { id: "CAS-4834", title: "Teams audio not working", created: "02/28 2:00 PM", resolved: "02/28 3:45 PM", sla: "met", responseTime: "1h 45m", fcr: true, priority: "Normal" },
      { id: "CAS-4838", title: "Software install request", created: "02/28 3:22 PM", resolved: null, sla: "at-risk", responseTime: "2h 38m", fcr: false, priority: "High" },
      { id: "CAS-4841", title: "Email delivery failure", created: "02/28 4:10 PM", resolved: "02/28 4:55 PM", sla: "met", responseTime: "45m", fcr: true, priority: "Normal" },
    ],
  },
  {
    name: "Herbert Felix", avatar: "HF", tier: "Tier 1",
    metrics: { casesCreated: 6, casesResolved: 6, slaCompliance: 100, fcr: 83, avgResponseTime: "52m", csatScore: 4.8, phoneCalls: 9, phoneAnswered: 8, phoneAbandoned: 1, avgTalkTime: "3m 45s" },
    cases: [
      { id: "CAS-4822", title: "Monitor flickering", created: "02/28 9:30 AM", resolved: "02/28 10:10 AM", sla: "met", responseTime: "40m", fcr: true, priority: "Normal" },
      { id: "CAS-4826", title: "New hire setup", created: "02/28 10:50 AM", resolved: "02/28 12:15 PM", sla: "met", responseTime: "1h 25m", fcr: false, priority: "Normal" },
      { id: "CAS-4828", title: "Calendar sync broken", created: "02/28 11:20 AM", resolved: "02/28 11:55 AM", sla: "met", responseTime: "35m", fcr: true, priority: "Normal" },
      { id: "CAS-4832", title: "Wi-Fi connectivity", created: "02/28 1:30 PM", resolved: "02/28 2:05 PM", sla: "met", responseTime: "35m", fcr: true, priority: "Normal" },
      { id: "CAS-4836", title: "MFA lockout", created: "02/28 2:45 PM", resolved: "02/28 3:10 PM", sla: "met", responseTime: "25m", fcr: true, priority: "High" },
      { id: "CAS-4840", title: "File recovery request", created: "02/28 3:55 PM", resolved: "02/28 5:00 PM", sla: "met", responseTime: "1h 05m", fcr: true, priority: "Normal" },
    ],
  },
  {
    name: "Oscar Martinez", avatar: "OM", tier: "Tier 1",
    metrics: { casesCreated: 5, casesResolved: 4, slaCompliance: 80, fcr: 60, avgResponseTime: "1h 48m", csatScore: 4.2, phoneCalls: 7, phoneAnswered: 5, phoneAbandoned: 2, avgTalkTime: "5m 30s" },
    cases: [
      { id: "CAS-4827", title: "Laptop overheating", created: "02/28 11:00 AM", resolved: "02/28 1:30 PM", sla: "met", responseTime: "2h 30m", fcr: false, priority: "Normal" },
      { id: "CAS-4833", title: "Dock station not detected", created: "02/28 1:45 PM", resolved: "02/28 2:50 PM", sla: "met", responseTime: "1h 05m", fcr: true, priority: "Normal" },
      { id: "CAS-4837", title: "OneDrive sync error", created: "02/28 3:00 PM", resolved: "02/28 4:15 PM", sla: "met", responseTime: "1h 15m", fcr: true, priority: "Normal" },
      { id: "CAS-4842", title: "Adobe license expired", created: "02/28 4:20 PM", resolved: null, sla: "breached", responseTime: "4h 40m+", fcr: false, priority: "High" },
      { id: "CAS-4843", title: "Webcam not working", created: "02/28 4:45 PM", resolved: "02/28 5:10 PM", sla: "met", responseTime: "25m", fcr: true, priority: "Normal" },
    ],
  },
  {
    name: "Diana Prince", avatar: "DP", tier: "Tier 2",
    metrics: { casesCreated: 4, casesResolved: 3, slaCompliance: 75, fcr: 50, avgResponseTime: "2h 15m", csatScore: 4.3, phoneCalls: 3, phoneAnswered: 3, phoneAbandoned: 0, avgTalkTime: "8m 20s" },
    cases: [
      { id: "CAS-4824", title: "API integration error", created: "02/28 9:45 AM", resolved: "02/28 1:20 PM", sla: "met", responseTime: "3h 35m", fcr: false, priority: "High" },
      { id: "CAS-4830", title: "Database timeout issue", created: "02/28 11:50 AM", resolved: "02/28 2:30 PM", sla: "met", responseTime: "2h 40m", fcr: true, priority: "Critical" },
      { id: "CAS-4835", title: "Custom report failing", created: "02/28 2:20 PM", resolved: null, sla: "breached", responseTime: "3h 40m+", fcr: false, priority: "High" },
      { id: "CAS-4839", title: "SSO config change", created: "02/28 3:40 PM", resolved: "02/28 4:25 PM", sla: "met", responseTime: "45m", fcr: true, priority: "Normal" },
    ],
  },
  {
    name: "Bruce Banner", avatar: "BB", tier: "Tier 2",
    metrics: { casesCreated: 3, casesResolved: 3, slaCompliance: 100, fcr: 33, avgResponseTime: "3h 05m", csatScore: 4.5, phoneCalls: 2, phoneAnswered: 2, phoneAbandoned: 0, avgTalkTime: "12m 10s" },
    cases: [
      { id: "CAS-4844", title: "Server migration prep", created: "02/28 10:00 AM", resolved: "02/28 2:00 PM", sla: "met", responseTime: "4h 00m", fcr: false, priority: "High" },
      { id: "CAS-4845", title: "Network VLAN config", created: "02/28 11:30 AM", resolved: "02/28 1:45 PM", sla: "met", responseTime: "2h 15m", fcr: true, priority: "Normal" },
      { id: "CAS-4846", title: "Firewall rule update", created: "02/28 2:30 PM", resolved: "02/28 5:30 PM", sla: "met", responseTime: "3h 00m", fcr: false, priority: "Critical" },
    ],
  },
  {
    name: "Natasha Romanov", avatar: "NR", tier: "Tier 3",
    metrics: { casesCreated: 2, casesResolved: 1, slaCompliance: 50, fcr: 0, avgResponseTime: "5h 20m", csatScore: 4.1, phoneCalls: 1, phoneAnswered: 1, phoneAbandoned: 0, avgTalkTime: "15m 00s" },
    cases: [
      { id: "CAS-4847", title: "ERP module deployment", created: "02/28 9:00 AM", resolved: "02/28 4:30 PM", sla: "met", responseTime: "7h 30m", fcr: false, priority: "Critical" },
      { id: "CAS-4848", title: "Data migration rollback", created: "02/28 1:00 PM", resolved: null, sla: "breached", responseTime: "8h+", fcr: false, priority: "Critical" },
    ],
  },
  {
    name: "Tony Stark", avatar: "TS", tier: "Tier 3",
    metrics: { casesCreated: 3, casesResolved: 2, slaCompliance: 67, fcr: 0, avgResponseTime: "4h 45m", csatScore: 4.4, phoneCalls: 2, phoneAnswered: 1, phoneAbandoned: 1, avgTalkTime: "10m 30s" },
    cases: [
      { id: "CAS-4849", title: "Cloud infra scaling", created: "02/28 8:30 AM", resolved: "02/28 12:00 PM", sla: "met", responseTime: "3h 30m", fcr: false, priority: "High" },
      { id: "CAS-4850", title: "CI/CD pipeline broken", created: "02/28 11:00 AM", resolved: "02/28 5:00 PM", sla: "met", responseTime: "6h 00m", fcr: false, priority: "Critical" },
      { id: "CAS-4851", title: "Security audit remediation", created: "02/28 2:00 PM", resolved: null, sla: "breached", responseTime: "6h+", fcr: false, priority: "Critical" },
    ],
  },
];

const DD_METRIC_DEFS = [
  { key: "all", label: "All Metrics", icon: "📊", unit: "", color: DC.cyan },
  { key: "slaCompliance", label: "SLA Compliance", icon: "⏱", unit: "%", target: 90, color: DC.green },
  { key: "fcr", label: "First Contact Resolution", icon: "🎯", unit: "%", target: 70, color: DC.blue },
  { key: "casesCreated", label: "Cases Created", icon: "📥", unit: "", color: DC.cyan },
  { key: "casesResolved", label: "Cases Resolved", icon: "✅", unit: "", color: DC.green },
  { key: "avgResponseTime", label: "Avg Response", icon: "⚡", unit: "", color: DC.orange },
  { key: "csatScore", label: "CSAT Score", icon: "⭐", unit: "/5", target: 4.5, color: DC.gold },
  { key: "phoneCalls", label: "Phone Metrics", icon: "📞", unit: "", color: DC.purple },
];

const ddGetSlaBadge = s => s === "met" ? { label: "Met", bg: DC.greenGlow, color: DC.green, icon: "✅" } : s === "breached" ? { label: "Breached", bg: DC.redGlow, color: DC.red, icon: "❌" } : { label: "At Risk", bg: DC.orangeGlow, color: DC.orange, icon: "⚠️" };
const ddGetPriStyle = p => p === "Critical" ? { bg: "rgba(239,68,68,0.2)", color: DC.red } : p === "High" ? { bg: "rgba(245,158,11,0.15)", color: DC.orange } : { bg: "rgba(148,163,184,0.1)", color: DC.muted };
const ddFilterCases = (cases, key) => key === "casesResolved" ? cases.filter(c => c.resolved) : cases;

function DDHBar({ label, value, max, color, unit = "", target }) {
  const pct = max > 0 ? Math.min(100, (value / max) * 100) : 0;
  const tPct = target && max > 0 ? Math.min(100, (target / max) * 100) : null;
  return (
    <div style={{ marginBottom: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
        <span style={{ fontSize: 12, color: DC.text, fontWeight: 500 }}>{label}</span>
        <span style={{ fontSize: 12, fontFamily: "'JetBrains Mono','Space Mono', monospace", color, fontWeight: 700 }}>{value}{unit}</span>
      </div>
      <div style={{ height: 8, borderRadius: 4, background: "rgba(255,255,255,0.05)", position: "relative", overflow: "hidden" }}>
        <div style={{ height: "100%", borderRadius: 4, background: `linear-gradient(90deg, ${color}88, ${color})`, width: `${pct}%`, transition: "width 0.5s" }} />
        {tPct && <div style={{ position: "absolute", top: -2, bottom: -2, left: `${tPct}%`, width: 2, background: DC.white, opacity: 0.5, borderRadius: 1 }} />}
      </div>
    </div>
  );
}

function DDDonut({ data, size = 120 }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  const r = size / 2 - 8;
  const circ = 2 * Math.PI * r;
  let offset = 0;
  return (
    <div style={{ position: "relative", width: size, height: size }}>
      <svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>
        {data.map((d, i) => {
          const dash = total > 0 ? (d.value / total) * circ : 0;
          const o = offset; offset += dash;
          return <circle key={i} cx={size/2} cy={size/2} r={r} fill="none" stroke={d.color} strokeWidth={Math.round(size * 0.1)} strokeDasharray={`${dash} ${circ - dash}`} strokeDashoffset={-o} />;
        })}
      </svg>
      <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
        <span style={{ fontSize: Math.round(size * 0.17), fontWeight: 700, color: DC.white, fontFamily: "'JetBrains Mono','Space Mono', monospace" }}>{total}</span>
        <span style={{ fontSize: Math.round(size * 0.075), color: DC.dim, fontWeight: 600 }}>TOTAL</span>
      </div>
    </div>
  );
}

function DDField({ label, value, color, mono }) {
  return (
    <div>
      <div style={{ fontSize: 10, color: DC.dim, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: mono ? 15 : 13, fontWeight: mono ? 700 : 500, color, fontFamily: mono ? "'JetBrains Mono','Space Mono', monospace" : "inherit" }}>{value}</div>
    </div>
  );
}

function DDSLabel({ icon, label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
      <span style={{ fontSize: 14 }}>{icon}</span>
      <span style={{ fontSize: 13, fontWeight: 600, color: DC.white }}>{label}</span>
    </div>
  );
}

export default function DrillDownView({ onBack }) {
  const [selectedTier, setSelectedTier] = useState("All Tiers");
  const [selectedAgent, setSelectedAgent] = useState(null);
  const [selectedMetric, setSelectedMetric] = useState("all");
  const [hoveredMetric, setHoveredMetric] = useState(null);
  const [expandedCase, setExpandedCase] = useState(null);

  const today = new Date().toISOString().split("T")[0];
  const [fromDate, setFromDate] = useState(today);
  const [toDate, setToDate] = useState(today);
  const [fromTime, setFromTime] = useState("00:00");
  const [toTime, setToTime] = useState("23:59");
  const [activePreset, setActivePreset] = useState(null);

  const presets = [{ label: "7D", days: 7 }, { label: "14D", days: 14 }, { label: "30D", days: 30 }, { label: "90D", days: 90 }, { label: "YTD", days: null }];
  const applyPreset = (p, idx) => {
    setActivePreset(idx);
    const to = new Date(today);
    const from = p.days ? new Date(to.getTime() - (p.days - 1) * 86400000) : new Date(to.getFullYear() + "-01-01");
    setFromDate(from.toISOString().split("T")[0]); setToDate(to.toISOString().split("T")[0]);
    setFromTime("00:00"); setToTime("23:59");
  };
  const fD = d => new Date(d + "T12:00:00").toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" });
  const fT = t => { const [h, m] = t.split(":"); const hr = parseInt(h); return `${hr === 0 ? 12 : hr > 12 ? hr - 12 : hr}:${m} ${hr >= 12 ? "PM" : "AM"}`; };

  const visibleAgents = selectedTier === "All Tiers" ? DD_AGENTS : DD_AGENTS.filter(a => a.tier === selectedTier);
  const agent = selectedAgent !== null ? visibleAgents[selectedAgent] : null;
  const tierCounts = { "Tier 1": DD_AGENTS.filter(a => a.tier === "Tier 1").length, "Tier 2": DD_AGENTS.filter(a => a.tier === "Tier 2").length, "Tier 3": DD_AGENTS.filter(a => a.tier === "Tier 3").length };

  const agg = useMemo(() => {
    const src = agent ? [agent] : visibleAgents;
    const t = src.reduce((s, a) => s + a.metrics.casesCreated, 0);
    const r = src.reduce((s, a) => s + a.metrics.casesResolved, 0);
    const sla = t > 0 ? Math.round(src.reduce((s, a) => s + a.metrics.slaCompliance * a.metrics.casesCreated, 0) / t) : 0;
    const fcr = t > 0 ? Math.round(src.reduce((s, a) => s + a.metrics.fcr * a.metrics.casesCreated, 0) / t) : 0;
    const csat = src.length > 0 ? (src.reduce((s, a) => s + a.metrics.csatScore, 0) / src.length).toFixed(1) : 0;
    const calls = src.reduce((s, a) => s + a.metrics.phoneCalls, 0);
    const ans = src.reduce((s, a) => s + a.metrics.phoneAnswered, 0);
    const abd = src.reduce((s, a) => s + a.metrics.phoneAbandoned, 0);
    return { casesCreated: t, casesResolved: r, slaCompliance: sla, fcr, avgResponseTime: "—", csatScore: parseFloat(csat), phoneCalls: calls, phoneAnswered: ans, phoneAbandoned: abd, avgTalkTime: "—" };
  }, [agent, visibleAgents]);

  const metrics = agent ? agent.metrics : agg;
  const allCases = agent ? agent.cases : visibleAgents.flatMap(a => a.cases.map(c => ({ ...c, agent: a.name, agentTier: a.tier })));
  const filtered = ddFilterCases(allCases, selectedMetric);
  const mDef = DD_METRIC_DEFS.find(m => m.key === selectedMetric);
  const isAll = selectedMetric === "all";
  const isPhone = selectedMetric === "phoneCalls";
  const showAgent = !agent;
  const cols = showAgent ? "90px 120px 1fr 120px 90px 100px 80px" : "90px 1fr 120px 90px 100px 80px";
  const chartAgents = agent ? [agent] : visibleAgents;

  const slaDonut = [
    { label: "Met", value: allCases.filter(c => c.sla === "met").length, color: DC.green },
    { label: "At Risk", value: allCases.filter(c => c.sla === "at-risk").length, color: DC.orange },
    { label: "Breached", value: allCases.filter(c => c.sla === "breached").length, color: DC.red },
  ];
  const fcrDonut = [
    { label: "FCR", value: allCases.filter(c => c.fcr).length, color: DC.blue },
    { label: "Multi-touch", value: allCases.filter(c => !c.fcr).length, color: DC.dim },
  ];
  const priDonut = [
    { label: "Normal", value: allCases.filter(c => c.priority === "Normal").length, color: DC.muted },
    { label: "High", value: allCases.filter(c => c.priority === "High").length, color: DC.orange },
    { label: "Critical", value: allCases.filter(c => c.priority === "Critical").length, color: DC.red },
  ];
  const phoneDonut = [
    { label: "Answered", value: metrics.phoneAnswered || 0, color: DC.green },
    { label: "Abandoned", value: metrics.phoneAbandoned || 0, color: DC.red },
  ];

  const iS = { background: DC.surface, border: `1px solid ${DC.border}`, borderRadius: 8, padding: "8px 12px", color: DC.text, fontSize: 13, fontFamily: "'JetBrains Mono','Space Mono', monospace", outline: "none", cursor: "pointer", colorScheme: "dark" };

  return (
    <div style={{ background: DC.bg, minHeight: "100vh", fontFamily: "'DM Sans', 'Segoe UI', system-ui, sans-serif", color: DC.text, padding: "24px" }}>
      <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ marginBottom: 24, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
            <button onClick={onBack} style={{ background: DC.surface, border: `1px solid ${DC.border}`, borderRadius: 8, padding: "6px 12px", cursor: "pointer", color: DC.text, fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>← Back</button>
            <span style={{ fontSize: 22 }}>🔍</span>
            <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: DC.white, letterSpacing: "-0.02em" }}>Agent Drill-Down</h1>
            <span style={{ background: "rgba(59,130,246,0.15)", color: DC.blue, fontSize: 11, fontWeight: 600, padding: "3px 10px", borderRadius: 20 }}>PREVIEW</span>
          </div>
          <p style={{ margin: 0, color: DC.dim, fontSize: 13 }}>Tier → Team Members → Metrics Report → Date Range → Results</p>
        </div>
      </div>

      {/* 1. TIER */}
      <DDSLabel icon="📋" label="Tier" />
      <div style={{ display: "flex", gap: 6, marginBottom: 20, flexWrap: "wrap" }}>
        {DD_TIERS.map(t => {
          const act = selectedTier === t;
          const tc = t === "All Tiers" ? DC.white : DD_TIER_COLORS[t];
          const cnt = t === "All Tiers" ? DD_AGENTS.length : tierCounts[t];
          return (
            <button key={t} onClick={() => { setSelectedTier(t); setSelectedAgent(null); setExpandedCase(null); }}
              style={{ display: "flex", alignItems: "center", gap: 6, background: act ? DC.surface : "transparent", border: `1.5px solid ${act ? tc : DC.border}`, borderRadius: 10, padding: "8px 14px", cursor: "pointer", transition: "all 0.2s", boxShadow: act ? `0 0 16px ${tc}18` : "none" }}>
              {t !== "All Tiers" && <span style={{ width: 8, height: 8, borderRadius: "50%", background: act ? tc : DC.dim, boxShadow: act ? `0 0 6px ${tc}` : "none" }} />}
              <span style={{ fontSize: 13, fontWeight: act ? 700 : 500, color: act ? tc : DC.muted }}>{t}</span>
              <span style={{ fontSize: 10, fontWeight: 700, background: act ? `${tc}22` : "rgba(148,163,184,0.08)", color: act ? tc : DC.dim, padding: "1px 6px", borderRadius: 6, minWidth: 18, textAlign: "center" }}>{cnt}</span>
            </button>
          );
        })}
      </div>

      {/* 2. TEAM MEMBERS */}
      <DDSLabel icon="👥" label="Team Members" />
      <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <button onClick={() => { setSelectedAgent(null); setExpandedCase(null); }}
          style={{ display: "flex", alignItems: "center", gap: 10, background: selectedAgent === null ? DC.surface : DC.card, border: `1.5px solid ${selectedAgent === null ? DC.cyan : DC.border}`, borderRadius: 12, padding: "10px 16px", cursor: "pointer", boxShadow: selectedAgent === null ? `0 0 20px ${DC.cyan}18` : "none" }}>
          <div style={{ width: 36, height: 36, borderRadius: "50%", background: selectedAgent === null ? `linear-gradient(135deg, ${DC.cyan}, ${DC.blue})` : DC.surface, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>👥</div>
          <div style={{ textAlign: "left" }}>
            <div style={{ fontSize: 14, fontWeight: 600, color: selectedAgent === null ? DC.white : DC.text }}>{selectedTier === "All Tiers" ? "All Agents" : selectedTier}</div>
            <div style={{ fontSize: 11, color: DC.dim }}>{visibleAgents.length} agents · {visibleAgents.reduce((s, a) => s + a.metrics.casesCreated, 0)} cases</div>
          </div>
        </button>
        {visibleAgents.map((a, i) => {
          const tc = DD_TIER_COLORS[a.tier] || DC.blue;
          return (
            <button key={a.name} onClick={() => { setSelectedAgent(i); setExpandedCase(null); }}
              style={{ display: "flex", alignItems: "center", gap: 10, background: selectedAgent === i ? DC.surface : DC.card, border: `1.5px solid ${selectedAgent === i ? tc : DC.border}`, borderRadius: 12, padding: "10px 16px", cursor: "pointer", boxShadow: selectedAgent === i ? `0 0 20px ${tc}18` : "none" }}>
              <div style={{ width: 36, height: 36, borderRadius: "50%", background: selectedAgent === i ? `linear-gradient(135deg, ${tc}, ${DC.purple})` : DC.surface, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, color: DC.white }}>{a.avatar}</div>
              <div style={{ textAlign: "left" }}>
                <div style={{ fontSize: 14, fontWeight: 600, color: selectedAgent === i ? DC.white : DC.text }}>{a.name}</div>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 4, background: `${tc}22`, color: tc }}>{a.tier}</span>
                  <span style={{ fontSize: 11, color: DC.dim }}>{a.metrics.casesCreated} cases</span>
                </div>
              </div>
            </button>
          );
        })}
      </div>

      {/* 3. METRICS REPORT */}
      <DDSLabel icon="📊" label="Metrics Report" />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, marginBottom: 20 }}>
        {DD_METRIC_DEFS.map(m => {
          const isAllM = m.key === "all";
          const isPhoneM = m.key === "phoneCalls";
          const val = isAllM ? allCases.length : isPhoneM ? metrics.phoneCalls : metrics[m.key];
          const sel = selectedMetric === m.key;
          const hov = hoveredMetric === m.key;
          const ok = m.target ? (typeof val === "number" ? val >= m.target : false) : true;
          return (
            <button key={m.key} onClick={() => { setSelectedMetric(m.key); setExpandedCase(null); }}
              onMouseEnter={() => setHoveredMetric(m.key)} onMouseLeave={() => setHoveredMetric(null)}
              style={{ background: sel ? DC.surface : DC.card, border: `1.5px solid ${sel ? m.color : hov ? DC.borderLight : DC.border}`, borderRadius: 12, padding: "14px 16px", cursor: "pointer", transition: "all 0.2s", textAlign: "left", position: "relative", overflow: "hidden", boxShadow: sel ? `0 0 24px ${m.color}22` : "none" }}>
              {sel && <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 2, background: `linear-gradient(90deg, transparent, ${m.color}, transparent)` }} />}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
                <span style={{ fontSize: 16 }}>{m.icon}</span>
                {m.target && !isAllM && <span style={{ fontSize: 9, fontWeight: 600, padding: "2px 6px", borderRadius: 6, background: ok ? DC.greenGlow : DC.redGlow, color: ok ? DC.green : DC.red }}>{ok ? "ON TARGET" : "BELOW"}</span>}
                {isAllM && sel && <span style={{ fontSize: 9, fontWeight: 600, padding: "2px 6px", borderRadius: 6, background: `${DC.cyan}22`, color: DC.cyan }}>ALL</span>}
              </div>
              <div style={{ fontSize: isAllM || isPhoneM ? 20 : 24, fontWeight: 700, color: sel ? m.color : DC.white, fontFamily: "'JetBrains Mono','Space Mono', monospace", lineHeight: 1.1, marginBottom: 2 }}>
                {isAllM ? `${allCases.length} cases` : isPhoneM ? `${val} calls` : `${val}${m.unit}`}
              </div>
              <div style={{ fontSize: 11, color: DC.dim, fontWeight: 500 }}>{m.label}</div>
            </button>
          );
        })}
      </div>

      {/* 4. DATE RANGE */}
      <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "16px 20px", marginBottom: 24 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
          <span style={{ fontSize: 14 }}>📅</span>
          <span style={{ fontSize: 13, fontWeight: 600, color: DC.white }}>Date Range</span>
          <span style={{ flex: 1 }} />
          <span style={{ fontSize: 12, color: DC.muted }}>{fD(fromDate)} {fT(fromTime)} — {fD(toDate)} {fT(toTime)}</span>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
          {presets.map((p, idx) => (
            <button key={p.label} onClick={() => applyPreset(p, idx)}
              style={{ padding: "5px 12px", borderRadius: 6, border: `1px solid ${activePreset === idx ? DC.cyan : DC.border}`, background: activePreset === idx ? `${DC.cyan}18` : "transparent", color: activePreset === idx ? DC.cyan : DC.dim, fontSize: 11, fontWeight: 600, cursor: "pointer", fontFamily: "'JetBrains Mono','Space Mono', monospace" }}>{p.label}</button>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 8 }}>
          <label><div style={{ fontSize: 10, color: DC.dim, marginBottom: 3, fontWeight: 600 }}>FROM DATE</div><input type="date" value={fromDate} onChange={e => { setFromDate(e.target.value); setActivePreset(null); }} style={iS} /></label>
          <label><div style={{ fontSize: 10, color: DC.dim, marginBottom: 3, fontWeight: 600 }}>FROM TIME</div><input type="time" value={fromTime} onChange={e => setFromTime(e.target.value)} style={iS} /></label>
          <label><div style={{ fontSize: 10, color: DC.dim, marginBottom: 3, fontWeight: 600 }}>TO DATE</div><input type="date" value={toDate} onChange={e => { setToDate(e.target.value); setActivePreset(null); }} style={iS} /></label>
          <label><div style={{ fontSize: 10, color: DC.dim, marginBottom: 3, fontWeight: 600 }}>TO TIME</div><input type="time" value={toTime} onChange={e => setToTime(e.target.value)} style={iS} /></label>
        </div>
      </div>

      {/* 5. CHARTS */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 24 }}>
        {/* SLA by Agent */}
        <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: DC.white, marginBottom: 14 }}>⏱ SLA Compliance by Agent</div>
          {chartAgents.map(a => <DDHBar key={a.name} label={a.name} value={a.metrics.slaCompliance} max={100} color={a.metrics.slaCompliance >= 90 ? DC.green : a.metrics.slaCompliance >= 70 ? DC.orange : DC.red} unit="%" target={90} />)}
        </div>
        {/* FCR by Agent */}
        <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: DC.white, marginBottom: 14 }}>🎯 FCR by Agent</div>
          {chartAgents.map(a => <DDHBar key={a.name} label={a.name} value={a.metrics.fcr} max={100} color={a.metrics.fcr >= 70 ? DC.blue : a.metrics.fcr >= 50 ? DC.orange : DC.red} unit="%" target={70} />)}
        </div>
        {/* SLA Distribution */}
        <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: DC.white, marginBottom: 14 }}>📊 SLA Distribution</div>
          <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
            <DDDonut data={slaDonut} />
            <div style={{ flex: 1 }}>
              {slaDonut.map(d => (
                <div key={d.label} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <span style={{ width: 10, height: 10, borderRadius: 3, background: d.color, flexShrink: 0 }} />
                  <span style={{ fontSize: 12, color: DC.muted, flex: 1 }}>{d.label}</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: d.color, fontFamily: "'JetBrains Mono','Space Mono', monospace" }}>{d.value}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
        {/* Priority + FCR Breakdown */}
        <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: DC.white, marginBottom: 14 }}>📈 Case Breakdown</div>
          <div style={{ display: "flex", gap: 20, justifyContent: "center" }}>
            <div style={{ textAlign: "center" }}><DDDonut data={priDonut} size={100} /><div style={{ fontSize: 10, color: DC.dim, fontWeight: 600, marginTop: 6 }}>BY PRIORITY</div></div>
            <div style={{ textAlign: "center" }}><DDDonut data={fcrDonut} size={100} /><div style={{ fontSize: 10, color: DC.dim, fontWeight: 600, marginTop: 6 }}>FCR RATE</div></div>
          </div>
          <div style={{ display: "flex", gap: 12, justifyContent: "center", marginTop: 12, flexWrap: "wrap" }}>
            {[...priDonut, ...fcrDonut].map(d => (
              <div key={d.label} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <span style={{ width: 8, height: 8, borderRadius: 2, background: d.color }} />
                <span style={{ fontSize: 10, color: DC.muted }}>{d.label}: {d.value}</span>
              </div>
            ))}
          </div>
        </div>
        {/* Phone Calls by Agent */}
        <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: DC.white, marginBottom: 14, display: "flex", alignItems: "center", gap: 6 }}>📞 Phone Calls by Agent</div>
          {chartAgents.map(a => <DDHBar key={a.name} label={a.name} value={a.metrics.phoneCalls} max={Math.max(...chartAgents.map(x => x.metrics.phoneCalls), 1)} color={DC.purple} />)}
        </div>
        {/* Phone Distribution */}
        <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: DC.white, marginBottom: 14 }}>📞 Phone Distribution</div>
          <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
            <DDDonut data={phoneDonut} />
            <div style={{ flex: 1 }}>
              {phoneDonut.map(d => (
                <div key={d.label} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <span style={{ width: 10, height: 10, borderRadius: 3, background: d.color, flexShrink: 0 }} />
                  <span style={{ fontSize: 12, color: DC.muted, flex: 1 }}>{d.label}</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: d.color, fontFamily: "'JetBrains Mono','Space Mono', monospace" }}>{d.value}</span>
                </div>
              ))}
              <div style={{ borderTop: `1px solid ${DC.border}`, marginTop: 8, paddingTop: 8 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12 }}>
                  <span style={{ color: DC.muted }}>Answer Rate</span>
                  <span style={{ fontWeight: 700, fontFamily: "'JetBrains Mono','Space Mono', monospace", color: metrics.phoneCalls > 0 && (metrics.phoneAnswered / metrics.phoneCalls) >= 0.8 ? DC.green : DC.orange }}>
                    {metrics.phoneCalls > 0 ? Math.round((metrics.phoneAnswered / metrics.phoneCalls) * 100) : 0}%
                  </span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginTop: 4 }}>
                  <span style={{ color: DC.muted }}>Avg Talk Time</span>
                  <span style={{ fontWeight: 700, fontFamily: "'JetBrains Mono','Space Mono', monospace", color: DC.purple }}>{metrics.avgTalkTime}</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* 6. PHONE SUMMARY */}
      {isPhone && (
        <div style={{ background: DC.card, border: `1px solid ${DC.purple}44`, borderRadius: 14, padding: "20px", marginBottom: 16, display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16 }}>
          {[
            { label: "Total Calls", value: metrics.phoneCalls, icon: "📞", color: DC.purple },
            { label: "Answered", value: metrics.phoneAnswered, icon: "✅", color: DC.green },
            { label: "Abandoned", value: metrics.phoneAbandoned, icon: "📵", color: DC.red },
            { label: "Answer Rate", value: metrics.phoneCalls > 0 ? `${Math.round((metrics.phoneAnswered / metrics.phoneCalls) * 100)}%` : "N/A", icon: "📊", color: metrics.phoneCalls > 0 && (metrics.phoneAnswered / metrics.phoneCalls) >= 0.8 ? DC.green : DC.orange },
          ].map(s => (
            <div key={s.label} style={{ textAlign: "center" }}>
              <div style={{ fontSize: 14, marginBottom: 4 }}>{s.icon}</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: s.color, fontFamily: "'JetBrains Mono','Space Mono', monospace" }}>{s.value}</div>
              <div style={{ fontSize: 11, color: DC.dim, fontWeight: 500 }}>{s.label}</div>
            </div>
          ))}
        </div>
      )}

      {/* 7. CASE DETAIL TABLE */}
      <div style={{ background: DC.card, border: `1px solid ${DC.border}`, borderRadius: 14, overflow: "hidden" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 20px", background: DC.surface, borderBottom: `1px solid ${DC.border}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 15 }}>{mDef.icon}</span>
            <span style={{ fontSize: 14, fontWeight: 600, color: DC.white }}>{agent ? agent.name : selectedTier}</span>
            <span style={{ color: DC.dim }}>›</span>
            <span style={{ fontSize: 14, fontWeight: 600, color: mDef.color }}>{mDef.label}</span>
          </div>
          <span style={{ fontSize: 11, fontWeight: 600, background: `${mDef.color}18`, color: mDef.color, padding: "4px 10px", borderRadius: 8 }}>{filtered.length} cases</span>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: cols, padding: "10px 20px", borderBottom: `1px solid ${DC.border}`, fontSize: 10, fontWeight: 600, color: DC.dim, letterSpacing: "0.06em", textTransform: "uppercase" }}>
          <div>Case ID</div>{showAgent && <div>Agent</div>}<div>Title</div><div>Created</div><div>Priority</div>
          <div style={{ color: !isAll && selectedMetric === "slaCompliance" ? mDef.color : DC.dim }}>SLA</div>
          <div style={{ color: !isAll && selectedMetric === "fcr" ? mDef.color : DC.dim }}>FCR</div>
        </div>
        {filtered.map((c, i) => {
          const sla = ddGetSlaBadge(c.sla); const pri = ddGetPriStyle(c.priority); const exp = expandedCase === c.id;
          return (
            <div key={c.id}>
              <div onClick={() => setExpandedCase(exp ? null : c.id)}
                style={{ display: "grid", gridTemplateColumns: cols, padding: "12px 20px", borderBottom: `1px solid ${DC.border}`, cursor: "pointer", transition: "background 0.15s", background: exp ? "rgba(59,130,246,0.06)" : "transparent", alignItems: "center" }}
                onMouseEnter={e => e.currentTarget.style.background = "rgba(59,130,246,0.06)"} onMouseLeave={e => e.currentTarget.style.background = exp ? "rgba(59,130,246,0.06)" : "transparent"}>
                <div style={{ fontSize: 12, fontFamily: "'JetBrains Mono','Space Mono', monospace", color: DC.blue, fontWeight: 500 }}>{c.id}</div>
                {showAgent && (
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <span style={{ fontSize: 12, fontWeight: 600, color: DC.text }}>{c.agent}</span>
                    <span style={{ fontSize: 8, fontWeight: 700, padding: "1px 4px", borderRadius: 3, background: `${DD_TIER_COLORS[c.agentTier] || DC.blue}22`, color: DD_TIER_COLORS[c.agentTier] || DC.blue }}>{c.agentTier?.replace("Tier ", "T")}</span>
                  </div>
                )}
                <div style={{ fontSize: 13, fontWeight: 500, color: DC.text }}>{c.title}</div>
                <div style={{ fontSize: 12, color: DC.muted, fontFamily: "'JetBrains Mono','Space Mono', monospace" }}>{c.created}</div>
                <div><span style={{ fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 6, background: pri.bg, color: pri.color }}>{c.priority}</span></div>
                <div><span style={{ fontSize: 11, fontWeight: 600, padding: "3px 8px", borderRadius: 6, background: sla.bg, color: sla.color, display: "inline-flex", alignItems: "center", gap: 4 }}>{sla.icon} {sla.label}</span></div>
                <div><span style={{ fontSize: 11, fontWeight: 600, padding: "3px 8px", borderRadius: 6, background: c.fcr ? DC.greenGlow : "rgba(148,163,184,0.1)", color: c.fcr ? DC.green : DC.dim }}>{c.fcr ? "Yes" : "No"}</span></div>
              </div>
              {exp && (
                <div style={{ padding: "16px 20px 16px 110px", background: "rgba(59,130,246,0.04)", borderBottom: `1px solid ${DC.border}`, display: "grid", gridTemplateColumns: showAgent ? "1fr 1fr 1fr 1fr 1fr" : "1fr 1fr 1fr 1fr", gap: 16 }}>
                  {showAgent && <DDField label="Agent" value={c.agent} color={DD_TIER_COLORS[c.agentTier] || DC.text} />}
                  <DDField label="Response Time" value={c.responseTime} color={DC.orange} mono />
                  <DDField label="Created" value={c.created} color={DC.text} />
                  <DDField label="Resolved" value={c.resolved || "⏳ Open"} color={c.resolved ? DC.green : DC.orange} />
                  <DDField label="SLA Target" value="4 hours" color={DC.muted} />
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div style={{ marginTop: 16, textAlign: "center", fontSize: 12, color: DC.dim, fontStyle: "italic" }}>
        Click any row to expand · Tier → Team → Metric → Date Range → Results · All data from D365 OData
      </div>
    </div>
  );
}


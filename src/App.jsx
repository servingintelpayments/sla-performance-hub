import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend, PieChart, Pie, Cell,
  AreaChart, Area
} from "recharts";

/* ═══════════════════════════════════════════════════════════════════
   SLA PERFORMANCE HUB v6
   Sidebar Layout (v4) + Real Service Desk Data (v5)
   Data: Dynamics 365 (Cases/SLA/CSAT) + 8x8 (Phone Metrics)
   ═══════════════════════════════════════════════════════════════════ */

/* ─── REAL TIER DEFINITIONS (from D365 casetypecode) ─── */
const TIERS = {
  1: {
    code: 1, label: "Tier 1", name: "Service Desk", icon: "🔵",
    color: "#2196F3", colorLight: "#e8f4fd", colorDark: "#1565c0",
    desc: "Front-line support — password resets, basic troubleshooting, general inquiries",
    d365Filter: "casetypecode eq 1", dateField: "createdon",
    metrics: ["sla_compliance", "fcr_rate", "escalation_rate", "avg_resolution_time", "total_cases"],
  },
  2: {
    code: 2, label: "Tier 2", name: "Programming Team", icon: "🟠",
    color: "#FF9800", colorLight: "#fff3e0", colorDark: "#e65100",
    desc: "Intermediate support — complex issues, escalations from Tier 1, technical cases",
    d365Filter: "casetypecode eq 2", dateField: "escalatedon",
    metrics: ["sla_compliance", "escalation_rate", "total_cases", "resolved"],
  },
  3: {
    code: 3, label: "Tier 3", name: "Relationship Managers", icon: "🟣",
    color: "#9C27B0", colorLight: "#f3e5f5", colorDark: "#7b1fa2",
    desc: "Advanced support — critical escalations, system-level issues, VIP accounts",
    d365Filter: "casetypecode eq 3", dateField: "escalatedon",
    metrics: ["sla_compliance", "total_cases", "resolved"],
  },
};

/* ─── REAL SLA TARGETS ─── */
const TARGETS = {
  sla_compliance: { value: 90, unit: "%", compare: "gte", label: "90%" },
  fcr_rate: { value: 90, unit: "%", compare: "gte", label: "90-95%" },
  escalation_rate: { value: 10, unit: "%", compare: "lt", label: "<10%" },
  answer_rate: { value: 95, unit: "%", compare: "gte", label: ">95%" },
  avg_phone_aht: { value: 6, unit: " min", compare: "lte", label: "<6 min" },
  csat_score: { value: 4.0, unit: "/5", compare: "gte", label: "4.0+" },
  email_sla: { value: 90, unit: "%", compare: "gte", label: "90%" },
};

function checkTarget(metricKey, value) {
  const t = TARGETS[metricKey];
  if (!t || value === null || value === undefined || value === "N/A") return "na";
  const v = parseFloat(value);
  if (isNaN(v)) return "na";
  const { compare, value: target } = t;
  if (compare === "gte") return v >= target ? "met" : v >= target * 0.9 ? "warn" : "miss";
  if (compare === "gt") return v > target ? "met" : v > target * 0.9 ? "warn" : "miss";
  if (compare === "lte") return v <= target ? "met" : v <= target * 1.15 ? "warn" : "miss";
  if (compare === "lt") return v < target ? "met" : v < target * 1.15 ? "warn" : "miss";
  return "na";
}

/* ─── REAL D365 OData QUERY MAP ─── */
const D365_QUERIES = {
  Get_Tier_1_Cases: (s, e) => `incidents?$filter=casetypecode eq 1 and createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_2_Cases: (s, e) => `incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_3_Cases: (s, e) => `incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_SLA_Met_Cases: (s, e) => `incidents?$filter=casetypecode eq 1 and resolvebyslastatus eq 4 and createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_FCR_Cases: (s, e) => `incidents?$filter=casetypecode eq 1 and firstresponsesent eq true and createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_Escalated_Cases: (s, e) => `incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_2_Escalated: (s, e) => `incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_2_Resolved: (s, e) => `incidents?$filter=casetypecode eq 2 and statecode eq 1 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_2_SLA_Met: (s, e) => `incidents?$filter=casetypecode eq 2 and resolvebyslastatus eq 4 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_3_Resolved: (s, e) => `incidents?$filter=casetypecode eq 3 and statecode eq 1 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_Tier_3_SLA_Met: (s, e) => `incidents?$filter=casetypecode eq 3 and resolvebyslastatus eq 4 and escalatedon ge ${s}T00:00:00Z and escalatedon le ${e}T23:59:59Z&$count=true`,
  Get_All_Cases: (s, e) => `incidents?$filter=createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_Resolved_Cases: (s, e) => `incidents?$filter=statecode eq 1 and modifiedon ge ${s}T00:00:00Z and modifiedon le ${e}T23:59:59Z&$count=true`,
  Get_Email_Cases: (s, e) => `incidents?$filter=caseorigincode eq 2 and createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_Email_Cases_Responded: (s, e) => `incidents?$filter=caseorigincode eq 2 and firstresponsesent eq true and createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_Email_Cases_Resolved: (s, e) => `incidents?$filter=caseorigincode eq 2 and statecode eq 1 and createdon ge ${s}T00:00:00Z and createdon le ${e}T23:59:59Z&$count=true`,
  Get_CSAT_Responses: (s, e) => `incidents?$filter=cr7fe_new_csatresponsereceived eq true and modifiedon ge ${s}T00:00:00Z and modifiedon le ${e}T23:59:59Z&$count=true`,
  Get_CSAT_Scores: (s, e) => `incidents?$filter=cr7fe_new_csatresponsereceived eq true and modifiedon ge ${s}T00:00:00Z and modifiedon le ${e}T23:59:59Z&$select=cr7fe_new_csatscore`,
  Get_Queue_Members: () => `queues?$expand=queuemembership_association($select=systemuserid,fullname)&$select=name,queueid`,
};

/* ─── COLORS ─── */
const C = {
  primary: "#1B2A4A", primaryDark: "#152d4a", accent: "#E8653A", accentLight: "#F09A7A",
  green: "#4CAF50", greenLight: "#e8f5e9", red: "#f44336", redLight: "#ffebee",
  orange: "#FF9800", orangeLight: "#fff3e0", blue: "#2196F3", blueLight: "#e3f2fd",
  purple: "#9C27B0", purpleLight: "#f3e5f5", gold: "#FFC107", goldLight: "#fff8e1",
  yellow: "#E6B422", gray: "#9e9e9e", grayLight: "#f5f5f5",
  bg: "#F4F1EC", card: "#FFFFFF", border: "#E2DDD5",
  textDark: "#1B2A4A", textMid: "#5A6578", textLight: "#8B95A5",
  d365: "#0078D4", e8x8: "#FF6B35",
};
const PIE_COLORS = [C.accent, "#2D9D78", C.blue, C.yellow, C.purple, C.accentLight, "#A3E4C8"];

/* ─── DEMO TEAM MEMBERS ─── */
const DEMO_TEAM_MEMBERS = [
  { id: "rjones", name: "Ryan Jones", role: "Agent I", avatar: "RJ", tier: 1 },
  { id: "tbrown", name: "Tyler Brown", role: "Agent I", avatar: "TB", tier: 1 },
  { id: "lnguyen", name: "Lisa Nguyen", role: "Agent I", avatar: "LN", tier: 1 },
  { id: "agarcia", name: "Ana Garcia", role: "Agent II", avatar: "AG", tier: 2 },
  { id: "mchen", name: "Michael Chen", role: "Agent II", avatar: "MC", tier: 2 },
  { id: "spatil", name: "Sanya Patil", role: "Agent II", avatar: "SP", tier: 2 },
  { id: "jsmith", name: "Jordan Smith", role: "Senior Agent", avatar: "JS", tier: 3 },
  { id: "kwilson", name: "Keisha Wilson", role: "Senior Agent", avatar: "KW", tier: 3 },
  { id: "dkim", name: "David Kim", role: "Senior Agent", avatar: "DK", tier: 3 },
];

/* ─── DEMO DATA GENERATOR ─── */
function rng(seed) { return (n) => { const x = Math.sin(seed + n) * 10000; return x - Math.floor(x); }; }
function seedFrom(str) { return str.split("").reduce((a, c) => a + c.charCodeAt(0), 0); }

function generateDemoData(startDate, endDate, selectedMembers) {
  const s = seedFrom(startDate + endDate + (selectedMembers?.join(",") || ""));
  const r = rng(s);
  const days = Math.max(1, Math.ceil((new Date(endDate) - new Date(startDate)) / 86400000));
  const scale = Math.max(1, Math.round(days / 1));

  const t1Cases = Math.round(12 * scale + r(1) * 8 * scale);
  const t1SLAMet = Math.round(t1Cases * (0.82 + r(2) * 0.16));
  const t1FCR = Math.round(t1Cases * (0.78 + r(3) * 0.18));
  const t1Escalated = Math.round(t1Cases * (0.03 + r(4) * 0.12));
  const t2Cases = Math.round(3 * scale + r(5) * 4 * scale);
  const t2Resolved = Math.round(t2Cases * (0.6 + r(6) * 0.35));
  const t2SLAMet = Math.round(t2Resolved * (0.75 + r(7) * 0.2));
  const t2Escalated = Math.round(t2Cases * (0.02 + r(8) * 0.08));
  const t3Cases = Math.round(1 * scale + r(9) * 3 * scale);
  const t3Resolved = Math.round(t3Cases * (0.5 + r(10) * 0.4));
  const t3SLAMet = Math.round(t3Resolved * (0.7 + r(11) * 0.25));

  const totalCalls = Math.round(30 * scale + r(12) * 20 * scale);
  const answered = Math.round(totalCalls * (0.9 + r(13) * 0.09));
  const abandoned = totalCalls - answered;
  const avgAHT = +(4 + r(14) * 5).toFixed(1);

  const emailCases = Math.round(5 * scale + r(15) * 6 * scale);
  const emailResponded = Math.round(emailCases * (0.8 + r(16) * 0.18));
  const emailResolved = Math.round(emailCases * (0.6 + r(17) * 0.35));

  const csatResponses = Math.round(2 * scale + r(18) * 4 * scale);
  const csatAvg = csatResponses > 0 ? +(3.2 + r(19) * 1.6).toFixed(1) : 0;

  const allCases = t1Cases + t2Cases + t3Cases;
  const allResolved = t1SLAMet + t2Resolved + t3Resolved;
  const avgResTime = +(1.5 + r(20) * 6).toFixed(1);

  const timeline = Array.from({ length: Math.min(days + 1, 90) }, (_, i) => {
    const d = new Date(startDate); d.setDate(d.getDate() + i);
    return {
      date: d.toLocaleDateString("en-US", { month: "short", day: "numeric" }),
      t1Cases: Math.round(8 + rng(s + i)(1) * 12),
      t2Cases: Math.round(1 + rng(s + i)(2) * 5),
      t3Cases: Math.round(rng(s + i)(3) * 3),
      calls: Math.round(20 + rng(s + i)(4) * 25),
      sla: Math.round(70 + rng(s + i)(5) * 28),
      csat: +(3 + rng(s + i)(6) * 1.8).toFixed(1),
    };
  });

  return {
    tier1: { total: t1Cases, slaMet: t1SLAMet, slaCompliance: t1Cases ? Math.round(t1SLAMet / t1Cases * 100) : 0, fcrRate: t1Cases ? Math.round(t1FCR / t1Cases * 100) : 0, escalationRate: t1Cases ? Math.round(t1Escalated / t1Cases * 100) : 0, avgResolutionTime: `${avgResTime} hrs`, escalated: t1Escalated },
    tier2: { total: t2Cases, resolved: t2Resolved, slaMet: t2SLAMet, slaCompliance: t2Resolved ? Math.round(t2SLAMet / t2Resolved * 100) : "N/A", escalationRate: t2Cases ? Math.round(t2Escalated / t2Cases * 100) : "N/A", escalated: t2Escalated },
    tier3: { total: t3Cases, resolved: t3Resolved, slaMet: t3SLAMet, slaCompliance: t3Resolved ? Math.round(t3SLAMet / t3Resolved * 100) : "N/A" },
    phone: { totalCalls, answered, abandoned, answerRate: totalCalls ? Math.round(answered / totalCalls * 100) : 0, avgAHT },
    email: { total: emailCases, responded: emailResponded, resolved: emailResolved, slaCompliance: emailResolved > 0 ? 100 : "N/A" },
    csat: { responses: csatResponses, avgScore: csatAvg || "N/A" },
    overall: { created: allCases, resolved: allResolved, csatResponses, answeredCalls: answered, abandonedCalls: abandoned },
    timeline,
  };
}

/* ─── AUTH STORE ─── */
const Auth = {
  getUsers() { try { return JSON.parse(localStorage.getItem("sla_users") || "[]"); } catch { return []; } },
  register(u, p, name) {
    const users = this.getUsers();
    if (users.find((x) => x.u === u.toLowerCase())) return { ok: false, err: "Username taken" };
    users.push({ u: u.toLowerCase(), p: btoa(p), name, at: Date.now() });
    localStorage.setItem("sla_users", JSON.stringify(users));
    return { ok: true };
  },
  login(u, p) {
    const user = this.getUsers().find((x) => x.u === u.toLowerCase() && x.p === btoa(p));
    if (!user) return { ok: false, err: "Invalid credentials" };
    const session = { u: user.u, name: user.name, at: Date.now() };
    localStorage.setItem("sla_session", JSON.stringify(session));
    return { ok: true, session };
  },
  session() { try { return JSON.parse(localStorage.getItem("sla_session")); } catch { return null; } },
  logout() { localStorage.removeItem("sla_session"); },
};

/* ═══════════════════════════════════════════════════════
   API CLIENTS
   ═══════════════════════════════════════════════════════ */

class Dynamics365Client {
  constructor(config) {
    this.orgUrl = config.orgUrl?.replace(/\/$/, "") || "";
    this.accessToken = config.accessToken || config.token || "";
    this.connected = false;
    this.lastError = null;
  }
  headers() {
    return {
      "Authorization": `Bearer ${this.accessToken}`,
      "OData-MaxVersion": "4.0", "OData-Version": "4.0",
      "Accept": "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "Prefer": "odata.include-annotations=*",
    };
  }
  async testConnection() {
    try {
      const res = await fetch(`${this.orgUrl}/api/data/v9.2/WhoAmI`, { headers: this.headers() });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      this.connected = true; this.lastError = null;
      return { success: true };
    } catch (err) {
      this.connected = false; this.lastError = err.message;
      return { success: false, error: err.message };
    }
  }
}

class EightByEightClient {
  constructor(config) {
    this.baseUrl = config.baseUrl?.replace(/\/$/, "") || "";
    this.tenantId = config.tenantId || "";
    this.apiKey = config.apiKey || "";
    this.connected = false;
    this.lastError = null;
  }
  headers() {
    return { "Authorization": `Bearer ${this.apiKey}`, "Content-Type": "application/json", "8x8-tenant": this.tenantId };
  }
  async testConnection() {
    try {
      const res = await fetch(`${this.baseUrl}/analytics/v2/status`, { headers: this.headers() });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      this.connected = true; this.lastError = null;
      return { success: true };
    } catch (err) {
      this.connected = false; this.lastError = err.message;
      return { success: false, error: err.message };
    }
  }
}

/* ═══════════════════════════════════════════════════════
   UI COMPONENTS — STATUS & METRIC DISPLAY
   ═══════════════════════════════════════════════════════ */

function StatusBadge({ status, value, unit, targetLabel }) {
  const colors = { met: { bg: C.green, icon: "✅" }, warn: { bg: C.orange, icon: "⚠️" }, miss: { bg: C.red, icon: "🔴" }, na: { bg: C.gray, icon: "➖" } };
  const c = colors[status] || colors.na;
  const display = value === "N/A" ? "N/A" : `${value}${unit || ""}`;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
      <span style={{ background: c.bg, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{display}</span>
      <span style={{ fontSize: 14 }}>{c.icon}</span>
      {targetLabel && <span style={{ color: C.textLight, fontSize: 11 }}>target: {targetLabel}</span>}
    </span>
  );
}

function MetricRow({ label, value, unit, metricKey, targetLabel, bold, big, bigColor }) {
  const status = metricKey ? checkTarget(metricKey, value) : null;
  return (
    <tr>
      <td style={{ color: "#333", padding: "8px 0", fontWeight: bold ? 700 : 400, fontSize: 14 }}>{label}</td>
      <td style={{ textAlign: "right", padding: "8px 0" }}>
        {big ? (
          <span style={{ fontWeight: 700, fontSize: 22, color: bigColor || C.textDark, fontFamily: "'Space Mono', monospace" }}>{value}</span>
        ) : status ? (
          <StatusBadge status={status} value={value} unit={unit} targetLabel={targetLabel} />
        ) : (
          <span style={{ background: C.blue, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{value}{unit || ""}</span>
        )}
      </td>
    </tr>
  );
}

function CTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: C.primary, color: "#fff", padding: "10px 14px", borderRadius: 10, fontSize: 12, fontFamily: "'DM Sans', sans-serif", boxShadow: "0 8px 24px rgba(0,0,0,0.2)" }}>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      {payload.map((p, i) => <div key={i} style={{ color: p.color || "#fff" }}>{p.name}: {p.value}</div>)}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   REPORT SECTIONS — Tier / Phone / Email / CSAT / Overall
   ═══════════════════════════════════════════════════════ */

function TierSection({ tier, data }) {
  const t = TIERS[tier]; if (!t) return null;
  const d = data[`tier${tier}`];
  return (
    <div style={{ background: t.colorLight, padding: "24px 28px", borderLeft: `5px solid ${t.color}`, borderRadius: "0 14px 14px 0", marginBottom: 4 }}>
      <h3 style={{ margin: "0 0 14px", color: t.colorDark, fontSize: 16, fontWeight: 700, textAlign: "center", letterSpacing: 0.5 }}>
        {t.icon} {t.label.toUpperCase()} — {t.name.toUpperCase()}
      </h3>
      <table cellPadding="0" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <tbody>
          {t.metrics.includes("sla_compliance") && <MetricRow label="SLA Compliance" value={d.slaCompliance} unit="%" metricKey="sla_compliance" targetLabel={TARGETS.sla_compliance.label} />}
          {t.metrics.includes("fcr_rate") && <MetricRow label="FCR Rate" value={d.fcrRate} unit="%" metricKey="fcr_rate" targetLabel={TARGETS.fcr_rate.label} />}
          {t.metrics.includes("escalation_rate") && <MetricRow label="Escalation Rate" value={d.escalationRate} unit="%" metricKey="escalation_rate" targetLabel={TARGETS.escalation_rate.label} />}
          {t.metrics.includes("avg_resolution_time") && (
            <tr>
              <td style={{ color: "#333", padding: "8px 0", fontSize: 14 }}>Avg Case Resolution Time</td>
              <td style={{ textAlign: "right", padding: "8px 0" }}>
                <span style={{ background: C.blue, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{d.avgResolutionTime}</span>
                <span style={{ marginLeft: 4 }}>⏱️</span>
              </td>
            </tr>
          )}
          {t.metrics.includes("total_cases") && <MetricRow label="Total Cases" value={d.total} bold big bigColor={t.colorDark} />}
          {t.metrics.includes("resolved") && (
            <tr>
              <td style={{ color: "#333", padding: "8px 0", fontSize: 14 }}>Resolved</td>
              <td style={{ textAlign: "right", padding: "8px 0" }}>
                <span style={{ background: C.gray, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{d.resolved}</span>
                <span style={{ marginLeft: 4 }}>📋</span>
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function PhoneSection({ data }) {
  const d = data.phone;
  return (
    <div style={{ background: C.greenLight, padding: "24px 28px", borderLeft: `5px solid ${C.green}`, borderRadius: "0 14px 14px 0", marginBottom: 4 }}>
      <h3 style={{ margin: "0 0 14px", color: "#2e7d32", fontSize: 16, fontWeight: 700, textAlign: "center" }}>📞 PHONE METRICS</h3>
      <table cellPadding="0" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <tbody>
          <MetricRow label="Total Calls" value={d.totalCalls} bold big bigColor="#2e7d32" />
          <tr>
            <td style={{ color: "#333", padding: "8px 0", fontSize: 14 }}>Answered Calls</td>
            <td style={{ textAlign: "right" }}><span style={{ background: C.green, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{d.answered}</span> ✅</td>
          </tr>
          <tr>
            <td style={{ color: "#333", padding: "8px 0", fontSize: 14 }}>Abandoned Calls</td>
            <td style={{ textAlign: "right" }}>
              <span style={{ background: d.abandoned === 0 ? C.green : C.red, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{d.abandoned}</span>
              {d.abandoned === 0 ? " ✅" : " 🔴"}
            </td>
          </tr>
          <MetricRow label="Answer Rate" value={d.answerRate} unit="%" metricKey="answer_rate" targetLabel={TARGETS.answer_rate.label} />
          <MetricRow label="Avg Phone AHT" value={d.avgAHT} unit=" min" metricKey="avg_phone_aht" targetLabel={TARGETS.avg_phone_aht.label} />
        </tbody>
      </table>
    </div>
  );
}

function EmailSection({ data }) {
  const d = data.email;
  return (
    <div style={{ background: C.blueLight, padding: "24px 28px", borderLeft: `5px solid #1976D2`, borderRadius: "0 14px 14px 0", marginBottom: 4 }}>
      <h3 style={{ margin: "0 0 14px", color: "#1565c0", fontSize: 16, fontWeight: 700, textAlign: "center" }}>📧 EMAIL METRICS</h3>
      <table cellPadding="0" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <tbody>
          <MetricRow label="Total Email Cases" value={d.total} bold big bigColor="#1565c0" />
          <tr>
            <td style={{ color: "#333", padding: "8px 0", fontSize: 14 }}>Responded</td>
            <td style={{ textAlign: "right" }}><span style={{ background: C.blue, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{d.responded}</span> 💬</td>
          </tr>
          <tr>
            <td style={{ color: "#333", padding: "8px 0", fontSize: 14 }}>Resolved</td>
            <td style={{ textAlign: "right" }}><span style={{ background: C.green, color: "#fff", padding: "3px 12px", borderRadius: 14, fontWeight: 700, fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{d.resolved}</span> ✅</td>
          </tr>
          <MetricRow label="SLA Compliance" value={d.slaCompliance} unit="%" metricKey="email_sla" targetLabel={TARGETS.email_sla.label} />
        </tbody>
      </table>
    </div>
  );
}

function CSATSection({ data }) {
  const d = data.csat;
  return (
    <div style={{ background: C.goldLight, padding: "24px 28px", borderLeft: `5px solid ${C.gold}`, borderRadius: "0 14px 14px 0", marginBottom: 4 }}>
      <h3 style={{ margin: "0 0 14px", color: "#f57f17", fontSize: 16, fontWeight: 700, textAlign: "center" }}>⭐ CSAT METRICS</h3>
      <table cellPadding="0" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <tbody>
          <MetricRow label="Total Responses" value={d.responses} bold big bigColor="#f57f17" />
          <MetricRow label="Avg Score" value={d.avgScore} unit="/5" metricKey="csat_score" targetLabel={TARGETS.csat_score.label} />
        </tbody>
      </table>
    </div>
  );
}

function OverallSummary({ data }) {
  const d = data.overall;
  const items = [
    { label: "Cases Created", value: d.created, color: "#4FC3F7" },
    { label: "Cases Resolved", value: d.resolved, color: "#81C784" },
    { label: "CSAT Responses", value: d.csatResponses, color: "#FFB74D" },
    { label: "Answered Calls", value: d.answeredCalls, color: "#81C784" },
    { label: "Abandoned", value: d.abandonedCalls, color: "#f44336" },
  ];
  return (
    <div style={{ background: C.primary, padding: "24px 28px", borderRadius: 14, marginBottom: 4 }}>
      <h3 style={{ margin: "0 0 20px", color: "#fff", fontSize: 16, fontWeight: 700, textAlign: "center" }}>📈 OVERALL SUMMARY</h3>
      <div style={{ display: "flex", justifyContent: "space-around", flexWrap: "wrap", gap: 12 }}>
        {items.map((it) => (
          <div key={it.label} style={{ textAlign: "center" }}>
            <div style={{ fontSize: 32, fontWeight: 700, color: it.color, fontFamily: "'Space Mono', monospace" }}>{it.value}</div>
            <div style={{ fontSize: 11, color: "#a8c6df", marginTop: 2 }}>{it.label}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Definitions() {
  const defs = [
    ["SLA Compliance", "Percentage of resolved cases meeting resolution time targets based on priority level"],
    ["FCR Rate", "First Contact Resolution — cases resolved without escalation or follow-up"],
    ["Escalation Rate", "Percentage of Tier 1 cases escalated to Tier 2 or Tier 3"],
    ["Avg Resolution Time", "Mean time from case creation to resolution for closed cases"],
    ["Answer Rate", "Percentage of calls answered vs. total calls"],
    ["AHT", "Average Handle Time — mean duration of phone calls"],
    ["CSAT Score", "Customer Satisfaction rating (1-5 scale). Target: 4.0+"],
  ];
  return (
    <div style={{ background: C.grayLight, padding: "24px 28px", borderTop: `1px solid ${C.border}`, borderRadius: "0 0 14px 14px" }}>
      <h4 style={{ margin: "0 0 14px", fontSize: 13, color: "#555", fontWeight: 700 }}>📝 DEFINITIONS & METHODOLOGY</h4>
      <table cellPadding="3" cellSpacing="0" style={{ width: "100%", fontSize: 12, color: "#666" }}>
        <tbody>
          {defs.map(([term, def]) => (
            <tr key={term}><td style={{ fontWeight: 700, width: 150, verticalAlign: "top", padding: "4px 0" }}>{term}</td><td style={{ padding: "4px 0" }}>{def}</td></tr>
          ))}
        </tbody>
      </table>
      <div style={{ borderTop: `1px solid ${C.border}`, marginTop: 14, paddingTop: 12, fontSize: 11, color: "#888", lineHeight: 1.8 }}>
        <div>✅ Target met &nbsp;|&nbsp; ⚠️ Approaching target &nbsp;|&nbsp; 🔴 Below target &nbsp;|&nbsp; ➖ N/A (no data)</div>
        <div>📊 <strong>Data Sources:</strong> Microsoft Dynamics 365 Customer Service (Cases) &nbsp;|&nbsp; 8x8 (Phone Metrics)</div>
        <div>⏰ <strong>Reporting Period:</strong> Selected date range</div>
        <div>📧 <strong>Note:</strong> CSAT responses may reflect cases resolved on previous days</div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   CHARTS PANEL
   ═══════════════════════════════════════════════════════ */
function ChartsPanel({ data }) {
  const tl = data.timeline;
  if (!tl || tl.length < 2) return null;
  const interval = Math.max(0, Math.floor(tl.length / 8));
  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginTop: 20 }}>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: `1px solid ${C.border}` }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>📊 Daily Cases by Tier</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={tl}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} />
            <YAxis fontSize={10} tick={{ fill: C.textLight }} />
            <Tooltip content={<CTooltip />} />
            <Legend iconType="circle" iconSize={7} formatter={(v) => <span style={{ fontSize: 10, color: C.textMid }}>{v}</span>} />
            <Bar dataKey="t1Cases" name="Tier 1" fill={TIERS[1].color} radius={[3,3,0,0]} barSize={14} />
            <Bar dataKey="t2Cases" name="Tier 2" fill={TIERS[2].color} radius={[3,3,0,0]} barSize={14} />
            <Bar dataKey="t3Cases" name="Tier 3" fill={TIERS[3].color} radius={[3,3,0,0]} barSize={14} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: `1px solid ${C.border}` }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>📈 SLA Compliance Trend</div>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={tl}>
            <defs><linearGradient id="slaG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.green} stopOpacity={0.3} /><stop offset="100%" stopColor={C.green} stopOpacity={0.02} /></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} />
            <YAxis fontSize={10} tick={{ fill: C.textLight }} domain={[50, 100]} />
            <Tooltip content={<CTooltip />} />
            <Area type="monotone" dataKey="sla" name="SLA %" stroke={C.green} fill="url(#slaG)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: `1px solid ${C.border}` }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>📞 Daily Call Volume</div>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={tl}>
            <defs><linearGradient id="callG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.blue} stopOpacity={0.3} /><stop offset="100%" stopColor={C.blue} stopOpacity={0.02} /></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} />
            <YAxis fontSize={10} tick={{ fill: C.textLight }} />
            <Tooltip content={<CTooltip />} />
            <Area type="monotone" dataKey="calls" name="Calls" stroke={C.blue} fill="url(#callG)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: `1px solid ${C.border}` }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>⭐ CSAT Score Trend</div>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={tl}>
            <defs><linearGradient id="csatG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.gold} stopOpacity={0.3} /><stop offset="100%" stopColor={C.gold} stopOpacity={0.02} /></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} />
            <YAxis fontSize={10} tick={{ fill: C.textLight }} domain={[1, 5]} />
            <Tooltip content={<CTooltip />} />
            <Area type="monotone" dataKey="csat" name="CSAT" stroke={C.gold} fill="url(#csatG)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   SIDEBAR COMPONENTS
   ═══════════════════════════════════════════════════════ */

function MultiMemberSelect({ selected, onChange, members }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef(null);
  const inputRef = useRef(null);
  useEffect(() => { const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }; document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h); }, []);
  useEffect(() => { if (open && inputRef.current) inputRef.current.focus(); }, [open]);

  const toggle = (id) => {
    if (id === "__all") { onChange(selected.length === members.length ? [] : members.map((m) => m.id)); }
    else { onChange(selected.includes(id) ? selected.filter((s) => s !== id) : [...selected, id]); }
  };

  const filtered = members.filter((m) => m.name.toLowerCase().includes(search.toLowerCase()));
  const allSelected = selected.length === members.length;
  const displayText = selected.length === 0 ? "Select team members..." : selected.length === members.length ? "All Team Members" : selected.length <= 2 ? selected.map((id) => members.find((m) => m.id === id)?.name).join(", ") : `${selected.length} members selected`;

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button onClick={() => { setOpen(!open); setSearch(""); }} style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${open ? C.accent : C.border}`, fontSize: 13, fontFamily: "'DM Sans', sans-serif", background: C.bg, color: selected.length ? C.textDark : C.textLight, cursor: "pointer", outline: "none", display: "flex", alignItems: "center", justifyContent: "space-between", textAlign: "left" }}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 500 }}>{displayText}</span>
        <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
          {selected.length > 0 && <span style={{ background: C.accent, color: "#fff", borderRadius: 10, padding: "1px 7px", fontSize: 10, fontWeight: 700, fontFamily: "'Space Mono', monospace" }}>{selected.length}</span>}
          <span style={{ transform: open ? "rotate(180deg)" : "rotate(0deg)", transition: "transform 0.2s", fontSize: 10, color: C.textLight }}>▼</span>
        </span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 6px)", left: 0, right: 0, background: C.card, borderRadius: 12, border: `1.5px solid ${C.border}`, boxShadow: "0 12px 40px rgba(27,42,74,0.15)", zIndex: 999, maxHeight: 340, display: "flex", flexDirection: "column", animation: "fadeIn 0.15s ease" }}>
          <div style={{ padding: "10px 12px", borderBottom: `1px solid ${C.border}` }}>
            <input ref={inputRef} type="text" placeholder="Search members..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ width: "100%", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12, fontFamily: "'DM Sans', sans-serif", outline: "none", background: C.bg, boxSizing: "border-box" }} />
          </div>
          <div style={{ overflowY: "auto", padding: "4px 0" }}>
            <button onClick={() => toggle("__all")} style={{ width: "100%", padding: "10px 14px", border: "none", background: allSelected ? C.primary + "0A" : "transparent", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, fontSize: 13, fontFamily: "'DM Sans', sans-serif", color: C.textDark, fontWeight: 600 }} onMouseEnter={(e) => e.currentTarget.style.background = C.bg} onMouseLeave={(e) => e.currentTarget.style.background = allSelected ? C.primary + "0A" : "transparent"}>
              <span style={{ width: 18, height: 18, borderRadius: 4, display: "flex", alignItems: "center", justifyContent: "center", border: allSelected ? `2px solid ${C.accent}` : `2px solid ${C.border}`, background: allSelected ? C.accent : "transparent", color: "#fff", fontSize: 11, fontWeight: 700 }}>{allSelected ? "✓" : ""}</span>
              Select All ({members.length})
            </button>
            <div style={{ height: 1, background: C.border, margin: "2px 14px" }} />
            {filtered.map((m) => { const checked = selected.includes(m.id); const idx = members.indexOf(m); const t = TIERS[m.tier]; return (
              <button key={m.id} onClick={() => toggle(m.id)} style={{ width: "100%", padding: "9px 14px", border: "none", background: checked ? C.primary + "08" : "transparent", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, fontSize: 13, fontFamily: "'DM Sans', sans-serif", color: C.textDark }} onMouseEnter={(e) => e.currentTarget.style.background = C.bg} onMouseLeave={(e) => e.currentTarget.style.background = checked ? C.primary + "08" : "transparent"}>
                <span style={{ width: 18, height: 18, borderRadius: 4, display: "flex", alignItems: "center", justifyContent: "center", border: checked ? `2px solid ${C.accent}` : `2px solid ${C.border}`, background: checked ? C.accent : "transparent", color: "#fff", fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{checked ? "✓" : ""}</span>
                <div style={{ width: 28, height: 28, borderRadius: 7, background: `linear-gradient(135deg, ${PIE_COLORS[idx % PIE_COLORS.length]}, ${PIE_COLORS[(idx + 2) % PIE_COLORS.length]})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700, color: "#fff", flexShrink: 0 }}>{m.avatar}</div>
                <div style={{ textAlign: "left" }}>
                  <div style={{ fontWeight: 500, display: "flex", alignItems: "center", gap: 5 }}>{m.name}<span style={{ fontSize: 8, fontWeight: 700, padding: "1px 5px", borderRadius: 3, background: (t?.color || C.blue) + "22", color: t?.color || C.blue }}>T{m.tier}</span></div>
                  <div style={{ fontSize: 10, color: C.textLight }}>{m.role}</div>
                </div>
              </button>
            ); })}
          </div>
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   CONNECTION BAR
   ═══════════════════════════════════════════════════════ */
function ConnectionBar({ config, onOpenSettings }) {
  const d365Ok = config.d365?.orgUrl && config.d365?.token;
  const e8x8Ok = config.e8x8?.baseUrl && config.e8x8?.apiKey;
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 28px", background: C.card, borderBottom: `1px solid ${C.border}`, fontSize: 11 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
        <span style={{ fontWeight: 600, color: C.textLight, fontSize: 10, textTransform: "uppercase", letterSpacing: 1 }}>Data Sources</span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
          <span style={{ width: 6, height: 6, borderRadius: "50%", background: d365Ok ? C.green : C.accent }} />
          <span style={{ fontWeight: 600, color: d365Ok ? C.green : C.accent }}>D365</span>
          <span style={{ color: C.textLight }}>{d365Ok ? config.d365.orgUrl.replace("https://", "").split(".")[0] : "Not configured"}</span>
        </span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
          <span style={{ width: 6, height: 6, borderRadius: "50%", background: e8x8Ok ? C.green : C.accent }} />
          <span style={{ fontWeight: 600, color: e8x8Ok ? C.green : C.accent }}>8x8</span>
          <span style={{ color: C.textLight }}>{e8x8Ok ? `Tenant: ${config.e8x8.tenantId}` : "Not configured"}</span>
        </span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
          <span style={{ width: 6, height: 6, borderRadius: "50%", background: config.live ? C.green : C.blue }} />
          <span style={{ fontWeight: 600, color: config.live ? C.green : C.blue }}>{config.live ? "Live" : "Demo"}</span>
        </span>
      </div>
      <button onClick={onOpenSettings} style={{ background: "none", border: "none", fontSize: 11, fontWeight: 600, color: C.primary, cursor: "pointer", textDecoration: "underline" }}>⚙️ Configure</button>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   SETTINGS MODAL
   ═══════════════════════════════════════════════════════ */
function SettingsModal({ show, onClose, config, onSave }) {
  const [local, setLocal] = useState(config);
  const [d365Status, setD365Status] = useState(null);
  const [e8x8Status, setE8x8Status] = useState(null);
  const [testing, setTesting] = useState(null);
  useEffect(() => { setLocal(config); }, [config]);
  if (!show) return null;
  const upd = (sec, key, val) => setLocal((p) => ({ ...p, [sec]: { ...p[sec], [key]: val } }));
  const iS = { width: "100%", padding: "10px 12px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 12, fontFamily: "'DM Sans',sans-serif", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" };
  const lS = { fontSize: 11, fontWeight: 600, color: C.textDark, marginBottom: 4, display: "block" };

  const testD365 = async () => { setTesting("d365"); const client = new Dynamics365Client(local.d365); const result = await client.testConnection(); setD365Status(result); setTesting(null); };
  const test8x8 = async () => { setTesting("8x8"); const client = new EightByEightClient(local.e8x8); const result = await client.testConnection(); setE8x8Status(result); setTesting(null); };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.55)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background: C.card, borderRadius: 20, width: 660, maxHeight: "90vh", overflow: "auto", boxShadow: "0 24px 80px rgba(0,0,0,0.25)" }}>
        <div style={{ padding: "24px 28px 18px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div><h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textDark }}>⚙️ Data Source Configuration</h2><p style={{ margin: "4px 0 0", fontSize: 12, color: C.textMid }}>Dynamics 365 + 8x8 Analytics connections</p></div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: C.textLight }}>✕</button>
        </div>
        <div style={{ padding: "24px 28px" }}>
          {/* D365 */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <div style={{ width: 32, height: 32, borderRadius: 8, background: C.d365, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 14 }}>D</div>
            <div><div style={{ fontSize: 14, fontWeight: 700, color: C.textDark }}>Microsoft Dynamics 365</div><div style={{ fontSize: 10, color: C.textMid }}>Web API v9.2 — Cases, SLA, CSAT, Queues</div></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <label><span style={lS}>Organization URL</span><input value={local.d365?.orgUrl || ""} onChange={(e) => upd("d365", "orgUrl", e.target.value)} placeholder="https://yourorg.crm.dynamics.com" style={iS} /></label>
            <label><span style={lS}>API Version</span><input value={local.d365?.apiVersion || "v9.2"} onChange={(e) => upd("d365", "apiVersion", e.target.value)} style={iS} /></label>
          </div>
          <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
            <label style={{ flex: 1 }}><span style={lS}>Access Token</span><input type="password" value={local.d365?.token || ""} onChange={(e) => upd("d365", "token", e.target.value)} placeholder="Bearer token..." style={iS} /></label>
            <div style={{ alignSelf: "flex-end" }}><button onClick={testD365} disabled={testing === "d365"} style={{ padding: "10px 16px", borderRadius: 8, border: `1px solid ${C.d365}`, background: "transparent", color: C.d365, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{testing === "d365" ? "Testing..." : "Test"}</button></div>
          </div>
          {d365Status && <div style={{ marginBottom: 16, padding: "8px 12px", borderRadius: 8, fontSize: 11, background: d365Status.success ? C.greenLight : C.redLight, color: d365Status.success ? C.green : C.red }}>{d365Status.success ? "✅ Connected!" : `❌ ${d365Status.error}`}</div>}

          {/* 8x8 */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <div style={{ width: 32, height: 32, borderRadius: 8, background: C.e8x8, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 14 }}>8</div>
            <div><div style={{ fontSize: 14, fontWeight: 700, color: C.textDark }}>8x8 Analytics</div><div style={{ fontSize: 10, color: C.textMid }}>Contact Center Analytics — Calls, AHT, Queue Stats</div></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <label><span style={lS}>API Base URL</span><input value={local.e8x8?.baseUrl || ""} onChange={(e) => upd("e8x8", "baseUrl", e.target.value)} placeholder="https://api.8x8.com" style={iS} /></label>
            <label><span style={lS}>Tenant ID</span><input value={local.e8x8?.tenantId || ""} onChange={(e) => upd("e8x8", "tenantId", e.target.value)} style={iS} /></label>
          </div>
          <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
            <label style={{ flex: 1 }}><span style={lS}>API Key</span><input type="password" value={local.e8x8?.apiKey || ""} onChange={(e) => upd("e8x8", "apiKey", e.target.value)} style={iS} /></label>
            <div style={{ alignSelf: "flex-end" }}><button onClick={test8x8} disabled={testing === "8x8"} style={{ padding: "10px 16px", borderRadius: 8, border: `1px solid ${C.e8x8}`, background: "transparent", color: C.e8x8, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{testing === "8x8" ? "Testing..." : "Test"}</button></div>
          </div>
          {e8x8Status && <div style={{ marginBottom: 16, padding: "8px 12px", borderRadius: 8, fontSize: 11, background: e8x8Status.success ? C.greenLight : C.redLight, color: e8x8Status.success ? C.green : C.red }}>{e8x8Status.success ? "✅ Connected!" : `❌ ${e8x8Status.error}`}</div>}

          {/* Mode Toggle */}
          <div style={{ background: C.bg, borderRadius: 10, padding: "14px 18px", border: `1px solid ${C.border}` }}>
            <label style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }}>
              <div onClick={() => setLocal((p) => ({ ...p, live: !p.live }))} style={{ width: 44, height: 24, borderRadius: 12, padding: 2, background: local.live ? C.green : C.border, transition: "background 0.2s", cursor: "pointer" }}>
                <div style={{ width: 20, height: 20, borderRadius: 10, background: "#fff", transform: local.live ? "translateX(20px)" : "translateX(0)", transition: "transform 0.2s", boxShadow: "0 1px 4px rgba(0,0,0,0.15)" }} />
              </div>
              <div><div style={{ fontSize: 13, fontWeight: 600, color: C.textDark }}>{local.live ? "🟢 Live Data Mode" : "🔵 Demo Data Mode"}</div><div style={{ fontSize: 11, color: C.textMid }}>{local.live ? "Pulling from D365 + 8x8 APIs" : "Simulated demo data"}</div></div>
            </label>
          </div>
        </div>
        <div style={{ padding: "16px 28px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button onClick={onClose} style={{ padding: "10px 22px", borderRadius: 10, border: `1px solid ${C.border}`, background: "transparent", fontSize: 13, fontWeight: 600, color: C.textMid, cursor: "pointer" }}>Cancel</button>
          <button onClick={() => { onSave(local); onClose(); }} style={{ padding: "10px 22px", borderRadius: 10, border: "none", background: C.primary, fontSize: 13, fontWeight: 600, color: "#fff", cursor: "pointer" }}>Save</button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   LOGIN PAGE
   ═══════════════════════════════════════════════════════ */
function LoginPage({ onLogin }) {
  const [mode, setMode] = useState("login");
  const [u, setU] = useState(""); const [p, setP] = useState(""); const [name, setName] = useState(""); const [cp, setCp] = useState("");
  const [err, setErr] = useState(""); const [ok, setOk] = useState(""); const [loading, setLoading] = useState(false);
  const [showPw, setShowPw] = useState(false); const [ready, setReady] = useState(false);
  useEffect(() => { const s = Auth.session(); if (s) onLogin?.(s); setTimeout(() => setReady(true), 100); }, []);

  const doLogin = async () => { setErr(""); if (!u || !p) return setErr("Fill in all fields"); setLoading(true); await new Promise(r => setTimeout(r, 600)); const res = Auth.login(u, p); setLoading(false); if (res.ok) { setOk("Welcome! Redirecting..."); setTimeout(() => onLogin?.(res.session), 500); } else setErr(res.err); };
  const doReg = async () => { setErr(""); if (!u || !p || !name) return setErr("Fill in all fields"); if (p.length < 4) return setErr("Password: min 4 chars"); if (p !== cp) return setErr("Passwords don't match"); setLoading(true); await new Promise(r => setTimeout(r, 500)); const res = Auth.register(u, p, name); setLoading(false); if (res.ok) { setOk("Account created!"); setTimeout(() => { const lr = Auth.login(u, p); if (lr.ok) onLogin?.(lr.session); }, 600); } else setErr(res.err); };
  const onKey = (e) => { if (e.key === "Enter") mode === "login" ? doLogin() : doReg(); };

  const iS = (f) => ({ width: "100%", padding: "14px 16px 14px 44px", borderRadius: 12, fontSize: 14, border: `2px solid ${f ? C.accent : C.border}`, background: "#fff", color: C.textDark, outline: "none", boxSizing: "border-box", fontFamily: "'DM Sans',sans-serif", transition: "all 0.2s" });

  return (
    <div style={{ minHeight: "100vh", display: "flex", fontFamily: "'DM Sans',sans-serif", background: `linear-gradient(135deg, ${C.primaryDark} 0%, ${C.primary} 40%, #2A3F6A 100%)`, position: "relative", overflow: "hidden" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Space+Mono:wght@400;700&family=Playfair+Display:wght@600;700;800&display=swap" rel="stylesheet" />
      <style>{`@keyframes fadeUp{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:translateY(0)}} @keyframes slideR{from{opacity:0;transform:translateX(-40px)}to{opacity:1;transform:translateX(0)}} @keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}} input::placeholder{color:${C.textLight}}`}</style>
      <div style={{ position: "absolute", inset: 0, opacity: 0.04, backgroundImage: `linear-gradient(${C.accent} 1px, transparent 1px), linear-gradient(90deg, ${C.accent} 1px, transparent 1px)`, backgroundSize: "60px 60px" }} />

      {/* Left branding */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 80px", position: "relative", zIndex: 2, animation: ready ? "slideR 0.8s ease" : "none", opacity: ready ? 1 : 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 60 }}>
          <div style={{ width: 56, height: 56, borderRadius: 16, background: `linear-gradient(135deg, ${C.accent}, ${C.gold})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28, fontWeight: 800, color: "#fff", boxShadow: `0 8px 32px rgba(232,101,58,0.35)` }}>S</div>
          <div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "#fff", fontFamily: "'Playfair Display',serif" }}>SLA Performance Hub</div>
            <div style={{ fontSize: 12, color: "#ffffff80", letterSpacing: 2, textTransform: "uppercase", fontWeight: 600, marginTop: 2 }}>Service Desk Analytics</div>
          </div>
        </div>
        <h1 style={{ fontSize: 52, fontWeight: 800, color: "#fff", lineHeight: 1.1, margin: "0 0 24px", fontFamily: "'Playfair Display',serif", maxWidth: 520 }}>
          Real-time SLA<br /><span style={{ background: `linear-gradient(135deg, ${C.accent}, ${C.gold})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Intelligence</span>
        </h1>
        <p style={{ fontSize: 17, color: "#ffffff90", lineHeight: 1.7, maxWidth: 460, margin: "0 0 48px" }}>Monitor your Service Desk, Programming Team, and Relationship Managers. Powered by Dynamics 365 and 8x8.</p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
          {[["⬥ Dynamics 365", C.d365], ["⬥ 8x8 Analytics", C.e8x8], ["🔵 Tier 1 Service Desk", null], ["🟠 Tier 2 Programming", null], ["🟣 Tier 3 Rel. Managers", null]].map(([l, c], i) => (
            <div key={i} style={{ padding: "8px 16px", borderRadius: 10, background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", fontSize: 13, fontWeight: 600, color: c || "#ffffffCC" }}>{l}</div>
          ))}
        </div>
      </div>

      {/* Right login form */}
      <div style={{ width: 480, display: "flex", alignItems: "center", justifyContent: "center", padding: "40px 60px", position: "relative", zIndex: 2, animation: ready ? "fadeUp 0.6s ease 0.2s both" : "none" }}>
        <div style={{ width: "100%", background: C.card, borderRadius: 24, padding: "44px 36px", boxShadow: "0 24px 80px rgba(0,0,0,0.3)" }}>
          <h2 style={{ margin: "0 0 4px", fontSize: 26, fontWeight: 800, color: C.textDark, fontFamily: "'Playfair Display',serif" }}>{mode === "login" ? "Welcome Back" : "Create Account"}</h2>
          <p style={{ margin: "0 0 28px", fontSize: 14, color: C.textMid }}>{mode === "login" ? "Sign in to your dashboard" : "Set up your SLA Hub access"}</p>

          {/* Tab */}
          <div style={{ display: "flex", gap: 0, marginBottom: 28, background: C.bg, borderRadius: 10, padding: 3 }}>
            {["login", "register"].map((m) => (
              <button key={m} onClick={() => { setMode(m); setErr(""); setOk(""); }} style={{ flex: 1, padding: "10px 0", borderRadius: 8, border: "none", background: mode === m ? C.card : "transparent", color: mode === m ? C.textDark : C.textLight, fontSize: 13, fontWeight: 600, cursor: "pointer", boxShadow: mode === m ? "0 2px 8px rgba(0,0,0,0.08)" : "none", transition: "all 0.2s", fontFamily: "'DM Sans',sans-serif" }}>
                {m === "login" ? "Sign In" : "Create Account"}
              </button>
            ))}
          </div>

          {mode === "register" && (
            <div style={{ marginBottom: 16, position: "relative" }}>
              <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontSize: 16, color: C.textLight }}>👤</span>
              <input placeholder="Full Name" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={onKey} style={iS(false)} />
            </div>
          )}
          <div style={{ marginBottom: 16, position: "relative" }}>
            <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontSize: 16, color: C.textLight }}>📧</span>
            <input placeholder="Username" value={u} onChange={(e) => setU(e.target.value)} onKeyDown={onKey} style={iS(false)} />
          </div>
          <div style={{ marginBottom: mode === "register" ? 16 : 8, position: "relative" }}>
            <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontSize: 16, color: C.textLight }}>🔒</span>
            <input type={showPw ? "text" : "password"} placeholder="Password" value={p} onChange={(e) => setP(e.target.value)} onKeyDown={onKey} style={iS(false)} />
            <span onClick={() => setShowPw(!showPw)} style={{ position: "absolute", right: 14, top: "50%", transform: "translateY(-50%)", fontSize: 14, cursor: "pointer", color: C.textLight }}>{showPw ? "🙈" : "👁️"}</span>
          </div>
          {mode === "register" && (
            <div style={{ marginBottom: 8, position: "relative" }}>
              <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontSize: 16, color: C.textLight }}>🔒</span>
              <input type="password" placeholder="Confirm Password" value={cp} onChange={(e) => setCp(e.target.value)} onKeyDown={onKey} style={iS(false)} />
            </div>
          )}

          {err && <div style={{ padding: "10px 14px", borderRadius: 10, background: C.redLight, color: C.red, fontSize: 13, fontWeight: 500, marginTop: 12, marginBottom: 4, display: "flex", alignItems: "center", gap: 8 }}>❌ {err}</div>}
          {ok && <div style={{ padding: "10px 14px", borderRadius: 10, background: C.greenLight, color: C.green, fontSize: 13, fontWeight: 500, marginTop: 12, marginBottom: 4, display: "flex", alignItems: "center", gap: 8 }}>✅ {ok}</div>}

          <button onClick={mode === "login" ? doLogin : doReg} disabled={loading} style={{ width: "100%", padding: "14px", borderRadius: 12, border: "none", background: `linear-gradient(135deg, ${C.accent}, ${C.gold})`, color: "#fff", fontSize: 16, fontWeight: 700, cursor: loading ? "wait" : "pointer", marginTop: 20, opacity: loading ? 0.7 : 1, boxShadow: "0 4px 20px rgba(232,101,58,0.35)", transition: "all 0.2s", fontFamily: "'DM Sans',sans-serif" }}>
            {loading ? "..." : mode === "login" ? "Sign In →" : "Create Account →"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   MAIN DASHBOARD — SIDEBAR LAYOUT
   ═══════════════════════════════════════════════════════ */
function Dashboard({ user, onLogout }) {
  const [teamMembers] = useState(DEMO_TEAM_MEMBERS);
  const [selectedMembers, setSelectedMembers] = useState([]);
  const [reportType, setReportType] = useState("daily");
  const [startDate, setStartDate] = useState(() => { const d = new Date(); return d.toISOString().split("T")[0]; });
  const [endDate, setEndDate] = useState(() => { const d = new Date(); return d.toISOString().split("T")[0]; });
  const [data, setData] = useState(null);
  const [hasRun, setHasRun] = useState(false);
  const [isRunning, setIsRunning] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [apiConfig, setApiConfig] = useState({ d365: { orgUrl: "", apiVersion: "v9.2", token: "" }, e8x8: { baseUrl: "", tenantId: "", apiKey: "" }, live: false });
  const reportRef = useRef(null);

  const canRun = selectedMembers.length > 0;

  const setPreset = (type) => {
    setReportType(type);
    const today = new Date();
    if (type === "daily") {
      setStartDate(today.toISOString().split("T")[0]);
      setEndDate(today.toISOString().split("T")[0]);
    } else if (type === "weekly") {
      const weekAgo = new Date(today); weekAgo.setDate(weekAgo.getDate() - 7);
      setStartDate(weekAgo.toISOString().split("T")[0]);
      setEndDate(today.toISOString().split("T")[0]);
    }
  };

  const handleRun = async () => {
    setIsRunning(true);
    await new Promise((r) => setTimeout(r, 800));
    const d = generateDemoData(startDate, endDate, selectedMembers);
    setData(d);
    setHasRun(true);
    setIsRunning(false);
  };

  const handleExportPDF = () => {
    if (reportRef.current) {
      window.print();
    }
  };

  const dateLabel = startDate === endDate
    ? new Date(startDate + "T12:00:00").toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" })
    : `${new Date(startDate + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })} — ${new Date(endDate + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}`;

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'DM Sans', sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Space+Mono:wght@400;700&family=Playfair+Display:wght@600;700;800&display=swap" rel="stylesheet" />
      <style>{`@keyframes fadeIn { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } } @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.5; } } @keyframes slideIn { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: translateY(0); } } @media print { .no-print { display: none !important; } }`}</style>

      <SettingsModal show={showSettings} onClose={() => setShowSettings(false)} config={apiConfig} onSave={setApiConfig} />

      {/* Header */}
      <div className="no-print" style={{ background: C.primary, padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ width: 38, height: 38, borderRadius: 9, background: `linear-gradient(135deg, ${C.accent}, ${C.yellow})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 700, color: "#fff" }}>S</div>
          <div>
            <h1 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: "#fff", fontFamily: "'Playfair Display', serif" }}>SLA Performance Hub</h1>
            <div style={{ fontSize: 11, color: "#B3D4F7", marginTop: 1, letterSpacing: 0.5 }}>Dynamics 365 + 8x8 Analytics · Service Desk</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 13, color: "#ffffff80", fontWeight: 500 }}>👤 {user?.name || "User"}</span>
          {hasRun && <button onClick={handleExportPDF} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}><span>📄</span> Export PDF</button>}
          <button onClick={() => setShowSettings(true)} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 14px", fontSize: 14, cursor: "pointer" }}>⚙️</button>
          <button onClick={onLogout} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Logout</button>
        </div>
      </div>

      <ConnectionBar config={apiConfig} onOpenSettings={() => setShowSettings(true)} />

      <div style={{ display: "flex", maxWidth: 1500, margin: "0 auto" }}>
        {/* ═══════ SIDEBAR ═══════ */}
        <div className="no-print" style={{ width: 310, minWidth: 310, background: C.card, borderRight: `1px solid ${C.border}`, padding: "24px 20px", minHeight: "calc(100vh - 110px)", display: "flex", flexDirection: "column" }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: C.textLight, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 14 }}>Configure Report</div>

          {/* Team Members */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}><span>👥</span> Team Members</div>
            <MultiMemberSelect selected={selectedMembers} onChange={setSelectedMembers} members={teamMembers} />
            {selectedMembers.length > 0 && <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 4 }}>
              {selectedMembers.slice(0, 4).map((id) => { const m = teamMembers.find((t) => t.id === id); const idx = teamMembers.indexOf(m); return <span key={id} style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: PIE_COLORS[idx % PIE_COLORS.length] + "18", color: PIE_COLORS[idx % PIE_COLORS.length], fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}>{m?.name?.split(" ")[0]}<span onClick={() => setSelectedMembers(selectedMembers.filter((s) => s !== id))} style={{ cursor: "pointer", opacity: 0.6, fontSize: 8 }}>✕</span></span>; })}
              {selectedMembers.length > 4 && <span style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: C.bg, color: C.textLight, fontWeight: 600 }}>+{selectedMembers.length - 4} more</span>}
            </div>}
          </div>

          {/* Report Type */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}><span>📊</span> Report Type</div>
            <div style={{ display: "flex", gap: 4 }}>
              {[["daily", "📋 Daily"], ["weekly", "📅 Weekly"]].map(([v, l]) => (
                <button key={v} onClick={() => setPreset(v)} style={{ flex: 1, padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${reportType === v ? C.accent : C.border}`, background: reportType === v ? C.accent + "10" : C.bg, color: reportType === v ? C.accent : C.textMid, fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>{l}</button>
              ))}
            </div>
          </div>

          {/* Tier Info Card */}
          <div style={{ marginBottom: 18, background: C.bg, borderRadius: 10, padding: "12px 14px", border: `1px solid ${C.border}` }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: C.textLight, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Queue Tiers</div>
            {Object.values(TIERS).map((t) => (
              <div key={t.code} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, padding: "4px 0" }}>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: t.color, flexShrink: 0 }} />
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: C.textDark }}>{t.icon} {t.label} — {t.name}</div>
                  <div style={{ fontSize: 9, color: C.textLight, lineHeight: 1.3 }}>{t.desc}</div>
                </div>
              </div>
            ))}
          </div>

          {/* Date Range */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}><span>📅</span> Date Range</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <label><div style={{ fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 }}>From</div><input type="date" value={startDate} onChange={(e) => { setStartDate(e.target.value); setReportType("custom"); }} style={{ width: "100%", padding: "8px 8px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 11, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} /></label>
              <label><div style={{ fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 }}>To</div><input type="date" value={endDate} onChange={(e) => { setEndDate(e.target.value); setReportType("custom"); }} style={{ width: "100%", padding: "8px 8px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 11, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} /></label>
            </div>
            <div style={{ display: "flex", gap: 5, marginTop: 8 }}>
              {[{ l: "7D", s: 7 }, { l: "14D", s: 14 }, { l: "30D", s: 30 }, { l: "90D", s: 90 }, { l: "YTD", s: -1 }].map((q) => <button key={q.l} onClick={() => { const e = new Date(), sD = new Date(); if (q.s === -1) { sD.setMonth(0); sD.setDate(1); } else { sD.setDate(sD.getDate() - q.s); } setStartDate(sD.toISOString().split("T")[0]); setEndDate(e.toISOString().split("T")[0]); setReportType("custom"); }} style={{ flex: 1, padding: "5px 0", borderRadius: 6, border: `1px solid ${C.border}`, background: "transparent", fontSize: 10, fontWeight: 600, color: C.textMid, cursor: "pointer", fontFamily: "'Space Mono', monospace" }}>{q.l}</button>)}
            </div>
          </div>

          <div style={{ flex: 1 }} />

          {/* Run Button */}
          <button onClick={handleRun} disabled={!canRun || isRunning} style={{ width: "100%", padding: "14px", borderRadius: 12, border: "none", background: canRun ? `linear-gradient(135deg, ${C.accent}, ${C.yellow})` : C.border, color: canRun ? "#fff" : C.textLight, fontSize: 15, fontWeight: 700, cursor: canRun ? "pointer" : "not-allowed", letterSpacing: 0.5, display: "flex", alignItems: "center", justifyContent: "center", gap: 10, boxShadow: canRun ? "0 4px 20px rgba(232,101,58,0.35)" : "none", opacity: isRunning ? 0.7 : 1 }}>
            {isRunning ? <><span style={{ animation: "pulse 1s infinite" }}>⏳</span> Generating Report...</> : <><span style={{ fontSize: 18 }}>▶</span> Run Report</>}
          </button>
          {!canRun && <div style={{ fontSize: 10, color: C.accent, textAlign: "center", marginTop: 6 }}>Select at least 1 team member</div>}
          {hasRun && <button onClick={handleExportPDF} style={{ width: "100%", padding: "12px", borderRadius: 10, border: `1.5px solid ${C.border}`, background: C.card, color: C.textDark, fontSize: 13, fontWeight: 600, cursor: "pointer", marginTop: 10, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><span>📄</span> Export to PDF</button>}
        </div>

        {/* ═══════ MAIN CONTENT ═══════ */}
        <div style={{ flex: 1, padding: "24px 28px", overflow: "auto", minHeight: "calc(100vh - 110px)" }}>
          {!hasRun ? (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", textAlign: "center" }}>
              <div style={{ width: 100, height: 100, borderRadius: 24, background: `linear-gradient(135deg, ${C.accent}15, ${C.yellow}15)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 44, marginBottom: 20 }}>📊</div>
              <h2 style={{ margin: "0 0 8px", fontSize: 22, fontWeight: 700, color: C.textDark, fontFamily: "'Playfair Display', serif" }}>Service Desk SLA Dashboard</h2>
              <p style={{ margin: 0, fontSize: 14, color: C.textMid, maxWidth: 440, lineHeight: 1.6 }}>
                Select your team members, choose a report type, set your date range, and hit <strong style={{ color: C.accent }}>Run Report</strong>.
                Data pulls from <strong style={{ color: C.d365 }}>Dynamics 365</strong> and <strong style={{ color: C.e8x8 }}>8x8 Analytics</strong>.
              </p>
              <div style={{ marginTop: 20, display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "center" }}>
                {[{ icon: "👥", label: `${selectedMembers.length} members`, ok: selectedMembers.length > 0 }, { icon: "📊", label: reportType === "daily" ? "Daily Report" : reportType === "weekly" ? "Weekly Report" : "Custom Range", ok: true }, { icon: "📅", label: `${startDate} → ${endDate}`, ok: startDate && endDate }].map((s, i) => <div key={i} style={{ padding: "10px 16px", borderRadius: 10, background: s.ok ? C.greenLight + "22" : C.accentLight + "22", border: `1px solid ${s.ok ? C.greenLight + "44" : C.accentLight + "44"}`, fontSize: 12, fontWeight: 600, color: s.ok ? C.green : C.accent, display: "flex", alignItems: "center", gap: 6 }}><span>{s.icon}</span> {s.label} {s.ok ? "✓" : "✗"}</div>)}
              </div>
              <div style={{ marginTop: 28, display: "flex", gap: 16 }}>
                {[["🔵 Tier 1", "Service Desk", TIERS[1].color], ["🟠 Tier 2", "Programming", TIERS[2].color], ["🟣 Tier 3", "Rel. Managers", TIERS[3].color], ["📞 Phone", "8x8", C.e8x8], ["📧 Email", "D365", C.d365]].map(([icon, sub, clr]) => (
                  <div key={icon} style={{ padding: "12px 18px", borderRadius: 10, background: C.card, border: `1px solid ${C.border}`, textAlign: "center" }}><div style={{ fontSize: 14, fontWeight: 700, color: C.textDark }}>{icon}</div><div style={{ fontSize: 10, color: clr, fontWeight: 600, marginTop: 2 }}>{sub}</div></div>
                ))}
              </div>
            </div>
          ) : data && (
            <div style={{ animation: "slideIn 0.4s ease" }} ref={reportRef}>
              {/* Report Header */}
              <div style={{ marginBottom: 24, display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
                <div>
                  <h2 style={{ margin: 0, fontSize: 24, fontWeight: 800, color: C.textDark, fontFamily: "'Playfair Display', serif" }}>📊 {reportType === "weekly" ? "Weekly" : reportType === "daily" ? "Daily" : "Custom"} Service Desk Report</h2>
                  <div style={{ fontSize: 12, color: C.textMid, marginTop: 4, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                    <span>👥 {selectedMembers.length} member{selectedMembers.length > 1 ? "s" : ""}</span>
                    <span>📅 {dateLabel}</span>
                    <span style={{ fontSize: 10, padding: "2px 8px", borderRadius: 4, background: apiConfig.live ? C.greenLight + "33" : "#0078D415", color: apiConfig.live ? C.green : "#0078D4", fontWeight: 600 }}>{apiConfig.live ? "🟢 Live" : "🔵 Demo"}</span>
                  </div>
                </div>
              </div>

              {/* Report Sections */}
              <TierSection tier={1} data={data} />
              <div style={{ height: 3, background: C.bg }} />
              <TierSection tier={2} data={data} />
              <div style={{ height: 3, background: C.bg }} />
              <TierSection tier={3} data={data} />
              <div style={{ height: 3, background: C.bg }} />
              <PhoneSection data={data} />
              <div style={{ height: 3, background: C.bg }} />
              <EmailSection data={data} />
              <div style={{ height: 3, background: C.bg }} />
              <CSATSection data={data} />
              <div style={{ height: 3, background: C.bg }} />
              <OverallSummary data={data} />
              <Definitions />

              {/* Footer */}
              <div style={{ background: C.primaryDark, padding: 14, textAlign: "center", borderRadius: "0 0 14px 14px" }}>
                <p style={{ margin: 0, color: "#a8c6df", fontSize: 11 }}>Report generated automatically by Service Desk SLA System</p>
              </div>

              {/* Charts */}
              <ChartsPanel data={data} />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   APP ROOT
   ═══════════════════════════════════════════════════════ */
export default function App() {
  const [user, setUser] = useState(null);
  const [checking, setChecking] = useState(true);
  useEffect(() => { const s = Auth.session(); if (s) setUser(s); setChecking(false); }, []);
  if (checking) return <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: C.primary, color: "#fff", fontFamily: "'DM Sans',sans-serif" }}><div style={{ textAlign: "center" }}><div style={{ width: 56, height: 56, borderRadius: 16, margin: "0 auto 16px", background: `linear-gradient(135deg, ${C.accent}, ${C.gold})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28, fontWeight: 800, color: "#fff" }}>S</div><div style={{ fontSize: 14, color: "#ffffff60" }}>Loading...</div></div></div>;
  if (!user) return <LoginPage onLogin={(s) => setUser(s)} />;
  return <Dashboard user={user} onLogout={() => { Auth.logout(); setUser(null); }} />;
}

import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend, PieChart, Pie, Cell,
  AreaChart, Area
} from "recharts";
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

/* ═══════════════════════════════════════════════════════════════════
   SERVICE AND OPERATIONS DASHBOARD v7
   Sidebar Layout + MSAL.js D365 Auth + Live OData
   Data: Dynamics 365 (Cases/SLA/CSAT/Phone Calls)
   ═══════════════════════════════════════════════════════════════════ */

/* ─── MSAL CONFIGURATION ─── */
const MSAL_CONFIG = {
  auth: {
    clientId: "0918449d-b73e-428a-8238-61723f2a2e7d",
    authority: "https://login.microsoftonline.com/1b0086bd-aeda-4c74-a15a-23adfe4d0693",
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

const D365_SCOPE = "https://servingintel.crm.dynamics.com/user_impersonation";
const D365_BASE = "https://servingintel.crm.dynamics.com/api/data/v9.2";
const GRAPH_SCOPE = "Mail.Send";

let msalInstance = null;
function getMsal() {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(MSAL_CONFIG);
  }
  return msalInstance;
}

/* ─── D365 TOKEN ACQUISITION ─── */
async function getD365Token() {
  const msal = getMsal();
  await msal.initialize();
  const accounts = msal.getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    const result = await msal.acquireTokenSilent({
      scopes: [D365_SCOPE],
      account: accounts[0],
    });
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      try {
        const result = await msal.acquireTokenPopup({ scopes: [D365_SCOPE] });
        return result.accessToken;
      } catch { return null; }
    }
    return null;
  }
}

/* ─── GRAPH API TOKEN (for sending email) ─── */
async function getGraphToken() {
  const msal = getMsal();
  await msal.initialize();
  const accounts = msal.getAllAccounts();
  if (accounts.length === 0) throw new Error("Not signed in. Sign in with Microsoft first.");
  try {
    const result = await msal.acquireTokenSilent({ scopes: [GRAPH_SCOPE], account: accounts[0] });
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const result = await msal.acquireTokenPopup({ scopes: [GRAPH_SCOPE] });
      return result.accessToken;
    }
    throw err;
  }
}

/* ─── SEND EMAIL VIA GRAPH API ─── */
async function sendEmailViaGraph(to, subject, htmlBody) {
  const token = await getGraphToken();
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: "HTML", content: htmlBody },
        toRecipients: to.split(",").map(e => ({ emailAddress: { address: e.trim() } })),
      },
    }),
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Email failed (${res.status}): ${txt.substring(0, 200)}`);
  }
  return true;
}

async function msalLogin() {
  const msal = getMsal();
  await msal.initialize();
  try {
    const result = await msal.loginPopup({
      scopes: [D365_SCOPE],
    });
    return result;
  } catch (err) {
    console.error("MSAL login error:", err);
    return null;
  }
}

async function msalLogoutD365() {
  const msal = getMsal();
  await msal.initialize();
  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    await msal.logoutPopup({ account: accounts[0] });
  }
}

function getMsalAccount() {
  try {
    const msal = getMsal();
    const accounts = msal.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  } catch { return null; }
}

/* ─── TIMEZONE: Convert Central Time to UTC (auto-detects CST/CDT) ─── */
function ctToUTC(dateStr, timeStr, endOfMinute = false) {
  const pad = (n) => String(n).padStart(2, "0");
  const [h, m] = (timeStr || "00:00").split(":").map(Number);
  const sec = endOfMinute ? 59 : 0;
  const [y, mo, d] = dateStr.split("-").map(Number);

  // Auto-detect DST: CDT (UTC-5) or CST (UTC-6)
  // US rule: CDT starts 2nd Sunday of March at 2AM, ends 1st Sunday of November at 2AM
  function getNthSunday(year, month, n) {
    const firstDay = new Date(year, month, 1).getDay();
    return 1 + ((7 - firstDay) % 7) + (n - 1) * 7;
  }
  const dstStart = new Date(y, 2, getNthSunday(y, 2, 2), 2, 0, 0);  // 2nd Sun March, 2:00 AM
  const dstEnd   = new Date(y, 10, getNthSunday(y, 10, 1), 2, 0, 0); // 1st Sun Nov, 2:00 AM
  const localProbe = new Date(y, mo - 1, d, h, m, sec);
  const offsetHours = (localProbe >= dstStart && localProbe < dstEnd) ? 5 : 6;

  // Convert CT wall clock → UTC by adding offset
  const dt = new Date(Date.UTC(y, mo - 1, d, h + offsetHours, m, sec));
  const utcDate = `${dt.getUTCFullYear()}-${pad(dt.getUTCMonth() + 1)}-${pad(dt.getUTCDate())}`;
  const utcTime = `${pad(dt.getUTCHours())}:${pad(dt.getUTCMinutes())}:${pad(dt.getUTCSeconds())}Z`;
  return { date: utcDate, time: utcTime };
}

/* ─── D365 API HELPER ─── */
async function d365Fetch(query) {
  const token = await getD365Token();
  if (!token) throw new Error("No D365 token — please sign in");
  const res = await fetch(`${D365_BASE}/${query}`, {
    headers: {
      "Authorization": `Bearer ${token}`,
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "Accept": "application/json",
      "Prefer": "odata.include-annotations=*,odata.maxpagesize=5000",
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`D365 API ${res.status}: ${text.substring(0, 200)}`);
  }
  return res.json();
}

async function d365Count(query) {
  const data = await d365Fetch(query);
  return data["@odata.count"] ?? data.value?.length ?? 0;
}

/* ─── REAL TIER DEFINITIONS (from D365 casetypecode) ─── */
const TIERS = {
  1: {
    code: 1, label: "Tier 1", name: "Service Desk", icon: "🔵",
    color: "#2196F3", colorLight: "#e8f4fd", colorDark: "#1565c0",
    desc: "Front-line support — password resets, basic troubleshooting, general inquiries",
    d365Filter: "casetypecode eq 1", dateField: "createdon",
    metrics: ["sla_compliance", "sla_response", "open_breach_rate", "fcr_rate", "escalation_rate", "avg_resolution_time", "total_cases"],
  },
  2: {
    code: 2, label: "Tier 2", name: "Programming Team", icon: "🟠",
    color: "#FF9800", colorLight: "#fff3e0", colorDark: "#e65100",
    desc: "Intermediate support — complex issues, escalations from Tier 1, technical cases",
    d365Filter: "casetypecode eq 2", dateField: "escalatedon",
    metrics: ["sla_compliance", "sla_response", "open_breach_rate", "escalation_rate", "total_cases", "resolved"],
  },
  3: {
    code: 3, label: "Tier 3", name: "Relationship Managers", icon: "🟣",
    color: "#9C27B0", colorLight: "#f3e5f5", colorDark: "#7b1fa2",
    desc: "Advanced support — critical escalations, system-level issues, VIP accounts",
    d365Filter: "casetypecode eq 3", dateField: "escalatedon",
    metrics: ["sla_compliance", "sla_response", "open_breach_rate", "total_cases", "resolved"],
  },
};

/* ─── REAL SLA TARGETS ─── */
const TARGETS = {
  sla_compliance: { value: 90, unit: "%", compare: "gte", label: "90%" },
  sla_response: { value: 90, unit: "%", compare: "gte", label: "90%" },
  open_breach_rate: { value: 5, unit: "%", compare: "lt", label: "<5%" },
  fcr_rate: { value: 90, unit: "%", compare: "gte", label: "90-95%" },
  escalation_rate: { value: 10, unit: "%", compare: "lt", label: "<10%" },
  answer_rate: { value: 95, unit: "%", compare: "gte", label: ">95%" },
  avg_phone_aht: { value: 6, unit: " min", compare: "lte", label: "<6 min" },
  csat_score: { value: 4.0, unit: "/5", compare: "gte", label: "4.0+" },
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

/* ─── COLORS ─── */
const C = {
  primary: "#1B2A4A", primaryDark: "#152d4a", accent: "#E8653A", accentLight: "#F09A7A",
  green: "#4CAF50", greenLight: "#e8f5e9", red: "#f44336", redLight: "#ffebee",
  orange: "#FF9800", orangeLight: "#fff3e0", blue: "#2196F3", blueLight: "#e3f2fd",
  purple: "#9C27B0", purpleLight: "#f3e5f5", gold: "#FFC107", goldLight: "#fff8e1",
  yellow: "#E6B422", gray: "#9e9e9e", grayLight: "#f5f5f5",
  bg: "#F4F1EC", card: "#F4F1EC", border: "#E2DDD5",
  textDark: "#1B2A4A", textMid: "#5A6578", textLight: "#8B95A5",
  d365: "#0078D4",
};
const PIE_COLORS = [C.accent, "#2D9D78", C.blue, C.yellow, C.purple, C.accentLight, "#A3E4C8"];

/* ─── DEMO TEAM MEMBERS (fallback) ─── */
const DEMO_TEAM_MEMBERS = [
  { id: "demo1", name: "Demo Agent 1", role: "Agent I", avatar: "D1", tier: 1 },
  { id: "demo2", name: "Demo Agent 2", role: "Agent I", avatar: "D2", tier: 1 },
  { id: "demo3", name: "Demo Agent 3", role: "Agent II", avatar: "D3", tier: 2 },
];

/* ─── FETCH QUEUES FROM D365 ─── */
async function fetchD365Queues() {
  try {
    const data = await d365Fetch(
      `queues?$filter=statecode eq 0&$select=queueid,name,description,emailaddress,queueviewtype&$orderby=name asc`
    );
    return (data.value || []).map(q => ({
      id: q.queueid,
      name: q.name + (q.queueviewtype === 1 ? " 🔒" : ""),
      description: q.description || "",
      email: q.emailaddress || "",
    }));
  } catch (err) {
    console.error("Failed to fetch queues:", err);
    try {
      const data = await d365Fetch(
        `queues?$select=queueid,name,description,emailaddress&$orderby=name asc&$top=50`
      );
      return (data.value || []).map(q => ({
        id: q.queueid,
        name: q.name,
        description: q.description || "",
        email: q.emailaddress || "",
      }));
    } catch {
      return [];
    }
  }
}

/* ─── FETCH QUEUE MEMBERS FROM D365 ─── */
async function fetchD365QueueMembers(queueId) {
  try {
    const data = await d365Fetch(
      `queues(${queueId})/queue_membership?$select=systemuserid,fullname,title,jobtitle,internalemailaddress&$filter=isdisabled eq false`
    );
    if (data.value?.length > 0) return mapD365Users(data.value);
  } catch (e) { console.log("Approach 1 failed:", e.message); }

  try {
    const data = await d365Fetch(
      `queues(${queueId})/queuemembership_association?$select=systemuserid,fullname,title,jobtitle,internalemailaddress`
    );
    if (data.value?.length > 0) return mapD365Users(data.value);
  } catch (e) { console.log("Approach 2 failed:", e.message); }

  try {
    const fetchXml = encodeURIComponent(`<fetch><entity name="systemuser"><attribute name="systemuserid"/><attribute name="fullname"/><attribute name="title"/><attribute name="jobtitle"/><attribute name="internalemailaddress"/><filter><condition attribute="isdisabled" operator="eq" value="0"/></filter><link-entity name="queuemembership" from="systemuserid" to="systemuserid" intersect="true"><link-entity name="queue" from="queueid" to="queueid"><filter><condition attribute="queueid" operator="eq" value="${queueId}"/></filter></link-entity></link-entity></entity></fetch>`);
    const data = await d365Fetch(`systemusers?fetchXml=${fetchXml}`);
    if (data.value?.length > 0) return mapD365Users(data.value);
  } catch (e) { console.log("Approach 3 failed:", e.message); }

  try {
    const data = await d365Fetch(
      `incidents?$filter=_queueid_value eq ${queueId}&$select=_ownerid_value&$top=200`
    );
    if (data.value?.length > 0) {
      const ownerIds = [...new Set(data.value.map(c => c._ownerid_value).filter(Boolean))];
      if (ownerIds.length > 0) {
        const filterParts = ownerIds.slice(0, 15).map(id => `systemuserid eq ${id}`).join(" or ");
        const users = await d365Fetch(
          `systemusers?$filter=(${filterParts}) and isdisabled eq false&$select=systemuserid,fullname,title,jobtitle,internalemailaddress`
        );
        if (users.value?.length > 0) return mapD365Users(users.value);
      }
    }
  } catch (e) { console.log("Approach 4 failed:", e.message); }

  try {
    const data = await d365Fetch(
      `queueitems?$filter=_queueid_value eq ${queueId}&$select=_workerid_value&$top=200`
    );
    if (data.value?.length > 0) {
      const workerIds = [...new Set(data.value.map(qi => qi._workerid_value).filter(Boolean))];
      if (workerIds.length > 0) {
        const filterParts = workerIds.slice(0, 15).map(id => `systemuserid eq ${id}`).join(" or ");
        const users = await d365Fetch(
          `systemusers?$filter=(${filterParts}) and isdisabled eq false&$select=systemuserid,fullname,title,jobtitle,internalemailaddress`
        );
        if (users.value?.length > 0) return mapD365Users(users.value);
      }
    }
  } catch (e) { console.log("Approach 5 failed:", e.message); }

  return [];
}

function mapD365Users(users) {
  return users
    .filter(u => u.fullname && !u.fullname.startsWith("#") && !u.fullname.includes("SYSTEM") && !u.fullname.includes("Integration") && !u.fullname.includes("Builder") && !u.fullname.includes("Apollo") && !u.fullname.includes("Tools"))
    .map(u => {
      const name = u.fullname || "Unknown";
      const parts = name.split(" ");
      const avatar = parts.length >= 2
        ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
        : name.substring(0, 2).toUpperCase();
      const titleLower = (u.title || u.jobtitle || "").toLowerCase();
      let tier = 1;
      if (titleLower.includes("senior") || titleLower.includes("manager") || titleLower.includes("relationship") || titleLower.includes("vip")) tier = 3;
      else if (titleLower.includes("program") || titleLower.includes("developer") || titleLower.includes("engineer") || titleLower.includes("ii") || titleLower.includes("level 2") || titleLower.includes("tier 2")) tier = 2;
      return {
        id: u.systemuserid,
        name,
        role: u.title || u.jobtitle || "Agent",
        avatar,
        tier,
        email: u.internalemailaddress || "",
      };
    });
}

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
  const t1Resolved = t1SLAMet + Math.round(t1Cases * (0.05 + r(21) * 0.1));
  const allResolved = t1Resolved + t2Resolved + t3Resolved;
  const avgResTime = +(1.5 + r(20) * 6).toFixed(1);

  // Demo SLA Response data
  const t1RespMet = Math.round(t1SLAMet * (0.85 + r(22) * 0.13));
  const t1RespMissed = Math.round(t1SLAMet * (0.02 + r(23) * 0.08));
  const t1Active = Math.max(1, t1Cases - t1Resolved);
  const t1BreachCount = Math.round(t1Active * (0.05 + r(24) * 0.15));
  const t2RespMet = Math.round(t2Resolved * (0.8 + r(25) * 0.15));
  const t2RespMissed = Math.round(t2Resolved * (0.03 + r(26) * 0.1));
  const t2Active = Math.max(1, t2Cases - t2Resolved);
  const t2BreachCount = Math.round(t2Active * (0.1 + r(27) * 0.2));
  const t3RespMet = Math.round(t3Resolved * (0.75 + r(28) * 0.2));
  const t3RespMissed = Math.round(t3Resolved * (0.05 + r(29) * 0.1));
  const t3Active = Math.max(1, t3Cases - t3Resolved);
  const t3BreachCount = Math.round(t3Active * (0.15 + r(30) * 0.25));

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
    tier1: { total: t1Cases, slaMet: t1SLAMet, slaMissed: Math.round(t1Cases * (0.05 + r(21) * 0.1)), slaCompliance: t1SLAMet ? Math.min(100, Math.round(t1SLAMet / (t1SLAMet + Math.round(t1Cases * (0.05 + r(21) * 0.1))) * 100)) : 0, slaResponseMet: t1RespMet, slaResponseMissed: t1RespMissed, slaResponseCompliance: (t1RespMet + t1RespMissed) > 0 ? Math.min(100, Math.round(t1RespMet / (t1RespMet + t1RespMissed) * 100)) : "N/A", openBreachCount: t1BreachCount, openBreachTotal: t1Active, openBreachRate: t1Active > 0 ? Math.min(100, Math.round(t1BreachCount / t1Active * 100)) : 0, fcrRate: t1Cases ? Math.min(100, Math.round(t1FCR / t1Cases * 100)) : 0, escalationRate: t1Cases ? Math.min(100, Math.round(t1Escalated / t1Cases * 100)) : 0, avgResolutionTime: `${avgResTime} hrs`, escalated: t1Escalated },
    tier2: { total: t2Cases, resolved: t2Resolved, slaMet: t2SLAMet, slaMissed: Math.max(0, t2Resolved - t2SLAMet), slaCompliance: t2Resolved ? Math.min(100, Math.round(t2SLAMet / t2Resolved * 100)) : "N/A", slaResponseMet: t2RespMet, slaResponseMissed: t2RespMissed, slaResponseCompliance: (t2RespMet + t2RespMissed) > 0 ? Math.min(100, Math.round(t2RespMet / (t2RespMet + t2RespMissed) * 100)) : "N/A", openBreachCount: t2BreachCount, openBreachTotal: t2Active, openBreachRate: t2Active > 0 ? Math.min(100, Math.round(t2BreachCount / t2Active * 100)) : 0, escalationRate: t2Cases ? Math.min(100, Math.round(t2Escalated / t2Cases * 100)) : "N/A", escalated: t2Escalated },
    tier3: { total: t3Cases, resolved: t3Resolved, slaMet: t3SLAMet, slaMissed: Math.max(0, t3Resolved - t3SLAMet), slaCompliance: t3Resolved ? Math.min(100, Math.round(t3SLAMet / t3Resolved * 100)) : "N/A", slaResponseMet: t3RespMet, slaResponseMissed: t3RespMissed, slaResponseCompliance: (t3RespMet + t3RespMissed) > 0 ? Math.min(100, Math.round(t3RespMet / (t3RespMet + t3RespMissed) * 100)) : "N/A", openBreachCount: t3BreachCount, openBreachTotal: t3Active, openBreachRate: t3Active > 0 ? Math.min(100, Math.round(t3BreachCount / t3Active * 100)) : 0 },
    phone: { totalCalls, answered, abandoned, answerRate: totalCalls ? Math.min(100, Math.round(answered / totalCalls * 100)) : 0, avgAHT },
    email: { total: emailCases, responded: emailResponded, resolved: emailResolved, slaCompliance: emailResolved > 0 ? 100 : "N/A" },
    csat: { responses: csatResponses, avgScore: csatAvg || "N/A" },
    overall: { created: allCases, resolved: allResolved, csatResponses, answeredCalls: answered, abandonedCalls: abandoned },
    timeline,
    source: "demo",
  };
}

/* ═══════════════════════════════════════════════════════
   LIVE DATA FETCHER — D365 OData
   ═══════════════════════════════════════════════════════ */

async function fetchMemberD365Data(member, startDate, endDate, onProgress, startTime, endTime) {
  const utcStart = ctToUTC(startDate, startTime || "00:00", false);
  const utcEnd = ctToUTC(endDate, endTime || "23:59", true);
  const s = utcStart.date, sT = utcStart.time, e = utcEnd.date, eT = utcEnd.time;
  const oid = member.id;
  const errors = [];
  const progress = (msg) => onProgress?.(`${member.name}: ${msg}`);

  async function safeCount(label, query) {
    try { progress(label); return await d365Count(query); }
    catch (err) { errors.push(`${member.name} — ${label}: ${err.message}`); return 0; }
  }

  async function safeFetchCount(label, query) {
    try {
      progress(label);
      const cleanQuery = query.replace(/&?\$count=true/g, '').replace(/&?\$top=\d+/g, '');
      const data = await d365Fetch(`${cleanQuery}&$top=5000`);
      return data.value?.length ?? 0;
    }
    catch (err) { errors.push(`${member.name} — ${label}: ${err.message}`); return 0; }
  }

  const totalCases = await safeCount("Total Cases",
    `incidents?$filter=_ownerid_value eq ${oid} and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const resolvedCases = await safeFetchCount("Resolved",
    `incidents?$filter=_ownerid_value eq ${oid} and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);

  // Fetch SLA KPI Instance status via $expand on resolved cases
  let slaMet = 0, slaMissed = 0;
  try {
    progress("Fetching SLA KPI...");
    const slaData = await d365Fetch(
      `incidents?$filter=_ownerid_value eq ${oid} and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$expand=resolvebykpiid($select=status)&$top=5000`
    );
    for (const rec of (slaData.value || [])) {
      const kpiStatus = rec.resolvebykpiid?.status;
      if (kpiStatus === 4) slaMet++;
      else if (kpiStatus === 1) slaMissed++;
    }
    console.log(`[D365 SLA] ${member.name}: met=${slaMet}, missed=${slaMissed}`);
  } catch (err) {
    errors.push(`${member.name} — SLA KPI: ${err.message}`);
  }
  const fcrCases = await safeFetchCount("FCR",
    `incidents?$filter=_ownerid_value eq ${oid} and cr7fe_new_fcr eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);
  const escalatedCases = await safeCount("Escalated",
    `incidents?$filter=_ownerid_value eq ${oid} and isescalated eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const activeCases = await safeCount("Active",
    `incidents?$filter=_ownerid_value eq ${oid} and statecode eq 0 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);

  const emailCases = await safeCount("Email Cases",
    `incidents?$filter=_ownerid_value eq ${oid} and caseorigincode eq 2 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const emailResolved = await safeFetchCount("Email Resolved",
    `incidents?$filter=_ownerid_value eq ${oid} and caseorigincode eq 2 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);

  const casesCreatedBy = await safeFetchCount("Cases Created",
    `incidents?$filter=_createdby_value eq ${oid} and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid`);

  const totalPhoneCalls = await safeFetchCount("Total Phone Calls",
    `phonecalls?$filter=_ownerid_value eq ${oid} and actualstart ge ${s}T${sT} and actualstart le ${e}T${eT}&$select=actualdurationminutes`);
  const answeredLive = await safeFetchCount("Answered Calls",
    `phonecalls?$filter=_ownerid_value eq ${oid} and actualstart ge ${s}T${sT} and actualstart le ${e}T${eT} and actualdurationminutes gt 0&$select=actualdurationminutes`);
  const abandonedCalls = await safeFetchCount("Abandoned Calls",
    `phonecalls?$filter=_ownerid_value eq ${oid} and actualstart ge ${s}T${sT} and actualstart le ${e}T${eT} and actualdurationminutes eq 0&$select=actualdurationminutes`);
  const incomingCalls = totalPhoneCalls;
  const outgoingCalls = 0;
  const voicemails = abandonedCalls;

  let memberAHT = "N/A";
  if (totalPhoneCalls > 0) {
    try {
      const ahtData = await d365Fetch(
        `phonecalls?$filter=_ownerid_value eq ${oid} and actualstart ge ${s}T${sT} and actualstart le ${e}T${eT}&$select=actualdurationminutes&$top=5000`
      );
      if (ahtData.value?.length > 0) {
        const durations = ahtData.value.map(r => parseFloat(r.actualdurationminutes) || 0).filter(n => !isNaN(n));
        if (durations.length > 0) {
          const avg = Math.round(durations.reduce((a, b) => a + b, 0) / durations.length / 60);
          memberAHT = `${avg} min`;
        }
      }
    } catch (err) { errors.push(`${member.name} — AHT: ${err.message}`); }
  }

  let csatResponses = 0, csatAvg = "N/A";
  try {
    progress("CSAT");
    const csatData = await d365Fetch(
      `incidents?$filter=_ownerid_value eq ${oid} and cr7fe_new_csatresponsereceived eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=cr7fe_new_csatscore`
    );
    csatResponses = csatData.value?.length || 0;
    if (csatResponses > 0) {
      const scores = csatData.value.map(r => parseFloat(r.cr7fe_new_csatscore)).filter(n => !isNaN(n));
      if (scores.length > 0) csatAvg = +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1);
    }
  } catch (err) { errors.push(`${member.name} — CSAT: ${err.message}`); }

  let avgResTime = "N/A";
  try {
    progress("Resolution time");
    const resolved = await d365Fetch(
      `incidents?$filter=_ownerid_value eq ${oid} and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid,cr7fe_new_handletime,createdon,modifiedon&$top=50&$orderby=modifiedon desc`
    );
    if (resolved.value?.length > 0) {
      const handleTimes = resolved.value.map(r => parseFloat(r.cr7fe_new_handletime)).filter(n => !isNaN(n) && n > 0);
      if (handleTimes.length > 0) {
        const avgRaw = handleTimes.reduce((a, b) => a + b, 0) / handleTimes.length;
        const avgMin = avgRaw > 1000 ? avgRaw / 60 : avgRaw;
        if (avgMin >= 60) {
          const h = Math.floor(avgMin / 60);
          const m = Math.round(avgMin % 60);
          avgResTime = `${h}h ${m}m`;
        } else {
          avgResTime = `${Math.round(avgMin)} min`;
        }
      } else {
        const times = resolved.value.map(r => (new Date(r.modifiedon) - new Date(r.createdon)) / (1000 * 60 * 60)).filter(h => h >= 0 && h < 168);
        if (times.length > 0) {
          const avg = times.reduce((a, b) => a + b, 0) / times.length;
          avgResTime = avg < 1 ? `${Math.round(avg * 60)} min` : `${avg.toFixed(1)} hrs`;
        }
      }
    }
  } catch (err) { errors.push(`${member.name} — ResTime: ${err.message}`); }

  const slaCompliance = (slaMet + slaMissed) > 0 ? Math.min(100, Math.round(slaMet / (slaMet + slaMissed) * 100)) : "N/A";
  const fcrRate = totalCases > 0 ? Math.min(100, Math.round(fcrCases / totalCases * 100)) : "N/A";
  const escalationRate = totalCases > 0 ? Math.min(100, Math.round(escalatedCases / totalCases * 100)) : "N/A";

  return {
    member,
    totalCases, resolvedCases, activeCases, slaMet, slaMissed, slaCompliance,
    casesCreatedBy,
    fcrCases, fcrRate, escalatedCases, escalationRate,
    emailCases, emailResolved,
    totalPhoneCalls, incomingCalls, outgoingCalls, answeredLive, voicemails, memberAHT,
    csatResponses, csatAvg,
    avgResTime: typeof avgResTime === "number" ? `${avgResTime} hrs` : avgResTime,
    errors,
  };
}

async function fetchLiveD365Data(startDate, endDate, onProgress, startTime, endTime) {
  const utcStart = ctToUTC(startDate, startTime || "00:00", false);
  const utcEnd = ctToUTC(endDate, endTime || "23:59", true);
  const s = utcStart.date, sT = utcStart.time, e = utcEnd.date, eT = utcEnd.time;
  const errors = [];
  const progress = (msg) => onProgress?.(`D365: ${msg}`);

  async function safeCount(label, query) {
    try { progress(`Fetching ${label}...`); return await d365Count(query); }
    catch (err) { errors.push(`${label}: ${err.message}`); return 0; }
  }

  async function safeFetch(label, query) {
    try { progress(`Fetching ${label}...`); return await d365Fetch(query); }
    catch (err) { errors.push(`${label}: ${err.message}`); return { value: [] }; }
  }

  async function safeFetchCount(label, query) {
    try {
      progress(`Fetching ${label}...`);
      const cleanQuery = query.replace(/&?\$count=true/g, '').replace(/&?\$top=\d+/g, '');
      const data = await d365Fetch(`${cleanQuery}&$top=5000`);
      const count = data.value?.length ?? 0;
      console.log(`[D365 Count] ${label}: ${count} records`);
      return count;
    } catch (err) {
      errors.push(`${label}: ${err.message}`);
      console.error(`[D365 Count] ${label} ERROR:`, err.message);
      return 0;
    }
  }

  async function safeFetchSLA(label, baseFilter) {
    try {
      progress(`Fetching ${label} SLA...`);
      const data = await d365Fetch(
        `incidents?$filter=${baseFilter}&$select=incidentid&$expand=resolvebykpiid($select=status)&$top=5000`
      );
      let met = 0, missed = 0;
      for (const rec of (data.value || [])) {
        const kpiStatus = rec.resolvebykpiid?.status;
        if (kpiStatus === 4) met++;
        else if (kpiStatus === 1) missed++;
      }
      console.log(`[D365 SLA] ${label}: met=${met}, missed=${missed}, total=${data.value?.length || 0}`);
      return { met, missed };
    } catch (err) {
      errors.push(`${label} SLA: ${err.message}`);
      console.error(`[D365 SLA] ${label} ERROR:`, err.message);
      return { met: 0, missed: 0 };
    }
  }

  async function safeFetchResponseSLA(label, baseFilter) {
    try {
      progress(`Fetching ${label} Response SLA...`);
      const data = await d365Fetch(
        `incidents?$filter=${baseFilter}&$select=incidentid&$expand=firstresponsebykpiid($select=status)&$top=5000`
      );
      let met = 0, missed = 0;
      for (const rec of (data.value || [])) {
        const kpiStatus = rec.firstresponsebykpiid?.status;
        if (kpiStatus === 4) met++;
        else if (kpiStatus === 1) missed++;
      }
      console.log(`[D365 Response SLA] ${label}: met=${met}, missed=${missed}, total=${data.value?.length || 0}`);
      return { met, missed };
    } catch (err) {
      errors.push(`${label} Response SLA: ${err.message}`);
      console.error(`[D365 Response SLA] ${label} ERROR:`, err.message);
      return { met: 0, missed: 0 };
    }
  }

  async function safeFetchOpenBreach(label, baseFilter) {
    try {
      progress(`Fetching ${label} Open Breaches...`);
      const data = await d365Fetch(
        `incidents?$filter=${baseFilter}&$select=incidentid&$expand=resolvebykpiid($select=status)&$top=5000`
      );
      let breached = 0, total = 0;
      for (const rec of (data.value || [])) {
        total++;
        const kpiStatus = rec.resolvebykpiid?.status;
        if (kpiStatus === 1) breached++;
      }
      console.log(`[D365 Open Breach] ${label}: breached=${breached}, active=${total}`);
      return { breached, total };
    } catch (err) {
      errors.push(`${label} Open Breach: ${err.message}`);
      console.error(`[D365 Open Breach] ${label} ERROR:`, err.message);
      return { breached: 0, total: 0 };
    }
  }

  const t1Cases = await safeCount("Tier 1 Cases",
    `incidents?$filter=casetypecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const t1Resolved = await safeFetchCount("T1 Resolved",
    `incidents?$filter=casetypecode eq 1 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t1SLA = await safeFetchSLA("T1",
    `casetypecode eq 1 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}`);
  const t1SLAMet = t1SLA.met;
  const t1SLAMissed = t1SLA.missed;
  const t1ResponseSLA = await safeFetchResponseSLA("T1",
    `casetypecode eq 1 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}`);
  const t1OpenBreach = await safeFetchOpenBreach("T1",
    `casetypecode eq 1 and statecode eq 0 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}`);
  const t1FCR = await safeFetchCount("FCR Cases",
    `incidents?$filter=casetypecode eq 1 and cr7fe_new_fcr eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t1Escalated = await safeCount("Tier 1 Escalated",
    `incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);

  const t2Cases = await safeCount("Tier 2 Cases",
    `incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);
  const t2Resolved = await safeFetchCount("Tier 2 Resolved",
    `incidents?$filter=casetypecode eq 2 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t2SLA = await safeFetchSLA("T2",
    `casetypecode eq 2 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}`);
  const t2SLAMet = t2SLA.met;
  const t2SLAMissed = t2SLA.missed;
  const t2ResponseSLA = await safeFetchResponseSLA("T2",
    `casetypecode eq 2 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}`);
  const t2OpenBreach = await safeFetchOpenBreach("T2",
    `casetypecode eq 2 and statecode eq 0 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}`);
  const t2Escalated = await safeCount("Tier 2 Escalated to T3",
    `incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);

  const t3Cases = await safeCount("Tier 3 Cases",
    `incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);
  const t3Resolved = await safeFetchCount("Tier 3 Resolved",
    `incidents?$filter=casetypecode eq 3 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t3SLA = await safeFetchSLA("T3",
    `casetypecode eq 3 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}`);
  const t3SLAMet = t3SLA.met;
  const t3SLAMissed = t3SLA.missed;
  const t3ResponseSLA = await safeFetchResponseSLA("T3",
    `casetypecode eq 3 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}`);
  const t3OpenBreach = await safeFetchOpenBreach("T3",
    `casetypecode eq 3 and statecode eq 0 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}`);

  const emailCases = await safeCount("Email Cases",
    `incidents?$filter=caseorigincode eq 2 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const emailResponded = await safeCount("Email Responded",
    `incidents?$filter=caseorigincode eq 2 and firstresponsesent eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const emailResolved = await safeFetchCount("Email Resolved",
    `incidents?$filter=caseorigincode eq 2 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);

  const csatResponses = await safeCount("CSAT Responses",
    `incidents?$filter=cr7fe_new_csatresponsereceived eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);

  let csatAvg = "N/A";
  if (csatResponses > 0) {
    try {
      progress("Fetching CSAT Scores...");
      const csatData = await d365Fetch(
        `incidents?$filter=cr7fe_new_csatresponsereceived eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=cr7fe_new_csatscore`
      );
      if (csatData.value?.length > 0) {
        const scores = csatData.value.map(r => parseFloat(r.cr7fe_new_csatscore)).filter(n => !isNaN(n));
        if (scores.length > 0) {
          csatAvg = +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1);
        }
      }
    } catch (err) { errors.push(`CSAT Scores: ${err.message}`); }
  }

  let avgResTime = "N/A";
  try {
    progress("Fetching resolution times...");
    const resolved = await d365Fetch(
      `incidents?$filter=casetypecode eq 1 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid,cr7fe_new_handletime,createdon,modifiedon&$top=5000&$orderby=modifiedon desc`
    );
    if (resolved.value?.length > 0) {
      const handleTimes = resolved.value.map(r => parseFloat(r.cr7fe_new_handletime)).filter(n => !isNaN(n) && n > 0);
      if (handleTimes.length > 0) {
        const avgRaw = handleTimes.reduce((a, b) => a + b, 0) / handleTimes.length;
        const avgMin = avgRaw > 1000 ? avgRaw / 60 : avgRaw;
        if (avgMin >= 60) {
          const h = Math.floor(avgMin / 60);
          const m = Math.round(avgMin % 60);
          avgResTime = `${h}h ${m}m`;
        } else {
          avgResTime = `${Math.round(avgMin)} min`;
        }
      } else {
        const times = resolved.value.map(r => {
          const created = new Date(r.createdon);
          const modified = new Date(r.modifiedon);
          return (modified - created) / (1000 * 60 * 60);
        }).filter(h => h >= 0 && h < 168);
        if (times.length > 0) {
          const avg = times.reduce((a, b) => a + b, 0) / times.length;
          if (avg < 1) { avgResTime = `${Math.round(avg * 60)} min`; }
          else { avgResTime = `${avg.toFixed(1)} hrs`; }
        }
      }
    }
  } catch (err) { errors.push(`Resolution time: ${err.message}`); }

  const phoneTotal = await safeFetchCount("Phone Total",
    `phonecalls?$filter=actualstart ge ${s}T${sT} and actualstart le ${e}T${eT}&$select=actualdurationminutes`);
  const phoneAnswered = await safeFetchCount("Phone Answered",
    `phonecalls?$filter=actualstart ge ${s}T${sT} and actualstart le ${e}T${eT} and actualdurationminutes gt 0&$select=actualdurationminutes`);
  const phoneAbandoned = await safeFetchCount("Phone Abandoned",
    `phonecalls?$filter=actualstart ge ${s}T${sT} and actualstart le ${e}T${eT} and actualdurationminutes eq 0&$select=actualdurationminutes`);
  const phoneAnswerRate = phoneTotal > 0 ? Math.min(100, Math.round(phoneAnswered / phoneTotal * 100)) : 0;

  let phoneAHT = "N/A";
  try {
    progress("Fetching Phone AHT...");
    const ahtData = await d365Fetch(
      `phonecalls?$filter=actualstart ge ${s}T${sT} and actualstart le ${e}T${eT}&$select=actualdurationminutes&$top=5000`
    );
    if (ahtData.value?.length > 0) {
      const durations = ahtData.value.map(r => parseFloat(r.actualdurationminutes) || 0).filter(n => !isNaN(n));
      if (durations.length > 0) {
        const avg = Math.round(durations.reduce((a, b) => a + b, 0) / durations.length / 60);
        phoneAHT = `${avg} min`;
      }
    }
  } catch (err) { errors.push(`Phone AHT: ${err.message}`); }

  let timelineData = [];
  try {
    const startD = new Date(s); const endD = new Date(e);
    const dayCount = Math.round((endD - startD) / (1000 * 60 * 60 * 24)) + 1;
    if (dayCount >= 2 && dayCount <= 90) {
      progress("Building timeline...");
      const [t1Raw, t2Raw, t3Raw, phoneRaw, csatRaw, slaRaw] = await Promise.all([
        d365Fetch(`incidents?$filter=casetypecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=createdon&$top=5000`),
        d365Fetch(`incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$select=escalatedon&$top=5000`),
        d365Fetch(`incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$select=escalatedon&$top=5000`),
        d365Fetch(`phonecalls?$filter=actualstart ge ${s}T${sT} and actualstart le ${e}T${eT}&$select=actualstart&$top=5000`),
        d365Fetch(`incidents?$filter=cr7fe_new_csatresponsereceived eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=createdon,cr7fe_new_csatscore&$top=5000`),
        d365Fetch(`incidents?$filter=casetypecode eq 1 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=createdon&$top=5000`),
      ]);
      const bucket = (records, dateField) => {
        const map = {};
        (records.value || []).forEach(r => { const d = (r[dateField] || "").slice(0, 10); map[d] = (map[d] || 0) + 1; });
        return map;
      };
      const bucketAvg = (records, dateField, valField) => {
        const map = {};
        (records.value || []).forEach(r => {
          const d = (r[dateField] || "").slice(0, 10);
          const v = parseFloat(r[valField]);
          if (!isNaN(v)) { if (!map[d]) map[d] = []; map[d].push(v); }
        });
        const avg = {};
        Object.keys(map).forEach(d => { avg[d] = +(map[d].reduce((a, b) => a + b, 0) / map[d].length).toFixed(1); });
        return avg;
      };
      const t1Map = bucket(t1Raw, "createdon");
      const t2Map = bucket(t2Raw, "escalatedon");
      const t3Map = bucket(t3Raw, "escalatedon");
      const phoneMap = bucket(phoneRaw, "actualstart");
      const csatMap = bucketAvg(csatRaw, "createdon", "cr7fe_new_csatscore");
      const slaMap = bucket(slaRaw, "createdon");
      for (let i = 0; i < dayCount; i++) {
        const d = new Date(startD); d.setDate(d.getDate() + i);
        const key = d.toISOString().slice(0, 10);
        const label = `${d.getMonth() + 1}/${d.getDate()}`;
        const t1 = t1Map[key] || 0;
        const slaDay = slaMap[key] || 0;
        timelineData.push({
          date: label, t1Cases: t1, t2Cases: t2Map[key] || 0, t3Cases: t3Map[key] || 0,
          sla: t1 > 0 ? Math.round(slaDay / t1 * 100) : 0,
          calls: phoneMap[key] || 0, csat: csatMap[key] || 0,
        });
      }
    }
  } catch (err) { errors.push(`Timeline: ${err.message}`); }

  const allCases = t1Cases + t2Cases + t3Cases;
  const allResolved = t1Resolved + t2Resolved + t3Resolved;

  return {
    tier1: { total: t1Cases, slaMet: t1SLAMet, slaMissed: t1SLAMissed, slaCompliance: (t1SLAMet + t1SLAMissed) > 0 ? Math.min(100, Math.round(t1SLAMet / (t1SLAMet + t1SLAMissed) * 100)) : "N/A", slaResponseMet: t1ResponseSLA.met, slaResponseMissed: t1ResponseSLA.missed, slaResponseCompliance: (t1ResponseSLA.met + t1ResponseSLA.missed) > 0 ? Math.min(100, Math.round(t1ResponseSLA.met / (t1ResponseSLA.met + t1ResponseSLA.missed) * 100)) : "N/A", openBreachCount: t1OpenBreach.breached, openBreachTotal: t1OpenBreach.total, openBreachRate: t1OpenBreach.total > 0 ? Math.min(100, Math.round(t1OpenBreach.breached / t1OpenBreach.total * 100)) : 0, fcrRate: t1Cases ? Math.min(100, Math.round(t1FCR / t1Cases * 100)) : 0, escalationRate: t1Cases ? Math.min(100, Math.round(t1Escalated / t1Cases * 100)) : 0, avgResolutionTime: avgResTime, escalated: t1Escalated },
    tier2: { total: t2Cases, resolved: t2Resolved, slaMet: t2SLAMet, slaMissed: t2SLAMissed, slaCompliance: (t2SLAMet + t2SLAMissed) > 0 ? Math.min(100, Math.round(t2SLAMet / (t2SLAMet + t2SLAMissed) * 100)) : "N/A", slaResponseMet: t2ResponseSLA.met, slaResponseMissed: t2ResponseSLA.missed, slaResponseCompliance: (t2ResponseSLA.met + t2ResponseSLA.missed) > 0 ? Math.min(100, Math.round(t2ResponseSLA.met / (t2ResponseSLA.met + t2ResponseSLA.missed) * 100)) : "N/A", openBreachCount: t2OpenBreach.breached, openBreachTotal: t2OpenBreach.total, openBreachRate: t2OpenBreach.total > 0 ? Math.min(100, Math.round(t2OpenBreach.breached / t2OpenBreach.total * 100)) : 0, escalationRate: t2Cases ? Math.min(100, Math.round(t2Escalated / t2Cases * 100)) : "N/A", escalated: t2Escalated },
    tier3: { total: t3Cases, resolved: t3Resolved, slaMet: t3SLAMet, slaMissed: t3SLAMissed, slaCompliance: (t3SLAMet + t3SLAMissed) > 0 ? Math.min(100, Math.round(t3SLAMet / (t3SLAMet + t3SLAMissed) * 100)) : "N/A", slaResponseMet: t3ResponseSLA.met, slaResponseMissed: t3ResponseSLA.missed, slaResponseCompliance: (t3ResponseSLA.met + t3ResponseSLA.missed) > 0 ? Math.min(100, Math.round(t3ResponseSLA.met / (t3ResponseSLA.met + t3ResponseSLA.missed) * 100)) : "N/A", openBreachCount: t3OpenBreach.breached, openBreachTotal: t3OpenBreach.total, openBreachRate: t3OpenBreach.total > 0 ? Math.min(100, Math.round(t3OpenBreach.breached / t3OpenBreach.total * 100)) : 0 },
    email: { total: emailCases, responded: emailResponded, resolved: emailResolved, slaCompliance: emailResolved > 0 ? 100 : (emailCases > 0 ? 0 : "N/A") },
    csat: { responses: csatResponses, avgScore: csatAvg },
    phone: { totalCalls: phoneTotal, incoming: phoneTotal, outgoing: 0, answered: phoneAnswered, abandoned: phoneAbandoned, voicemails: 0, answerRate: phoneAnswerRate, avgAHT: phoneAHT },
    overall: { created: allCases, resolved: allResolved, csatResponses, answeredCalls: phoneAnswered, abandonedCalls: phoneAbandoned },
    timeline: timelineData,
    source: "d365",
    errors,
  };
}

async function fetchLiveData(config, startDate, endDate, onProgress, startTime, endTime) {
  const progress = (msg) => onProgress?.(msg);
  progress("Connecting to Dynamics 365...");
  const d365Data = await fetchLiveD365Data(startDate, endDate, progress, startTime, endTime);
  progress("Compiling report...");
  return {
    ...d365Data,
    phone: d365Data.phone || { totalCalls: 0, answered: 0, abandoned: 0, answerRate: 0, avgAHT: 0 },
    source: "live",
    errors: d365Data.errors || [],
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
   UI COMPONENTS
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

function MetricCard({ label, value, target, unit, inverse, color }) {
  const numVal = parseFloat(value);
  const numTarget = parseFloat(target);
  const isNA = value === "N/A" || (typeof value === "string" && value === "N/A") || (typeof value !== "string" && isNaN(numVal));
  const noTarget = target === null || target === undefined;

  if (noTarget) {
    const displayVal = isNA ? "N/A" : (typeof value === "string" ? value : `${value}${unit || ""}`);
    return (
      <div style={{ background: C.card, borderRadius: 12, border: "none", padding: "18px 20px", position: "relative", overflow: "hidden",  }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.textDark }}>{label}</div>
        </div>
        <div style={{ height: 10, marginBottom: 8 }} />
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
          <span style={{ fontSize: 20, fontWeight: 800, color: C.textDark, fontFamily: "'Space Mono', monospace" }}>{displayVal}</span>
        </div>
      </div>
    );
  }

  let pct = isNA ? 0 : inverse ? Math.min(100, (numTarget / Math.max(numVal, 0.01)) * 100) : Math.min(100, (numVal / Math.max(numTarget, 0.01)) * 100);
  if (unit === " min" || unit === " hrs") pct = isNA ? 0 : numVal <= numTarget ? 100 : Math.max(0, 100 - ((numVal - numTarget) / numTarget) * 100);
  const met = isNA ? null : inverse ? numVal <= numTarget : numVal >= numTarget;
  const barColor = isNA ? C.gray : met ? "#2D9D78" : numVal >= numTarget * (inverse ? 0.85 : 0.8) ? C.orange : "#E5544B";
  const pillColor = isNA ? C.gray : met ? "#2D9D78" : "#E5544B";
  const displayVal = isNA ? "N/A" : `${value}${unit || ""}`;

  return (
    <div style={{ background: C.card, borderRadius: 12, border: "none", padding: "18px 20px", position: "relative", overflow: "hidden",  }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: C.textDark }}>{label}</div>
        <div style={{ padding: "3px 10px", borderRadius: 10, background: pillColor, color: "#fff", fontSize: 11, fontWeight: 700, fontFamily: "'Space Mono', monospace" }}>{met === null ? "N/A" : met ? "Met" : "Miss"}</div>
      </div>
      <div style={{ height: 10, background: `${barColor}18`, borderRadius: 5, overflow: "hidden", marginBottom: 8 }}>
        <div style={{ height: "100%", width: `${Math.min(100, Math.max(2, pct))}%`, background: barColor, borderRadius: 5, transition: "width 0.6s ease" }} />
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
        <span style={{ fontSize: 20, fontWeight: 800, color: C.textDark, fontFamily: "'Space Mono', monospace" }}>{displayVal}</span>
        <span style={{ fontSize: 11, color: C.textLight }}>Target: {target}{unit || ""}</span>
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, color }) {
  return (
    <div style={{ flex: 1, minWidth: 100, background: C.card, borderRadius: 10, border: "none", padding: "14px 16px", textAlign: "center",  }}>
      <div style={{ fontSize: 10, fontWeight: 600, color: C.textLight, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 800, color: color || C.textDark, fontFamily: "'Space Mono', monospace", lineHeight: 1.1 }}>{value}</div>
      {sub && <div style={{ fontSize: 10, color: C.textLight, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function TierSection({ tier, data, members }) {
  const t = TIERS[tier]; if (!t) return null;
  const d = data[`tier${tier}`]; if (!d) return null;
  const tierMembers = (members || []).filter(m => m.tier === tier);
  const slaRate = d.slaCompliance;
  const slaMet = d.slaMet || 0;
  const slaMissed = d.slaMissed || 0;
  const slaTotal = slaMet + slaMissed;
  const metrics = [];
  if (t.metrics.includes("sla_compliance")) metrics.push({ label: "SLA Compliance", value: slaRate, target: 90, unit: "%" });
  if (t.metrics.includes("sla_response")) metrics.push({ label: "Response SLA", value: d.slaResponseCompliance, target: 90, unit: "%" });
  if (t.metrics.includes("open_breach_rate")) metrics.push({ label: "Open SLA Breach", value: d.openBreachRate, target: 5, unit: "%", inverse: true, sub: d.openBreachCount > 0 ? `${d.openBreachCount} of ${d.openBreachTotal} active` : null });
  if (t.metrics.includes("fcr_rate")) metrics.push({ label: "First Call Resolution", value: d.fcrRate, target: 90, unit: "%" });
  if (t.metrics.includes("escalation_rate")) metrics.push({ label: "Escalation Rate", value: d.escalationRate, target: 10, unit: "%", inverse: true });
  if (t.metrics.includes("avg_resolution_time")) {
    const raw = d.avgResolutionTime || "N/A";
    const num = parseFloat(raw);
    metrics.push({ label: "Avg Resolution Time", value: isNaN(num) ? "N/A" : raw, target: null, unit: "", rawDisplay: true });
  }
  if (tier === 1 && data.csat) {
    metrics.push({ label: "CSAT Score", value: data.csat.avgScore, target: 4.0, unit: "/5" });
  }
  return (
    <div style={{ marginBottom: 24 }}>
      <div style={{ background: `linear-gradient(135deg, ${t.color}, ${t.colorDark})`, borderRadius: "14px 14px 0 0", padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          <h3 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: "#fff" }}>{t.label} {t.name}</h3>
          <div style={{ fontSize: 12, color: "rgba(255,255,255,0.7)", marginTop: 4 }}>{t.desc}</div>
        </div>
        {tierMembers.length > 0 && (
          <div style={{ display: "flex", gap: -8 }}>
            {tierMembers.slice(0, 5).map((m, i) => (
              <div key={m.id} style={{ width: 38, height: 38, borderRadius: 10, background: "rgba(255,255,255,0.2)", border: "2px solid rgba(255,255,255,0.4)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, color: "#fff", marginLeft: i > 0 ? -6 : 0 }}>{m.avatar}</div>
            ))}
            {tierMembers.length > 5 && <div style={{ width: 38, height: 38, borderRadius: 10, background: "rgba(255,255,255,0.15)", border: "2px solid rgba(255,255,255,0.3)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 600, color: "#fff", marginLeft: -6 }}>+{tierMembers.length - 5}</div>}
          </div>
        )}
      </div>
      <div className="stat-row" style={{ display: "flex", gap: 10, padding: "16px 0", overflowX: "auto" }}>
        <StatCard label={`${t.label} SLA Rate`} value={slaRate === "N/A" ? "N/A" : `${slaRate}%`} color={slaRate !== "N/A" && slaRate >= 90 ? "#2D9D78" : "#E5544B"} />
        <StatCard label="SLAs Met" value={`${slaMet}/${slaTotal}`} color="#2D9D78" />
        <StatCard label="SLAs Missed" value={`${slaMissed}/${slaTotal}`} color={slaMissed > 0 ? "#E5544B" : "#2D9D78"} />
        <StatCard label="Total Cases" value={d.total} color={t.colorDark} />
        <StatCard label="Metrics" value={metrics.length} color={C.textMid} />
      </div>
      <div style={{ fontSize: 14, fontWeight: 600, color: C.textDark, marginBottom: 12 }}>{t.label} SLA Status — All Metrics</div>
      <div className="metric-grid-2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {metrics.map((m, i) => (<MetricCard key={i} label={m.label} value={m.value} target={m.target} unit={m.unit} inverse={m.inverse} />))}
      </div>
      {tier === 1 && data.phone && (() => {
        const totalCalls = data.phone.totalCalls ?? 0;
        const answered = data.phone.answered ?? 0;
        const abandoned = data.phone.abandoned ?? 0;
        const answerRate = data.phone.answerRate ?? 0;
        const avgAHT = data.phone.avgAHT ?? "N/A";
        const MR = ({ icon, label, value, accent, badge }) => (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 0", borderBottom: `1px solid ${C.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 15 }}>{icon}</span><span style={{ fontSize: 13, color: C.textMid }}>{label}</span></div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 18, fontWeight: 700, color: accent || C.textDark, fontFamily: "'Space Mono', monospace" }}>{value}</span>
              {badge && <span style={{ fontSize: 11, width: 18, height: 18, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", background: badge === "met" ? "#2D9D78" : "#E5544B", color: "#fff" }}>{badge === "met" ? "✓" : "!"}</span>}
            </div>
          </div>
        );
        return (<>
          <div style={{ marginTop: 16, background: C.card, borderRadius: 12, border: "none", padding: "18px 20px",  }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#E91E63", marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 18 }}>📞</span> PHONE METRICS</div>
            <MR icon="📞" label="Total Calls" value={totalCalls} accent={C.textDark} />
            <MR icon="✅" label="Answered Calls" value={answered} accent="#2D9D78" badge="met" />
            <MR icon="❌" label="Abandoned Calls" value={abandoned} accent="#E5544B" badge={abandoned > 0 ? "miss" : "met"} />
            <MR icon="📊" label="Answer Rate" value={`${answerRate}%`} accent={answerRate >= 95 ? "#2D9D78" : "#E5544B"} badge={answerRate >= 95 ? "met" : "miss"} />
            <MR icon="⏱️" label="Avg Phone AHT" value={avgAHT} accent={C.textMid} />
          </div>
          {data.email && (
            <div style={{ marginTop: 16, background: C.card, borderRadius: 12, border: "none", padding: "18px 20px",  }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.blue, marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 18 }}>📧</span> EMAIL METRICS</div>
              <MR icon="📨" label="Total Email Cases" value={data.email.total ?? 0} accent={C.textDark} />
              <MR icon="💬" label="Responded" value={data.email.responded ?? 0} accent={C.orange} badge={(data.email.responded ?? 0) > 0 ? "miss" : "met"} />
              <MR icon="✅" label="Resolved" value={data.email.resolved ?? 0} accent="#2D9D78" badge="met" />
            </div>
          )}
        </>);
      })()}
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
    { label: "Abandoned", value: d.abandonedCalls, color: d.abandonedCalls > 0 ? "#f44336" : "#81C784" },
  ];
  return (
    <div style={{ background: C.primary, padding: "24px 28px", borderRadius: 14, marginBottom: 4 }}>
      <h3 style={{ margin: "0 0 20px", color: "#fff", fontSize: 16, fontWeight: 700, textAlign: "center" }}>📈 OVERALL SUMMARY</h3>
      <div style={{ display: "flex", justifyContent: "space-around", flexWrap: "wrap", gap: 12 }}>
        {items.map((it) => (<div key={it.label} style={{ textAlign: "center" }}><div style={{ fontSize: 32, fontWeight: 700, color: it.color, fontFamily: "'Space Mono', monospace" }}>{it.value}</div><div style={{ fontSize: 11, color: "#a8c6df", marginTop: 2 }}>{it.label}</div></div>))}
      </div>
    </div>
  );
}

function Definitions() {
  const defs = [
    ["SLA Compliance", "Percentage of resolved cases that were completed within the allowed time frame. Formula: SLAs Met ÷ (SLAs Met + SLAs Missed) × 100. Target: 90%"],
    ["SLAs Met", "Cases that were resolved before the deadline based on their priority level (e.g. High = 24hrs, Normal = 48hrs)"],
    ["SLAs Missed", "Cases that were resolved but took longer than the allowed time for their priority level"],
    ["Response SLA", "Percentage of cases where a first response was sent to the customer within the required time frame. Target: 90%"],
    ["Open SLA Breach", "Percentage of currently active cases that have already exceeded their resolution deadline — these are overdue right now. Target: below 5%"],
    ["FCR Rate", "First Contact Resolution — cases resolved without escalation or follow-up"],
    ["Escalation Rate", "Percentage of Tier 1 cases escalated to Tier 2 or Tier 3"],
    ["Avg Resolution Time", "Mean time from case creation to resolution for closed cases"],
    ["Answer Rate", "Percentage of calls answered vs. total calls"],
    ["AHT", "Average Handle Time — mean duration of phone calls"],
    ["CSAT Score", "Customer Satisfaction rating (1-5 scale). Target: 4.0+"],
  ];
  return (
    <div style={{ background: C.grayLight, padding: "24px 28px", borderTop: `1px solid ${C.border}`, borderRadius: "0 0 14px 14px", marginTop: 16 }}>
      <h4 style={{ margin: "0 0 14px", fontSize: 13, color: "#555", fontWeight: 700 }}>📝 DEFINITIONS & METHODOLOGY</h4>
      <table cellPadding="3" cellSpacing="0" style={{ width: "100%", fontSize: 12, color: "#666" }}>
        <tbody>{defs.map(([term, def]) => (<tr key={term}><td style={{ fontWeight: 700, width: 150, verticalAlign: "top", padding: "4px 0" }}>{term}</td><td style={{ padding: "4px 0" }}>{def}</td></tr>))}</tbody>
      </table>
      <div style={{ borderTop: `1px solid ${C.border}`, marginTop: 14, paddingTop: 12, fontSize: 11, color: "#888", lineHeight: 1.8 }}>
        <div>✅ Met &nbsp;|&nbsp; ⚠️ Approaching &nbsp;|&nbsp; 🔴 Miss &nbsp;|&nbsp; ➖ N/A</div>
        <div>📊 <strong>Data Sources:</strong> Microsoft Dynamics 365 Customer Service</div>
      </div>
    </div>
  );
}

function MemberSection({ memberData, index }) {
  const d = memberData;
  const m = d.member;
  const isTier1 = m.tier === 1;
  const colors = [C.blue, C.accent, C.purple, "#2D9D78", C.gold, "#E91E63", "#00BCD4", "#795548"];
  const color = colors[index % colors.length];
  const colorDark = color + "DD";
  const slaMet = d.slaMet || 0;
  const slaMissed = d.slaMissed || 0;
  const slaTotal = slaMet + slaMissed;
  const metrics = isTier1 ? [
    { label: "SLA Compliance", value: d.slaCompliance, target: 90, unit: "%" },
    { label: "First Call Resolution", value: d.fcrRate, target: 90, unit: "%" },
    { label: "Escalation Rate", value: d.escalationRate, target: 10, unit: "%", inverse: true },
    { label: "Avg Resolution Time", value: d.avgResTime || "N/A", target: null, unit: "", rawDisplay: true },
    { label: "CSAT Score", value: d.csatAvg, target: 4.0, unit: "/5" },
  ] : [
    { label: "SLA Compliance", value: d.slaCompliance, target: 90, unit: "%" },
    { label: "Escalation Rate", value: d.escalationRate, target: 10, unit: "%", inverse: true },
    { label: "Avg Resolution Time", value: d.avgResTime || "N/A", target: null, unit: "", rawDisplay: true },
  ];
  const PhoneStat = ({ icon, label, value, accent }) => (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.border}` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 16 }}>{icon}</span><span style={{ fontSize: 13, color: C.textMid }}>{label}</span></div>
      <span style={{ fontSize: 18, fontWeight: 700, color: accent || C.textDark, fontFamily: "'Space Mono', monospace" }}>{value}</span>
    </div>
  );
  return (
    <div style={{ marginBottom: 24 }}>
      <div style={{ background: `linear-gradient(135deg, ${color}, ${colorDark})`, borderRadius: "14px 14px 0 0", padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ width: 48, height: 48, borderRadius: 12, background: "rgba(255,255,255,0.2)", border: "2px solid rgba(255,255,255,0.4)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 700, color: "#fff" }}>{m.avatar}</div>
          <div><div style={{ fontSize: 20, fontWeight: 700, color: "#fff" }}>{m.name}</div><div style={{ fontSize: 12, color: "rgba(255,255,255,0.7)", marginTop: 2 }}>{m.role}{m.email ? ` · ${m.email}` : ""}</div></div>
        </div>
      </div>
      <div className="stat-row" style={{ display: "flex", gap: 10, padding: "16px 0", overflowX: "auto" }}>
        <StatCard label="Cases Owned" value={d.totalCases} color={color} />
        <StatCard label="Cases Created" value={d.casesCreatedBy ?? "—"} color={C.blue} />
        <StatCard label="Resolved" value={d.resolvedCases} color="#2D9D78" />
        <StatCard label="Active" value={d.activeCases} color={C.blue} />
        <StatCard label="SLAs Met" value={`${slaMet}/${slaTotal}`} color={slaMet > 0 ? "#2D9D78" : "#E5544B"} />
      </div>
      <div style={{ fontSize: 14, fontWeight: 600, color: C.textDark, marginBottom: 12 }}>Performance Metrics</div>
      <div className="metric-grid-2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {metrics.map((mt, i) => (<MetricCard key={i} label={mt.label} value={mt.value} target={mt.target} unit={mt.unit} inverse={mt.inverse} />))}
      </div>
      {isTier1 && (
        <div style={{ marginTop: 16, background: C.card, borderRadius: 12, border: "none", padding: "18px 20px",  }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#E91E63", marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 18 }}>📞</span> Phone Activity</div>
          <PhoneStat icon="📞" label="Total Calls" value={d.totalPhoneCalls ?? 0} accent={C.textDark} />
          <PhoneStat icon="✅" label="Answered Calls" value={d.answeredLive ?? 0} accent="#2D9D78" />
          <PhoneStat icon="❌" label="Abandoned Calls" value={d.voicemails ?? 0} accent="#E5544B" />
          <PhoneStat icon="⏱️" label="Avg Phone AHT" value={d.memberAHT ?? "N/A"} accent={C.textMid} />
        </div>
      )}
    </div>
  );
}

function TeamSummary({ memberDataList }) {
  const totals = memberDataList.reduce((acc, d) => ({
    totalCases: acc.totalCases + d.totalCases, resolved: acc.resolved + d.resolvedCases,
    active: acc.active + d.activeCases, escalated: acc.escalated + d.escalatedCases,
    slaMet: acc.slaMet + d.slaMet, slaMissed: acc.slaMissed + (d.slaMissed || 0), fcrCases: acc.fcrCases + d.fcrCases,
    emailCases: acc.emailCases + d.emailCases, emailResolved: acc.emailResolved + d.emailResolved,
    csatResponses: acc.csatResponses + d.csatResponses,
    csatTotal: acc.csatTotal + (d.csatAvg !== "N/A" ? d.csatAvg * d.csatResponses : 0),
  }), { totalCases: 0, resolved: 0, active: 0, escalated: 0, slaMet: 0, slaMissed: 0, fcrCases: 0, emailCases: 0, emailResolved: 0, csatResponses: 0, csatTotal: 0 });
  const items = [
    { label: "Total Cases", value: totals.totalCases, color: "#4FC3F7" },
    { label: "Resolved", value: totals.resolved, color: "#81C784" },
    { label: "Active", value: totals.active, color: "#64B5F6" },
    { label: "Escalated", value: totals.escalated, color: totals.escalated > 0 ? "#f44336" : "#81C784" },
    { label: "SLA Met", value: totals.slaMet, color: "#81C784" },
  ];
  return (
    <div style={{ background: C.primary, padding: "24px 28px", borderRadius: 14, marginBottom: 16 }}>
      <h3 style={{ margin: "0 0 20px", color: "#fff", fontSize: 16, fontWeight: 700, textAlign: "center" }}>📈 TEAM SUMMARY — {memberDataList.length} Member{memberDataList.length > 1 ? "s" : ""}</h3>
      <div style={{ display: "flex", justifyContent: "space-around", flexWrap: "wrap", gap: 12 }}>
        {items.map((it) => (<div key={it.label} style={{ textAlign: "center" }}><div style={{ fontSize: 32, fontWeight: 700, color: it.color, fontFamily: "'Space Mono', monospace" }}>{it.value}</div><div style={{ fontSize: 11, color: "#a8c6df", marginTop: 2 }}>{it.label}</div></div>))}
      </div>
      {totals.totalCases > 0 && (
        <div style={{ display: "flex", justifyContent: "center", gap: 24, marginTop: 18, paddingTop: 14, borderTop: "1px solid rgba(255,255,255,0.1)" }}>
          <div style={{ textAlign: "center" }}><StatusBadge status={checkTarget("sla_compliance", (totals.slaMet + totals.slaMissed) > 0 ? Math.min(100, Math.round(totals.slaMet / (totals.slaMet + totals.slaMissed) * 100)) : 0)} value={(totals.slaMet + totals.slaMissed) > 0 ? Math.min(100, Math.round(totals.slaMet / (totals.slaMet + totals.slaMissed) * 100)) : "N/A"} unit="%" /><div style={{ fontSize: 10, color: "#a8c6df", marginTop: 4 }}>Team SLA</div></div>
          <div style={{ textAlign: "center" }}><StatusBadge status={checkTarget("fcr_rate", totals.totalCases ? Math.min(100, Math.round(totals.fcrCases / totals.totalCases * 100)) : 0)} value={totals.totalCases ? Math.min(100, Math.round(totals.fcrCases / totals.totalCases * 100)) : 0} unit="%" /><div style={{ fontSize: 10, color: "#a8c6df", marginTop: 4 }}>Team FCR</div></div>
          {totals.csatResponses > 0 && <div style={{ textAlign: "center" }}><StatusBadge status={checkTarget("csat_score", +(totals.csatTotal / totals.csatResponses).toFixed(1))} value={+(totals.csatTotal / totals.csatResponses).toFixed(1)} unit="/5" /><div style={{ fontSize: 10, color: "#a8c6df", marginTop: 4 }}>Team CSAT</div></div>}
        </div>
      )}
    </div>
  );
}

function ChartsPanel({ data }) {
  const tl = data.timeline;
  if (!tl || tl.length < 2) return null;
  const interval = Math.max(0, Math.floor(tl.length / 8));
  return (
    <div className="chart-grid-2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginTop: 20 }}>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: "none",  }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>📊 Daily Cases by Tier</div>
        <ResponsiveContainer width="100%" height={220}><BarChart data={tl}><CartesianGrid strokeDasharray="3 3" stroke={C.border} /><XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} /><YAxis fontSize={10} tick={{ fill: C.textLight }} /><Tooltip content={<CTooltip />} /><Legend iconType="circle" iconSize={7} formatter={(v) => <span style={{ fontSize: 10, color: C.textMid }}>{v}</span>} /><Bar dataKey="t1Cases" name="Tier 1" fill={TIERS[1].color} radius={[3,3,0,0]} barSize={14} /><Bar dataKey="t2Cases" name="Tier 2" fill={TIERS[2].color} radius={[3,3,0,0]} barSize={14} /><Bar dataKey="t3Cases" name="Tier 3" fill={TIERS[3].color} radius={[3,3,0,0]} barSize={14} /></BarChart></ResponsiveContainer>
      </div>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: "none",  }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>📈 SLA Compliance Trend</div>
        <ResponsiveContainer width="100%" height={220}><AreaChart data={tl}><defs><linearGradient id="slaG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.green} stopOpacity={0.3} /><stop offset="100%" stopColor={C.green} stopOpacity={0.02} /></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke={C.border} /><XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} /><YAxis fontSize={10} tick={{ fill: C.textLight }} domain={[50, 100]} /><Tooltip content={<CTooltip />} /><Area type="monotone" dataKey="sla" name="SLA %" stroke={C.green} fill="url(#slaG)" strokeWidth={2} /></AreaChart></ResponsiveContainer>
      </div>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: "none",  }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>📞 Daily Call Volume</div>
        <ResponsiveContainer width="100%" height={220}><AreaChart data={tl}><defs><linearGradient id="callG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.blue} stopOpacity={0.3} /><stop offset="100%" stopColor={C.blue} stopOpacity={0.02} /></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke={C.border} /><XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} /><YAxis fontSize={10} tick={{ fill: C.textLight }} /><Tooltip content={<CTooltip />} /><Area type="monotone" dataKey="calls" name="Calls" stroke={C.blue} fill="url(#callG)" strokeWidth={2} /></AreaChart></ResponsiveContainer>
      </div>
      <div style={{ background: C.card, borderRadius: 14, padding: 20, border: "none",  }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textDark, marginBottom: 14 }}>⭐ CSAT Score Trend</div>
        <ResponsiveContainer width="100%" height={220}><AreaChart data={tl}><defs><linearGradient id="csatG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.gold} stopOpacity={0.3} /><stop offset="100%" stopColor={C.gold} stopOpacity={0.02} /></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke={C.border} /><XAxis dataKey="date" fontSize={9} tick={{ fill: C.textLight }} interval={interval} /><YAxis fontSize={10} tick={{ fill: C.textLight }} domain={[1, 5]} /><Tooltip content={<CTooltip />} /><Area type="monotone" dataKey="csat" name="CSAT" stroke={C.gold} fill="url(#csatG)" strokeWidth={2} /></AreaChart></ResponsiveContainer>
      </div>
    </div>
  );
}

function MultiMemberSelect({ selected, onChange, members }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef(null);
  const inputRef = useRef(null);
  useEffect(() => { const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }; document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h); }, []);
  useEffect(() => { if (open && inputRef.current) inputRef.current.focus(); }, [open]);
  const toggle = (id) => { if (id === "__all") { onChange(selected.length === members.length ? [] : members.map((m) => m.id)); } else { onChange(selected.includes(id) ? selected.filter((s) => s !== id) : [...selected, id]); } };
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
                <div style={{ textAlign: "left" }}><div style={{ fontWeight: 500, display: "flex", alignItems: "center", gap: 5 }}>{m.name}<span style={{ fontSize: 8, fontWeight: 700, padding: "1px 5px", borderRadius: 3, background: (t?.color || C.blue) + "22", color: t?.color || C.blue }}>T{m.tier}</span></div><div style={{ fontSize: 10, color: C.textLight }}>{m.role}</div></div>
              </button>
            ); })}
          </div>
        </div>
      )}
    </div>
  );
}

function ConnectionBar({ d365Connected, isLive, onOpenSettings }) {
  return (
    <div className="conn-bar" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 28px", background: C.card, borderBottom: `1px solid ${C.border}`, fontSize: 11 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
        <span style={{ fontWeight: 600, color: C.textLight, fontSize: 10, textTransform: "uppercase", letterSpacing: 1 }}>Data Sources</span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 6, height: 6, borderRadius: "50%", background: d365Connected ? C.green : C.accent }} /><span style={{ fontWeight: 600, color: d365Connected ? C.green : C.accent }}>D365</span><span style={{ color: C.textLight }}>{d365Connected ? "Connected" : "Not connected"}</span></span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 6, height: 6, borderRadius: "50%", background: C.green }} /><span style={{ fontWeight: 600, color: C.green }}>Live</span></span>
      </div>
      <button onClick={onOpenSettings} style={{ background: "none", border: "none", fontSize: 11, fontWeight: 600, color: C.primary, cursor: "pointer", textDecoration: "underline" }}>⚙️ Configure</button>
    </div>
  );
}

function SettingsModal({ show, onClose, config, onSave, d365Account, onD365Login, onD365Logout }) {
  const [local, setLocal] = useState(config);
  const [d365Status, setD365Status] = useState(null);
  const [signingIn, setSigningIn] = useState(false);
  useEffect(() => { setLocal(config); }, [config]);
  if (!show) return null;
  const handleD365SignIn = async () => {
    setSigningIn(true); setD365Status(null);
    const result = await onD365Login();
    if (result) { setD365Status({ success: true, name: result.account?.name }); }
    else { setD365Status({ success: false, error: "Sign-in cancelled or failed" }); }
    setSigningIn(false);
  };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.55)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background: C.card, borderRadius: 20, width: 660, maxHeight: "90vh", overflow: "auto", boxShadow: "0 24px 80px rgba(0,0,0,0.25)" }}>
        <div style={{ padding: "24px 28px 18px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div><h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textDark }}>⚙️ Data Source Configuration</h2><p style={{ margin: "4px 0 0", fontSize: 12, color: C.textMid }}>Connect to Dynamics 365</p></div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: C.textLight }}>✕</button>
        </div>
        <div style={{ padding: "24px 28px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <div style={{ width: 32, height: 32, borderRadius: 8, background: C.d365, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 14 }}>D</div>
            <div><div style={{ fontSize: 14, fontWeight: 700, color: C.textDark }}>Microsoft Dynamics 365</div><div style={{ fontSize: 10, color: C.textMid }}>servingintel.crm.dynamics.com — MSAL Authentication</div></div>
          </div>
          {d365Account ? (
            <div style={{ background: C.greenLight, borderRadius: 10, padding: "14px 18px", marginBottom: 16, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div><div style={{ fontSize: 13, fontWeight: 600, color: C.green }}>✅ Connected as {d365Account.name || d365Account.username}</div><div style={{ fontSize: 10, color: C.textMid, marginTop: 2 }}>{d365Account.username}</div></div>
              <button onClick={onD365Logout} style={{ padding: "6px 14px", borderRadius: 6, border: `1px solid ${C.border}`, background: "transparent", fontSize: 11, fontWeight: 600, color: C.textMid, cursor: "pointer" }}>Disconnect</button>
            </div>
          ) : (
            <div style={{ marginBottom: 16 }}>
              <button onClick={handleD365SignIn} disabled={signingIn} style={{ width: "100%", padding: "14px", borderRadius: 10, border: `2px solid ${C.d365}`, background: `${C.d365}08`, color: C.d365, fontSize: 14, fontWeight: 700, cursor: signingIn ? "wait" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 10, fontFamily: "'DM Sans',sans-serif" }}>
                <svg width="20" height="20" viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>
                {signingIn ? "Signing in..." : "Sign in with Microsoft"}
              </button>
              {d365Status && !d365Status.success && (<div style={{ marginTop: 8, padding: "8px 12px", borderRadius: 8, fontSize: 11, background: C.redLight, color: C.red }}>❌ {d365Status.error}</div>)}
            </div>
          )}
          <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 16, marginTop: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: "#0078D4", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 14 }}>✉️</div>
              <div><div style={{ fontSize: 14, fontWeight: 700, color: C.textDark }}>Email Reports</div><div style={{ fontSize: 10, color: C.textMid }}>Send reports via Microsoft Graph API using your signed-in account</div></div>
            </div>
            <div style={{ fontSize: 11, color: C.textMid, lineHeight: 1.6, padding: "8px 12px", background: C.bg, borderRadius: 8 }}>
              ✅ Uses your Microsoft sign-in to send emails directly — no Power Automate needed.<br/>
              💡 Requires <strong>Mail.Send</strong> permission on your Azure App Registration. A consent popup will appear on first use.
            </div>
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

function SendReportModal({ show, onClose, onSend, dateLabel }) {
  const [email, setEmail] = useState("");
  const [note, setNote] = useState("");
  const [status, setStatus] = useState(null); // null | "sending" | "sent" | "error"
  const [error, setError] = useState("");

  if (!show) return null;

  const handleSend = async () => {
    if (!email.trim() || !email.includes("@")) { setError("Enter a valid email address"); return; }
    setStatus("sending"); setError("");
    try {
      await onSend(email.trim(), note.trim());
      setStatus("sent");
      setTimeout(() => { onClose(); setStatus(null); setEmail(""); setNote(""); }, 2000);
    } catch (err) {
      setStatus("error"); setError(err.message);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.55)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background: C.card, borderRadius: 20, width: 480, boxShadow: "0 24px 80px rgba(0,0,0,0.25)" }}>
        <div style={{ padding: "24px 28px 16px", borderBottom: `1px solid ${C.border}` }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textDark }}>📤 Send Report via Email</h2>
          <p style={{ margin: "4px 0 0", fontSize: 12, color: C.textMid }}>{dateLabel}</p>
        </div>
        <div style={{ padding: "20px 28px" }}>
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6 }}>Recipient Email(s) *</div>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="manager@company.com, team@company.com"
              style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${error && !email.includes("@") ? "#E5544B" : C.border}`, fontSize: 13, fontFamily: "'DM Sans', sans-serif", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} />
            <div style={{ fontSize: 10, color: C.textLight, marginTop: 4 }}>Separate multiple emails with commas</div>
          </div>
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6 }}>Note <span style={{ fontWeight: 400, color: C.textLight }}>(optional)</span></div>
            <textarea value={note} onChange={e => setNote(e.target.value)} placeholder="Add a note to the email..." rows={3}
              style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${C.border}`, fontSize: 13, fontFamily: "'DM Sans', sans-serif", background: C.bg, color: C.textDark, outline: "none", resize: "vertical", boxSizing: "border-box" }} />
          </div>
          {status === "sent" && (
            <div style={{ padding: "10px 14px", borderRadius: 10, background: C.greenLight, fontSize: 12, fontWeight: 600, color: C.green, textAlign: "center", marginBottom: 12 }}>✅ Report sent successfully!</div>
          )}
          {status === "error" && (
            <div style={{ padding: "10px 14px", borderRadius: 10, background: C.redLight, fontSize: 12, fontWeight: 600, color: C.red, marginBottom: 12 }}>❌ {error}</div>
          )}
          <div style={{ background: C.bg, borderRadius: 10, padding: "12px 14px", marginBottom: 12 }}>
            <div style={{ fontSize: 10, fontWeight: 600, color: C.textLight, textTransform: "uppercase", marginBottom: 6 }}>How It Works</div>
            <div style={{ fontSize: 11, color: C.textMid, lineHeight: 1.6 }}>
              Sends a styled HTML email directly from your Microsoft account via Graph API. Includes all Tier 1–3 metrics, Phone, Email, CSAT, and Overall Summary.
            </div>
          </div>
        </div>
        <div style={{ padding: "16px 28px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button onClick={() => { onClose(); setStatus(null); setError(""); }} style={{ padding: "10px 22px", borderRadius: 10, border: `1px solid ${C.border}`, background: "transparent", fontSize: 13, fontWeight: 600, color: C.textMid, cursor: "pointer" }}>Cancel</button>
          <button onClick={handleSend} disabled={status === "sending" || status === "sent"}
            style={{ padding: "10px 22px", borderRadius: 10, border: "none", background: status === "sent" ? C.green : `linear-gradient(135deg, ${C.accent}, ${C.yellow})`, fontSize: 13, fontWeight: 600, color: "#fff", cursor: status === "sending" ? "wait" : "pointer", opacity: status === "sending" ? 0.7 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            {status === "sending" ? "⏳ Sending..." : status === "sent" ? "✅ Sent!" : "📤 Send Report"}
          </button>
        </div>
      </div>
    </div>
  );
}


const AUTOREPORT_PY_CONTENT = "\"\"\"Auto KPI Report - Queries D365 + Sends styled email via Graph API\"\"\"\nimport os, sys, json\nfrom datetime import datetime, timedelta, timezone\nfrom zoneinfo import ZoneInfo\nimport requests\nfrom msal import ConfidentialClientApplication\n\nD365_TENANT = os.environ[\"D365_TENANT_ID\"]\nD365_CLIENT = os.environ[\"D365_CLIENT_ID\"]\nD365_SECRET = os.environ[\"D365_CLIENT_SECRET\"]\nORG_URL = os.environ[\"D365_ORG_URL\"].rstrip(\"/\")\nGRAPH_TENANT = os.environ.get(\"GRAPH_TENANT_ID\", D365_TENANT)\nGRAPH_CLIENT = os.environ.get(\"GRAPH_CLIENT_ID\", D365_CLIENT)\nGRAPH_SECRET = os.environ.get(\"GRAPH_CLIENT_SECRET\", D365_SECRET)\nSEND_FROM = os.environ[\"SEND_FROM\"]\nSEND_TO = os.environ[\"SEND_TO\"]\nLOOKBACK = int(os.environ.get(\"LOOKBACK_HOURS\", \"24\"))\nTIERS = os.environ.get(\"REPORT_TIERS\", \"ALL\")\nMEMBERS = [m.strip() for m in os.environ.get(\"REPORT_MEMBERS\", \"\").split(\",\") if m.strip()]\nTIME_FROM = os.environ.get(\"TIME_FROM\", \"00:00\")\nTIME_TO = os.environ.get(\"TIME_TO\", \"23:59\")\nAPI_BASE = f\"{ORG_URL}/api/data/v9.2\"\nCT = ZoneInfo(\"America/Chicago\")\n\ndef get_d365_token():\n    app = ConfidentialClientApplication(D365_CLIENT, authority=f\"https://login.microsoftonline.com/{D365_TENANT}\", client_credential=D365_SECRET)\n    r = app.acquire_token_for_client(scopes=[f\"{ORG_URL}/.default\"])\n    if \"access_token\" not in r: print(f\"D365 auth failed: {r}\"); sys.exit(1)\n    return r[\"access_token\"]\n\ndef get_graph_token():\n    app = ConfidentialClientApplication(GRAPH_CLIENT, authority=f\"https://login.microsoftonline.com/{GRAPH_TENANT}\", client_credential=GRAPH_SECRET)\n    r = app.acquire_token_for_client(scopes=[\"https://graph.microsoft.com/.default\"])\n    if \"access_token\" not in r: print(f\"Graph auth failed: {r}\"); sys.exit(1)\n    return r[\"access_token\"]\n\nS = requests.Session()\n\ndef init_d365(token):\n    S.headers.update({\"Authorization\": f\"Bearer {token}\", \"OData-MaxVersion\": \"4.0\", \"OData-Version\": \"4.0\", \"Accept\": \"application/json\", \"Prefer\": \"odata.include-annotations=*,odata.maxpagesize=5000\"})\n\ndef d365_get(q):\n    url = f\"{API_BASE}/{q}\"\n    rows = []\n    while url:\n        r = S.get(url); r.raise_for_status(); d = r.json()\n        rows.extend(d.get(\"value\", [])); url = d.get(\"@odata.nextLink\")\n    return rows\n\ndef d365_count(q):\n    r = S.get(f\"{API_BASE}/{q}\"); r.raise_for_status(); d = r.json()\n    return d.get(\"@odata.count\", len(d.get(\"value\", [])))\n\ndef get_range():\n    now = datetime.now(CT)\n    start = now - timedelta(hours=LOOKBACK)\n    fh, fm2 = map(int, TIME_FROM.split(\":\"))\n    th, tm2 = map(int, TIME_TO.split(\":\"))\n    start = start.replace(hour=fh, minute=fm2, second=0, microsecond=0)\n    end = now.replace(hour=th, minute=tm2, second=59, microsecond=0)\n    s = start.astimezone(timezone.utc).strftime(\"%Y-%m-%dT%H:%M:%SZ\")\n    e = end.astimezone(timezone.utc).strftime(\"%Y-%m-%dT%H:%M:%SZ\")\n    label = f\"{start.strftime('%b %d, %I:%M %p')} - {end.strftime('%b %d, %I:%M %p %Z')}\"\n    return s, e, label\n\ndef count_kpi(rows, field):\n    met = missed = 0\n    for r in rows:\n        st = (r.get(field) or {}).get(\"status\")\n        if st == 4: met += 1\n        elif st == 1: missed += 1\n    return met, missed\n\ndef pct(n, d): return round(n/d*100) if d > 0 else \"N/A\"\ndef sla_pct(m, x): t = m+x; return round(m/t*100) if t > 0 else \"N/A\"\ndef ic(v, tgt, inv=False):\n    if v == \"N/A\": return \"\u2796\"\n    return \"\u2705\" if (inv and v < tgt) or (not inv and v >= tgt) else \"\ud83d\udd34\"\ndef fm(v): return f\"{v}%\" if v != \"N/A\" else \"N/A\"\n\ndef fetch_tier(code, df, s, e):\n    base = f\"casetypecode eq {code} and {df} ge {s} and {df} le {e}\"\n    total = d365_count(f\"incidents?$filter={base}&$count=true&$top=1\")\n    resolved = d365_get(f\"incidents?$filter={base} and statecode eq 1&$select=incidentid&$expand=resolvebykpiid($select=status)\")\n    resp = d365_get(f\"incidents?$filter={base} and statecode eq 1&$select=incidentid&$expand=firstresponsebykpiid($select=status)\")\n    active = d365_get(f\"incidents?$filter={base} and statecode eq 0&$select=incidentid&$expand=resolvebykpiid($select=status)\")\n    sm, sx = count_kpi(resolved, \"resolvebykpiid\")\n    rm, rx = count_kpi(resp, \"firstresponsebykpiid\")\n    breached = sum(1 for r in active if (r.get(\"resolvebykpiid\") or {}).get(\"status\") == 1)\n    return {\"total\": total, \"resolved\": len(resolved), \"sla_met\": sm, \"sla_missed\": sx, \"sla\": sla_pct(sm, sx),\n            \"resp_met\": rm, \"resp_missed\": rx, \"resp_sla\": sla_pct(rm, rx),\n            \"breach\": breached, \"breach_total\": len(active), \"breach_rate\": pct(breached, len(active)) if len(active) > 0 else 0}\n\ndef row(label, val, indent=False):\n    pad = \"padding-left:14px;font-size:12px;\" if indent else \"\"\n    return f'<tr><td style=\"padding:6px 0;color:#555;{pad}\">{label}</td><td style=\"padding:6px 0;text-align:right;font-weight:700;\">{val}</td></tr>'\n\ndef build_html(tiers_data, ph, em, cs, label, config_label):\n    colors = {1: (\"#1565c0\",\"#2196F3\"), 2: (\"#e65100\",\"#FF9800\"), 3: (\"#7b1fa2\",\"#9C27B0\")}\n    tier_html = \"\"\n    for td in tiers_data:\n        c1, c2 = colors.get(td[\"code\"], (\"#333\",\"#666\"))\n        t = td[\"data\"]\n        tier_html += f'''<div style=\"background:linear-gradient(135deg,{c1},{c2});color:#fff;padding:14px 20px;border-radius:10px 10px 0 0;\"><strong>{td[\"name\"]}</strong></div>\n<div style=\"background:#fff;padding:16px 20px;border-radius:0 0 10px 10px;margin-bottom:16px;\"><table style=\"width:100%;border-collapse:collapse;font-size:13px;\">\n{row(\"SLA Compliance\", f\"{fm(t['sla'])} {ic(t['sla'],90)}\")}\n{row(f\"  {t['sla_met']} met / {t['sla_missed']} missed\", \"\", True)}\n{row(\"Response SLA\", f\"{fm(t['resp_sla'])} {ic(t['resp_sla'],90)}\")}\n{row(\"Open Breach\", f\"{t['breach_rate']}% {ic(t['breach_rate'],5,True)} ({t['breach']}/{t['breach_total']})\")}\n{row(\"FCR\", f\"{fm(t.get('fcr_rate','N/A'))} {ic(t.get('fcr_rate','N/A'),90)}\")}\n{row(\"Escalation\", f\"{fm(t.get('esc_rate','N/A'))} {ic(t.get('esc_rate','N/A'),10,True)}\")}\n{row(\"Total\", t[\"total\"])}{row(\"Resolved\", t[\"resolved\"])}\n</table></div>'''\n    tc = sum(t[\"data\"][\"total\"] for t in tiers_data)\n    tr = sum(t[\"data\"][\"resolved\"] for t in tiers_data)\n    return f'''<div style=\"font-family:Segoe UI,Arial,sans-serif;max-width:700px;margin:0 auto;background:#f4f0eb;padding:20px;\">\n<div style=\"background:linear-gradient(135deg,#1a2332,#2d4a6f);color:#fff;padding:24px 28px;border-radius:12px;text-align:center;margin-bottom:16px;\">\n<h1 style=\"margin:0;font-size:20px;\">Auto Report</h1>\n<p style=\"margin:4px 0 0;font-size:13px;opacity:0.85;\">{label}</p>\n<p style=\"margin:2px 0 0;font-size:10px;opacity:0.6;\">{config_label}</p></div>\n{tier_html}\n<table style=\"width:100%;border-collapse:separate;border-spacing:12px 0;margin-bottom:16px;\"><tr>\n<td style=\"width:50%;vertical-align:top;background:#fff;border-radius:10px;padding:14px 16px;\"><strong style=\"color:#2D9D78;\">Phone</strong><table style=\"width:100%;font-size:12px;margin-top:8px;\">{row(\"Total\",ph[\"total\"])}{row(\"Answered\",ph[\"answered\"])}{row(\"Abandoned\",ph[\"abandoned\"])}{row(\"Rate\",f\"{fm(ph['rate'])} {ic(ph['rate'],95)}\")}</table></td>\n<td style=\"width:50%;vertical-align:top;background:#fff;border-radius:10px;padding:14px 16px;\"><strong style=\"color:#2196F3;\">Email</strong><table style=\"width:100%;font-size:12px;margin-top:8px;\">{row(\"Total\",em[\"total\"])}{row(\"Responded\",em[\"responded\"])}{row(\"Resolved\",em[\"resolved\"])}</table></td>\n</tr></table>\n<div style=\"background:linear-gradient(135deg,#1a2332,#2d4a6f);color:#fff;padding:20px 24px;border-radius:12px;text-align:center;\">\n<strong>OVERALL</strong><table style=\"width:100%;margin-top:12px;\"><tr>\n<td style=\"text-align:center\"><div style=\"font-size:24px;font-weight:700;color:#4FC3F7;\">{tc}</div><div style=\"font-size:10px;color:#a8c6df;\">Created</div></td>\n<td style=\"text-align:center\"><div style=\"font-size:24px;font-weight:700;color:#81C784;\">{tr}</div><div style=\"font-size:10px;color:#a8c6df;\">Resolved</div></td>\n<td style=\"text-align:center\"><div style=\"font-size:24px;font-weight:700;color:#81C784;\">{ph[\"answered\"]}</div><div style=\"font-size:10px;color:#a8c6df;\">Answered</div></td>\n<td style=\"text-align:center\"><div style=\"font-size:24px;font-weight:700;color:#f44336;\">{ph[\"abandoned\"]}</div><div style=\"font-size:10px;color:#a8c6df;\">Abandoned</div></td>\n</tr></table></div>\n<p style=\"text-align:center;font-size:10px;color:#999;margin-top:16px;\">Auto-generated from Dynamics 365</p></div>'''\n\ndef send_email(html, label):\n    token = get_graph_token()\n    recipients = [{\"emailAddress\": {\"address\": e.strip()}} for e in SEND_TO.split(\",\") if e.strip()]\n    payload = {\"message\": {\"subject\": f\"Auto Report - {label}\", \"body\": {\"contentType\": \"HTML\", \"content\": html}, \"toRecipients\": recipients, \"from\": {\"emailAddress\": {\"address\": SEND_FROM}}}}\n    r = requests.post(f\"https://graph.microsoft.com/v1.0/users/{SEND_FROM}/sendMail\", headers={\"Authorization\": f\"Bearer {token}\", \"Content-Type\": \"application/json\"}, json=payload)\n    if r.status_code not in (200, 202): print(f\"Email failed ({r.status_code}): {r.text[:300]}\"); sys.exit(1)\n    print(f\"Email sent to {SEND_TO}\")\n\ndef main():\n    print(\"Authenticating D365...\")\n    init_d365(get_d365_token())\n    s, e, label = get_range()\n    print(f\"Report: {label}\")\n    tier_configs = [\n        {\"code\": 1, \"df\": \"createdon\", \"name\": \"Tier 1 - Service Desk\"},\n        {\"code\": 2, \"df\": \"escalatedon\", \"name\": \"Tier 2 - Programming\"},\n        {\"code\": 3, \"df\": \"escalatedon\", \"name\": \"Tier 3 - Relationship Mgrs\"},\n    ]\n    if TIERS != \"ALL\":\n        tier_codes = [int(t.strip()) for t in TIERS.split(\",\")]\n        tier_configs = [tc for tc in tier_configs if tc[\"code\"] in tier_codes]\n    tiers_data = []\n    for tc in tier_configs:\n        print(f\"Querying {tc['name']}...\")\n        t = fetch_tier(tc[\"code\"], tc[\"df\"], s, e)\n        if tc[\"code\"] == 1:\n            fcr = d365_count(f\"incidents?$filter=casetypecode eq 1 and cr7fe_new_fcr eq true and createdon ge {s} and createdon le {e}&$count=true&$top=1\")\n            esc = d365_count(f\"incidents?$filter=casetypecode eq 2 and escalatedon ge {s} and escalatedon le {e}&$count=true&$top=1\")\n            t[\"fcr_rate\"] = pct(fcr, t[\"total\"]); t[\"esc_rate\"] = pct(esc, t[\"total\"])\n        elif tc[\"code\"] == 2:\n            t2e = d365_count(f\"incidents?$filter=casetypecode eq 3 and escalatedon ge {s} and escalatedon le {e}&$count=true&$top=1\")\n            t[\"esc_rate\"] = pct(t2e, t[\"total\"])\n        tiers_data.append({**tc, \"data\": t})\n    print(\"Phone...\")\n    try:\n        pb = f\"createdon ge {s} and createdon le {e} and directioncode eq true\"\n        pt = d365_count(f\"phonecalls?$filter={pb}&$count=true&$top=1\")\n        pa = d365_count(f\"phonecalls?$filter={pb} and statecode eq 1&$count=true&$top=1\")\n        ph = {\"total\": pt, \"answered\": pa, \"abandoned\": pt-pa, \"rate\": pct(pa, pt)}\n    except: ph = {\"total\":0,\"answered\":0,\"abandoned\":0,\"rate\":\"N/A\"}\n    print(\"Email...\")\n    eb = f\"caseorigincode eq 2 and createdon ge {s} and createdon le {e}\"\n    em = {\"total\": d365_count(f\"incidents?$filter={eb}&$count=true&$top=1\"),\n          \"responded\": d365_count(f\"incidents?$filter={eb} and firstresponsesent eq true&$count=true&$top=1\"),\n          \"resolved\": d365_count(f\"incidents?$filter={eb} and statecode eq 1&$count=true&$top=1\")}\n    print(\"CSAT...\")\n    try:\n        cr = d365_get(f\"cr7fe_new_csats?$filter=createdon ge {s} and createdon le {e}&$select=cr7fe_new_rating\")\n        sc = [r[\"cr7fe_new_rating\"] for r in cr if r.get(\"cr7fe_new_rating\") is not None]\n        cs = {\"count\": len(sc), \"avg\": round(sum(sc)/len(sc),1) if sc else \"N/A\"}\n    except: cs = {\"count\":0,\"avg\":\"N/A\"}\n    tier_label = \"All Tiers\" if TIERS == \"ALL\" else \", \".join(tc[\"name\"] for tc in tiers_data)\n    member_label = f\"{len(MEMBERS)} members\" if MEMBERS else \"\"\n    config_label = tier_label + (\" | \" + member_label if member_label else \"\")\n    html = build_html(tiers_data, ph, em, cs, label, config_label)\n    print(\"Sending email...\"); send_email(html, label)\n    print(\"Done!\")\n\nif __name__ == \"__main__\": main()\n";

function AutoReportModal({ show, onClose, queues, d365Account }) {
  const [emails, setEmails] = useState("");
  const [intervalHours, setIntervalHours] = useState(24);
  const [selectedTier, setSelectedTier] = useState("");
  const [autoMembers, setAutoMembers] = useState([]);
  const [autoMembersList, setAutoMembersList] = useState([]);
  const [loadingAutoMembers, setLoadingAutoMembers] = useState(false);
  const [startDate, setStartDate] = useState(() => { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().split("T")[0]; });
  const [endDate, setEndDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [fromTime, setFromTime] = useState("00:00");
  const [toTime, setToTime] = useState("23:59");
  const [generating, setGenerating] = useState(false);
  const [generated, setGenerated] = useState(false);

  useEffect(() => {
    if (selectedTier && selectedTier !== "all" && d365Account) {
      setLoadingAutoMembers(true); setAutoMembers([]);
      fetchD365QueueMembers(selectedTier).then(m => { setAutoMembersList(m); setLoadingAutoMembers(false); }).catch(() => { setAutoMembersList([]); setLoadingAutoMembers(false); });
    } else { setAutoMembersList([]); setAutoMembers([]); }
  }, [selectedTier, d365Account]);

  if (!show) return null;

  const cronExpr = intervalHours <= 1 ? "0 * * * *"
    : intervalHours <= 12 ? `0 */${intervalHours} * * *`
    : intervalHours === 24 ? "1 6 * * *"
    : intervalHours === 168 ? "1 6 * * 1"
    : `0 */${intervalHours} * * *`;

  const cronLabel = intervalHours === 1 ? "Every hour"
    : intervalHours === 24 ? "Daily at 12:01 AM CT"
    : intervalHours === 168 ? "Weekly (Monday 12:01 AM CT)"
    : `Every ${intervalHours} hours`;

  const tierObj = queues.find(q => q.id === selectedTier);
  const tierLabel = selectedTier === "all" ? "All Tiers" : (tierObj?.tierLabel || "Not selected");
  const tierNum = tierObj?.tierNum || null;
  const memberNames = autoMembers.map(id => autoMembersList.find(m => m.id === id)?.name || id);

  const setQuickRange = (days) => {
    const end = new Date();
    const start = new Date(); start.setDate(start.getDate() - days);
    setStartDate(start.toISOString().split("T")[0]);
    setEndDate(end.toISOString().split("T")[0]);
  };

  const daysDiff = Math.max(1, Math.round((new Date(endDate) - new Date(startDate)) / 86400000));

  const handleGenerate = () => {
    if (!emails.trim() || !selectedTier) return;
    setGenerating(true);
    const tierConfig = selectedTier === "all" ? "ALL" : String(tierNum || 1);
    const memberIds = autoMembers.length > 0 ? autoMembers.join(",") : "";
    const lookbackHours = daysDiff * 24;

    const yamlLines = [
      "name: Auto KPI Report",
      "on:",
      "  schedule:",
      "    - cron: '" + cronExpr + "'",
      "  workflow_dispatch:",
      "",
      "jobs:",
      "  send-report:",
      "    runs-on: ubuntu-latest",
      "    steps:",
      "      - uses: actions/checkout@v4",
      "      - uses: actions/setup-python@v5",
      "        with:",
      "          python-version: '3.11'",
      "      - run: pip install requests msal",
      "      - name: Generate and send report",
      "        env:",
      "          D365_TENANT_ID: ${{secrets.D365_TENANT_ID}}",
      "          D365_CLIENT_ID: ${{secrets.D365_CLIENT_ID}}",
      "          D365_CLIENT_SECRET: ${{secrets.D365_CLIENT_SECRET}}",
      "          D365_ORG_URL: ${{secrets.D365_ORG_URL}}",
      "          GRAPH_TENANT_ID: ${{secrets.GRAPH_TENANT_ID}}",
      "          GRAPH_CLIENT_ID: ${{secrets.GRAPH_CLIENT_ID}}",
      "          GRAPH_CLIENT_SECRET: ${{secrets.GRAPH_CLIENT_SECRET}}",
      "          SEND_FROM: ${{secrets.SEND_FROM}}",
      '          SEND_TO: "' + emails.trim() + '"',
      '          LOOKBACK_HOURS: "' + lookbackHours + '"',
      '          REPORT_TIERS: "' + tierConfig + '"',
      '          REPORT_MEMBERS: "' + memberIds + '"',
      '          TIME_FROM: "' + fromTime + '"',
      '          TIME_TO: "' + toTime + '"',
      "        run: python auto_report.py",
    ];
    const workflowYaml = yamlLines.join("\n");

    const pythonScript = AUTOREPORT_PY_CONTENT;

    const readmeLines = [
      "# Auto KPI Report", "",
      "Automated service desk report via GitHub Actions.", "",
      "## Configuration",
      "- Tier: " + tierLabel,
      "- Members: " + (memberNames.length > 0 ? memberNames.join(", ") : "All"),
      "- Schedule: " + cronLabel,
      "- Lookback: " + daysDiff + " day(s)",
      "- Time window: " + fromTime + " - " + toTime,
      "- Recipients: " + emails.trim(), "",
      "## Setup", "",
      "### 1. Azure App Registration",
      "Add **Microsoft Graph > Application > Mail.Send** + Grant admin consent", "",
      "### 2. GitHub Secrets",
      "| Secret | Value |",
      "|--------|-------|",
      "| D365_TENANT_ID | 1b0086bd-aeda-4c74-a15a-23adfe4d0693 |",
      "| D365_CLIENT_ID | 0918449d-b73e-428a-8238-61723f2a2e7d |",
      "| D365_CLIENT_SECRET | Your app client secret |",
      "| D365_ORG_URL | https://servingintel.crm.dynamics.com |",
      "| GRAPH_TENANT_ID | (same as D365_TENANT_ID) |",
      "| GRAPH_CLIENT_ID | (same as D365_CLIENT_ID) |",
      "| GRAPH_CLIENT_SECRET | (same as D365_CLIENT_SECRET) |",
      "| SEND_FROM | your-email@servingintel.com |", "",
      "### 3. Push and Run",
      'git add . && git commit -m "auto report" && git push',
      "Go to Actions > Run workflow to test",
    ];
    const readmeText = readmeLines.join("\n");

    [
      { name: "auto-report.yml", content: workflowYaml },
      { name: "auto_report.py", content: pythonScript },
      { name: "README.md", content: readmeText },
    ].forEach((f, i) => {
      setTimeout(() => {
        const blob = new Blob([f.content], { type: "text/plain" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = f.name; a.click();
        URL.revokeObjectURL(url);
      }, i * 500);
    });
    setGenerating(false);
    setGenerated(true);
  };

  const inputSt = { width: "100%", padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${C.border}`, fontSize: 13, fontFamily: "'DM Sans', sans-serif", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" };
  const labelSt = { fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 };
  const subSt = { fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.55)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background: C.card, borderRadius: 20, width: 560, maxHeight: "90vh", overflow: "auto", boxShadow: "0 24px 80px rgba(0,0,0,0.25)" }}>
        <div style={{ padding: "24px 28px 16px", borderBottom: `1px solid ${C.border}` }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textDark }}>⏰ Auto Report Setup</h2>
          <p style={{ margin: "4px 0 0", fontSize: 12, color: C.textMid }}>Configure automated reports via GitHub Actions</p>
        </div>
        <div style={{ padding: "20px 28px" }}>

          {/* TIER */}
          <div style={{ marginBottom: 16 }}>
            <div style={labelSt}><span>🏢</span> Tier</div>
            <select value={selectedTier} onChange={(e) => { setSelectedTier(e.target.value); setAutoMembers([]); }} style={{ ...inputSt, cursor: "pointer", appearance: "auto" }}>
              <option value="">Select a tier...</option>
              <option value="all">All Tiers</option>
              {queues.map(q => <option key={q.id} value={q.id}>{q.tierLabel || q.name}</option>)}
            </select>
          </div>

          {/* TEAM MEMBERS */}
          {selectedTier && selectedTier !== "all" && (
            <div style={{ marginBottom: 16 }}>
              <div style={labelSt}><span>👥</span> Team Members {loadingAutoMembers && <span style={{ fontSize: 10, color: C.accent }}>Loading...</span>}</div>
              <MultiMemberSelect selected={autoMembers} onChange={setAutoMembers} members={autoMembersList} />
              {autoMembers.length > 0 && <div style={{ marginTop: 6, display: "flex", flexWrap: "wrap", gap: 4 }}>
                {autoMembers.slice(0, 4).map(id => { const m = autoMembersList.find(t => t.id === id); return <span key={id} style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: `${C.accent}15`, color: C.accent, fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}>{m?.name?.split(" ")[0]}<span onClick={() => setAutoMembers(prev => prev.filter(s => s !== id))} style={{ cursor: "pointer", opacity: 0.6, fontSize: 8 }}>✕</span></span>; })}
                {autoMembers.length > 4 && <span style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: C.bg, color: C.textLight, fontWeight: 600 }}>+{autoMembers.length - 4} more</span>}
              </div>}
            </div>
          )}

          {/* DATE RANGE */}
          <div style={{ marginBottom: 16 }}>
            <div style={labelSt}><span>📅</span> Date Range</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 6 }}>
              <label><div style={subSt}>From</div><input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} style={{ ...inputSt, fontSize: 11, fontFamily: "'Space Mono', monospace", padding: "8px 8px" }} /></label>
              <label><div style={subSt}>To</div><input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} style={{ ...inputSt, fontSize: 11, fontFamily: "'Space Mono', monospace", padding: "8px 8px" }} /></label>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
              <label><div style={subSt}>From Time</div><input type="time" value={fromTime} onChange={e => setFromTime(e.target.value)} style={{ ...inputSt, fontSize: 11, fontFamily: "'Space Mono', monospace", padding: "8px 8px" }} /></label>
              <label><div style={subSt}>To Time</div><input type="time" value={toTime} onChange={e => setToTime(e.target.value)} style={{ ...inputSt, fontSize: 11, fontFamily: "'Space Mono', monospace", padding: "8px 8px" }} /></label>
            </div>
            <div style={{ display: "flex", gap: 5 }}>
              {[{ l: "7D", d: 7 }, { l: "14D", d: 14 }, { l: "30D", d: 30 }, { l: "90D", d: 90 }, { l: "YTD", d: Math.round((new Date() - new Date(new Date().getFullYear(), 0, 1)) / 86400000) }].map(q =>
                <button key={q.l} onClick={() => setQuickRange(q.d)} style={{ flex: 1, padding: "5px 0", borderRadius: 6, border: `1px solid ${daysDiff === q.d ? C.accent : C.border}`, background: daysDiff === q.d ? C.accentLight : "transparent", fontSize: 10, fontWeight: 600, color: daysDiff === q.d ? C.accent : C.textMid, cursor: "pointer", fontFamily: "'Space Mono', monospace" }}>{q.l}</button>
              )}
            </div>
          </div>

          {/* SEND EVERY */}
          <div style={{ marginBottom: 16 }}>
            <div style={labelSt}><span>🔄</span> Send Every</div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <input type="number" min="1" max="168" value={intervalHours} onChange={e => setIntervalHours(Math.max(1, parseInt(e.target.value) || 1))}
                style={{ width: 70, padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${C.border}`, fontSize: 14, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", textAlign: "center" }} />
              <span style={{ fontSize: 13, color: C.textMid, fontWeight: 600 }}>hours</span>
            </div>
            <div style={{ display: "flex", gap: 4 }}>
              {[{ l: "1h", v: 1 }, { l: "4h", v: 4 }, { l: "8h", v: 8 }, { l: "12h", v: 12 }, { l: "24h", v: 24 }].map(p =>
                <button key={p.l} onClick={() => setIntervalHours(p.v)} style={{ padding: "3px 8px", borderRadius: 6, border: `1px solid ${intervalHours === p.v ? C.accent : C.border}`, background: intervalHours === p.v ? C.accentLight : "transparent", fontSize: 10, fontWeight: 600, color: intervalHours === p.v ? C.accent : C.textLight, cursor: "pointer", fontFamily: "'Space Mono', monospace" }}>{p.l}</button>
              )}
            </div>
          </div>

          {/* RECIPIENTS */}
          <div style={{ marginBottom: 16 }}>
            <div style={labelSt}><span>📧</span> Recipient Email(s) *</div>
            <input type="text" value={emails} onChange={e => setEmails(e.target.value)} placeholder="manager@company.com, team@company.com" style={inputSt} />
            <div style={{ fontSize: 10, color: C.textLight, marginTop: 4 }}>Separate multiple emails with commas</div>
          </div>

          {/* SUMMARY */}
          <div style={{ background: C.bg, borderRadius: 10, padding: "14px 16px", marginBottom: 12 }}>
            <div style={{ fontSize: 10, fontWeight: 600, color: C.textLight, textTransform: "uppercase", marginBottom: 8 }}>Summary</div>
            <div style={{ fontSize: 12, color: C.textMid, lineHeight: 1.7 }}>
              <div>🏢 <strong>Tier:</strong> {tierLabel}</div>
              {autoMembers.length > 0 && <div>👥 <strong>Members:</strong> {memberNames.slice(0, 3).join(", ")}{autoMembers.length > 3 ? ` +${autoMembers.length - 3} more` : ""}</div>}
              <div>📅 <strong>Date range:</strong> {startDate} to {endDate} ({fromTime} — {toTime})</div>
              <div>🔄 <strong>Frequency:</strong> {cronLabel}</div>
              <div>📧 <strong>Recipients:</strong> {emails || "(enter emails above)"}</div>
              <div>📊 <strong>Includes:</strong> SLA, Response SLA, Open Breach, FCR, Phone, Email, CSAT</div>
            </div>
          </div>

          {generated && (
            <div style={{ padding: "12px 14px", borderRadius: 10, background: C.greenLight, fontSize: 12, color: C.green, marginBottom: 12, lineHeight: 1.6 }}>
              ✅ <strong>3 files downloaded!</strong><br/>
              1. <code>auto-report.yml</code> → put in <code>.github/workflows/</code><br/>
              2. <code>auto_report.py</code> → put in repo root<br/>
              3. <code>README.md</code> → setup instructions
            </div>
          )}
        </div>
        <div style={{ padding: "16px 28px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button onClick={onClose} style={{ padding: "10px 22px", borderRadius: 10, border: `1px solid ${C.border}`, background: "transparent", fontSize: 13, fontWeight: 600, color: C.textMid, cursor: "pointer" }}>Close</button>
          <button onClick={handleGenerate} disabled={generating || !emails.trim() || !selectedTier}
            style={{ padding: "10px 22px", borderRadius: 10, border: "none", background: (emails.trim() && selectedTier) ? `linear-gradient(135deg, ${C.accent}, ${C.yellow})` : C.border, fontSize: 13, fontWeight: 600, color: (emails.trim() && selectedTier) ? "#fff" : C.textLight, cursor: (emails.trim() && selectedTier) ? "pointer" : "not-allowed", display: "flex", alignItems: "center", gap: 6 }}>
            {generating ? "⏳ Generating..." : "⬇️ Download Auto Report Package"}
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
  const [queues, setQueues] = useState([]);
  const [selectedQueue, setSelectedQueue] = useState(null);
  const [loadingQueues, setLoadingQueues] = useState(false);
  const [teamMembers, setTeamMembers] = useState(DEMO_TEAM_MEMBERS);
  const [loadingMembers, setLoadingMembers] = useState(false);
  const [selectedMembers, setSelectedMembers] = useState([]);
  const [reportType, setReportType] = useState("daily");
  const [startDate, setStartDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [endDate, setEndDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [startTime, setStartTime] = useState("00:00");
  const [endTime, setEndTime] = useState("23:59");
  const [data, setData] = useState(null);
  const [hasRun, setHasRun] = useState(false);
  const [isRunning, setIsRunning] = useState(false);
  const [runProgress, setRunProgress] = useState("");
  const [showSettings, setShowSettings] = useState(false);
  const [showSendModal, setShowSendModal] = useState(false);
  const [showAutoReport, setShowAutoReport] = useState(false);
  const [apiConfig, setApiConfig] = useState({ live: true });
  const [d365Account, setD365Account] = useState(null);
  const [liveErrors, setLiveErrors] = useState([]);
  const reportRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const msal = getMsal();
        await msal.initialize();
        await msal.handleRedirectPromise();
        const account = getMsalAccount();
        if (account) setD365Account(account);
      } catch (err) { console.log("MSAL init:", err); }
    })();
  }, []);

  const TIER_DEFS = [
    { tier: 1, label: "Tier 1 — Service Desk", match: (name) => { const n = name.toLowerCase(); if (n === "<service desk>") return true; return false; } },
    { tier: 2, label: "Tier 2 — Programming Team", match: (name) => name.toLowerCase().includes("programming team") },
    { tier: 3, label: "Tier 3 — Relationship Managers", match: (name) => name.toLowerCase().includes("relationship manager") },
  ];

  useEffect(() => {
    if (d365Account) {
      setLoadingQueues(true);
      fetchD365Queues().then(q => {
        const filtered = [];
        for (const tierDef of TIER_DEFS) {
          const cleanQ = q.map(queue => ({ ...queue, cleanName: queue.name.replace(" 🔒", "").trim() }));
          let matches = cleanQ.filter(queue => tierDef.match(queue.cleanName));
          if (matches.length > 1 && tierDef.prefer) { matches.sort(tierDef.prefer); }
          if (matches.length > 0) { filtered.push({ ...matches[0], tierLabel: tierDef.label, tierNum: tierDef.tier }); }
        }
        console.log("Tier queues matched:", filtered.map(f => `${f.tierLabel} → "${f.name}" (${f.id})`));
        setQueues(filtered);
        setLoadingQueues(false);
      }).catch(() => setLoadingQueues(false));
    } else {
      setQueues([]); setSelectedQueue(null); setTeamMembers(DEMO_TEAM_MEMBERS); setSelectedMembers([]);
    }
  }, [d365Account]);

  useEffect(() => {
    if (selectedQueue && selectedQueue !== "all" && d365Account) {
      setLoadingMembers(true); setSelectedMembers([]);
      fetchD365QueueMembers(selectedQueue).then(members => { setTeamMembers(members.length > 0 ? members : []); setLoadingMembers(false); }).catch(() => { setTeamMembers([]); setLoadingMembers(false); });
    } else if (selectedQueue === "all" && d365Account) {
      setSelectedMembers([]); setTeamMembers([]);
    } else if (!d365Account) { setTeamMembers(DEMO_TEAM_MEMBERS); }
  }, [selectedQueue, d365Account]);

  const canRun = selectedMembers.length > 0 || (d365Account && selectedQueue);
  const isLive = !!d365Account;

  const setPreset = (type) => {
    setReportType(type);
    const today = new Date();
    if (type === "daily") { setStartDate(today.toISOString().split("T")[0]); setEndDate(today.toISOString().split("T")[0]); }
    else if (type === "weekly") { const w = new Date(today); w.setDate(w.getDate() - 7); setStartDate(w.toISOString().split("T")[0]); setEndDate(today.toISOString().split("T")[0]); }
  };

  const handleD365Login = async () => { const result = await msalLogin(); if (result?.account) { setD365Account(result.account); return result; } return null; };
  const handleD365Logout = async () => { await msalLogoutD365(); setD365Account(null); };

  const [memberData, setMemberData] = useState([]);

  const handleRun = async () => {
    setIsRunning(true); setRunProgress(""); setLiveErrors([]); setMemberData([]);
    try {
      if (selectedMembers.length > 0) {
        const results = []; const allErrors = [];
        for (const memberId of selectedMembers) {
          const member = teamMembers.find(m => m.id === memberId);
          if (!member) continue;
          setRunProgress(`Fetching data for ${member.name}...`);
          const memberResult = await fetchMemberD365Data(member, startDate, endDate, setRunProgress, startTime, endTime);
          results.push(memberResult); allErrors.push(...(memberResult.errors || []));
        }
        setMemberData(results);
        const combined = buildCombinedData(results);
        setData({ ...combined, source: "live" });
        if (allErrors.length > 0) setLiveErrors(allErrors);
      } else {
        const d = await fetchLiveData(apiConfig, startDate, endDate, setRunProgress, startTime, endTime);
        setData(d);
        if (d.errors?.length > 0) setLiveErrors(d.errors);
      }
    } catch (err) { setLiveErrors([err.message]); }
    setHasRun(true); setIsRunning(false); setRunProgress("");
  };

  function buildCombinedData(results) {
    const totals = results.reduce((acc, d) => ({
      totalCases: acc.totalCases + d.totalCases, resolved: acc.resolved + d.resolvedCases,
      slaMet: acc.slaMet + d.slaMet, emailCases: acc.emailCases + d.emailCases,
      emailResolved: acc.emailResolved + d.emailResolved, csatResponses: acc.csatResponses + d.csatResponses,
    }), { totalCases: 0, resolved: 0, slaMet: 0, emailCases: 0, emailResolved: 0, csatResponses: 0 });
    return {
      overall: { created: totals.totalCases, resolved: totals.resolved, csatResponses: totals.csatResponses, answeredCalls: 0, abandonedCalls: 0 },
      email: { total: totals.emailCases, responded: 0, resolved: totals.emailResolved, slaCompliance: totals.emailCases ? Math.min(100, Math.round(totals.emailResolved / totals.emailCases * 100)) : "N/A" },
      phone: { totalCalls: 0, answered: 0, abandoned: 0, answerRate: 0, avgAHT: 0 },
      timeline: [],
    };
  }

  const handleExportPDF = () => { if (reportRef.current) window.print(); };

  function buildEmailHTML(recipientNote) {
    if (!data) return "";
    const t1 = data.tier1 || {};
    const t2 = data.tier2 || {};
    const t3 = data.tier3 || {};
    const ph = data.phone || {};
    const em = data.email || {};
    const cs = data.csat || {};
    const ov = data.overall || {};
    const ic = (val, target, inv) => val === "N/A" || val === undefined ? "➖" : (inv ? (val < target ? "✅" : "🔴") : (val >= target ? "✅" : "🔴"));
    const fm = (v) => v === "N/A" || v === undefined ? "N/A" : `${v}%`;
    const row = (label, val, indent) => `<tr><td style="padding:6px 0;color:#555;${indent ? 'padding-left:14px;font-size:12px;' : ''}">${label}</td><td style="padding:6px 0;text-align:right;font-weight:700;">${val}</td></tr>`;

    return `<div style="font-family:Segoe UI,Arial,sans-serif;max-width:700px;margin:0 auto;background:#f4f0eb;padding:20px;">
  <div style="background:linear-gradient(135deg,#1a2332,#2d4a6f);color:#fff;padding:24px 28px;border-radius:12px;text-align:center;margin-bottom:16px;">
    <h1 style="margin:0;font-size:20px;">📊 Service &amp; Operations Report</h1>
    <p style="margin:4px 0 0;font-size:13px;opacity:0.85;">${dateLabel}</p>
  </div>
  ${recipientNote ? `<div style="background:#fff;border-radius:10px;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#555;border-left:4px solid #6264A7;">💬 ${recipientNote}</div>` : ""}
  <div style="background:linear-gradient(135deg,#1565c0,#2196F3);color:#fff;padding:14px 20px;border-radius:10px 10px 0 0;">
    <strong style="font-size:15px;">🔵 Tier 1 — Service Desk</strong>
  </div>
  <div style="background:#fff;padding:16px 20px;border-radius:0 0 10px 10px;margin-bottom:16px;">
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      ${row("SLA Compliance", `${fm(t1.slaCompliance)} ${ic(t1.slaCompliance, 90, false)}`)}
      ${row(`↳ ${t1.slaMet||0} met · ${t1.slaMissed||0} missed of ${(t1.slaMet||0)+(t1.slaMissed||0)} evaluated`, "", true)}
      ${row("Response SLA", `${fm(t1.slaResponseCompliance)} ${ic(t1.slaResponseCompliance, 90, false)}`)}
      ${row("Open SLA Breach", `${t1.openBreachRate||0}% ${ic(t1.openBreachRate||0, 5, true)}<br/><span style="font-size:11px;font-weight:normal;color:#888;">${t1.openBreachCount||0} of ${t1.openBreachTotal||0} active</span>`)}
      ${row("FCR Rate", `${fm(t1.fcrRate)} ${ic(t1.fcrRate, 90, false)}`)}
      ${row("Escalation Rate", `${fm(t1.escalationRate)} ${ic(t1.escalationRate, 10, true)}`)}
      ${row("Avg Resolution Time", `${t1.avgResolutionTime || "N/A"} ⏱️`)}
      ${row("Total Cases", t1.total || 0)}
      ${row("CSAT Score", `${cs.avgScore || "N/A"}/5 ${cs.avgScore != null && cs.avgScore !== "N/A" && cs.avgScore >= 4 ? "✅" : (cs.avgScore === "N/A" ? "➖" : "🔴")}`)}
    </table>
  </div>

  <table style="width:100%;border-collapse:separate;border-spacing:12px 0;margin-bottom:16px;"><tr>
    <td style="width:50%;vertical-align:top;background:#fff;border-radius:10px;padding:14px 16px;">
      <strong style="font-size:12px;color:#2D9D78;">📞 Phone</strong>
      <table style="width:100%;font-size:12px;margin-top:8px;">
        ${row("Total", ph.totalCalls||0)}
        ${row("Answered", `<span style="color:#2D9D78">${ph.answered||0}</span>`)}
        ${row("Abandoned", `<span style="color:#E5544B">${ph.abandoned||0}</span>`)}
        ${row("Answer Rate", `${fm(ph.answerRate)} ${ic(ph.answerRate, 95, false)}`)}
        ${row("Avg AHT", `${ph.avgAHT||0} min`)}
      </table>
    </td>
    <td style="width:50%;vertical-align:top;background:#fff;border-radius:10px;padding:14px 16px;">
      <strong style="font-size:12px;color:#2196F3;">📧 Email</strong>
      <table style="width:100%;font-size:12px;margin-top:8px;">
        ${row("Total", em.total||0)}
        ${row("Responded", `<span style="color:#FF9800">${em.responded||0}</span>`)}
        ${row("Resolved", `<span style="color:#2D9D78">${em.resolved||0}</span>`)}
      </table>
    </td>
  </tr></table>

  <table style="width:100%;border-collapse:separate;border-spacing:12px 0;margin-bottom:16px;"><tr>
    <td style="width:50%;vertical-align:top;">
      <div style="background:linear-gradient(135deg,#e65100,#FF9800);color:#fff;padding:10px 16px;border-radius:10px 10px 0 0;">
        <strong style="font-size:13px;">🟠 Tier 2 — Programming</strong>
      </div>
      <div style="background:#fff;padding:12px 16px;border-radius:0 0 10px 10px;">
        <table style="width:100%;font-size:12px;">
          ${row("SLA", `${fm(t2.slaCompliance)} ${ic(t2.slaCompliance, 90, false)}`)}
          ${row("Response SLA", `${fm(t2.slaResponseCompliance)} ${ic(t2.slaResponseCompliance, 90, false)}`)}
          ${row("Breach", `${t2.openBreachRate||0}% ${ic(t2.openBreachRate||0, 5, true)}`)}
          ${row("Cases", t2.total||0)}
          ${row("Resolved", t2.resolved||0)}
        </table>
      </div>
    </td>
    <td style="width:50%;vertical-align:top;">
      <div style="background:linear-gradient(135deg,#7b1fa2,#9C27B0);color:#fff;padding:10px 16px;border-radius:10px 10px 0 0;">
        <strong style="font-size:13px;">🟣 Tier 3 — Relationship Mgrs</strong>
      </div>
      <div style="background:#fff;padding:12px 16px;border-radius:0 0 10px 10px;">
        <table style="width:100%;font-size:12px;">
          ${row("SLA", `${fm(t3.slaCompliance)} ${ic(t3.slaCompliance, 90, false)}`)}
          ${row("Response SLA", `${fm(t3.slaResponseCompliance)} ${ic(t3.slaResponseCompliance, 90, false)}`)}
          ${row("Breach", `${t3.openBreachRate||0}% ${ic(t3.openBreachRate||0, 5, true)}`)}
          ${row("Cases", t3.total||0)}
          ${row("Resolved", t3.resolved||0)}
        </table>
      </div>
    </td>
  </tr></table>

  <div style="background:linear-gradient(135deg,#1a2332,#2d4a6f);color:#fff;padding:20px 24px;border-radius:12px;text-align:center;">
    <strong style="font-size:13px;">📈 OVERALL SUMMARY</strong>
    <table style="width:100%;margin-top:12px;"><tr>
      <td style="text-align:center"><div style="font-size:24px;font-weight:700;color:#4FC3F7;">${ov.created||0}</div><div style="font-size:10px;color:#a8c6df;">Created</div></td>
      <td style="text-align:center"><div style="font-size:24px;font-weight:700;color:#81C784;">${ov.resolved||0}</div><div style="font-size:10px;color:#a8c6df;">Resolved</div></td>
      <td style="text-align:center"><div style="font-size:24px;font-weight:700;color:#FFB74D;">${cs.responses||0}</div><div style="font-size:10px;color:#a8c6df;">CSAT</div></td>
      <td style="text-align:center"><div style="font-size:24px;font-weight:700;color:#81C784;">${ph.answered||0}</div><div style="font-size:10px;color:#a8c6df;">Answered</div></td>
      <td style="text-align:center"><div style="font-size:24px;font-weight:700;color:#f44336;">${ph.abandoned||0}</div><div style="font-size:10px;color:#a8c6df;">Abandoned</div></td>
    </tr></table>
  </div>
  <p style="text-align:center;font-size:10px;color:#999;margin-top:16px;">Report generated from live Dynamics 365 data · Service and Operations Dashboard</p>
</div>`;
  }

  const handleSendReport = async (recipientEmail, recipientNote) => {
    if (!d365Account) throw new Error("Sign in with Microsoft first.");
    const html = buildEmailHTML(recipientNote);
    if (!html) throw new Error("No report data. Run the report first.");
    const subject = `📊 Service & Operations Report — ${dateLabel}`;
    await sendEmailViaGraph(recipientEmail, subject, html);
    return true;
  };

  const timeLabel = (startTime !== "00:00" || endTime !== "23:59") ? ` · ${startTime} — ${endTime}` : "";
  const dateLabel = startDate === endDate
    ? new Date(startDate + "T12:00:00").toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" }) + timeLabel
    : `${new Date(startDate + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })} — ${new Date(endDate + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}` + timeLabel;

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'DM Sans', sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Space+Mono:wght@400;700&family=Playfair+Display:wght@600;700;800&display=swap" rel="stylesheet" />
      <style>{`
@keyframes fadeIn { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
@keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.5; } }
@keyframes slideIn { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: translateY(0); } }
@media print { .no-print { display: none !important; } }
@media (max-width: 900px) {
  .dash-header { flex-wrap: wrap; gap: 8px !important; padding: 14px 16px !important; }
  .dash-header-title { font-size: 14px !important; }
  .dash-header-sub { display: none !important; }
  .dash-header-actions { gap: 6px !important; }
  .dash-header-actions > span:first-child { display: none !important; }
  .dash-layout { flex-direction: column !important; }
  .dash-sidebar { width: 100% !important; min-width: unset !important; min-height: unset !important; border-right: none !important; border-bottom: 1px solid ${C.border} !important; padding: 16px !important; }
  .dash-main { padding: 16px !important; min-height: unset !important; }
  .dash-sidebar-toggle { display: flex !important; }
  .metric-grid-2 { grid-template-columns: 1fr !important; }
  .chart-grid-2 { grid-template-columns: 1fr !important; }
  .stat-row { flex-wrap: wrap !important; }
  .conn-bar { flex-wrap: wrap; gap: 6px !important; padding: 6px 16px !important; font-size: 10px !important; }
}
`}</style>
      <SettingsModal show={showSettings} onClose={() => setShowSettings(false)} config={apiConfig} onSave={setApiConfig} d365Account={d365Account} onD365Login={handleD365Login} onD365Logout={handleD365Logout} />
      <SendReportModal show={showSendModal} onClose={() => setShowSendModal(false)} onSend={handleSendReport} dateLabel={dateLabel} />
      <AutoReportModal show={showAutoReport} onClose={() => setShowAutoReport(false)} queues={queues} d365Account={d365Account} />
      <div className="no-print dash-header" style={{ background: C.primary, padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, minWidth: 0 }}>
          <div style={{ width: 38, height: 38, borderRadius: 9, background: `linear-gradient(135deg, ${C.accent}, ${C.yellow})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 700, color: "#fff", flexShrink: 0 }}>S</div>
          <div style={{ minWidth: 0 }}><h1 className="dash-header-title" style={{ margin: 0, fontSize: 18, fontWeight: 700, color: "#fff", fontFamily: "'Playfair Display', serif", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>Service and Operations Dashboard</h1><div className="dash-header-sub" style={{ fontSize: 11, color: "#B3D4F7", marginTop: 1, letterSpacing: 0.5 }}>Dynamics 365 · Operations</div></div>
        </div>
        <div className="dash-header-actions" style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          <span style={{ fontSize: 13, color: "#ffffff80", fontWeight: 500, whiteSpace: "nowrap" }}>👤 {user?.name || "User"}</span>
          {d365Account && <span style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: "#4CAF5030", color: "#81C784", fontWeight: 600 }}>🟢 D365</span>}
          {hasRun && <button onClick={handleExportPDF} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}><span>📄</span> Export PDF</button>}
          {hasRun && d365Account && <button onClick={() => setShowSendModal(true)} style={{ background: "linear-gradient(135deg, #0078D440, #0078D420)", color: "#fff", border: "1px solid #0078D450", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}><span>📤</span> Send</button>}
          <button onClick={() => setShowSettings(true)} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 14px", fontSize: 14, cursor: "pointer" }}>⚙️</button>
          <button onClick={onLogout} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Logout</button>
        </div>
      </div>
      <ConnectionBar d365Connected={!!d365Account} isLive={isLive} onOpenSettings={() => setShowSettings(true)} />
      <div className="dash-layout" style={{ display: "flex", maxWidth: 1500, margin: "0 auto" }}>
        <div className="no-print dash-sidebar" style={{ width: 310, minWidth: 310, background: C.card, borderRight: `1px solid ${C.border}`, padding: "24px 20px", minHeight: "calc(100vh - 110px)", display: "flex", flexDirection: "column" }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: C.textLight, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 14 }}>Configure Report</div>
          {d365Account && (
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}><span>🏢</span> Tier {loadingQueues && <span style={{ fontSize: 10, color: C.accent, animation: "pulse 1s infinite" }}>Loading...</span>}</div>
              <select value={selectedQueue || ""} onChange={(e) => setSelectedQueue(e.target.value || null)} style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1.5px solid ${C.border}`, fontSize: 13, fontFamily: "'DM Sans', sans-serif", background: C.bg, color: C.textDark, outline: "none", cursor: "pointer", appearance: "auto" }}>
                <option value="">Select a tier...</option>
                <option value="all">All Tiers</option>
                {queues.map(q => (<option key={q.id} value={q.id}>{q.tierLabel || q.name}</option>))}
              </select>
              {selectedQueue && (<div style={{ marginTop: 6, fontSize: 10, color: C.textLight }}>{queues.find(q => q.id === selectedQueue)?.description || ""}</div>)}
            </div>
          )}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}><span>👥</span> Team Members {loadingMembers && <span style={{ fontSize: 10, color: C.accent, animation: "pulse 1s infinite" }}>Loading from D365...</span>}</div>
            <MultiMemberSelect selected={selectedMembers} onChange={setSelectedMembers} members={teamMembers} />
            {d365Account && selectedQueue && selectedQueue !== "all" && !loadingMembers && teamMembers.length === 0 && (<div style={{ marginTop: 6, fontSize: 11, color: C.blue, padding: "8px 10px", background: C.blueLight, borderRadius: 8 }}>No individual members found — you can still run the report using the tier's case data.</div>)}
            {selectedMembers.length > 0 && <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 4 }}>
              {selectedMembers.slice(0, 4).map((id) => { const m = teamMembers.find((t) => t.id === id); const idx = teamMembers.indexOf(m); return <span key={id} style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: PIE_COLORS[idx % PIE_COLORS.length] + "18", color: PIE_COLORS[idx % PIE_COLORS.length], fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}>{m?.name?.split(" ")[0]}<span onClick={() => setSelectedMembers(selectedMembers.filter((s) => s !== id))} style={{ cursor: "pointer", opacity: 0.6, fontSize: 8 }}>✕</span></span>; })}
              {selectedMembers.length > 4 && <span style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: C.bg, color: C.textLight, fontWeight: 600 }}>+{selectedMembers.length - 4} more</span>}
            </div>}
          </div>
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textDark, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}><span>📅</span> Date Range</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <label><div style={{ fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 }}>From</div><input type="date" value={startDate} onChange={(e) => { setStartDate(e.target.value); setReportType("custom"); }} style={{ width: "100%", padding: "8px 8px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 11, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} /></label>
              <label><div style={{ fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 }}>To</div><input type="date" value={endDate} onChange={(e) => { setEndDate(e.target.value); setReportType("custom"); }} style={{ width: "100%", padding: "8px 8px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 11, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} /></label>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 8 }}>
              <label><div style={{ fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 }}>From Time</div><input type="time" value={startTime} onChange={(e) => setStartTime(e.target.value)} style={{ width: "100%", padding: "8px 8px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 11, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} /></label>
              <label><div style={{ fontSize: 10, color: C.textLight, marginBottom: 3, fontWeight: 500 }}>To Time</div><input type="time" value={endTime} onChange={(e) => setEndTime(e.target.value)} style={{ width: "100%", padding: "8px 8px", borderRadius: 8, border: `1.5px solid ${C.border}`, fontSize: 11, fontFamily: "'Space Mono', monospace", background: C.bg, color: C.textDark, outline: "none", boxSizing: "border-box" }} /></label>
            </div>
            <div style={{ display: "flex", gap: 5, marginTop: 8 }}>
              {[{ l: "7D", s: 7 }, { l: "14D", s: 14 }, { l: "30D", s: 30 }, { l: "90D", s: 90 }, { l: "YTD", s: -1 }].map((q) => <button key={q.l} onClick={() => { const e = new Date(), sD = new Date(); if (q.s === -1) { sD.setMonth(0); sD.setDate(1); } else { sD.setDate(sD.getDate() - q.s); } setStartDate(sD.toISOString().split("T")[0]); setEndDate(e.toISOString().split("T")[0]); setStartTime("00:00"); setEndTime("23:59"); setReportType("custom"); }} style={{ flex: 1, padding: "5px 0", borderRadius: 6, border: `1px solid ${C.border}`, background: "transparent", fontSize: 10, fontWeight: 600, color: C.textMid, cursor: "pointer", fontFamily: "'Space Mono', monospace" }}>{q.l}</button>)}
            </div>
          </div>
          <div style={{ flex: 1 }} />
          <button onClick={handleRun} disabled={!canRun || isRunning} style={{ width: "100%", padding: "14px", borderRadius: 12, border: "none", background: canRun ? `linear-gradient(135deg, ${C.accent}, ${C.yellow})` : C.border, color: canRun ? "#fff" : C.textLight, fontSize: 15, fontWeight: 700, cursor: canRun ? "pointer" : "not-allowed", letterSpacing: 0.5, display: "flex", alignItems: "center", justifyContent: "center", gap: 10, boxShadow: canRun ? "0 4px 20px rgba(232,101,58,0.35)" : "none", opacity: isRunning ? 0.7 : 1 }}>
            {isRunning ? <><span style={{ animation: "pulse 1s infinite" }}>⏳</span> {runProgress || "Generating..."}</> : <><span style={{ fontSize: 18 }}>▶</span> Run Report (Live)</>}
          </button>
          {!canRun && <div style={{ fontSize: 10, color: C.accent, textAlign: "center", marginTop: 6 }}>{d365Account ? "Select a tier to run report" : "Select at least 1 team member"}</div>}
          {hasRun && <button onClick={handleExportPDF} style={{ width: "100%", padding: "12px", borderRadius: 10, border: `1.5px solid ${C.border}`, background: C.card, color: C.textDark, fontSize: 13, fontWeight: 600, cursor: "pointer", marginTop: 10, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><span>📄</span> Export to PDF</button>}
          {hasRun && <button onClick={() => setShowSendModal(true)} style={{ width: "100%", padding: "12px", borderRadius: 10, border: `1.5px solid ${d365Account ? "#0078D4" : C.border}`, background: d365Account ? "#0078D410" : C.card, color: d365Account ? "#0078D4" : C.textLight, fontSize: 13, fontWeight: 600, cursor: d365Account ? "pointer" : "not-allowed", marginTop: 6, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><span>📤</span> Send Report{!d365Account && <span style={{ fontSize: 9, opacity: 0.7 }}>(sign in first)</span>}</button>}
          <button onClick={() => setShowAutoReport(true)} style={{ width: "100%", padding: "12px", borderRadius: 10, border: `1.5px solid #FF980060`, background: "#FF980008", color: "#e65100", fontSize: 13, fontWeight: 600, cursor: "pointer", marginTop: 6, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><span>⏰</span> Auto Reports</button>
        </div>
        <div className="dash-main" style={{ flex: 1, padding: "24px 28px", overflow: "auto", minHeight: "calc(100vh - 110px)" }}>
          {!hasRun ? (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", textAlign: "center" }}>
              <div style={{ width: 100, height: 100, borderRadius: 24, background: `linear-gradient(135deg, ${C.accent}15, ${C.yellow}15)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 44, marginBottom: 20 }}>📊</div>
              <h2 style={{ margin: "0 0 8px", fontSize: 22, fontWeight: 700, color: C.textDark, fontFamily: "'Playfair Display', serif" }}>Service and Operations Dashboard</h2>
              <p style={{ margin: 0, fontSize: 14, color: C.textMid, maxWidth: 440, lineHeight: 1.6 }}>Select your team members, choose a report type, set your date range, and hit <strong style={{ color: C.accent }}>Run Report</strong>. Live data from <strong style={{ color: C.d365 }}>Dynamics 365</strong>.</p>
              <div style={{ marginTop: 20, display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "center" }}>
                {[{ icon: "👥", label: `${selectedMembers.length} members`, ok: selectedMembers.length > 0 }, { icon: "📊", label: reportType === "daily" ? "Daily Report" : reportType === "weekly" ? "Weekly Report" : "Custom Range", ok: true }, { icon: "📅", label: `${startDate} ${startTime} → ${endDate} ${endTime}`, ok: startDate && endDate }].map((s, i) => <div key={i} style={{ padding: "10px 16px", borderRadius: 10, background: s.ok ? C.greenLight + "22" : C.accentLight + "22", border: `1px solid ${s.ok ? C.greenLight + "44" : C.accentLight + "44"}`, fontSize: 12, fontWeight: 600, color: s.ok ? C.green : C.accent, display: "flex", alignItems: "center", gap: 6 }}><span>{s.icon}</span> {s.label} {s.ok ? "✓" : "✗"}</div>)}
              </div>
              {!d365Account && (
                <div style={{ marginTop: 24, padding: "14px 20px", borderRadius: 12, background: `${C.d365}08`, border: `1px solid ${C.d365}20`, maxWidth: 440 }}>
                  <div style={{ fontSize: 12, color: C.textMid, lineHeight: 1.5 }}>💡 Click <strong>⚙️ Settings</strong> and <strong style={{ color: C.d365 }}>Sign in with Microsoft</strong> to connect D365 for live data.</div>
                </div>
              )}
            </div>
          ) : data && (
            <div style={{ animation: "slideIn 0.4s ease" }} ref={reportRef}>
              <div style={{ marginBottom: 24, display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
                <div>
                  <h2 style={{ margin: 0, fontSize: 24, fontWeight: 800, color: C.textDark, fontFamily: "'Playfair Display', serif" }}>📊 {memberData.length > 0 ? "Individual Performance Report" : `${reportType === "weekly" ? "Weekly" : reportType === "daily" ? "Daily" : "Custom"} Operations Report`}</h2>
                  <div style={{ fontSize: 12, color: C.textMid, marginTop: 4, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                    {selectedQueue && <span>🏢 {selectedQueue === "all" ? "All Tiers" : (queues.find(q => q.id === selectedQueue)?.tierLabel?.replace(" 🔒", "") || "Tier")}</span>}
                    <span>👥 {selectedMembers.length > 0 ? `${selectedMembers.length} member${selectedMembers.length > 1 ? "s" : ""}` : "All tier members"}</span>
                    <span>📅 {dateLabel}</span>
                    <span style={{ fontSize: 10, padding: "2px 8px", borderRadius: 4, background: C.greenLight + "33", color: C.green, fontWeight: 600 }}>🟢 Live Data</span>
                  </div>
                </div>
              </div>
              {liveErrors.length > 0 && (
                <div style={{ marginBottom: 16, padding: "12px 16px", borderRadius: 10, background: C.orangeLight, border: `1px solid ${C.orange}30` }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.orange, marginBottom: 4 }}>⚠️ Some data could not be fetched ({liveErrors.length} issue{liveErrors.length > 1 ? "s" : ""})</div>
                  {liveErrors.slice(0, 5).map((e, i) => <div key={i} style={{ fontSize: 10, color: C.textMid, lineHeight: 1.5 }}>• {e}</div>)}
                  {liveErrors.length > 5 && <div style={{ fontSize: 10, color: C.textLight }}>...and {liveErrors.length - 5} more</div>}
                </div>
              )}
              {memberData.length > 0 ? (
                <>
                  {memberData.map((md, i) => (<MemberSection key={md.member.id} memberData={md} index={i} />))}
                  <TeamSummary memberDataList={memberData} />
                  <Definitions />
                </>
              ) : (
                <>
                  {(() => {
                    if (selectedQueue === "all") { return [1, 2, 3].map(t => <TierSection key={t} tier={t} data={data} members={teamMembers} />); }
                    const selectedTierNum = queues.find(q => q.id === selectedQueue)?.tierNum;
                    if (selectedTierNum) { return <TierSection tier={selectedTierNum} data={data} members={teamMembers} />; }
                    return [1, 2, 3].map(t => <TierSection key={t} tier={t} data={data} members={teamMembers} />);
                  })()}
                  <OverallSummary data={data} />
                  <Definitions />
                </>
              )}
              <div style={{ background: C.primaryDark, padding: 14, textAlign: "center", borderRadius: "0 0 14px 14px" }}>
                <p style={{ margin: 0, color: "#a8c6df", fontSize: 11 }}>Report generated from live Dynamics 365 data by Service and Operations Dashboard</p>
              </div>
              <ChartsPanel data={data} />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════
   APP ROOT — CHANGED: No more LoginPage, logout reloads to MSAL landing
   ═══════════════════════════════════════════════════════ */
export default function App() {
  const [user, setUser] = useState(null);
  const [checking, setChecking] = useState(true);
  useEffect(() => { const s = Auth.session(); if (s) setUser(s); setChecking(false); }, []);
  if (checking) return <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: C.primary, color: "#fff", fontFamily: "'DM Sans',sans-serif" }}><div style={{ textAlign: "center" }}><div style={{ width: 56, height: 56, borderRadius: 16, margin: "0 auto 16px", background: `linear-gradient(135deg, ${C.accent}, ${C.gold})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28, fontWeight: 800, color: "#fff" }}>S</div><div style={{ fontSize: 14, color: "#ffffff60" }}>Loading...</div></div></div>;
  if (!user) {
    Auth.logout();
    sessionStorage.clear();
    window.location.reload();
    return null;
  }
  return <Dashboard user={user} onLogout={() => { Auth.logout(); sessionStorage.clear(); window.location.reload(); }} />;
}

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
    tier1: { total: t1Cases, slaMet: t1SLAMet, slaCompliance: t1Cases ? Math.min(100, Math.round(t1SLAMet / t1Cases * 100)) : 0, fcrRate: t1Cases ? Math.min(100, Math.round(t1FCR / t1Cases * 100)) : 0, escalationRate: t1Cases ? Math.min(100, Math.round(t1Escalated / t1Cases * 100)) : 0, avgResolutionTime: `${avgResTime} hrs`, escalated: t1Escalated },
    tier2: { total: t2Cases, resolved: t2Resolved, slaMet: t2SLAMet, slaCompliance: t2Resolved ? Math.min(100, Math.round(t2SLAMet / t2Resolved * 100)) : "N/A", escalationRate: t2Cases ? Math.min(100, Math.round(t2Escalated / t2Cases * 100)) : "N/A", escalated: t2Escalated },
    tier3: { total: t3Cases, resolved: t3Resolved, slaMet: t3SLAMet, slaCompliance: t3Resolved ? Math.min(100, Math.round(t3SLAMet / t3Resolved * 100)) : "N/A" },
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
  const slaMet = resolvedCases;
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

  const slaCompliance = totalCases > 0 ? Math.min(100, Math.round(slaMet / totalCases * 100)) : "N/A";
  const fcrRate = totalCases > 0 ? Math.min(100, Math.round(fcrCases / totalCases * 100)) : "N/A";
  const escalationRate = totalCases > 0 ? Math.min(100, Math.round(escalatedCases / totalCases * 100)) : "N/A";

  return {
    member,
    totalCases, resolvedCases, activeCases, slaMet, slaCompliance,
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

  const t1Cases = await safeCount("Tier 1 Cases",
    `incidents?$filter=casetypecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$count=true&$top=1`);
  const t1SLAMet = await safeFetchCount("SLA Met Cases",
    `incidents?$filter=casetypecode eq 1 and statecode eq 1 and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t1FCR = await safeFetchCount("FCR Cases",
    `incidents?$filter=casetypecode eq 1 and cr7fe_new_fcr eq true and createdon ge ${s}T${sT} and createdon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t1Escalated = await safeCount("Tier 1 Escalated",
    `incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);

  const t2Cases = await safeCount("Tier 2 Cases",
    `incidents?$filter=casetypecode eq 2 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);
  const t2Resolved = await safeFetchCount("Tier 2 Resolved",
    `incidents?$filter=casetypecode eq 2 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t2SLAMet = t2Resolved;
  const t2Escalated = await safeCount("Tier 2 Escalated to T3",
    `incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);

  const t3Cases = await safeCount("Tier 3 Cases",
    `incidents?$filter=casetypecode eq 3 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$count=true&$top=1`);
  const t3Resolved = await safeFetchCount("Tier 3 Resolved",
    `incidents?$filter=casetypecode eq 3 and statecode eq 1 and escalatedon ge ${s}T${sT} and escalatedon le ${e}T${eT}&$select=incidentid&$count=true`);
  const t3SLAMet = t3Resolved;

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
  const allResolved = t1SLAMet + t2Resolved + t3Resolved;

  return {
    tier1: { total: t1Cases, slaMet: t1SLAMet, slaCompliance: t1Cases ? Math.min(100, Math.round(t1SLAMet / t1Cases * 100)) : 0, fcrRate: t1Cases ? Math.min(100, Math.round(t1FCR / t1Cases * 100)) : 0, escalationRate: t1Cases ? Math.min(100, Math.round(t1Escalated / t1Cases * 100)) : 0, avgResolutionTime: avgResTime, escalated: t1Escalated },
    tier2: { total: t2Cases, resolved: t2Resolved, slaMet: t2SLAMet, slaCompliance: t2Resolved ? Math.min(100, Math.round(t2SLAMet / t2Resolved * 100)) : "N/A", escalationRate: t2Cases ? Math.min(100, Math.round(t2Escalated / t2Cases * 100)) : "N/A", escalated: t2Escalated },
    tier3: { total: t3Cases, resolved: t3Resolved, slaMet: t3SLAMet, slaCompliance: t3Resolved ? Math.min(100, Math.round(t3SLAMet / t3Resolved * 100)) : "N/A" },
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
  const slaMet = Math.min(d.slaMet || 0, d.total);
  const slaMissed = Math.max(0, d.total - slaMet);
  const metrics = [];
  if (t.metrics.includes("sla_compliance")) metrics.push({ label: "SLA Compliance", value: slaRate, target: 90, unit: "%" });
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
        <StatCard label="SLAs Met" value={`${slaMet}/${d.total}`} color="#2D9D78" />
        <StatCard label="SLAs Missed" value={`${slaMissed}/${d.total}`} color={slaMissed > 0 ? "#E5544B" : "#2D9D78"} />
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
    ["SLA Compliance", "Percentage of resolved cases meeting resolution time targets based on priority level"],
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
  const slaMet = Math.min(d.slaMet || 0, d.totalCases);
  const slaMissed = Math.max(0, d.totalCases - slaMet);
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
        <StatCard label="SLAs Met" value={`${slaMet}/${d.totalCases}`} color={slaMet > 0 ? "#2D9D78" : "#E5544B"} />
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
    slaMet: acc.slaMet + d.slaMet, fcrCases: acc.fcrCases + d.fcrCases,
    emailCases: acc.emailCases + d.emailCases, emailResolved: acc.emailResolved + d.emailResolved,
    csatResponses: acc.csatResponses + d.csatResponses,
    csatTotal: acc.csatTotal + (d.csatAvg !== "N/A" ? d.csatAvg * d.csatResponses : 0),
  }), { totalCases: 0, resolved: 0, active: 0, escalated: 0, slaMet: 0, fcrCases: 0, emailCases: 0, emailResolved: 0, csatResponses: 0, csatTotal: 0 });
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
          <div style={{ textAlign: "center" }}><StatusBadge status={checkTarget("sla_compliance", totals.totalCases ? Math.min(100, Math.round(totals.slaMet / totals.totalCases * 100)) : 0)} value={totals.totalCases ? Math.min(100, Math.round(totals.slaMet / totals.totalCases * 100)) : 0} unit="%" /><div style={{ fontSize: 10, color: "#a8c6df", marginTop: 4 }}>Team SLA</div></div>
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
      <div className="no-print dash-header" style={{ background: C.primary, padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, minWidth: 0 }}>
          <div style={{ width: 38, height: 38, borderRadius: 9, background: `linear-gradient(135deg, ${C.accent}, ${C.yellow})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 700, color: "#fff", flexShrink: 0 }}>S</div>
          <div style={{ minWidth: 0 }}><h1 className="dash-header-title" style={{ margin: 0, fontSize: 18, fontWeight: 700, color: "#fff", fontFamily: "'Playfair Display', serif", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>Service and Operations Dashboard</h1><div className="dash-header-sub" style={{ fontSize: 11, color: "#B3D4F7", marginTop: 1, letterSpacing: 0.5 }}>Dynamics 365 · Operations</div></div>
        </div>
        <div className="dash-header-actions" style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          <span style={{ fontSize: 13, color: "#ffffff80", fontWeight: 500, whiteSpace: "nowrap" }}>👤 {user?.name || "User"}</span>
          {d365Account && <span style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: "#4CAF5030", color: "#81C784", fontWeight: 600 }}>🟢 D365</span>}
          {hasRun && <button onClick={handleExportPDF} style={{ background: "linear-gradient(135deg, #fff2, #fff1)", color: "#fff", border: "1px solid #fff3", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}><span>📄</span> Export PDF</button>}
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

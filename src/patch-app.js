#!/usr/bin/env node
/**
 * patch-app.js — Apply DrillDownView integration to App.jsx
 * 
 * Usage:
 *   node patch-app.js App.jsx
 *   → Creates App.jsx.bak (backup) and patches App.jsx in place
 * 
 * 4 Changes Applied:
 *   1. Import DrillDownView component
 *   2. Add viewMode state in Dashboard
 *   3. Add View Mode toggle in sidebar
 *   4. Wrap main content with viewMode conditional
 */

const fs = require("fs");
const path = require("path");

const file = process.argv[2] || "App.jsx";
if (!fs.existsSync(file)) {
  console.error(`❌ File not found: ${file}`);
  console.log("Usage: node patch-app.js <path-to-App.jsx>");
  process.exit(1);
}

let content = fs.readFileSync(file, "utf-8");
const original = content;
let changes = 0;

// ═══ CHANGE 1: Add import ═══
const importMarker = 'import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";';
if (content.includes(importMarker) && !content.includes('import DrillDownView')) {
  content = content.replace(
    importMarker,
    importMarker + '\nimport DrillDownView from "./DrillDownView";'
  );
  changes++;
  console.log("✅ Change 1: Added DrillDownView import");
} else if (content.includes('import DrillDownView')) {
  console.log("⏩ Change 1: DrillDownView import already exists");
} else {
  console.log("❌ Change 1: Could not find import marker");
}

// ═══ CHANGE 2: Add viewMode state ═══
const stateMarker = 'const [memberData, setMemberData] = useState([]);';
if (content.includes(stateMarker) && !content.includes('viewMode')) {
  content = content.replace(
    stateMarker,
    stateMarker + '\n  const [viewMode, setViewMode] = useState("report");'
  );
  changes++;
  console.log("✅ Change 2: Added viewMode state");
} else if (content.includes('viewMode')) {
  console.log("⏩ Change 2: viewMode state already exists");
} else {
  console.log("❌ Change 2: Could not find state marker");
}

// ═══ CHANGE 3: Add View Mode toggle in sidebar ═══
const sidebarMarker = `          {/* 8x8 Agent Status Monitor */}\n          <AgentMonitor teamMembers={teamMembers} />`;
if (content.includes(sidebarMarker) && !content.includes('VIEW MODE TOGGLE')) {
  const viewToggle = `          {/* VIEW MODE TOGGLE */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: C.textLight, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 8 }}>View Mode</div>
            <div style={{ display: "flex", gap: 6 }}>
              <button onClick={() => setViewMode("report")} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, border: \`1.5px solid \${viewMode === "report" ? C.accent : C.border}\`, background: viewMode === "report" ? \`\${C.accent}12\` : "transparent", color: viewMode === "report" ? C.accent : C.textMid, fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, fontFamily: "'DM Sans', sans-serif" }}>
                <span>📊</span> Report
              </button>
              <button onClick={() => setViewMode("drilldown")} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, border: \`1.5px solid \${viewMode === "drilldown" ? "#3B82F6" : C.border}\`, background: viewMode === "drilldown" ? "rgba(59,130,246,0.08)" : "transparent", color: viewMode === "drilldown" ? "#3B82F6" : C.textMid, fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, fontFamily: "'DM Sans', sans-serif" }}>
                <span>🔍</span> Drill Down
              </button>
            </div>
          </div>

          {/* 8x8 Agent Status Monitor */}
          <AgentMonitor teamMembers={teamMembers} />`;
  content = content.replace(sidebarMarker, viewToggle);
  changes++;
  console.log("✅ Change 3: Added View Mode toggle in sidebar");
} else if (content.includes('VIEW MODE TOGGLE')) {
  console.log("⏩ Change 3: View Mode toggle already exists");
} else {
  console.log("❌ Change 3: Could not find sidebar marker");
}

// ═══ CHANGE 4: Wrap main content with viewMode conditional ═══
const mainMarker = `        <div className="dash-main" style={{ flex: 1, padding: "32px 40px", overflow: "auto", minHeight: "calc(100vh - 110px)" }}>\n          {!hasRun ? (`;
if (content.includes(mainMarker) && !content.includes('DrillDownView onBack')) {
  const newMain = `        <div className="dash-main" style={{ flex: 1, padding: viewMode === "drilldown" ? "0" : "32px 40px", overflow: "auto", minHeight: "calc(100vh - 110px)" }}>
          {viewMode === "drilldown" ? (
            <DrillDownView onBack={() => setViewMode("report")} />
          ) : !hasRun ? (`;
  content = content.replace(mainMarker, newMain);
  changes++;
  console.log("✅ Change 4: Wrapped main content with viewMode conditional");
} else if (content.includes('DrillDownView onBack')) {
  console.log("⏩ Change 4: viewMode conditional already exists");
} else {
  console.log("❌ Change 4: Could not find main content marker");
}

// ═══ SAVE ═══
if (changes > 0) {
  // Backup original
  fs.writeFileSync(file + ".bak", original);
  console.log(`\n💾 Backup saved: ${file}.bak`);
  
  // Write patched file
  fs.writeFileSync(file, content);
  console.log(`✅ Patched ${file} — ${changes} change${changes !== 1 ? "s" : ""} applied`);
  console.log(`\n📋 Don't forget to place DrillDownView.jsx in the same directory as App.jsx!`);
} else {
  console.log("\n⚠️  No changes applied — file may already be patched or markers not found.");
}

// src/LandingPage.jsx
import React from "react";

const styles = {
    page: {
        fontFamily: "'DM Sans', sans-serif",
        background: "#0d0f14",
        color: "#f0f0f0",
        minHeight: "100vh",
        position: "relative",
        overflow: "hidden",
    },
    glowOrange: {
        position: "fixed", width: 400, height: 400, borderRadius: "50%",
        filter: "blur(120px)", opacity: 0.15, background: "#e8922a",
        top: -100, right: -100, pointerEvents: "none", zIndex: 0,
    },
    glowBlue: {
        position: "fixed", width: 400, height: 400, borderRadius: "50%",
        filter: "blur(120px)", opacity: 0.08, background: "#4a9eff",
        bottom: -100, left: -100, pointerEvents: "none", zIndex: 0,
    },
    header: {
        position: "sticky", top: 0, zIndex: 100,
        background: "rgba(13,15,20,0.85)", backdropFilter: "blur(20px)",
        WebkitBackdropFilter: "blur(20px)",
        borderBottom: "1px solid rgba(255,255,255,0.06)",
        padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between",
    },
    logoGroup: { display: "flex", alignItems: "center", gap: 10 },
    logoIcon: {
        width: 36, height: 36,
        background: "linear-gradient(135deg, #e8922a, #d47a15)",
        borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center",
        fontWeight: 800, fontSize: 18, color: "white", flexShrink: 0,
    },
    logoTitle: { fontWeight: 700, fontSize: 13, color: "#f0f0f0", lineHeight: 1.15 },
    logoSubtitle: {
        fontSize: 9, fontWeight: 500, letterSpacing: 2,
        textTransform: "uppercase", color: "#5a5e72", lineHeight: 1.15,
    },
    signInBtn: {
        background: "linear-gradient(135deg, #e8922a, #d47a15)",
        color: "white", border: "none", padding: "10px 20px", borderRadius: 10,
        fontFamily: "'DM Sans', sans-serif", fontWeight: 700, fontSize: 14,
        cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
        whiteSpace: "nowrap",
    },
    signInBtnLarge: {
        background: "linear-gradient(135deg, #e8922a, #d47a15)",
        color: "white", border: "none", padding: "14px 32px", borderRadius: 14,
        fontFamily: "'DM Sans', sans-serif", fontWeight: 700, fontSize: 16,
        cursor: "pointer", display: "flex", alignItems: "center", gap: 8,
        boxShadow: "0 4px 20px rgba(232,146,42,0.25)", margin: "0 auto 32px",
    },
    hero: {
        padding: "48px 20px 36px", textAlign: "center", position: "relative", zIndex: 1,
    },
    silhouette: {
        position: "absolute", top: "50%", left: "50%",
        transform: "translate(-50%, -50%)",
        width: 380, height: 380, opacity: 0.12,
        pointerEvents: "none", zIndex: -1, objectFit: "contain",
    },
    heroLogoIcon: {
        width: 64, height: 64,
        background: "linear-gradient(135deg, #e8922a, #d47a15)",
        borderRadius: 16, display: "flex", alignItems: "center", justifyContent: "center",
        fontWeight: 800, fontSize: 32, color: "white",
        boxShadow: "0 8px 32px rgba(232,146,42,0.25)",
        margin: "0 auto 14px",
    },
    heroTitle: {
        fontFamily: "'Plus Jakarta Sans', sans-serif",
        fontWeight: 800, fontSize: "clamp(32px, 8vw, 56px)", lineHeight: 1.05, marginBottom: 16,
    },
    highlight: {
        background: "linear-gradient(135deg, #e8922a, #f5c842)",
        WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text",
    },
    heroDesc: {
        color: "#8b8fa3", fontSize: 15, lineHeight: 1.6, maxWidth: 380, margin: "0 auto 28px",
    },
    tierGrid: {
        display: "flex", flexDirection: "column", gap: 12,
        maxWidth: 480, margin: "0 auto 40px", padding: "0 20px",
    },
    tierCard: {
        background: "#161922", border: "1px solid rgba(255,255,255,0.06)",
        borderRadius: 14, padding: "16px 18px", display: "flex", alignItems: "center", gap: 14,
    },
    tierDot: (color) => ({
        width: 10, height: 10, borderRadius: "50%", flexShrink: 0, background: color,
    }),
    tierName: { fontWeight: 700, fontSize: 14 },
    tierDesc: { fontSize: 12, color: "#8b8fa3" },
    tierBadge: {
        marginLeft: "auto", fontSize: 10, fontWeight: 600, letterSpacing: 1,
        textTransform: "uppercase", padding: "4px 10px", borderRadius: 6,
        background: "rgba(255,255,255,0.04)", color: "#5a5e72",
    },
    featuresGrid: {
        display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10,
        maxWidth: 480, margin: "0 auto 48px", padding: "0 20px",
    },
    featureItem: {
        background: "#161922", border: "1px solid rgba(255,255,255,0.06)",
        borderRadius: 12, padding: 16, textAlign: "center",
    },
    featIcon: { fontSize: 22, marginBottom: 8 },
    featText: { fontSize: 12, fontWeight: 600, color: "#8b8fa3", lineHeight: 1.4 },
    featLabel: {
        fontSize: 10, fontWeight: 600, letterSpacing: 2.5, textTransform: "uppercase",
        color: "#5a5e72", marginBottom: 16, textAlign: "center", padding: "0 20px",
    },
    footer: {
        textAlign: "center", padding: "24px 20px 32px",
        borderTop: "1px solid rgba(255,255,255,0.06)",
    },
    footerText: { fontSize: 11, color: "#5a5e72" },
    error: { color: "#e85a5a", fontSize: 13, textAlign: "center", marginBottom: 16 },
};

const SignInIcon = () => (
    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2.5">
        <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15m3 0l3-3m0 0l-3-3m3 3H9" />
    </svg>
);

const tiers = [
    { name: "Dynamics 365", desc: "Live data integration", color: "#e8922a", badge: "Source" },
    { name: "Tier 1 Service Desk", desc: "First response & triage", color: "#4a9eff", badge: "SLA" },
    { name: "Tier 2 Programming", desc: "Escalations & dev", color: "#f5a623", badge: "SLA" },
    { name: "Tier 3 Rel. Managers", desc: "Client operations", color: "#b07dd6", badge: "SLA" },
];

const features = [
    { icon: "📊", text: "Real-time KPIs" },
    { icon: "⏱️", text: "Avg Resolution Time" },
    { icon: "📞", text: "8x8 Call Metrics" },
    { icon: "🔒", text: "Azure AD Auth" },
];

export default function LandingPage({ onSignIn, loading, error }) {
    return (
        <div style={styles.page}>
            <div style={styles.glowOrange} />
            <div style={styles.glowBlue} />

            <header style={styles.header}>
                <div style={styles.logoGroup}>
                    <div style={styles.logoIcon}>S</div>
                    <div>
                        <div style={styles.logoTitle}>Service and Operations Dashboard</div>
                        <div style={styles.logoSubtitle}>Performance Analytics</div>
                    </div>
                </div>
                <button style={styles.signInBtn} onClick={onSignIn} disabled={loading}>
                    <SignInIcon /> {loading ? "Signing in..." : "Sign In"}
                </button>
            </header>

            <section style={styles.hero}>
                <img
                    src="./logo_bg.png"
                    alt=""
                    style={styles.silhouette}
                    onError={(e) => { e.target.style.display = "none"; }}
                />
                <div style={styles.heroLogoIcon}>S</div>
                <div style={{ fontWeight: 700, fontSize: 18, marginBottom: 6 }}>
                    Service and Operations Dashboard
                </div>
                <div style={{ fontSize: 10, fontWeight: 600, letterSpacing: 2.5, textTransform: "uppercase", color: "#5a5e72", marginBottom: 28 }}>
                    Performance Analytics
                </div>
                <h1 style={styles.heroTitle}>
                    Real-time SLA<br />
                    <span style={styles.highlight}>Intelligence</span>
                </h1>
                <p style={styles.heroDesc}>
                    Monitor your Service Desk, Programming Team, and Relationship Managers. Powered by Dynamics 365.
                </p>

                {error && <div style={styles.error}>⚠️ {error}</div>}

                <button style={styles.signInBtnLarge} onClick={onSignIn} disabled={loading}>
                    <SignInIcon /> {loading ? "Signing in..." : "Sign In with Microsoft"}
                </button>
            </section>

            <div style={styles.tierGrid}>
                {tiers.map((t) => (
                    <div key={t.name} style={styles.tierCard}>
                        <div style={styles.tierDot(t.color)} />
                        <div>
                            <div style={styles.tierName}>{t.name}</div>
                            <div style={styles.tierDesc}>{t.desc}</div>
                        </div>
                        <span style={styles.tierBadge}>{t.badge}</span>
                    </div>
                ))}
            </div>

            <div style={styles.featLabel}>What you get</div>
            <div style={styles.featuresGrid}>
                {features.map((f) => (
                    <div key={f.text} style={styles.featureItem}>
                        <div style={styles.featIcon}>{f.icon}</div>
                        <div style={styles.featText}>{f.text}</div>
                    </div>
                ))}
            </div>

            <footer style={styles.footer}>
                <p style={styles.footerText}>© 2026 ServingIntel · Service and Operations Dashboard</p>
            </footer>
        </div>
    );
}

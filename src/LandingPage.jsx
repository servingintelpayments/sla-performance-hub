// src/LandingPage.jsx
import React from "react";

var SignInIcon = function() {
    return React.createElement("svg", {
        width: 18, height: 18, fill: "none", viewBox: "0 0 24 24",
        stroke: "currentColor", strokeWidth: 2.5
    }, React.createElement("path", {
        strokeLinecap: "round", strokeLinejoin: "round",
        d: "M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15m3 0l3-3m0 0l-3-3m3 3H9"
    }));
};

var tiers = [
    { name: "Dynamics 365", desc: "Live data integration", color: "#e8922a", badge: "Source" },
    { name: "Tier 1 Service Desk", desc: "First response & triage", color: "#4a9eff", badge: "SLA" },
    { name: "Tier 2 Programming", desc: "Escalations & dev", color: "#f5a623", badge: "SLA" },
    { name: "Tier 3 Rel. Managers", desc: "Client operations", color: "#b07dd6", badge: "SLA" },
];

var features = [
    { icon: "\u{1F4CA}", text: "Real-time KPIs" },
    { icon: "\u23F1\uFE0F", text: "Avg Resolution Time" },
    { icon: "\u{1F4DE}", text: "8x8 Call Metrics" },
    { icon: "\u{1F512}", text: "Azure AD Auth" },
];

// Nuclear border kill — every property that could possibly create a visible edge
var kill = { border: "none", borderWidth: 0, borderStyle: "none", borderColor: "transparent", borderTop: "none", borderBottom: "none", borderLeft: "none", borderRight: "none", outline: "none", outlineWidth: 0, outlineStyle: "none", boxShadow: "none", WebkitAppearance: "none" };

export default function LandingPage(props) {
    var onSignIn = props.onSignIn;
    var loading = props.loading;
    var error = props.error;

    return React.createElement("div", {
        style: Object.assign({}, kill, {
            fontFamily: "'DM Sans', sans-serif",
            background: "#0d0f14",
            color: "#f0f0f0",
            minHeight: "100vh",
            position: "relative",
            overflow: "hidden",
        })
    },
        // Background glows
        React.createElement("div", { style: {
            position: "fixed", width: 400, height: 400, borderRadius: "50%",
            filter: "blur(120px)", opacity: 0.15, background: "#e8922a",
            top: -100, right: -100, pointerEvents: "none", zIndex: 0,
        }}),
        React.createElement("div", { style: {
            position: "fixed", width: 400, height: 400, borderRadius: "50%",
            filter: "blur(120px)", opacity: 0.08, background: "#4a9eff",
            bottom: -100, left: -100, pointerEvents: "none", zIndex: 0,
        }}),

        // Header
        React.createElement("header", { style: Object.assign({}, kill, {
            position: "sticky", top: 0, zIndex: 100,
            background: "rgba(13,15,20,0.95)", backdropFilter: "blur(20px)",
            WebkitBackdropFilter: "blur(20px)",
            padding: "14px 20px", display: "flex", alignItems: "center",
        })},
            React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 10 } },
                React.createElement("div", { style: {
                    width: 36, height: 36,
                    background: "linear-gradient(135deg, #e8922a, #d47a15)",
                    borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center",
                    fontWeight: 800, fontSize: 18, color: "white", flexShrink: 0,
                }}, "S"),
                React.createElement("div", null,
                    React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: "#f0f0f0", lineHeight: 1.15 } }, "Service and Operations Dashboard"),
                    React.createElement("div", { style: { fontSize: 9, fontWeight: 500, letterSpacing: 2, textTransform: "uppercase", color: "#5a5e72", lineHeight: 1.15 } }, "Performance Analytics")
                )
            )
        ),

        // Hero section
        React.createElement("section", { style: Object.assign({}, kill, {
            padding: "60px 20px 36px", textAlign: "center", position: "relative", zIndex: 1,
            minHeight: 400,
        })},
            React.createElement("img", {
                src: "./logo_bg.png",
                alt: "",
                style: Object.assign({}, kill, {
                    position: "absolute", top: "50%", left: "50%",
                    transform: "translate(-50%, -50%)",
                    width: "110vw", maxWidth: 900, height: "auto",
                    opacity: 0.07, pointerEvents: "none", zIndex: 0,
                    filter: "grayscale(100%) brightness(2)",
                }),
                onError: function(e) { e.target.style.display = "none"; },
            }),

            React.createElement("div", { style: { position: "relative", zIndex: 1 } },
                React.createElement("h1", { style: {
                    fontFamily: "'Plus Jakarta Sans', sans-serif",
                    fontWeight: 800, fontSize: "clamp(36px, 8vw, 60px)", lineHeight: 1.05, marginBottom: 20,
                }},
                    "Real-time SLA",
                    React.createElement("br"),
                    React.createElement("span", { style: {
                        background: "linear-gradient(135deg, #e8922a, #f5c842)",
                        WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text",
                    }}, "Intelligence")
                ),
                React.createElement("p", { style: {
                    color: "#8b8fa3", fontSize: 16, lineHeight: 1.6, maxWidth: 420, margin: "0 auto 32px",
                }}, "Monitor your Service Desk, Programming Team, and Relationship Managers. Powered by Dynamics 365."),

                error && React.createElement("div", { style: Object.assign({}, kill, {
                    color: "#e85a5a", fontSize: 13, textAlign: "center", marginBottom: 16,
                    background: "rgba(232,90,90,0.1)", padding: "8px 16px", borderRadius: 8,
                    display: "inline-block",
                })}, "\u26A0\uFE0F " + error),

                React.createElement("div", null,
                    React.createElement("button", {
                        style: Object.assign({}, kill, {
                            background: "linear-gradient(135deg, #e8922a, #d47a15)",
                            color: "white", padding: "16px 36px", borderRadius: 14,
                            fontFamily: "'DM Sans', sans-serif", fontWeight: 700, fontSize: 17,
                            cursor: loading ? "not-allowed" : "pointer",
                            display: "inline-flex", alignItems: "center", gap: 10,
                            boxShadow: "0 4px 24px rgba(232,146,42,0.3)", marginBottom: 48,
                            opacity: loading ? 0.7 : 1,
                        }),
                        onClick: onSignIn, disabled: loading,
                    }, React.createElement(SignInIcon), loading ? "Signing in..." : "Sign In with Microsoft")
                )
            )
        ),

        // Tier cards
        React.createElement("div", { style: Object.assign({}, kill, {
            display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
            gap: 10, maxWidth: 560, margin: "0 auto 40px", padding: "0 20px", position: "relative", zIndex: 1,
        })},
            tiers.map(function(t) {
                return React.createElement("div", {
                    key: t.name,
                    style: Object.assign({}, kill, {
                        background: "transparent", borderRadius: 14, padding: "16px 18px",
                        display: "flex", alignItems: "center", gap: 14,
                    })
                },
                    React.createElement("div", { style: {
                        width: 10, height: 10, borderRadius: "50%", flexShrink: 0, background: t.color,
                    }}),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14 } }, t.name),
                        React.createElement("div", { style: { fontSize: 12, color: "#8b8fa3" } }, t.desc)
                    ),
                    React.createElement("span", { style: Object.assign({}, kill, {
                        marginLeft: "auto", fontSize: 10, fontWeight: 600, letterSpacing: 1,
                        textTransform: "uppercase", padding: "4px 10px", borderRadius: 6,
                        background: "rgba(255,255,255,0.04)", color: "#5a5e72",
                    })}, t.badge)
                );
            })
        ),

        // Features label
        React.createElement("div", { style: Object.assign({}, kill, {
            fontSize: 10, fontWeight: 600, letterSpacing: 2.5, textTransform: "uppercase",
            color: "#5a5e72", marginBottom: 16, textAlign: "center", position: "relative", zIndex: 1,
        })}, "What you get"),

        // Features grid
        React.createElement("div", { style: Object.assign({}, kill, {
            display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10,
            maxWidth: 480, margin: "0 auto 48px", padding: "0 20px", position: "relative", zIndex: 1,
        })},
            features.map(function(f) {
                return React.createElement("div", {
                    key: f.text,
                    style: Object.assign({}, kill, {
                        background: "transparent", borderRadius: 12, padding: 16, textAlign: "center",
                    })
                },
                    React.createElement("div", { style: { fontSize: 22, marginBottom: 8 } }, f.icon),
                    React.createElement("div", { style: { fontSize: 12, fontWeight: 600, color: "#8b8fa3", lineHeight: 1.4 } }, f.text)
                );
            })
        ),

        // Footer
        React.createElement("footer", { style: Object.assign({}, kill, {
            textAlign: "center", padding: "24px 20px 32px", position: "relative", zIndex: 1,
        })},
            React.createElement("p", { style: { fontSize: 11, color: "#5a5e72" } }, "\u00A9 2026 ServingIntel \u00B7 Service and Operations Dashboard")
        )
    );
}

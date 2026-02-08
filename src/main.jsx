// src/main.jsx
import React, { useState, useEffect } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import LandingPage from "./LandingPage";
import { getMsalInstance, loginRequest } from "./authConfig";

function Root() {
    var _s1 = useState(null), msalRef = _s1[0], setMsalRef = _s1[1];
    var _s2 = useState(null), account = _s2[0], setAccount = _s2[1];
    var _s3 = useState(false), signInLoading = _s3[0], setSignInLoading = _s3[1];
    var _s4 = useState(null), error = _s4[0], setError = _s4[1];
    var _s5 = useState(true), initializing = _s5[0], setInitializing = _s5[1];

    // Initialize MSAL and check for redirect/existing session
    useEffect(function() {
        getMsalInstance()
            .then(function(instance) {
                setMsalRef(instance);
                var active = instance.getActiveAccount();
                if (active) {
                    // Auto-create session for App.jsx's Auth system
                    var session = { u: active.username, name: active.name || active.username, at: Date.now() };
                    localStorage.setItem("sla_session", JSON.stringify(session));
                    setAccount(active);
                }
                setInitializing(false);
            })
            .catch(function(err) {
                console.warn("MSAL init:", err.message);
                setInitializing(false);
            });
    }, []);

    // Sign in — redirect to Microsoft
    var handleSignIn = function() {
        if (signInLoading) return;
        setSignInLoading(true);
        setError(null);

        function doLogin(instance) {
            instance.loginRedirect(loginRequest).catch(function(err) {
                setSignInLoading(false);
                if (err && err.errorCode === "user_cancelled") return;
                setError("Sign in failed: " + (err && err.message ? err.message : "Unknown error"));
            });
        }

        if (msalRef) {
            doLogin(msalRef);
        } else {
            getMsalInstance()
                .then(function(instance) {
                    setMsalRef(instance);
                    doLogin(instance);
                })
                .catch(function(err) {
                    setSignInLoading(false);
                    setError("Auth failed to load. Please refresh the page.");
                });
        }
    };

    // Show nothing while checking for redirect response
    if (initializing) {
        return React.createElement("div", {
            style: {
                background: "#0d0f14", color: "#5a5e72", minHeight: "100vh",
                display: "flex", alignItems: "center", justifyContent: "center",
                fontFamily: "'DM Sans', sans-serif", fontSize: 14,
            }
        }, "Loading...");
    }

    // Not signed in → landing page
    if (!account) {
        return React.createElement(LandingPage, {
            onSignIn: handleSignIn,
            loading: signInLoading,
            error: error,
        });
    }

    // Signed in → go straight to App (which will find the sla_session we set)
    return React.createElement(App, null);
}

ReactDOM.createRoot(document.getElementById("root")).render(
    React.createElement(React.StrictMode, null,
        React.createElement(Root)
    )
);

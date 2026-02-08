// src/main.jsx
import React, { useState, useEffect, useRef } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import LandingPage from "./LandingPage";
import { getMsalInstance, loginRequest } from "./authConfig";

function Root() {
    var _s1 = useState(null), msalRef = _s1[0], setMsalRef = _s1[1];
    var _s2 = useState(null), account = _s2[0], setAccount = _s2[1];
    var _s3 = useState(false), signInLoading = _s3[0], setSignInLoading = _s3[1];
    var _s4 = useState(null), error = _s4[0], setError = _s4[1];

    // Load MSAL in background — never blocks UI
    useEffect(function() {
        getMsalInstance()
            .then(function(instance) {
                setMsalRef(instance);
                var active = instance.getActiveAccount();
                if (active) setAccount(active);
            })
            .catch(function(err) {
                console.warn("MSAL background init:", err.message);
            });
    }, []);

    // Sign in handler
    var handleSignIn = function() {
        if (signInLoading) return; // prevent double click
        setSignInLoading(true);
        setError(null);

        // 15s safety net — ALWAYS reset button
        var safetyTimer = setTimeout(function() {
            setSignInLoading(false);
            setError("Sign in timed out. Try again or allow popups for this site.");
        }, 15000);

        function done() {
            clearTimeout(safetyTimer);
            setSignInLoading(false);
        }

        function tryLogin(instance) {
            instance.loginPopup(loginRequest)
                .then(function(response) {
                    if (response && response.account) {
                        instance.setActiveAccount(response.account);
                        setAccount(response.account);
                    }
                    done();
                })
                .catch(function(err) {
                    done();
                    var code = err && err.errorCode;
                    if (code === "user_cancelled") return;
                    if (code === "popup_window_error" || code === "empty_window_error") {
                        setError("Popup blocked. Please allow popups for this site, then try again.");
                        return;
                    }
                    setError("Sign in failed: " + (err && err.message ? err.message : "Unknown error"));
                    console.error("Login error:", err);
                });
        }

        if (msalRef) {
            tryLogin(msalRef);
        } else {
            getMsalInstance()
                .then(function(instance) {
                    setMsalRef(instance);
                    tryLogin(instance);
                })
                .catch(function(err) {
                    done();
                    setError("Auth failed to load. Please refresh the page.");
                });
        }
    };

    // Always show landing page if not signed in — never show "Loading..."
    if (!account) {
        return React.createElement(LandingPage, {
            onSignIn: handleSignIn,
            loading: signInLoading,
            error: error,
        });
    }

    return React.createElement(App, {
        msalAccount: account,
        msalInstance: msalRef,
    });
}

ReactDOM.createRoot(document.getElementById("root")).render(
    React.createElement(React.StrictMode, null,
        React.createElement(Root)
    )
);

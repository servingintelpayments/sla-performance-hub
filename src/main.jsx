// src/main.jsx
import React, { useState, useEffect } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import LandingPage from "./LandingPage";
import { getMsalInstance, loginRequest } from "./authConfig";

function Root() {
    var _s1 = useState(null), msalInstance = _s1[0], setMsalInstance = _s1[1];
    var _s2 = useState(null), account = _s2[0], setAccount = _s2[1];
    var _s3 = useState(false), signInLoading = _s3[0], setSignInLoading = _s3[1];
    var _s4 = useState(null), error = _s4[0], setError = _s4[1];

    // Load MSAL in background
    useEffect(function() {
        getMsalInstance()
            .then(function(instance) {
                setMsalInstance(instance);
                var active = instance.getActiveAccount();
                if (active) {
                    setAccount(active);
                }
            })
            .catch(function(err) {
                console.error("MSAL init error:", err);
            });
    }, []);

    // Sign in handler
    var handleSignIn = function() {
        setSignInLoading(true);
        setError(null);

        // Safety timeout — reset after 15 seconds no matter what
        var timeout = setTimeout(function() {
            setSignInLoading(false);
            setError("Sign in timed out. Please try again.");
        }, 15000);

        // Get or create MSAL instance
        var go = msalInstance ? Promise.resolve(msalInstance) : getMsalInstance();

        go.then(function(instance) {
            setMsalInstance(instance);
            return instance.loginPopup(loginRequest);
        }).then(function(response) {
            clearTimeout(timeout);
            var instance = msalInstance || response.account;
            if (response && response.account) {
                getMsalInstance().then(function(inst) {
                    inst.setActiveAccount(response.account);
                    setAccount(response.account);
                    setSignInLoading(false);
                });
            }
        }).catch(function(err) {
            clearTimeout(timeout);
            setSignInLoading(false);
            if (err && err.errorCode === "user_cancelled") {
                // User closed popup — no error needed
                return;
            }
            if (err && (err.errorCode === "popup_window_error" || err.errorCode === "empty_window_error")) {
                setError("Popup was blocked. Please allow popups for this site.");
                return;
            }
            setError("Sign in failed: " + (err && err.message ? err.message : "Unknown error"));
            console.error("Login error:", err);
        });
    };

    // Not signed in → show landing page
    if (!account) {
        return React.createElement(LandingPage, {
            onSignIn: handleSignIn,
            loading: signInLoading,
            error: error,
        });
    }

    // Signed in → show dashboard
    return React.createElement(App, {
        msalAccount: account,
        msalInstance: msalInstance,
    });
}

ReactDOM.createRoot(document.getElementById("root")).render(
    React.createElement(React.StrictMode, null,
        React.createElement(Root)
    )
);

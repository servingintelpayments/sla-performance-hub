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

    // Load MSAL in background and check for redirect response
    useEffect(function() {
        getMsalInstance()
            .then(function(instance) {
                setMsalRef(instance);
                // Check if user is already signed in (from redirect or session)
                var active = instance.getActiveAccount();
                if (active) {
                    setAccount(active);
                }
            })
            .catch(function(err) {
                console.warn("MSAL background init:", err.message);
            });
    }, []);

    // Sign in handler — uses REDIRECT (not popup) for reliability
    var handleSignIn = function() {
        if (signInLoading) return;
        setSignInLoading(true);
        setError(null);

        function doLogin(instance) {
            // loginRedirect will navigate away from the page to Microsoft login
            // When done, it comes back and handleRedirectPromise in authConfig picks it up
            instance.loginRedirect(loginRequest).catch(function(err) {
                setSignInLoading(false);
                if (err && err.errorCode === "user_cancelled") return;
                setError("Sign in failed: " + (err && err.message ? err.message : "Unknown error"));
                console.error("Login error:", err);
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

    // Always show landing page if not signed in
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

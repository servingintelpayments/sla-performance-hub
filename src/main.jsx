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

    // Load MSAL in background — don't block the page
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
                // Don't show error yet — show it only when user tries to sign in
            });
    }, []);

    // Sign in handler
    var handleSignIn = async function() {
        setSignInLoading(true);
        setError(null);

        // If MSAL isn't ready yet, try loading it now
        var instance = msalInstance;
        if (!instance) {
            try {
                instance = await getMsalInstance();
                setMsalInstance(instance);
            } catch (err) {
                setError("Auth failed to load. Please refresh the page.");
                setSignInLoading(false);
                return;
            }
        }

        try {
            var response = await instance.loginPopup(loginRequest);
            instance.setActiveAccount(response.account);
            setAccount(response.account);
        } catch (err) {
            if (err.errorCode === "popup_window_error" || err.errorCode === "empty_window_error") {
                try {
                    await instance.loginRedirect(loginRequest);
                } catch (redirectErr) {
                    setError("Login failed: " + redirectErr.message);
                }
            } else if (err.errorCode !== "user_cancelled") {
                setError("Sign in failed: " + (err.message || "Unknown error"));
                console.error("Login error:", err);
            }
        } finally {
            setSignInLoading(false);
        }
    };

    // Not signed in → show landing page immediately
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

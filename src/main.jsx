// src/main.jsx
import React, { useState, useEffect } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import LandingPage from "./LandingPage";
import { getMsalInstance, loginRequest } from "./authConfig";

function Root() {
    const [msalInstance, setMsalInstance] = useState(null);
    const [account, setAccount] = useState(null);
    const [loading, setLoading] = useState(true);
    const [signInLoading, setSignInLoading] = useState(false);
    const [error, setError] = useState(null);

    // Initialize MSAL on mount
    useEffect(() => {
        getMsalInstance()
            .then((instance) => {
                setMsalInstance(instance);
                const active = instance.getActiveAccount();
                if (active) {
                    setAccount(active);
                }
                setLoading(false);
            })
            .catch((err) => {
                console.error("MSAL init error:", err);
                setError("Auth failed to load: " + err.message);
                setLoading(false);
            });
    }, []);

    // Sign in handler
    const handleSignIn = async () => {
        if (!msalInstance) {
            setError("Auth not ready. Please refresh.");
            return;
        }
        setSignInLoading(true);
        setError(null);

        try {
            const response = await msalInstance.loginPopup(loginRequest);
            msalInstance.setActiveAccount(response.account);
            setAccount(response.account);
        } catch (err) {
            if (err.errorCode === "popup_window_error" || err.errorCode === "empty_window_error") {
                // Popup blocked (mobile) — fallback to redirect
                try {
                    await msalInstance.loginRedirect(loginRequest);
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

    // Loading state
    if (loading) {
        return (
            <div style={{
                background: "#0d0f14", color: "#8b8fa3", minHeight: "100vh",
                display: "flex", alignItems: "center", justifyContent: "center",
                fontFamily: "'DM Sans', sans-serif", fontSize: 14,
            }}>
                Loading...
            </div>
        );
    }

    // Not signed in → show landing page
    if (!account) {
        return (
            <LandingPage
                onSignIn={handleSignIn}
                loading={signInLoading}
                error={error}
            />
        );
    }

    // Signed in → show dashboard
    return <App msalAccount={account} msalInstance={msalInstance} />;
}

ReactDOM.createRoot(document.getElementById("root")).render(
    <React.StrictMode>
        <Root />
    </React.StrictMode>
);

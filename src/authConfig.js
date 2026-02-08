// src/authConfig.js
// MSAL Configuration for Azure AD Authentication

var CLIENT_ID = "3df6c0f7-b009-47a3-87f2-82172d866bdf";
var TENANT_ID = "1b0086bd-aeda-4c74-a15a-23adfe4d0693";

export var msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + TENANT_ID,
        redirectUri: window.location.origin + window.location.pathname,
        postLogoutRedirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
};

export var loginRequest = {
    scopes: ["User.Read"],
};

// Dynamically load MSAL from CDN
function loadMsalScript() {
    return new Promise(function(resolve, reject) {
        // Already loaded?
        if (typeof window.msal !== "undefined") {
            resolve(window.msal);
            return;
        }

        // Check if script tag already exists
        var existing = document.querySelector('script[src*="msal-browser"]');
        if (existing) {
            // Script tag exists but hasn't finished loading — wait for it
            existing.addEventListener("load", function() { resolve(window.msal); });
            existing.addEventListener("error", function() { reject(new Error("MSAL script failed to load")); });
            // If it already loaded but window.msal isn't set, wait a bit
            setTimeout(function() {
                if (typeof window.msal !== "undefined") resolve(window.msal);
            }, 500);
            return;
        }

        // Inject script dynamically
        var script = document.createElement("script");
        script.src = "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js";
        script.onload = function() {
            if (typeof window.msal !== "undefined") {
                resolve(window.msal);
            } else {
                reject(new Error("MSAL loaded but not available"));
            }
        };
        script.onerror = function() {
            reject(new Error("Failed to load MSAL from CDN"));
        };
        document.head.appendChild(script);
    });
}

// Create singleton MSAL instance
var msalInstance = null;

export async function getMsalInstance() {
    if (msalInstance) return msalInstance;

    var msalLib = await loadMsalScript();
    msalInstance = new msalLib.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    // Handle redirect (mobile fallback)
    var response = await msalInstance.handleRedirectPromise();
    if (response) {
        msalInstance.setActiveAccount(response.account);
    } else {
        var accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            msalInstance.setActiveAccount(accounts[0]);
        }
    }

    return msalInstance;
}

export { msalInstance };

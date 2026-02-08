// src/authConfig.js
// MSAL Configuration for Azure AD Authentication

const CLIENT_ID = "3df6c0f7-b009-47a3-87f2-82172d866bdf";
const TENANT_ID = "1b0086bd-aeda-4c74-a15a-23adfe4d0693";

export const msalConfig = {
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

export const loginRequest = {
    scopes: ["User.Read"],
};

// Create singleton MSAL instance (uses CDN-loaded msal)
let msalInstance = null;

export async function getMsalInstance() {
    if (msalInstance) return msalInstance;

    if (typeof window.msal === "undefined") {
        throw new Error("MSAL library not loaded");
    }

    msalInstance = new window.msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    // Handle redirect (mobile fallback)
    const response = await msalInstance.handleRedirectPromise();
    if (response) {
        msalInstance.setActiveAccount(response.account);
    } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            msalInstance.setActiveAccount(accounts[0]);
        }
    }

    return msalInstance;
}

export { msalInstance };

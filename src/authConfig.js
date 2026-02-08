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

// Wait for MSAL CDN to load (max 10 seconds)
function waitForMsal(timeout = 10000) {
    return new Promise(function(resolve, reject) {
        if (typeof window.msal !== "undefined") {
            resolve(window.msal);
            return;
        }
        var start = Date.now();
        var check = setInterval(function() {
            if (typeof window.msal !== "undefined") {
                clearInterval(check);
                resolve(window.msal);
            } else if (Date.now() - start > timeout) {
                clearInterval(check);
                reject(new Error("MSAL library failed to load. Please refresh the page."));
            }
        }, 100);
    });
}

// Create singleton MSAL instance
var msalInstance = null;

export async function getMsalInstance() {
    if (msalInstance) return msalInstance;

    var msalLib = await waitForMsal();
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

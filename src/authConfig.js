// src/authConfig.js
import { PublicClientApplication } from "@azure/msal-browser";

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

// Singleton MSAL instance using npm package
var msalInstance = null;

export async function getMsalInstance() {
    if (msalInstance) return msalInstance;

    msalInstance = new PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    // Handle redirect response (from loginRedirect)
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

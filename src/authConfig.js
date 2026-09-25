import { LogLevel } from "@azure/msal-browser";

export const msalConfig = {
    auth: {
        clientId: "563f3b76-e7c5-417d-94b2-68ec5bfdc9c0",
        authority: "https://login.microsoftonline.com/6c5abb9a-09e5-470d-9973-e54cf9ad6475",
        redirectUri: "http://localhost:3000",
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) return;
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }
            },
        },
    },
};

export const loginRequest = {
    scopes: ["api://5d67c44c-524d-4958-921f-fe75b912a037/.default"]
};
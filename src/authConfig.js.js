

export const msalConfig = {
    auth: {
        clientId: "TU_CLIENT_ID_FRONTEND", // ID de aplicación (cliente) de MesaTech-React-Frontend
        authority: "https://login.microsoftonline.com/TU_TENANT_ID", // ID de directorio (inquilino)
        redirectUri: "http://localhost:3000",
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const loginRequest = {
    scopes: ["api://5d67c44c-524d-4958-921f-fe75b912a037/Solicitudes.ReadWrite"]
};
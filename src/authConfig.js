export const msalConfig = {
  auth: {
    clientId: "b3e11022-2087-4ef5-a811-2a1cf3fab194",
    authority: "https://login.microsoftonline.com/f1450f14-3d44-4c62-b16d-f6df64edd6c6",
    redirectUri: "https://localhost:3000", // Ensure this matches the URI in Azure portal
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

export const loginRequest = {
  scopes: ["User.Read"],
};
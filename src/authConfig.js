export const msalConfig = {
  auth: {
    clientId: "",
    authority: "",
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
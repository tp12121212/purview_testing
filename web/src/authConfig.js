export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_M365_CLIENT_ID,
    authority: `${import.meta.env.VITE_M365_AUTHORITY_HOST ?? 'https://login.microsoftonline.com'}/common`,
    redirectUri: import.meta.env.VITE_M365_REDIRECT_URI ?? window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: (import.meta.env.VITE_M365_SCOPES ?? '').split(',').map((scope) => scope.trim()).filter(Boolean)
};

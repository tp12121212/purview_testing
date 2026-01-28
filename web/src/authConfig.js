export const authorityHost = (import.meta.env.VITE_M365_AUTHORITY_HOST ?? 'https://login.microsoftonline.com')
  .replace(/\/+$/, '');

export const authorityTenant = (import.meta.env.VITE_M365_AUTHORITY_TENANT ?? 'organizations')
  .replace(/^\/+|\/+$/g, '');

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_M365_CLIENT_ID,
    authority: `${authorityHost}/${authorityTenant || 'organizations'}`,
    redirectUri: import.meta.env.VITE_M365_REDIRECT_URI ?? window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
};

export const clientId = msalConfig.auth.clientId;

const parseScopes = (value) => (value ?? '')
  .split(',')
  .map((scope) => scope.trim())
  .filter(Boolean);

const rawScopes = parseScopes(import.meta.env.VITE_M365_SCOPES);

export const loginRequest = {
  scopes: parseScopes(import.meta.env.VITE_LOGIN_SCOPES ?? 'openid,profile,email')
};

export const exoScope = import.meta.env.VITE_EXO_SCOPE
  ?? rawScopes.find((scope) => scope.includes('outlook.office365.com'))
  ?? 'https://outlook.office365.com/.default';

export const complianceScope = import.meta.env.VITE_COMPLIANCE_SCOPE
  ?? rawScopes.find((scope) => scope.includes('compliance.microsoft.com'))
  ?? 'https://compliance.microsoft.com/.default';

const buildAdminConsentScopes = () => {
  const scopes = new Set();
  const addIfResourceScope = (scope) => {
    if (!scope) {
      return;
    }
    if (scope.includes('://')) {
      scopes.add(scope);
    }
  };

  if (rawScopes.length) {
    rawScopes.forEach(addIfResourceScope);
    return Array.from(scopes);
  }

  addIfResourceScope(exoScope);
  addIfResourceScope(complianceScope);

  return Array.from(scopes);
};

export const adminConsentScopes = buildAdminConsentScopes();

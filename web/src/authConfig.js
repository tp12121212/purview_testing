const trimTrailingSlash = (value) => (value ?? '').replace(/\/+$/, '');

const buildAuthority = (config) => {
  const host = trimTrailingSlash(config.authorityHost);
  const tenant = (config.authorityTenant ?? 'organizations').replace(/^\/+|\/+$/g, '');
  return `${host}/${tenant || 'organizations'}`;
};

const normalizeLoginScopes = (scopes) => {
  const allowed = new Set(['openid', 'profile', 'email', 'offline_access']);
  return (Array.isArray(scopes) ? scopes : [])
    .map((scope) => scope.trim())
    .filter((scope) => allowed.has(scope));
};

export const buildMsalConfig = (config) => ({
  auth: {
    clientId: config.clientId,
    authority: buildAuthority(config),
    redirectUri: config.redirectUri ?? window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
});

export const buildLoginRequest = (config) => {
  const scopes = normalizeLoginScopes(config.loginScopes);
  return {
    scopes: scopes.length ? scopes : ['openid', 'profile', 'email']
  };
};

export const buildGraphAdminConsentUrl = (config, tenantId) => {
  if (!config.clientId) {
    return null;
  }
  const tenantSegment = tenantId || config.authorityTenant || 'organizations';
  const redirectUri = encodeURIComponent(config.redirectUri ?? window.location.origin);
  const scopeParam = encodeURIComponent('https://graph.microsoft.com/.default');
  return `${trimTrailingSlash(config.authorityHost)}/${tenantSegment}/v2.0/adminconsent?client_id=${config.clientId}&scope=${scopeParam}&redirect_uri=${redirectUri}`;
};

export const formatScopeList = (scopes) => {
  if (!Array.isArray(scopes)) {
    return scopes ? String(scopes) : '';
  }
  return scopes.filter(Boolean).join(', ');
};

export const buildScopeList = (values) => Array.from(new Set(values.filter(Boolean)));

export const summarizeScopeSet = (scopeSet) => scopeSet.length ? scopeSet.join(', ') : 'None';

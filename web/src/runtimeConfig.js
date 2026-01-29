const STORAGE_KEY = 'purview_runtime_config_v2';

const DEFAULT_RUNTIME_CONFIG = {
  clientId: '',
  authorityHost: 'https://login.microsoftonline.com',
  authorityTenant: 'organizations',
  redirectUri: '',
  graphBaseUrl: 'https://graph.microsoft.com',
  graphApiVersion: 'v1.0',
  infoProtectionApiVersion: 'beta',
  loginScopes: ['openid', 'profile', 'email'],
  profileScope: 'User.Read',
  mailReadScope: 'Mail.Read',
  filesReadScope: 'Files.Read',
  sitesReadScope: 'Sites.Read.All',
  labelsReadScope: 'InformationProtectionPolicy.Read'
};

const sanitizeHost = (value) => {
  if (!value || typeof value !== 'string') {
    return '';
  }
  return value.trim().replace(/\/+$/, '');
};

const sanitizeTenant = (value) => {
  if (!value || typeof value !== 'string') {
    return '';
  }
  return value.trim().replace(/^\/+|\/+$/g, '');
};

const sanitizeUrl = (value) => {
  if (!value || typeof value !== 'string') {
    return '';
  }
  return value.trim();
};

const sanitizeBaseUrl = (value) => {
  if (!value || typeof value !== 'string') {
    return '';
  }
  return value.trim().replace(/\/+$/, '');
};

const sanitizeVersion = (value, fallback) => {
  if (!value || typeof value !== 'string') {
    return fallback;
  }
  const trimmed = value.trim();
  return trimmed || fallback;
};

const normalizeStringList = (value) => {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.map((item) => String(item).trim()).filter(Boolean);
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return [];
    }
    const separator = trimmed.includes(',') ? ',' : /\s+/;
    return trimmed.split(separator).map((item) => item.trim()).filter(Boolean);
  }
  return [];
};

const normalizeLoginScopes = (value) => {
  const allowed = new Set(['openid', 'profile', 'email', 'offline_access']);
  const scopes = normalizeStringList(value).filter((scope) => allowed.has(scope));
  return scopes.length ? scopes : DEFAULT_RUNTIME_CONFIG.loginScopes;
};

const parseJsonSafely = (value) => {
  if (!value) {
    return null;
  }
  try {
    return JSON.parse(value);
  } catch (error) {
    return null;
  }
};

const normalizeRuntimeConfig = (config) => {
  const authorityHost = sanitizeHost(config.authorityHost) || DEFAULT_RUNTIME_CONFIG.authorityHost;
  const authorityTenant = sanitizeTenant(config.authorityTenant) || DEFAULT_RUNTIME_CONFIG.authorityTenant;
  const redirectUri = sanitizeUrl(config.redirectUri) || window.location.origin;
  const graphBaseUrl = sanitizeBaseUrl(config.graphBaseUrl) || DEFAULT_RUNTIME_CONFIG.graphBaseUrl;

  return {
    clientId: (config.clientId ?? '').trim(),
    authorityHost,
    authorityTenant,
    redirectUri,
    graphBaseUrl,
    graphApiVersion: sanitizeVersion(config.graphApiVersion, DEFAULT_RUNTIME_CONFIG.graphApiVersion),
    infoProtectionApiVersion: sanitizeVersion(config.infoProtectionApiVersion, DEFAULT_RUNTIME_CONFIG.infoProtectionApiVersion),
    loginScopes: normalizeLoginScopes(config.loginScopes),
    profileScope: (config.profileScope ?? DEFAULT_RUNTIME_CONFIG.profileScope).trim(),
    mailReadScope: (config.mailReadScope ?? DEFAULT_RUNTIME_CONFIG.mailReadScope).trim(),
    filesReadScope: (config.filesReadScope ?? DEFAULT_RUNTIME_CONFIG.filesReadScope).trim(),
    sitesReadScope: (config.sitesReadScope ?? DEFAULT_RUNTIME_CONFIG.sitesReadScope).trim(),
    labelsReadScope: (config.labelsReadScope ?? DEFAULT_RUNTIME_CONFIG.labelsReadScope).trim()
  };
};

const readStoredConfig = () => {
  if (typeof window === 'undefined') {
    return null;
  }
  const raw = window.localStorage.getItem(STORAGE_KEY);
  return parseJsonSafely(raw);
};

const saveRuntimeConfig = (config) => {
  if (typeof window === 'undefined') {
    return;
  }
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
};

const resetRuntimeConfig = () => {
  if (typeof window === 'undefined') {
    return;
  }
  window.localStorage.removeItem(STORAGE_KEY);
};

const parseQueryConfig = () => {
  if (typeof window === 'undefined') {
    return {};
  }
  const params = new URLSearchParams(window.location.search);
  const queryConfig = {};

  const assign = (key, value) => {
    if (value === null || value === undefined || value === '') {
      return;
    }
    queryConfig[key] = value;
  };

  assign('clientId', params.get('clientId'));
  assign('authorityHost', params.get('authorityHost'));
  assign('authorityTenant', params.get('authorityTenant'));
  assign('authorityTenant', params.get('tenant'));
  assign('redirectUri', params.get('redirectUri'));
  assign('graphBaseUrl', params.get('graphBaseUrl'));
  assign('graphApiVersion', params.get('graphApiVersion'));
  assign('infoProtectionApiVersion', params.get('infoProtectionApiVersion'));
  assign('loginScopes', params.get('loginScopes'));
  assign('profileScope', params.get('profileScope'));
  assign('mailReadScope', params.get('mailReadScope'));
  assign('filesReadScope', params.get('filesReadScope'));
  assign('sitesReadScope', params.get('sitesReadScope'));
  assign('labelsReadScope', params.get('labelsReadScope'));

  return queryConfig;
};

const loadConfigFile = async () => {
  try {
    const response = await fetch('/runtime-config.json', { cache: 'no-store' });
    if (!response.ok) {
      return { config: null, error: null };
    }
    const data = await response.json();
    return { config: data, error: null };
  } catch (error) {
    return { config: null, error: 'Failed to load runtime-config.json.' };
  }
};

const loadRuntimeConfig = async () => {
  const { config: fileConfig, error: fileError } = await loadConfigFile();
  const storedConfig = readStoredConfig();
  const queryConfig = parseQueryConfig();
  const merged = {
    ...DEFAULT_RUNTIME_CONFIG,
    ...(fileConfig ?? {}),
    ...(storedConfig ?? {}),
    ...queryConfig
  };
  const normalized = normalizeRuntimeConfig(merged);

  if (Object.keys(queryConfig).length > 0) {
    saveRuntimeConfig(normalized);
  }

  return { config: normalized, warnings: fileError ? [fileError] : [] };
};

const buildConfigQueryString = (config) => {
  const params = new URLSearchParams();
  const assign = (key, value) => {
    if (!value) {
      return;
    }
    params.set(key, value);
  };

  assign('clientId', config.clientId);
  assign('authorityHost', config.authorityHost);
  assign('authorityTenant', config.authorityTenant);
  assign('redirectUri', config.redirectUri);
  assign('graphBaseUrl', config.graphBaseUrl);
  assign('graphApiVersion', config.graphApiVersion);
  assign('infoProtectionApiVersion', config.infoProtectionApiVersion);
  if (config.loginScopes?.length) {
    params.set('loginScopes', config.loginScopes.join(','));
  }
  assign('profileScope', config.profileScope);
  assign('mailReadScope', config.mailReadScope);
  assign('filesReadScope', config.filesReadScope);
  assign('sitesReadScope', config.sitesReadScope);
  assign('labelsReadScope', config.labelsReadScope);

  const queryString = params.toString();
  return queryString ? `?${queryString}` : '';
};

export {
  DEFAULT_RUNTIME_CONFIG,
  normalizeRuntimeConfig,
  loadRuntimeConfig,
  saveRuntimeConfig,
  resetRuntimeConfig,
  buildConfigQueryString
};

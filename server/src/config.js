const parseList = (value, fallback = []) => {
  const items = (value ?? '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
  return items.length > 0 ? items : fallback;
};

const normalizeHosts = (hosts) => hosts.map((host) => host.replace(/\/+$/, ''));

const defaultExchangeAudiences = [
  'https://outlook.office365.com',
  'https://outlook.office.com'
];

const defaultComplianceAudiences = [
  'https://compliance.microsoft.com'
];

const defaultAuthorityHosts = [
  'https://login.microsoftonline.com',
  'https://login.microsoftonline.us',
  'https://login.partner.microsoftonline.cn',
  'https://login.microsoftonline.de'
];

const defaultContentTypes = [
  'application/pdf',
  'message/rfc822',
  'application/vnd.ms-outlook',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
];

export const config = {
  port: Number(process.env.PORT ?? 4000),
  allowedTenants: parseList(process.env.M365_ALLOWED_TENANTS ?? ''),
  allowedAuthorityHosts: normalizeHosts(parseList(process.env.M365_ALLOWED_AUTHORITY_HOSTS, defaultAuthorityHosts)),
  exchangeAudiences: parseList(process.env.M365_EXCHANGE_AUDIENCES, defaultExchangeAudiences),
  complianceAudiences: parseList(process.env.M365_COMPLIANCE_AUDIENCES, defaultComplianceAudiences),
  fileUploadLimitMb: Number(process.env.FILE_UPLOAD_LIMIT_MB ?? 25),
  tempDir: process.env.UPLOAD_TEMP_DIR ?? '/tmp/purview_uploads',
  allowedContentTypes: parseList(process.env.ALLOWED_CONTENT_TYPES, defaultContentTypes),
  logLevel: process.env.LOG_LEVEL ?? 'info',
  powershellPath: process.env.PWSH_PATH ?? 'pwsh',
  powershellScriptsDir: process.env.PWSH_SCRIPTS_DIR ?? new URL('../scripts', import.meta.url).pathname
};

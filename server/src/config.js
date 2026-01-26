import dotenv from 'dotenv';

dotenv.config();

const required = (name, fallback = undefined) => {
  const value = process.env[name] ?? fallback;
  if (value === undefined || value === '') {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
};

export const config = {
  port: Number(process.env.PORT ?? 4000),
  clientId: required('M365_CLIENT_ID'),
  authorityHost: process.env.M365_AUTHORITY_HOST ?? 'https://login.microsoftonline.com',
  allowedTenants: (process.env.M365_ALLOWED_TENANTS ?? '').split(',').map((tenant) => tenant.trim()).filter(Boolean),
  apiScopes: (process.env.M365_API_SCOPES ?? '').split(',').map((scope) => scope.trim()).filter(Boolean),
  fileUploadLimitMb: Number(process.env.FILE_UPLOAD_LIMIT_MB ?? 25),
  tempDir: process.env.UPLOAD_TEMP_DIR ?? '/tmp/purview_uploads',
  allowedContentTypes: (process.env.ALLOWED_CONTENT_TYPES ?? 'application/pdf,message/rfc822,application/vnd.ms-outlook,application/vnd.openxmlformats-officedocument.wordprocessingml.document').split(',').map((value) => value.trim()).filter(Boolean),
  logLevel: process.env.LOG_LEVEL ?? 'info',
  powershellPath: process.env.PWSH_PATH ?? 'pwsh',
  powershellScriptsDir: process.env.PWSH_SCRIPTS_DIR ?? new URL('../scripts', import.meta.url).pathname
};

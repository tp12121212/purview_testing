# Purview Extraction & Classification App

This repository now includes a full-stack application for running Microsoft Purview `Test-TextExtraction` and `Test-DataClassification` against uploaded documents. It ships a backend service that hosts PowerShell 7 cmdlets and a responsive frontend for Microsoft 365 sign-in, uploads, and results visualization.

## Repository Layout
- `server/`: Node.js API for running Exchange Online and Purview compliance cmdlets via PowerShell 7.
- `web/`: React frontend with MSAL sign-in and UI for uploads + results.
- `textExctraction.ps1`: Standalone text extraction script.
- `testDataclassification.ps1`: Standalone extraction + classification script (non-interactive SIT input).

## Prerequisites
- PowerShell 7 (`pwsh`).
- Exchange Online PowerShell module (`ExchangeOnlineManagement`).
- Access to Microsoft Purview compliance cmdlets (`Connect-IPPSSession`, `Test-DataClassification`).
- Node.js 18+ for local development.

## Microsoft Entra ID App Registration (MSAL mode)
### Required permissions (delegated)
- **Microsoft Graph**: `openid`, `profile`, `email` (basic sign-in).
- **Office 365 Exchange Online**: `full_access_as_user` or `EWS.AccessAsUser.All` (required for `Connect-ExchangeOnline -AccessToken`).
- **Microsoft 365 compliance**: `Compliance.Read` (minimum) or `Compliance.ReadWrite` (if you need to change compliance data).

> ⚠️ Permission labels can differ slightly by tenant/licensing. If you cannot find the exact names, search by keyword (e.g., “Exchange Online”, “Compliance”, “Purview”, “DLP”). Ensure you add **Delegated** permissions.
> ℹ️ The Microsoft Purview portal is now at `https://purview.microsoft.com`. The OAuth resource used by the compliance cmdlets still targets `https://compliance.microsoft.com`.

### Manual steps (Azure portal)
1. Create a **multi-tenant** app registration.
2. Add **SPA Redirect URI(s)** (e.g., `http://localhost:5173` and your production URL).
3. Add the **delegated permissions** above.
4. Grant **admin consent** in the target tenant.
5. Copy the **Application (client) ID** into `server/.env` and `web/.env`.

> ⚠️ The signed-in user must also be assigned the appropriate **Purview/Compliance role group** permissions in their tenant (for example DLP or Compliance admin roles) to run the compliance cmdlets.

### Automated setup (PowerShell + Microsoft Graph)
The script below creates a **multi-tenant** app registration, adds the delegated permissions, grants admin consent (optional), and updates `server/.env` + `web/.env` with the client ID, auth mode, redirect URI, and scopes.

```powershell
pwsh ./scripts/create-entra-app.ps1 `
  -AppName "Purview Extraction & Classification" `
  -RedirectUris @("http://localhost:5173","https://yourdomain.example") `
  -ExchangeScopeValue "full_access_as_user" `
  -ComplianceScopeValue "Compliance.Read" `
  -GrantAdminConsent `
  -UpdateEnvFiles
```

Prerequisites:
- Install Microsoft Graph PowerShell: `Install-Module Microsoft.Graph -Scope CurrentUser` (or `Install-PSResource Microsoft.Graph` if you use PSResourceGet)
- You must sign in with an account that can create app registrations.
- To auto-grant admin consent, your account must have permission to grant tenant-wide consent (the script requests `DelegatedPermissionGrant.ReadWrite.All`).

Manual fallback steps (if auto-consent fails):
1. Open the v2.0 admin consent URL printed by the script in a browser as a tenant admin (it contains only resource `/.default` scopes).
2. Confirm consent.

If the script reports that it cannot find the Exchange Online delegated scope:
- Re-run with `-ExchangeScopeValue` set to one of the values printed under “Available Exchange Online delegated scopes.”

## Authentication
The app runs **MSAL-only** with browser-based sign-in (authorization code + PKCE). This is required for public, multi-tenant deployments so all authentication happens in the user’s browser session.

## Environment Configuration
Create `.env` files using the examples below (or copy from `server/.env.example.msal` and `web/.env.example.msal`):

### Backend (`server/.env`)
```bash
PORT=4000
AUTH_MODE=msal
M365_CLIENT_ID=your-client-id
M365_AUTHORITY_HOST=https://login.microsoftonline.com
M365_ALLOWED_TENANTS=
M365_API_SCOPES=https://outlook.office365.com/.default,https://compliance.microsoft.com/.default
FILE_UPLOAD_LIMIT_MB=25
ALLOWED_CONTENT_TYPES=application/pdf,message/rfc822,application/vnd.ms-outlook,application/vnd.openxmlformats-officedocument.wordprocessingml.document
UPLOAD_TEMP_DIR=/tmp/purview_uploads
LOG_LEVEL=info
PWSH_PATH=pwsh
PWSH_SCRIPTS_DIR=/app/scripts
```

### Frontend (`web/.env`)
```bash
VITE_M365_CLIENT_ID=your-client-id
VITE_M365_AUTHORITY_HOST=https://login.microsoftonline.com
VITE_M365_AUTHORITY_TENANT=organizations
VITE_M365_REDIRECT_URI=http://localhost:5173
VITE_LOGIN_SCOPES=openid,profile,email
VITE_EXO_SCOPE=https://outlook.office365.com/.default
VITE_COMPLIANCE_SCOPE=https://compliance.microsoft.com/.default
VITE_M365_SCOPES=https://outlook.office365.com/.default,https://compliance.microsoft.com/.default
VITE_API_BASE_URL=http://localhost:4000
```

## Local Development

### Backend
```bash
cd server
npm install
pwsh -NoProfile -File ./scripts/install-modules.ps1
npm start
```

### Frontend
```bash
cd web
npm install
npm run dev
```

Navigate to `http://localhost:5173` to sign in and run extraction/classification.

## Docker Deployment
```bash
docker compose up --build
```

- Backend: `http://localhost:4000`
- Frontend: `http://localhost:5173`

## API Endpoints
- `GET /api/health` — health check.
- `GET /api/sensitive-information-types` — list SIT catalog from Purview.
- `POST /api/extraction` — run `Test-TextExtraction` with file upload.
- `POST /api/classification` — run `Test-DataClassification` with file upload and SIT selection.

## Security Notes
- File uploads are validated by content type and size.
- Access tokens are validated against Entra ID JWKS for multi-tenant sign-in.
- Audit logs are emitted for authentication, extraction, and classification events.

## Standalone Scripts
### Text extraction only
```powershell
./textExctraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf"
```

### Text extraction + data classification
```powershell
./testDataclassification.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.msg" -DataClassification -SensitiveInformationTypes "U.S. Social Security Number"
```

## Notes
- Provide a full file path; `~/` is not expanded by PowerShell in all contexts.
- `testDataclassification.ps1` now accepts SIT names/IDs via `-SensitiveInformationTypes` or `-AllSensitiveInformationTypes`.

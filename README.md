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
1. Create a **multi-tenant** app registration.
2. Add **Redirect URI** for SPA: `http://localhost:5173` (for local dev).
3. Add API permissions:
   - **Microsoft Graph (Delegated)**:
     - `openid`, `profile`, `email` (basic sign-in and profile access).
   - **Office 365 Exchange Online (Delegated)**:
     - Search for **“Office 365 Exchange Online”** in the APIs list.
     - Add **`Exchange.ManageAsApp`** only if you intend to use **app-only** auth (not used in MSAL SPA mode).
     - For delegated user auth (recommended for MSAL SPA), add **`full_access_as_user`** (Exchange Online) so `Connect-ExchangeOnline -AccessToken` can run.
   - **Microsoft Purview compliance**:
     - Search for **“Microsoft Purview”** or **“Microsoft 365 compliance”** (the label can differ by tenant).
     - Add delegated permissions that include **DLP/SIT read** access (e.g., `Compliance.Read`, `Compliance.ReadWrite` if available in your tenant).
4. Grant **admin consent** for the tenant (required for compliance/Exchange delegated scopes).
5. Copy the **Application (client) ID**.

> ⚠️ Permission names can differ by tenant and licensing. If you cannot find the exact permission above, open the API picker and search by keyword (e.g., “Exchange Online”, “Compliance”, “Purview”, “DLP”). Ensure the permissions you add are **Delegated** for the SPA flow.

> ⚠️ The backend expects access tokens issued for the client ID configured in `server/.env` when `AUTH_MODE=msal`.

## Authentication Modes
The app supports two modes:

1. **MSAL (default)**: uses the SPA + API token flow and requires your own Entra ID app registration.
2. **Interactive (device code/user)**: skips MSAL and prompts for user auth on the server host (device code/web login) when PowerShell connects. This mode is intended for local or trusted environments and **does not** require an app registration, but it does require someone to complete interactive sign-in on the server.

## Environment Configuration
Create `.env` files using the examples below (or copy from `server/.env.example` and `web/.env.example`):

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
VITE_AUTH_MODE=msal
VITE_M365_AUTHORITY_HOST=https://login.microsoftonline.com
VITE_M365_REDIRECT_URI=http://localhost:5173
VITE_M365_SCOPES=openid,profile,email,https://outlook.office365.com/.default,https://compliance.microsoft.com/.default
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

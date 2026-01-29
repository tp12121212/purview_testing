# Purview Extraction & Classification (Browser-Only)

This repo now contains a **browser-only** web app that approximates the outcomes of:
- `Test-TextExtraction`
- `Test-DataClassification`

The app runs entirely in the browser using **MSAL (OAuth 2.0 Authorization Code + PKCE)** and **Microsoft Graph**. No secrets, no server-side token exchange, and no runtime `.env` files are required.

## What This App Does
- **Text extraction** from:
  - Local files (PDF, DOCX, TXT, CSV, MD, JSON, EML) using client-side parsing.
  - Outlook messages via Microsoft Graph (Mail.Read).
  - OneDrive recent files via Microsoft Graph (Files.Read).
- **Data classification** using client-side regex detectors (read-only mode).
- **Optional label evaluation** using Microsoft Graph Information Protection (beta) if you provide `sensitiveTypeId` values and have `InformationProtectionPolicy.Read`.

> Support for additional file types (MSG, image scans, PDF scans, ZIP/7z/RAR containers) and `.msg` attachments is planned once the required parsing/OCR packages can be installed in this environment.

## Sensitive Information Types (SIT) Rule Packs
To evaluate extracted text against built-in or custom SITs, export a rule pack XML from your tenant and import it in the app.

Example (PowerShell):
```powershell
$rulePack = Get-DlpSensitiveInformationTypeRulePackage -Identity "Microsoft Rule Package"
[System.IO.File]::WriteAllBytes("./rulepack.xml", $rulePack.SerializedClassificationRuleCollection)
```

> The app parses regex patterns from rule packs. Keyword lists and functions may not be fully represented client-side.
> The parser now understands supporting elements such as regex nodes, keyword lists/dictionaries, and `<Any>` group logic (min/max matches) so the same SITs you load in Purview can be evaluated in the browser. Unsupported supporting elements surface as warnings on import.

## Reality Check / Gaps
Microsoft Graph does **not** provide direct replacements for the Exchange Online PowerShell cmdlets `Test-TextExtraction` or `Test-DataClassification`. This app implements the closest achievable behavior using:
- Local (in-browser) extraction for files.
- Graph for message/file retrieval and label evaluation.

If you need exact cmdlet parity, you must still use PowerShell or a backend service with those cmdlets.

## Prerequisites
- Node.js 18+ for local dev.
- A multi-tenant Entra ID **SPA** app registration.

## Entra ID App Registration (Multi-Tenant SPA)
Create a **multi-tenant** app registration:
1. Add **SPA Redirect URI(s)** (e.g., `http://localhost:5173`).
2. Add **delegated** Microsoft Graph permissions as needed:

### Required Delegated Permissions
| Feature | Permission | Why |
| --- | --- | --- |
| Graph access check | `User.Read` | Used to validate Graph access after consent. |
| Outlook source | `Mail.Read` | Read message body for extraction. |
| OneDrive source | `Files.Read` | Read user files for extraction. |
| SharePoint source (optional) | `Sites.Read.All` | Read SharePoint files via Graph. |
| Label evaluation | `InformationProtectionPolicy.Read` | List labels + evaluate classification results (beta). |

> Login scopes are **not** Graph scopes. Login only uses `openid`, `profile`, `email`.

### Admin Consent
Tenant admins can grant organization-wide consent using the in-app **Admin consent link**. This uses the multi-tenant admin consent flow for Microsoft Graph.

## Runtime Configuration (No .env)
The app loads configuration in this order:
1. `web/public/runtime-config.json`
2. Browser `localStorage` (saved from Runtime Settings)
3. URL parameters

### `web/public/runtime-config.json`
```json
{
  "clientId": "YOUR_CLIENT_ID",
  "authorityHost": "https://login.microsoftonline.com",
  "authorityTenant": "organizations",
  "redirectUri": "http://localhost:5173",
  "graphBaseUrl": "https://graph.microsoft.com",
  "graphApiVersion": "v1.0",
  "infoProtectionApiVersion": "beta",
  "loginScopes": ["openid", "profile", "email"],
  "profileScope": "User.Read",
  "mailReadScope": "Mail.Read",
  "filesReadScope": "Files.Read",
  "sitesReadScope": "Sites.Read.All",
  "labelsReadScope": "InformationProtectionPolicy.Read"
}
```

### URL Parameter Overrides
```
?clientId=...&tenant=organizations&graphBaseUrl=...&graphApiVersion=v1.0&infoProtectionApiVersion=beta&loginScopes=openid,profile,email
```

## Local Development
```bash
cd web
npm install
npm run dev
```

Navigate to `http://localhost:5173`, open **Runtime Settings**, and enter your client ID + scopes.

## Docker Deployment
```bash
docker compose up --build
```

- Frontend: `http://localhost:5173`

## Smoke Test
```bash
node scripts/smoke-test.mjs
```

## Troubleshooting
| Error | Likely Cause | Fix |
| --- | --- | --- |
| `AADSTS70011` scope invalid | Graph scopes placed in **Login Scopes** | Use only `openid,profile,email` for Login Scopes. |
| `AADSTS65001` admin consent required | Tenant admin approval missing | Use Admin consent link or ask tenant admin. |
| 401/403 from Graph | Missing delegated permission | Grant the exact scope for the feature. |
| 403 on `/security/*` endpoints | Missing Security Reader/Admin role | Assign Security Reader or Security Administrator role in Entra ID. |
| 429 Graph throttling | Too many requests | Retry after the `Retry-After` delay (handled automatically). |

## Legacy Server
The `server/` folder remains in the repo for reference but is **not used** by the current browser-only flows.

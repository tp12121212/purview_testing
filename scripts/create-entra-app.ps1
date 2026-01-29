[CmdletBinding()]
param(
    [string]$AppName = "Purview Extraction & Classification",
    [string[]]$RedirectUris = @("http://localhost:5173"),
    [string]$ExchangeScopeValue = "EWS.AccessAsUser.All",
    [string]$ComplianceScopeValue = "Compliance.Read",
    [switch]$IncludeComplianceReadWrite,
    [switch]$GrantAdminConsent,
    [switch]$UpdateEnvFiles,
    [switch]$UpdateRuntimeConfig,
    [string]$ServerEnvPath = "../server/.env",
    [string]$WebEnvPath = "../web/.env",
    [string]$RuntimeConfigPath = "../web/public/runtime-config.json"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft.Graph module not found. Install it with: Install-PSResource -Name Microsoft.Graph -Scope CurrentUser"
    exit 1
}

Import-Module Microsoft.Graph -ErrorAction Stop
if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) {
    Select-MgProfile -Name "v1.0" | Out-Null
}

$graphScopes = @(
    "Application.ReadWrite.All",
    "Directory.ReadWrite.All"
)
if ($GrantAdminConsent) {
    $graphScopes += "DelegatedPermissionGrant.ReadWrite.All"
}

Connect-MgGraph -Scopes $graphScopes | Out-Null

function Set-EnvValue {
    param(
        [string]$Path,
        [string]$Key,
        [string]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Warning "Env file not found: $Path"
        return
    }

    $content = Get-Content -LiteralPath $Path -Raw
    if ($content -match "^(?m)$Key=") {
        $content = [regex]::Replace($content, "^(?m)$Key=.*$", "$Key=$Value")
    }
    else {
        if (-not $content.EndsWith("`n")) {
            $content += "`n"
        }
        $content += "$Key=$Value`n"
    }
    Set-Content -LiteralPath $Path -Value $content
}

function Write-RuntimeConfig {
    param(
        [string]$Path,
        [string]$ClientId,
        [string]$RedirectUri,
        [string]$ExchangeScopeValue,
        [string]$ComplianceScopeValue
    )

    $exchangeScope = $ExchangeScopeValue
    if ($ExchangeScopeValue -notmatch '://') {
        $exchangeScope = "https://outlook.office365.com/$ExchangeScopeValue"
    }

    $complianceScope = $ComplianceScopeValue
    if ($ComplianceScopeValue -notmatch '://') {
        $complianceScope = "https://compliance.microsoft.com/$ComplianceScopeValue"
    }

    $config = [ordered]@{
        clientId = $ClientId
        authorityHost = "https://login.microsoftonline.com"
        authorityTenant = "organizations"
        redirectUri = $RedirectUri
        loginScopes = @("openid", "profile", "email")
        exchangeScope = $exchangeScope
        complianceScope = $complianceScope
        apiBaseUrl = "http://localhost:4000"
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Path
}

$graphAppId = "00000003-0000-0000-c000-000000000000"
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property "id,appId,displayName,oauth2PermissionScopes"
if (-not $graphSp) {
    Write-Error "Could not find the Microsoft Graph service principal."
    exit 1
}

$graphScopeValues = @("openid", "profile", "email")
$graphAccess = @()
foreach ($value in $graphScopeValues) {
    $scope = $graphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $value -and $_.IsEnabled } | Select-Object -First 1
    if ($scope) {
        $graphAccess += @{
            Id = $scope.Id
            Type = "Scope"
        }
    }
}

$exoAppId = "00000002-0000-0ff1-ce00-000000000000"
$exoSp = Get-MgServicePrincipal -Filter "appId eq '$exoAppId'" -Property "id,appId,displayName,oauth2PermissionScopes"
if (-not $exoSp) {
    Write-Error "Could not find the Exchange Online service principal (appId $exoAppId) in this tenant."
    exit 1
}

$exoScope = $exoSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $ExchangeScopeValue -and $_.IsEnabled } | Select-Object -First 1
if (-not $exoScope) {
    $availableScopes = @()
    if ($exoSp.Oauth2PermissionScopes) {
        $availableScopes = $exoSp.Oauth2PermissionScopes | Where-Object { $_.IsEnabled } | Select-Object -ExpandProperty Value
    }
    Write-Error "Could not find delegated scope '$ExchangeScopeValue' on Office 365 Exchange Online."
    if ($availableScopes.Count -gt 0) {
        Write-Host "Available Exchange Online delegated scopes:"
        $availableScopes | Sort-Object | ForEach-Object { Write-Host " - $_" }
        Write-Host "Re-run with: -ExchangeScopeValue '<value from list>'"
    }
    exit 1
}

$complianceAppId = "80ccca67-54bd-44ab-8625-4b79c4dc7775"
$complianceSp = Get-MgServicePrincipal -Filter "appId eq '$complianceAppId'" -Property "id,appId,displayName,oauth2PermissionScopes"
if (-not $complianceSp) {
    $complianceSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft 365 compliance'" -Property "id,appId,displayName,oauth2PermissionScopes"
}
if (-not $complianceSp) {
    $candidates = Get-MgServicePrincipal -Filter "startswith(displayName,'Microsoft 365')" -Property "id,appId,displayName,oauth2PermissionScopes" -All
    $complianceSp = $candidates | Where-Object { $_.Oauth2PermissionScopes.Value -contains $ComplianceScopeValue } | Select-Object -First 1
}

if (-not $complianceSp) {
    Write-Error "Could not locate the Microsoft 365 compliance service principal with scope '$ComplianceScopeValue'."
    Write-Host "Search your tenant for the compliance resource and choose an available scope."
    exit 1
}

$complianceAccess = @()
$scopeValues = @($ComplianceScopeValue)
if ($IncludeComplianceReadWrite -and $ComplianceScopeValue -ne "Compliance.ReadWrite") {
    $scopeValues += "Compliance.ReadWrite"
}

foreach ($value in $scopeValues) {
    $scope = $complianceSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $value -and $_.IsEnabled } | Select-Object -First 1
    if (-not $scope) {
        Write-Warning "Scope '$value' not found on '$($complianceSp.DisplayName)'."
        continue
    }
    $complianceAccess += @{
        Id = $scope.Id
        Type = "Scope"
    }
}

if ($complianceAccess.Count -eq 0) {
    Write-Error "No compliance scopes were added. Available scopes on '$($complianceSp.DisplayName)':"
    if ($complianceSp.Oauth2PermissionScopes) {
        $complianceSp.Oauth2PermissionScopes | Select-Object Value, Id, IsEnabled | Format-Table | Out-Host
    }
    else {
        Write-Host "No delegated scopes were found on this tenant's compliance service principal."
        Write-Host "Ensure Microsoft Purview/Compliance is provisioned and licensed for the tenant, then retry."
        Write-Host "You can also sign in to the Microsoft Purview portal (https://purview.microsoft.com) as a tenant admin to trigger provisioning."
    }
    exit 1
}

$requiredResourceAccess = @()
if ($graphAccess.Count -gt 0) {
    $requiredResourceAccess += @{
        ResourceAppId = $graphSp.AppId
        ResourceAccess = $graphAccess
    }
}
$requiredResourceAccess += @(
    @{
        ResourceAppId = $exoSp.AppId
        ResourceAccess = @(
            @{
                Id = $exoScope.Id
                Type = "Scope"
            }
        )
    },
    @{
        ResourceAppId = $complianceSp.AppId
        ResourceAccess = $complianceAccess
    }
)

$app = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMultipleOrgs" -Spa @{ RedirectUris = $RedirectUris } -RequiredResourceAccess $requiredResourceAccess
New-MgServicePrincipal -AppId $app.AppId | Out-Null

$tenantId = (Get-MgContext).TenantId
$redirectUri = $RedirectUris[0]
$consentScopes = @("https://outlook.office365.com/.default")
if ($complianceAccess.Count -gt 0) {
    $consentScopes += "https://compliance.microsoft.com/.default"
}
$consentScopeString = ($consentScopes | Where-Object { $_ } | Select-Object -Unique) -join " "
$consentScopeEncoded = [uri]::EscapeDataString($consentScopeString)
$consentUrl = "https://login.microsoftonline.com/$tenantId/v2.0/adminconsent?client_id=$($app.AppId)&scope=$consentScopeEncoded&redirect_uri=$([uri]::EscapeDataString($redirectUri))"

if ($GrantAdminConsent) {
    $clientSp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -Property "id,appId,displayName"
    if ($clientSp) {
        $grants = @(
            @{
                Resource = $exoSp
                Scopes = @($ExchangeScopeValue)
            },
            @{
                Resource = $complianceSp
                Scopes = $scopeValues
            }
        )

        foreach ($grant in $grants) {
            $resourceId = $grant.Resource.Id
            $scopeString = ($grant.Scopes | Where-Object { $_ } | Select-Object -Unique) -join " "
            if (-not $scopeString) {
                continue
            }
            $existing = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($clientSp.Id)' and resourceId eq '$resourceId' and consentType eq 'AllPrincipals'" -All | Select-Object -First 1
            if ($existing) {
                $existingScopes = @()
                if ($existing.Scope) {
                    $existingScopes = $existing.Scope -split " "
                }
                $merged = ($existingScopes + ($scopeString -split " ")) | Where-Object { $_ } | Select-Object -Unique
                $mergedString = $merged -join " "
                if ($mergedString -ne $existing.Scope) {
                    Update-MgOauth2PermissionGrant -Oauth2PermissionGrantId $existing.Id -Scope $mergedString | Out-Null
                }
            }
            else {
                New-MgOauth2PermissionGrant -ClientId $clientSp.Id -ConsentType "AllPrincipals" -ResourceId $resourceId -Scope $scopeString | Out-Null
            }
        }
    }
    else {
        Write-Warning "Could not locate service principal for the new app. Admin consent not granted."
    }
}

if ($UpdateEnvFiles) {
    Set-EnvValue -Path $ServerEnvPath -Key "AUTH_MODE" -Value "msal"
    Set-EnvValue -Path $ServerEnvPath -Key "M365_CLIENT_ID" -Value $app.AppId
    Set-EnvValue -Path $ServerEnvPath -Key "M365_AUTHORITY_HOST" -Value "https://login.microsoftonline.com"
    Set-EnvValue -Path $ServerEnvPath -Key "M365_API_SCOPES" -Value "https://outlook.office365.com/.default,https://compliance.microsoft.com/.default"
    Set-EnvValue -Path $ServerEnvPath -Key "M365_ALLOWED_TENANTS" -Value ""

    Set-EnvValue -Path $WebEnvPath -Key "VITE_M365_CLIENT_ID" -Value $app.AppId
    Set-EnvValue -Path $WebEnvPath -Key "VITE_M365_AUTHORITY_HOST" -Value "https://login.microsoftonline.com"
    Set-EnvValue -Path $WebEnvPath -Key "VITE_M365_AUTHORITY_TENANT" -Value "organizations"
    Set-EnvValue -Path $WebEnvPath -Key "VITE_M365_REDIRECT_URI" -Value $redirectUri
    Set-EnvValue -Path $WebEnvPath -Key "VITE_LOGIN_SCOPES" -Value "openid,profile,email"
    Set-EnvValue -Path $WebEnvPath -Key "VITE_EXO_SCOPE" -Value "https://outlook.office365.com/$ExchangeScopeValue"
    Set-EnvValue -Path $WebEnvPath -Key "VITE_COMPLIANCE_SCOPE" -Value "https://compliance.microsoft.com/$ComplianceScopeValue"
    Set-EnvValue -Path $WebEnvPath -Key "VITE_M365_SCOPES" -Value "https://outlook.office365.com/.default,https://compliance.microsoft.com/.default"
}

if ($UpdateRuntimeConfig) {
    Write-RuntimeConfig -Path $RuntimeConfigPath -ClientId $app.AppId -RedirectUri $redirectUri -ExchangeScopeValue $ExchangeScopeValue -ComplianceScopeValue $ComplianceScopeValue
}

Write-Host ""
Write-Host "App registration created."
Write-Host "App (client) ID: $($app.AppId)"
Write-Host "Tenant ID: $tenantId"
Write-Host "Admin consent URL (run in a browser as a tenant admin if consent was not granted automatically):"
Write-Host $consentUrl
if ($UpdateRuntimeConfig) {
    Write-Host "Runtime config updated: $RuntimeConfigPath"
}

Disconnect-MgGraph | Out-Null

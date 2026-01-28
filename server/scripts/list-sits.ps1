[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$AccessToken
)

try {
    $ippsCommand = Get-Command Connect-IPPSSession -ErrorAction Stop
    $deviceParam = $null
    foreach ($candidate in @("UseDeviceAuthentication", "Device")) {
        if ($ippsCommand.Parameters.ContainsKey($candidate)) {
            $deviceParam = $candidate
            break
        }
    }

    if ($AccessToken) {
        Connect-IPPSSession -AccessToken $AccessToken -ShowBanner:$false -ErrorAction Stop
    }
    elseif ($UserPrincipalName) {
        Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
    }
    elseif ($deviceParam) {
        $deviceParams = @{
            ShowBanner = $false
            ErrorAction = "Stop"
        }
        $deviceParams[$deviceParam] = $true
        Connect-IPPSSession @deviceParams
    }
    else {
        throw "Either AccessToken or UserPrincipalName must be provided (device authentication is not supported by this module version)."
    }

    $sitResults = @(Get-DlpSensitiveInformationType -ErrorAction Stop)
    $sitCatalog = @()
    foreach ($sit in $sitResults) {
        if (-not $sit) {
            continue
        }

        $displayName = $null
        foreach ($prop in @("Name", "DisplayName", "Identity", "Id")) {
            if ($sit.PSObject.Properties.Name -contains $prop) {
                $value = $sit.$prop
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $displayName = $value
                    break
                }
            }
        }

        $idValue = $null
        foreach ($prop in @("Id", "Identity")) {
            if ($sit.PSObject.Properties.Name -contains $prop) {
                $value = $sit.$prop
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $idValue = $value
                    break
                }
            }
        }

        if ($displayName -and $idValue) {
            $sitCatalog += [pscustomobject]@{
                Display = $displayName
                Id = $idValue
            }
        }
    }

    $sitCatalog | Sort-Object Display | ConvertTo-Json -Depth 4
}
catch {
    Write-Error $_
    exit 1
}
finally {
    if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
        Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue
    }
}

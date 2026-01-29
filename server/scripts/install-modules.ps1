[CmdletBinding()]
param(
    [string[]]$Modules = @(
        "ExchangeOnlineManagement",
        "Microsoft.PowerShell.SecretManagement",
        "Microsoft.PowerShell.SecretStore"
    )
)

$ErrorActionPreference = "Stop"

function Ensure-PSResourceGet {
    if (-not (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue)) {
        $module = Get-Module -ListAvailable -Name Microsoft.PowerShell.PSResourceGet | Select-Object -First 1
        if ($module) {
            Import-Module Microsoft.PowerShell.PSResourceGet -ErrorAction SilentlyContinue | Out-Null
        }
    }

    if (-not (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue)) {
        Write-Error "PowerShellGet v3 (Microsoft.PowerShell.PSResourceGet) is required. Install it and re-run this script."
        throw
    }
}

function Ensure-Tls12 {
    try {
        $current = [Net.ServicePointManager]::SecurityProtocol
        if (($current -band [Net.SecurityProtocolType]::Tls12) -eq 0) {
            [Net.ServicePointManager]::SecurityProtocol = $current -bor [Net.SecurityProtocolType]::Tls12
        }
    }
    catch {
        # Ignore if not supported.
    }
}

function Ensure-PSGallery {
    $repo = Get-PSResourceRepository -Name "PSGallery" -ErrorAction SilentlyContinue
    if (-not $repo) {
        Write-Host "Registering PSGallery repository..."
        Register-PSResourceRepository -PSGallery | Out-Null
    }
    Set-PSResourceRepository -Name "PSGallery" -Trusted
}

Ensure-Tls12
Ensure-PSResourceGet
Ensure-PSGallery

foreach ($module in $Modules) {
    if (-not (Get-InstalledPSResource -Name $module -ErrorAction SilentlyContinue)) {
        Install-PSResource -Name $module -Scope CurrentUser -TrustRepository
    }
}

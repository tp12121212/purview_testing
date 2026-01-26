[CmdletBinding()]
param(
    [string[]]$Modules = @(
        "ExchangeOnlineManagement",
        "Microsoft.PowerShell.SecretManagement",
        "Microsoft.PowerShell.SecretStore"
    )
)

Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

foreach ($module in $Modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -Scope AllUsers
    }
}

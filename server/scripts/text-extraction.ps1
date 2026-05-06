[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$FilePath,

    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$AccessToken
)

try {
    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "File not found: $FilePath"
    }

    $exoCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop
    $deviceParam = $null
    foreach ($candidate in @("UseDeviceAuthentication", "Device")) {
        if ($exoCommand.Parameters.ContainsKey($candidate)) {
            $deviceParam = $candidate
            break
        }
    }

    if ($AccessToken) {
        Connect-ExchangeOnline -AccessToken $AccessToken -ShowBanner:$false -ErrorAction Stop
    }
    elseif ($UserPrincipalName) {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
    }
    elseif ($deviceParam) {
        $deviceParams = @{
            ShowBanner = $false
            ErrorAction = "Stop"
        }
        $deviceParams[$deviceParam] = $true
        Connect-ExchangeOnline @deviceParams
    }
    else {
        throw "Either AccessToken or UserPrincipalName must be provided (device authentication is not supported by this module version)."
    }

    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $extractionResult = Test-TextExtraction -FileData $fileBytes -ErrorAction Stop

    $extractionResult.ExtractedResults | ConvertTo-Json -Depth 9
}
catch {
    Write-Error $_
    exit 1
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

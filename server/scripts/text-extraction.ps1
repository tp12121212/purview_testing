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

    if ($AccessToken) {
        Connect-ExchangeOnline -AccessToken $AccessToken -ShowBanner:$false -ErrorAction Stop
    }
    elseif ($UserPrincipalName) {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
    }
    else {
        throw "Either AccessToken or UserPrincipalName must be provided."
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

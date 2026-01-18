# Works on macOS and Windows

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$WinFile,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$MacFile
)

Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop

try {
    # Detect OS and pick the right path
    if ($IsWindows) {
        $FilePath = $WinFile
    }
    elseif ($IsMacOS) {
        $FilePath = $MacFile
    }
    else {
        throw "Unsupported OS. This script currently supports Windows and macOS only."
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "File not found: $FilePath"
    }

    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $extractionResult = Test-TextExtraction -FileData $fileBytes

    $extractionResult.ExtractedResults | ConvertTo-Json -Depth 9
}
catch {
    Write-Error $_
}
finally {
    # Optional: clean disconnect (comment out if you want to keep the session)
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

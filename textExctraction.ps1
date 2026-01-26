<#
.SYNOPSIS
Connects to Exchange Online and runs Test-TextExtraction against a file.

.DESCRIPTION
Authenticates to Exchange Online using a User Principal Name (UPN), then reads a local file
(PDF, etc.) and submits its bytes to Test-TextExtraction.

You can provide either:
- -WinFile (typically when running on Windows), OR
- -MacFile (typically when running on macOS)

Only ONE file path is required. If you provide both, the script prefers the one that matches
the current OS.

.PARAMETER UserPrincipalName
The Exchange Online sign-in identity (UPN), e.g. admin@contoso.com, used by Connect-ExchangeOnline.

.PARAMETER WinFile
Optional. Full path to the file on Windows (e.g. C:\Temp\document.pdf).
Required only when running on Windows IF -MacFile is not provided.

.PARAMETER MacFile
Optional. Full path to the file on macOS (e.g. /Users/user/temp/document.pdf), do not use ~/Temp/doc.pdf.
Required only when running on macOS IF -WinFile is not provided.

.EXAMPLE
# Windows (provide WinFile only)
.\Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf"

.EXAMPLE
# macOS (provide MacFile only)
pwsh ./Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -MacFile "$HOME/temp/document.pdf"

.EXAMPLE
# Either OS (provide both; script uses the OS-appropriate one)
.\Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf" -MacFile "$HOME/temp/document.pdf"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "UPN used to authenticate to Exchange Online (e.g. admin@contoso.com).")]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false, HelpMessage = "Optional Windows file path (e.g. C:\Temp\document.pdf). Required on Windows if MacFile is not provided.")]
    [ValidateNotNullOrEmpty()]
    [string]$WinFile,

    [Parameter(Mandatory = $false, HelpMessage = "Optional macOS file path (e.g. ~/temp/document.pdf). Required on macOS if WinFile is not provided.")]
    [ValidateNotNullOrEmpty()]
    [string]$MacFile
)

Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop

try {
    # Pick the path that matches the current OS (fallback to the other if provided).
    if ($IsWindows) {
        if (-not [string]::IsNullOrWhiteSpace($WinFile)) {
            $FilePath = $WinFile
        }
        elseif (-not [string]::IsNullOrWhiteSpace($MacFile)) {
            $FilePath = $MacFile
        }
        else {
            throw "On Windows you must provide -WinFile (or provide -MacFile as an override)."
        }
    }
    elseif ($IsMacOS) {
        if (-not [string]::IsNullOrWhiteSpace($MacFile)) {
            $FilePath = $MacFile
        }
        elseif (-not [string]::IsNullOrWhiteSpace($WinFile)) {
            $FilePath = $WinFile
        }
        else {
            throw "On macOS you must provide -MacFile (or provide -WinFile as an override)."
        }
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
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
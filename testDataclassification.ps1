<#
.SYNOPSIS
Connects to Exchange Online and runs Test-TextExtraction against a file.

.DESCRIPTION
Authenticates to Exchange Online using a User Principal Name (UPN), then reads a local file
(PDF, etc.) and submits its bytes to Test-TextExtraction.

Optionally, connects to the Microsoft Purview compliance session (IPPS) and runs
Test-DataClassification against the extracted text streams.

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

.PARAMETER DataClassification
Optional. If set, connects via Connect-IPPSSession and runs Test-DataClassification against
each ExtractedStreamText returned by Test-TextExtraction.

.EXAMPLE
# Windows (provide WinFile only)
.\Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf"

.EXAMPLE
# macOS (provide MacFile only)
pwsh ./Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -MacFile "$HOME/temp/document.pdf"

.EXAMPLE
# Either OS (provide both; script uses the OS-appropriate one)
.\Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf" -MacFile "$HOME/temp/document.pdf"

.EXAMPLE
# Run extraction + data classification on extracted text streams
.\Test-Extraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.msg" -DataClassification

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
    [string]$MacFile,

    [Parameter(Mandatory = $false, HelpMessage = "Run Test-DataClassification against extracted text via Connect-IPPSSession.")]
    [switch]$DataClassification
)

try {
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$true -ErrorAction Stop

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
    $extractionResult = Test-TextExtraction -FileData $fileBytes -ErrorAction Stop
    if (-not $extractionResult) {
        throw "Test-TextExtraction returned no result."
    }

    $extractedResults = $extractionResult.ExtractedResults
    if (-not $extractedResults) {
        throw "Test-TextExtraction returned no ExtractedResults."
    }

    if (-not $DataClassification) {
        $extractedResults | ConvertTo-Json -Depth 9
        return
    }

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    $selectedSensitiveTypes = $null
    Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$true -ErrorAction Stop

    $testDataClassificationCommand = Get-Command Test-DataClassification -ErrorAction Stop
    $textParamName = $null
    foreach ($candidate in @("TextToClassify", "Text", "Content", "ContentText", "TextContent")) {
        if ($testDataClassificationCommand.Parameters.ContainsKey($candidate)) {
            $textParamName = $candidate
            break
        }
    }

    $fileDataParamName = $null
    foreach ($candidate in @("FileData", "FileBytes")) {
        if ($testDataClassificationCommand.Parameters.ContainsKey($candidate)) {
            $fileDataParamName = $candidate
            break
        }
    }

    $fileNameParamName = $null
    foreach ($candidate in @("FileName", "Name")) {
        if ($testDataClassificationCommand.Parameters.ContainsKey($candidate)) {
            $fileNameParamName = $candidate
            break
        }
    }

    $testTextExtractionParamName = if ($testDataClassificationCommand.Parameters.ContainsKey("TestTextExtractionResults")) {
        "TestTextExtractionResults"
    }
    else {
        $null
    }

    if (-not $textParamName -and -not $fileDataParamName -and -not $testTextExtractionParamName) {
        $availableParams = ($testDataClassificationCommand.Parameters.Keys | Sort-Object) -join ", "
        throw "Test-DataClassification does not expose a supported text, file data, or TestTextExtractionResults parameter. Available parameters: $availableParams"
    }

    $scopeParamName = if ($testDataClassificationCommand.Parameters.ContainsKey("ClassificationNames")) {
        "ClassificationNames"
    }
    elseif ($testDataClassificationCommand.Parameters.ContainsKey("SensitiveType")) {
        "SensitiveType"
    }
    elseif ($testDataClassificationCommand.Parameters.ContainsKey("SensitiveInformationType")) {
        "SensitiveInformationType"
    }
    elseif ($testDataClassificationCommand.Parameters.ContainsKey("SensitiveInfoType")) {
        "SensitiveInfoType"
    }
    else {
        $null
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

    $sitCatalog = $sitCatalog | Sort-Object Display
    if ($sitCatalog.Count -gt 0) {
        $runAllChoice = Read-Host "Run Test-DataClassification against all Sensitive Information Types? (Y/n)"
        if ($runAllChoice -match '^(n|no)$') {
            Write-Host "Available Sensitive Information Types:"
            for ($i = 0; $i -lt $sitCatalog.Count; $i++) {
                $index = $i + 1
                Write-Host ("[{0}] {1}" -f $index, $sitCatalog[$i].Display)
            }

            while ($true) {
                $selectionInput = Read-Host "Select one or more SIT numbers separated by commas (e.g. 1 or 1,5,7)"
                if ([string]::IsNullOrWhiteSpace($selectionInput)) {
                    Write-Warning "No selections provided. Defaulting to all SITs."
                    $selectedSensitiveTypes = $null
                    break
                }

                $selectedIndexes = $selectionInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ } | Select-Object -Unique
                $invalidIndexes = $selectedIndexes | Where-Object { $_ -lt 1 -or $_ -gt $sitCatalog.Count }
                if ($invalidIndexes.Count -gt 0 -or $selectedIndexes.Count -eq 0) {
                    Write-Warning "Selection contains invalid numbers. Please try again."
                    continue
                }

                if ($scopeParamName -eq "ClassificationNames") {
                    $selectedSensitiveTypes = $selectedIndexes | ForEach-Object { $sitCatalog[$_ - 1].Id } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                }
                else {
                    $selectedSensitiveTypes = $selectedIndexes | ForEach-Object { $sitCatalog[$_ - 1].Display } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                }
                if (-not $selectedSensitiveTypes -or $selectedSensitiveTypes.Count -eq 0) {
                    Write-Warning "Selections did not resolve to valid SITs. Defaulting to all SITs."
                    $selectedSensitiveTypes = $null
                }
                break
            }
        }
    }
    else {
        Write-Warning "No Sensitive Information Types returned by Get-DlpSensitiveInformationType. Defaulting to all SITs."
    }

    $sourceFileName = if ($FilePath) { [System.IO.Path]::GetFileName($FilePath) } else { $null }
    $sourceExtension = if ($FilePath) { [System.IO.Path]::GetExtension($FilePath) } else { $null }
    $isEmailSource = $false
    if ($sourceExtension) {
        $isEmailSource = @(".msg", ".eml") -contains $sourceExtension.ToLowerInvariant()
    }

    $streamedText = @()
    $streamIndex = 0
    foreach ($item in $extractedResults) {
        $streamTexts = @()
        if ($item.PSObject.Properties.Name -contains "ExtractedStreamText") {
            $streamTexts = @($item.ExtractedStreamText)
        }
        elseif ($item.PSObject.Properties.Name -contains "ExtractedText") {
            $streamTexts = @($item.ExtractedText)
        }
        else {
            $fallbackTextProps = $item.PSObject.Properties | Where-Object { $_.Name -match 'Extracted.*Text' }
            foreach ($prop in $fallbackTextProps) {
                if ($prop.Value) {
                    $streamTexts += @($prop.Value)
                }
            }
        }

        if (-not $streamTexts -or $streamTexts.Count -eq 0) {
            continue
        }

        $streamKind = "Unknown"
        $streamName = $null

        foreach ($prop in @("AttachmentName", "AttachmentFileName", "Attachment", "AttachmentFile")) {
            if ($item.PSObject.Properties.Name -contains $prop) {
                $value = $item.$prop
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $streamKind = "Attachment"
                    $streamName = $value
                    break
                }
            }
        }

        if ($streamKind -eq "Unknown") {
            if ($item.PSObject.Properties.Name -contains "IsAttachment" -and $item.IsAttachment) {
                $streamKind = "Attachment"
            }
        }

        if (-not $streamName) {
            foreach ($prop in @("StreamName", "ItemName", "FileName", "Name")) {
                if ($item.PSObject.Properties.Name -contains $prop) {
                    $value = $item.$prop
                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        $streamName = $value
                        break
                    }
                }
            }
        }

        if ($streamKind -eq "Unknown" -and $streamName) {
            if ($streamName -match '(?i)attachment') {
                $streamKind = "Attachment"
            }
            elseif ($streamName -match '(?i)body|message') {
                $streamKind = "Body"
            }
        }

        if ($streamKind -eq "Unknown") {
            if ($isEmailSource) {
                $streamKind = "Body"
                if (-not $streamName) {
                    $streamName = "Body"
                }
            }
            else {
                $streamKind = "Document"
                if (-not $streamName) {
                    $streamName = $sourceFileName
                }
            }
        }

        if (-not $streamName) {
            $streamName = "Stream"
        }

        foreach ($text in $streamTexts) {
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            $streamIndex++
            $streamedText += [pscustomobject]@{
                StreamIndex = $streamIndex
                Kind = $streamKind
                Name = $streamName
                SourceFile = $sourceFileName
                Text = $text
            }
        }
    }

    if ($textParamName -or $fileDataParamName) {
        $dataClassificationResults = @()
        foreach ($stream in $streamedText) {
            $dcParams = @{
                ErrorAction = "Stop"
            }
            if ($textParamName) {
                $dcParams[$textParamName] = $stream.Text
            }
            else {
                $dcParams[$fileDataParamName] = [System.Text.Encoding]::UTF8.GetBytes($stream.Text)
                if ($fileNameParamName) {
                    $dcParams[$fileNameParamName] = "extracted_stream_$($stream.StreamIndex).txt"
                }
            }

            if ($selectedSensitiveTypes -and $scopeParamName) {
                $dcParams[$scopeParamName] = $selectedSensitiveTypes
            }
            elseif ($selectedSensitiveTypes -and -not $scopeParamName) {
                Write-Warning "Selected SITs were provided but Test-DataClassification does not expose a classification/sensitive type parameter. Scope will be ignored."
            }

            $dcResult = Test-DataClassification @dcParams
            $dataClassificationResults += [pscustomobject]@{
                StreamIndex = $stream.StreamIndex
                Kind = $stream.Kind
                Name = $stream.Name
                SourceFile = $stream.SourceFile
                Result = $dcResult
            }
        }
    }
    elseif ($testTextExtractionParamName) {
        $dcParams = @{
            ErrorAction = "Stop"
        }
        if ($extractedResults -is [string] -or $extractedResults -is [string[]]) {
            throw "Test-TextExtraction results are strings, not TestTextExtractionResult objects. Cannot pass to Test-DataClassification."
        }
        $dcParams[$testTextExtractionParamName] = @($extractedResults)

        if ($selectedSensitiveTypes -and $scopeParamName) {
            $dcParams[$scopeParamName] = $selectedSensitiveTypes
        }
        elseif ($selectedSensitiveTypes -and -not $scopeParamName) {
            Write-Warning "Selected SITs were provided but Test-DataClassification does not expose a classification/sensitive type parameter. Scope will be ignored."
        }

        $dataClassificationResults = [pscustomobject]@{
            Mode = "TestTextExtractionResults"
            SourceFile = $sourceFileName
            Result = (Test-DataClassification @dcParams)
        }
    }

    [pscustomobject]@{
        SourceFile = $sourceFileName
        Streams = @($streamedText | Select-Object StreamIndex, Kind, Name, SourceFile)
        Extraction = $extractedResults
        DataClassification = $dataClassificationResults
    } | ConvertTo-Json -Depth 9
}
catch {
    Write-Error $_
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
        Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue
    }
}

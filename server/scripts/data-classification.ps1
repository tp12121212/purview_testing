[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$FilePath,

    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$AccessToken,

    [Parameter(Mandatory = $false)]
    [string]$SensitiveInformationTypes,

    [Parameter(Mandatory = $false)]
    [switch]$AllSensitiveInformationTypes
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
    if (-not $extractionResult) {
        throw "Test-TextExtraction returned no result."
    }

    $extractedResults = $extractionResult.ExtractedResults
    if (-not $extractedResults) {
        throw "Test-TextExtraction returned no ExtractedResults."
    }

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    if ($AccessToken) {
        Connect-IPPSSession -AccessToken $AccessToken -ShowBanner:$false -ErrorAction Stop
    }
    elseif ($UserPrincipalName) {
        Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
    }
    else {
        throw "Either AccessToken or UserPrincipalName must be provided for Purview compliance."
    }

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

    $testTextExtractionParamName = if ($testDataClassificationCommand.Parameters.ContainsKey("TestTextExtractionResults")) { "TestTextExtractionResults" } else { $null }

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

    $selectedSensitiveTypes = $null
    if (-not $AllSensitiveInformationTypes) {
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

        if ($SensitiveInformationTypes) {
            $requested = $SensitiveInformationTypes -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $matched = @()
            foreach ($request in $requested) {
                $hit = $sitCatalog | Where-Object { $_.Display -eq $request -or $_.Id -eq $request }
                if ($hit) {
                    $matched += $hit
                }
            }

            if ($matched.Count -gt 0) {
                if ($scopeParamName -eq "ClassificationNames") {
                    $selectedSensitiveTypes = $matched.Id | Select-Object -Unique
                }
                else {
                    $selectedSensitiveTypes = $matched.Display | Select-Object -Unique
                }
            }
        }
    }

    $sourceFileName = [System.IO.Path]::GetFileName($FilePath)
    $sourceExtension = [System.IO.Path]::GetExtension($FilePath)
    $isEmailSource = $false
    if ($sourceExtension) {
        $isEmailSource = @(".msg", ".eml") -contains $sourceExtension.ToLowerInvariant()
    }

    $resolveValue = {
        param(
            [object]$obj,
            [string]$name
        )
        if ($null -eq $obj) {
            return $null
        }
        if ($obj -is [System.Collections.IDictionary]) {
            if ($obj.Contains($name)) {
                return $obj[$name]
            }
            return $null
        }
        if ($obj.PSObject.Properties.Name -contains $name) {
            return $obj.$name
        }
        return $null
    }

    $streamedText = @()
    $streamIndex = 0
    foreach ($item in $extractedResults) {
        $streamTexts = @()
        $extractedStreamText = & $resolveValue $item "ExtractedStreamText"
        if ($extractedStreamText) {
            $streamTexts = @($extractedStreamText)
        }
        else {
            $extractedText = & $resolveValue $item "ExtractedText"
            if ($extractedText) {
                $streamTexts = @($extractedText)
            }
            else {
                if ($item -is [System.Collections.IDictionary]) {
                    $fallbackTextProps = $item.Keys | Where-Object { $_ -match 'Extracted.*Text' }
                    foreach ($propName in $fallbackTextProps) {
                        $value = $item[$propName]
                        if ($value) {
                            $streamTexts += @($value)
                        }
                    }
                }
                else {
                    $fallbackTextProps = $item.PSObject.Properties | Where-Object { $_.Name -match 'Extracted.*Text' }
                    foreach ($prop in $fallbackTextProps) {
                        if ($prop.Value) {
                            $streamTexts += @($prop.Value)
                        }
                    }
                }
            }
        }

        if (-not $streamTexts -or $streamTexts.Count -eq 0) {
            continue
        }

        $streamKind = "Unknown"
        $streamName = $null

        foreach ($prop in @("AttachmentName", "AttachmentFileName", "Attachment", "AttachmentFile")) {
            $value = & $resolveValue $item $prop
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $streamKind = "Attachment"
                $streamName = $value
                break
            }
        }

        if ($streamKind -eq "Unknown") {
            $isAttachment = & $resolveValue $item "IsAttachment"
            if ($isAttachment) {
                $streamKind = "Attachment"
            }
        }

        if (-not $streamName) {
            foreach ($prop in @("StreamName", "ItemName", "FileName", "Name")) {
                $value = & $resolveValue $item $prop
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $streamName = $value
                    break
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

    if (-not $streamedText -or $streamedText.Count -eq 0) {
        $fallbackText = $null
        if ($extractedResults -is [string]) {
            $fallbackText = $extractedResults
        }
        elseif ($extractedResults -is [string[]]) {
            $fallbackText = ($extractedResults -join [Environment]::NewLine)
        }
        else {
            $candidateTexts = @()
            foreach ($item in @($extractedResults)) {
                if ($item -is [string]) {
                    $candidateTexts += $item
                    continue
                }

                $extractedStreamText = & $resolveValue $item "ExtractedStreamText"
                if ($extractedStreamText) {
                    $candidateTexts += @($extractedStreamText)
                    continue
                }

                $extractedText = & $resolveValue $item "ExtractedText"
                if ($extractedText) {
                    $candidateTexts += @($extractedText)
                    continue
                }
            }

            if ($candidateTexts.Count -gt 0) {
                $fallbackText = ($candidateTexts -join [Environment]::NewLine)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($fallbackText)) {
            $fallbackKind = if ($isEmailSource) { "Body" } else { "Document" }
            $fallbackName = if ($isEmailSource) { "Body" } else { $sourceFileName }
            $streamedText = @([pscustomobject]@{
                StreamIndex = 1
                Kind = $fallbackKind
                Name = $fallbackName
                SourceFile = $sourceFileName
                Text = $fallbackText
            })
        }
    }

    if ($textParamName -or $fileDataParamName) {
        $dataClassificationResults = @()
        foreach ($stream in $streamedText) {
            $dcParams = @{ ErrorAction = "Stop" }
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
        $dcParams = @{ ErrorAction = "Stop" }
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
    exit 1
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
        Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue
    }
}

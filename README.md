# purview_testing
Scripts and other code to test and validate Microsoft Purview features, including DLP and text extraction.

## Contents
- `textExctraction.ps1`: Runs `Test-TextExtraction` against a local file, returning extracted results as JSON.
- `testDataclassification.ps1`: Runs `Test-TextExtraction`, then (optionally) `Test-DataClassification` against extracted streams with interactive SIT selection and structured output.

## Prerequisites
- PowerShell 7+ (`pwsh`) on macOS or Windows PowerShell on Windows.
- Exchange Online PowerShell module (`ExchangeOnlineManagement`).
- An account that can authenticate to Exchange Online and run `Test-TextExtraction`.
- An account with access to Microsoft Purview compliance cmdlets for `Connect-IPPSSession` and `Test-DataClassification` (only required for `testDataclassification.ps1`).

## Usage
### Text extraction only
Windows example:
```powershell
.\textExctraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf"
```

macOS example:
```powershell
pwsh ./textExctraction.ps1 -UserPrincipalName "admin@contoso.com" -MacFile "$HOME/temp/document.pdf"
```

Either OS (provide both; script uses the OS-appropriate one):
```powershell
.\textExctraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf" -MacFile "$HOME/temp/document.pdf"
```

### Text extraction + data classification
Run extraction + data classification on extracted text streams:
```powershell
.\testDataclassification.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.msg" -DataClassification
```

## Notes
- Provide a full file path; `~/` is not expanded by PowerShell in all contexts.
- Both scripts read the file as bytes and submit it to `Test-TextExtraction`.
- `testDataclassification.ps1` disconnects from Exchange Online after extraction and connects to IPPS to run `Test-DataClassification`.
- When `-DataClassification` is set, the script prompts to run against all SITs or lets you select specific SITs by number from `Get-DlpSensitiveInformationType`.
- If `ClassificationNames` is available, the script passes SIT GUIDs; otherwise it falls back to the display name parameter exposed by `Test-DataClassification`.
- The data-classification output includes stream context (`Kind`, `Name`, `SourceFile`) to distinguish message body vs attachment content for container formats (e.g., `.msg`, `.eml`).

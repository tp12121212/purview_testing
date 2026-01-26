# purview_testing
Scripts and other code to test and validate Microsoft Purview features, including DLP and text extraction.

## Contents
- `textExtraction_test.ps1`: Connects to Exchange Online and runs `Test-TextExtraction` against a local file, returning extracted results as JSON.

## Prerequisites
- PowerShell 7+ (`pwsh`) on macOS or Windows PowerShell on Windows.
- Exchange Online PowerShell module (`ExchangeOnlineManagement`).
- An account that can authenticate to Exchange Online and run `Test-TextExtraction`.

## Usage
Windows example:
```powershell
.\textExtraction_test.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf"
```

macOS example:
```powershell
pwsh ./textExtraction_test.ps1 -UserPrincipalName "admin@contoso.com" -MacFile "$HOME/temp/document.pdf"
```

Either OS (provide both; script uses the OS-appropriate one):
```powershell
.\textExtraction_test.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf" -MacFile "$HOME/temp/document.pdf"
```

Run extraction + data classification on extracted text streams:
```powershell
.\textExtraction_test.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.msg" -DataClassification
```

## Notes
- Provide a full file path; `~/` is not expanded by PowerShell in all contexts.
- The script reads the file as bytes and submits it to `Test-TextExtraction`.
- Output is JSON from `ExtractedResults` for easy inspection or piping. When `-DataClassification` is set,
  the script returns a composite JSON object with `Extraction` and `DataClassification`.
- When `-DataClassification` is set, the script prompts to run against all SITs or lets you select specific
  SITs by number from `Get-DlpSensitiveInformationType`.

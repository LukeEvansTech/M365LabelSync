# LabelHelpers Module

`LabelHelpers.psm1` is a shared PowerShell module imported by both `Export-Labels.ps1` and `Import-Labels.ps1`.

## Functions

### Connect-ComplianceTenant

Wraps `Connect-IPPSSession` with logging and connection verification.

```powershell
Connect-ComplianceTenant -UserPrincipalName admin@contoso.com
```

Verifies the connection by calling `Get-Label` after connecting. Throws if the connection or verification fails.

### Write-Log / Save-Log

Structured logging with colored console output and a log buffer.

```powershell
Write-Log "Something happened" -Level INFO    # Cyan
Write-Log "Watch out" -Level WARN             # Yellow
Write-Log "It broke" -Level ERROR             # Red
Write-Log "All good" -Level SUCCESS           # Green
Write-Log "Already exists" -Level SKIP        # DarkGray

Save-Log -Path ".\export\export-log.txt"
```

### ConvertTo-LabelExportObject

Converts a label object from `Get-Label` into a clean PSObject for JSON export. Resolves `ParentId` GUIDs to DisplayNames and generates the `_LabelPath` compound key.

Accepts pipeline input and a `GuidToDisplayName` hashtable for parent resolution.

### ConvertTo-PolicyExportObject

Converts a policy object from `Get-LabelPolicy` into a clean PSObject. Resolves label references (which may be GUIDs or Names) to `LabelPaths` using both `GuidToLabelPath` and `NameToLabelPath` lookups.

### Build-NewLabelParameters

Builds a splatting hashtable for `New-Label` from an exported label object. Supports `-SkipEncryption` switch and `-ParentId` for sub-labels.

### Build-NewLabelPolicyParameters

Builds a splatting hashtable for `New-LabelPolicy` from an exported policy object. Resolves `LabelPaths` to target-tenant GUIDs using a `LabelPathToGuid` hashtable.

## Exported functions

```powershell
Write-Log
Save-Log
Connect-ComplianceTenant
ConvertTo-LabelExportObject
ConvertTo-PolicyExportObject
Build-NewLabelParameters
Build-NewLabelPolicyParameters
```

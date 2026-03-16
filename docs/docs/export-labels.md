# Export-Labels.ps1

Exports all sensitivity labels and label policies from an M365 tenant to JSON files.

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-UserPrincipalName` | Yes | — | UPN of a Compliance Administrator in the source tenant |
| `-OutputDir` | No | `.\export` | Directory for exported JSON and log files |

## Usage

```powershell
# Default output to ./export/
.\Export-Labels.ps1 -UserPrincipalName admin@contoso.com

# Custom output directory
.\Export-Labels.ps1 -UserPrincipalName admin@contoso.com -OutputDir C:\migration\export
```

## Output files

### labels.json

Array of label objects, parents first then sub-labels, each containing:

- **Identity** — `DisplayName`, `Name`, `Comment`, `Tooltip`, `Priority`, `ContentType`
- **Hierarchy** — `ParentLabelDisplayName`, `_LabelPath` (compound key)
- **Content marking** — header, footer, and watermark settings
- **Encryption** — all encryption parameters, flagged with `_HasEncryption`
- **AdvancedSettings** — hashtable of advanced settings
- **Metadata** — `_SourceTenantLabelGuid`, `_ExportedAt`

### label-policies.json

Array of policy objects containing:

- **Identity** — `Name`, `Comment`, `Enabled`, `Priority`
- **Label references** — `LabelPaths` array with resolved compound paths
- **Scoping** — `ExchangeLocation`, `ModernGroupLocation` and exceptions
- **AdvancedSettings** — hashtable of advanced policy settings
- **Metadata** — `_SourceTenantPolicyGuid`, `_ExportedAt`

## Behaviour

1. Connects to the source tenant via `Connect-IPPSSession`
2. Retrieves all labels with `Get-Label -IncludeDetailedLabelActions`
3. Builds GUID and Name lookup tables for resolving parent labels and policy references
4. Separates parent and child labels, sorts by priority
5. Converts each label to a clean export object via `ConvertTo-LabelExportObject`
6. Retrieves all policies with `Get-LabelPolicy`
7. Resolves policy label references using both GUID and Name lookups
8. Logs warnings for labels with encryption enabled

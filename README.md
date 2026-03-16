# M365LabelSync

Export and import sensitivity labels and label policies between Microsoft 365 tenants using Security & Compliance PowerShell.

## Prerequisites

- **PowerShell 7+**
- **ExchangeOnlineManagement module v3.4+**
  ```powershell
  Install-Module ExchangeOnlineManagement -MinimumVersion 3.4.0
  ```
- **Compliance Administrator** role in both source and target tenants

## Quick Start

### 1. Export from source tenant

```powershell
.\Export-Labels.ps1 -UserPrincipalName admin@source-tenant.com
```

Outputs to `./export/`:
- `labels.json` — all sensitivity labels (parents first, then sub-labels)
- `label-policies.json` — all label policies
- `export-log.txt` — operation log

### 2. Review exported data

Inspect `labels.json` for encryption warnings and `label-policies.json` for user/group scoping that will need manual remapping.

### 3. Preview import (WhatIf)

```powershell
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -WhatIf
```

### 4. Import to target tenant

```powershell
# Recommended: skip encryption on first pass
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -SkipEncryption

# Or import everything (only if RMS/encryption is configured in target)
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com
```

### 5. Verify

Re-run the import to confirm idempotency — all labels should show as `SKIP`.

## Parameters

### Export-Labels.ps1

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-UserPrincipalName` | Yes | — | UPN of Compliance Admin in source tenant |
| `-OutputDir` | No | `.\export` | Directory for exported files |

### Import-Labels.ps1

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-UserPrincipalName` | Yes | — | UPN of Compliance Admin in target tenant |
| `-InputDir` | No | `.\export` | Directory containing exported JSON files |
| `-SkipEncryption` | No | `$false` | Skip encryption settings (recommended for cross-tenant) |
| `-SkipPolicies` | No | `$false` | Skip policy import |
| `-WhatIf` | No | `$false` | Preview without making changes |

## Cross-Tenant Caveats

- **Encryption**: RMS keys and templates are tenant-specific. Use `-SkipEncryption` unless you've configured encryption in the target tenant. Labels with encryption are flagged during export.
- **GUIDs**: Label and policy GUIDs are tenant-specific and cannot be preserved. Labels are matched by `_LabelPath` (a compound key like `Parent\Child`) across tenants, which correctly handles sub-labels that share a DisplayName under different parents.
- **User/group scoping**: Policy scoping that references specific users or groups will be logged but not automatically remapped (GUIDs differ between tenants). Remap manually in the Purview compliance portal.
- **Policy sync delay**: After import, allow up to 24 hours for policies to propagate to all M365 services.
- **AdvancedSettings**: Applied via `Set-Label`/`Set-LabelPolicy` after creation, as `New-Label` does not support them directly.
- **Auto-labeling conditions**: Not currently exported/imported. Can be added in a future version.

## Import Phases

1. **Parent labels** — Created first so sub-labels can reference them
2. **Sub-labels** — Created with `ParentId` resolved from the target tenant
3. **Policies** — Created with label GUIDs resolved from the target tenant

All phases are idempotent — existing labels/policies (matched by `_LabelPath`/`Name`) are skipped.

## File Structure

```
M365LabelSync/
  LabelHelpers.psm1     # Shared module (connection, logging, property mapping)
  Export-Labels.ps1      # Export labels + policies to JSON
  Import-Labels.ps1      # Import labels + policies from JSON
  README.md              # This file
  CLAUDE.md              # Project context for AI assistants
  export/                # Default output directory
    labels.json          # Exported labels (parents first, then sub-labels)
    label-policies.json  # Exported label policies
    export-log.txt       # Operation log
```

## How It Works

Labels are identified across tenants using a **`_LabelPath`** compound key (`Parent\Child` for sub-labels, or just `DisplayName` for parent labels). This handles the common case where sub-labels like "Anyone (not protected)" or "All Employees" appear under multiple parent labels.

Policy label references are resolved by looking up both label GUIDs and label Names (some labels use a Name that matches their DisplayName rather than a GUID format).

# Quick Start

## 1. Export from source tenant

```powershell
.\Export-Labels.ps1 -UserPrincipalName admin@source-tenant.com
```

This creates an `export/` directory containing:

- `labels.json` — all sensitivity labels (parents first, then sub-labels)
- `label-policies.json` — all label policies with resolved label paths
- `export-log.txt` — operation log with timestamps

## 2. Review exported data

Open `labels.json` and check for:

- Labels flagged with `_HasEncryption: true` — these need attention for cross-tenant migration
- The `_LabelPath` field — confirms the compound key used for matching (e.g. `Confidential\All Employees`)

Open `label-policies.json` and check:

- `LabelPaths` array — all label references should be resolved (no `UNRESOLVED:` prefixes)
- `ExchangeLocation` / `ModernGroupLocation` — scoping that may need manual remapping

## 3. Preview import (WhatIf)

```powershell
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -WhatIf
```

This connects to the target tenant and logs what _would_ happen without making any changes.

## 4. Import to target tenant

```powershell
# Recommended: skip encryption on first pass
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -SkipEncryption
```

!!! warning "Encryption"
    Use `-SkipEncryption` unless you've specifically configured RMS/encryption in the target tenant. RMS keys and templates are tenant-specific and don't transfer.

## 5. Verify

Re-run the import command. All labels should show as `SKIP`, confirming idempotency:

```powershell
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -SkipEncryption
```

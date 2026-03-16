# Import-Labels.ps1

Imports sensitivity labels and label policies into an M365 tenant from exported JSON files.

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-UserPrincipalName` | Yes | — | UPN of a Compliance Administrator in the target tenant |
| `-InputDir` | No | `.\export` | Directory containing `labels.json` and `label-policies.json` |
| `-SkipEncryption` | No | `$false` | Skip encryption settings when creating labels |
| `-SkipPolicies` | No | `$false` | Skip importing label policies |
| `-WhatIf` | No | `$false` | Preview changes without creating anything |

## Usage

```powershell
# Preview changes
.\Import-Labels.ps1 -UserPrincipalName admin@fabrikam.com -WhatIf

# Import without encryption (recommended for cross-tenant)
.\Import-Labels.ps1 -UserPrincipalName admin@fabrikam.com -SkipEncryption

# Import everything (only if encryption is configured in target)
.\Import-Labels.ps1 -UserPrincipalName admin@fabrikam.com

# Import labels only, skip policies
.\Import-Labels.ps1 -UserPrincipalName admin@fabrikam.com -SkipEncryption -SkipPolicies
```

## Import phases

### Phase 1: Parent labels

Creates top-level labels first. Each label is matched against the target tenant by `_LabelPath` (which equals `DisplayName` for parents). If a match is found, the label is skipped.

### Phase 2: Sub-labels

Creates child labels, resolving `ParentId` by looking up the parent's DisplayName in the target tenant. Matched by `_LabelPath` (e.g. `Confidential\All Employees`).

### Phase 3: Policies

Creates label policies, resolving `LabelPaths` to target-tenant GUIDs. Skips if a policy with the same `Name` already exists.

## AdvancedSettings

`New-Label` and `New-LabelPolicy` do not support the `-AdvancedSettings` parameter. After creating each label/policy, the script applies advanced settings via `Set-Label` or `Set-LabelPolicy`.

## Idempotency

The import is fully idempotent. Re-running against the same target tenant will skip all existing labels and policies, logging each as `SKIP`.

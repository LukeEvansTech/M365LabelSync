# M365LabelSync

**Export and import M365 sensitivity labels between tenants.**

No native Microsoft tool exists for migrating sensitivity labels and label policies between M365 tenants. M365LabelSync fills this gap with a simple PowerShell 7+ solution using Security & Compliance cmdlets.

## What it does

- **Exports** all sensitivity labels and label policies from a source tenant to JSON
- **Imports** them into a target tenant, preserving hierarchy, content marking, and advanced settings
- **Handles duplicates** — sub-labels that share a DisplayName across different parents are correctly matched using compound `_LabelPath` keys
- **Idempotent** — safe to re-run; existing labels/policies are skipped

## Workflow

```
Source Tenant                          Target Tenant
┌─────────────┐    labels.json       ┌─────────────┐
│  Get-Label   │──────────────────▶  │  New-Label   │
│  Get-Policy  │──────────────────▶  │  New-Policy  │
└─────────────┘  label-policies.json └─────────────┘
```

1. **Export** — Run `Export-Labels.ps1` against the source tenant
2. **Review** — Inspect the JSON, check encryption warnings
3. **Preview** — Run `Import-Labels.ps1 -WhatIf` against the target tenant
4. **Import** — Run `Import-Labels.ps1 -SkipEncryption` against the target tenant
5. **Verify** — Re-run import to confirm all labels show as `SKIP`

## Quick example

```powershell
# Export
.\Export-Labels.ps1 -UserPrincipalName admin@source.onmicrosoft.com

# Import (skip encryption — RMS keys don't transfer)
.\Import-Labels.ps1 -UserPrincipalName admin@target.onmicrosoft.com -SkipEncryption
```

## Files

| File | Purpose |
|------|---------|
| `LabelHelpers.psm1` | Shared module — connection, logging, property mapping |
| `Export-Labels.ps1` | Export labels + policies to JSON |
| `Import-Labels.ps1` | Import labels + policies from JSON |

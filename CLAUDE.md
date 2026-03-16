# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
PowerShell 7+ tool for migrating sensitivity labels and label policies between M365 tenants using Security & Compliance PowerShell cmdlets (Get-Label, New-Label, Set-Label, Get-LabelPolicy, New-LabelPolicy, Set-LabelPolicy).

## Running

```powershell
# Export from source tenant
.\Export-Labels.ps1 -UserPrincipalName admin@source-tenant.com

# Preview import (dry run)
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -WhatIf

# Import to target tenant (skip encryption for cross-tenant)
.\Import-Labels.ps1 -UserPrincipalName admin@target-tenant.com -SkipEncryption
```

Requires: PowerShell 7+, ExchangeOnlineManagement module v3.4+, Compliance Administrator role.

## Architecture

Three files, no build step:

- **`LabelHelpers.psm1`** — Shared module imported by both scripts. Contains: connection helper (`Connect-ComplianceTenant`), structured logging (`Write-Log`/`Save-Log`), export converters (`ConvertTo-LabelExportObject`/`ConvertTo-PolicyExportObject`), and parameter builders (`Build-NewLabelParameters`/`Build-NewLabelPolicyParameters`).
- **`Export-Labels.ps1`** — Reads all labels/policies from source tenant, builds GUID/Name-to-LabelPath lookup tables, outputs JSON to `export/`.
- **`Import-Labels.ps1`** — Three-phase import: (1) parent labels, (2) sub-labels (with ParentId resolved from target tenant), (3) policies (with label GUIDs resolved from target tenant).

## Key Design Decisions
- **`_LabelPath` as cross-tenant key** — Compound key (`Parent\Child`) instead of bare DisplayName, because sub-labels can share DisplayNames across different parents (e.g. "Anyone (not protected)" exists under both Confidential and Highly Confidential).
- **Parents before children** — Import creates parent labels first so ParentId can be resolved for sub-labels.
- **`-SkipEncryption` flag** — RMS keys/templates are tenant-specific and don't transfer.
- **AdvancedSettings applied post-creation** — `New-Label` doesn't support AdvancedSettings; must use `Set-Label` after.
- **Policy label references** — Policy.Labels array may contain GUIDs or Name values depending on label type; export resolves both via `GuidToLabelPath` and `NameToLabelPath` lookups.
- **`-IncludeDetailedLabelActions`** — Is a switch parameter (no `$true` argument), not a boolean parameter.
- **Idempotent** — Import skips existing labels/policies matched by _LabelPath/Name.

## Tested Against
- 10 labels (5 parents, 5 sub-labels), 1 policy
- 2 labels with encryption (Manual encryption, HC\All Employees)

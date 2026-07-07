# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**M365LabelSync** — a small PowerShell tool for exporting sensitivity labels and label policies from one Microsoft 365 tenant and importing them into another, via the Security & Compliance (`Connect-IPPSSession`) PowerShell surface.

Public docs: <https://lukeevansTech.github.io/M365LabelSync/>
Repository: <https://github.com/LukeEvansTech/M365LabelSync>

The full user-facing reference is in `README.md`. This file captures the _why_ and non-obvious internals that aren't derivable by reading the scripts.

## Architecture in one paragraph

Three files do the work: `Export-Labels.ps1` dumps labels and policies to JSON; `Import-Labels.ps1` reads that JSON and recreates them in a target tenant;
`LabelHelpers.psm1` holds the shared logic — connection, logging, and the property-mapping functions (`ConvertTo-LabelExportObject`, `ConvertTo-PolicyExportObject`, `Build-NewLabelParameters`, `Build-NewLabelPolicyParameters`). Output lands in `./export/` by default.

## Key design decisions (non-obvious)

- **`_LabelPath` compound key** — labels are matched across tenants by `Parent\Child` (or just `DisplayName` for parents), _not_ by GUID. This is because label GUIDs are tenant-specific and cannot be preserved, and because sub-label `DisplayName`s like "Anyone (not protected)" or "All Employees" regularly repeat under different parents. See `ConvertTo-LabelExportObject` in `LabelHelpers.psm1`.
- **Policy label references resolve by both GUID and Name** — `Get-LabelPolicy` can return either in `Labels`, so the export builds two lookup tables (`guidToLabelPath` + `nameToLabelPath`) and the resolver falls back from one to the other. See `Export-Labels.ps1` and `ConvertTo-PolicyExportObject`.
- **Parents before children** — import phases are: parent labels → sub-labels (with `ParentId` resolved in target) → policies (with label GUIDs resolved in target). Export sorts output the same way so the file is human-readable.
- **Idempotency** — both scripts are safe to re-run. Existing labels/policies (matched by `_LabelPath` / `Name`) are skipped. Re-running import after a successful run should produce all `SKIP` lines.
- **`AdvancedSettings` applied post-create** — `New-Label`/`New-LabelPolicy` don't accept `AdvancedSettings` directly, so they're applied via `Set-Label`/`Set-LabelPolicy` after creation.
- **Single-object JSON unwrap** — `Import-Labels.ps1` re-wraps the deserialized data in `@(...)` (see lines around 62 and 207) because `ConvertFrom-Json` returns a bare object, not a 1-element array, when the source JSON contains exactly one label or policy. Removing those guards silently breaks single-item imports.
- **`Get-Label -IncludeDetailedLabelActions` is load-bearing** — the flag in `Export-Labels.ps1` is what populates content-marking and encryption properties. Without it the export still succeeds but silently produces labels with empty `HeaderText`/`WatermarkText`/`Encryption*` fields. Don't drop it.

## Cross-tenant caveats to remember

- **Encryption / RMS is tenant-specific.** Labels with encryption are flagged during export. Default recommendation: first-pass import with `-SkipEncryption`, then re-apply encryption settings in the target tenant once RMS templates are in place.
- **User/group scoping on policies does not port.** GUIDs differ between tenants. The export logs the scoping for manual remapping in the Purview portal; the import does not attempt to remap automatically.
- **Auto-labelling conditions are not currently exported or imported.** Flagged in `README.md` as a possible future addition.
- **Policy sync delay** — allow up to 24 hours after import for policies to propagate across M365 services.

## Testing / verification

There are no automated tests — no Pester suite, no linter configured. Verification path for any change:

- `-WhatIf` on `Import-Labels.ps1` for a dry run preview before touching a real tenant.
- Round-trip: export from a test tenant → import into a second tenant → re-run the import and confirm every line shows `SKIP` (idempotency check).
- The saved log files (`export-log.txt` / `import-log.txt`) are the primary signal; `Write-Log` entries are both printed and buffered for `Save-Log` to flush at the end.

## Current context (transient — verify before relying on)

**Active scenario**: Using this tool to stand up the Barclays label taxonomy in the Tesco Bank tenant. Standard Barclays labels: Secret, Restricted-Internal, Restricted-External, Unrestricted. BU-specific (likely not applicable to Tesco Bank): Banking Secrecy, Strictly Private and Confidential, Confidential Supervisory Information.

**Architectural pushback in flight (May 2026)**: The tactical "replicate Barclays labels into Tesco" framing is being challenged in an architecture review. The challenge is technically correct — labels recreated cross-tenant are separate technical objects, with implications for DLP enforcement, Vontu/Zscaler config, and rework risk if tenant strategy lands on full migration.
Background, ground truth, and meeting talking points are in `meeting-prep-cross-tenant-labels.md`. Treat any further label-replication work as scope-sensitive: prefer minimum-viable subsets, keep scope defensible, and don't expand into encryption / user-group scoping / auto-labelling without explicit sign-off.

## Conventions

- PowerShell 7+ only (`#Requires -Version 7.0` at the top of each script).
- Requires `ExchangeOnlineManagement` v3.4+ (provides `Connect-IPPSSession`). Not enforced by the scripts themselves — install with `Install-Module ExchangeOnlineManagement -MinimumVersion 3.4.0`.
- `Write-Log` is the only acceptable way to surface status — do not use bare `Write-Host`. It writes coloured output _and_ buffers entries for `Save-Log`.
- When adding a new label property to the export, update both `ConvertTo-LabelExportObject` (export side) and `Build-NewLabelParameters` (import side) together.
- Both scripts load the shared module via `Import-Module ... LabelHelpers.psm1 -Force`. The `-Force` is load-bearing — PowerShell caches imported modules per-session, so edits to `LabelHelpers.psm1` won't take effect without it.
- `Connect-ComplianceTenant` runs a no-op `Get-Label` after `Connect-IPPSSession` as a verification probe — it surfaces a missing Compliance Administrator role immediately rather than letting downstream cmdlets fail with less-obvious errors. Keep the probe when refactoring connection logic.

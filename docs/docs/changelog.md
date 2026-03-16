# Changelog

## v1.0.0 — 2026-03-16

Initial release.

- Export sensitivity labels and label policies from M365 tenants to JSON
- Import labels and policies into target tenants with idempotency
- `_LabelPath` compound key for cross-tenant matching (handles duplicate sub-label DisplayNames)
- Three-phase import: parent labels, sub-labels, then policies
- `-SkipEncryption` flag for cross-tenant migration where RMS keys don't transfer
- `-WhatIf` mode for previewing changes
- `-SkipPolicies` flag to import labels only
- AdvancedSettings applied via `Set-Label`/`Set-LabelPolicy` after creation
- Dual policy label resolution (GUID and Name lookups)
- Structured logging with colored console output and log file export

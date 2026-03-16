# Cross-Tenant Caveats

## Encryption

RMS keys and encryption templates are tenant-specific. Labels with encryption enabled will be flagged during export with `_HasEncryption: true`.

Use the `-SkipEncryption` flag on import unless you have specifically configured encryption in the target tenant. The import script will log a warning for each skipped encryption configuration.

## GUIDs

All GUIDs (label, policy, user, group) are tenant-specific and cannot be preserved. M365LabelSync matches labels across tenants using the `_LabelPath` compound key instead.

## User and group scoping

Policy scoping that references specific users or groups (`ExchangeLocation`, `ModernGroupLocation`, etc.) is exported for reference but **not automatically remapped**. The import logs these for manual attention:

```
WARN: Policy 'Sales Policy': ExchangeLocation scoping needs manual remapping
```

After import, configure scoping manually in the Microsoft Purview compliance portal.

## Policy sync delay

After importing policies, allow **up to 24 hours** for them to propagate to all M365 services (Exchange, SharePoint, Teams, etc.).

## Auto-labeling conditions

Auto-labeling conditions are **not currently exported or imported**. These can be configured manually in the target tenant after migration, or support may be added in a future version.

## Duplicate sub-label DisplayNames

Sub-labels can share a DisplayName across different parent labels (e.g. "All Employees" under both "Confidential" and "Highly Confidential"). M365LabelSync handles this correctly using `_LabelPath` compound keys, but be aware of it when reviewing exported JSON.

## EncryptionRightsDefinitions

The `EncryptionRightsDefinitions` field may contain tenant-specific identity references (e.g. `contoso.onmicrosoft.com`). When importing with encryption enabled, these references will need to be updated to match the target tenant's domain.

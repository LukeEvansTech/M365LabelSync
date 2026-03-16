# Prerequisites

## PowerShell 7+

M365LabelSync requires PowerShell 7 or later. Check your version:

```powershell
$PSVersionTable.PSVersion
```

## ExchangeOnlineManagement Module

Version 3.4.0 or later is required for REST-based Security & Compliance cmdlets.

```powershell
# Install
Install-Module ExchangeOnlineManagement -MinimumVersion 3.4.0

# Or update if already installed
Update-Module ExchangeOnlineManagement
```

## Compliance Administrator Role

You need the **Compliance Administrator** (or **Compliance Data Administrator**) role in both the source and target tenants. This role grants access to:

- `Get-Label` / `New-Label` / `Set-Label`
- `Get-LabelPolicy` / `New-LabelPolicy` / `Set-LabelPolicy`

The role is assigned in the Microsoft Entra admin center or via PowerShell.

## Network Access

The scripts connect to Security & Compliance PowerShell via `Connect-IPPSSession`. Ensure your network allows outbound connections to Microsoft 365 endpoints.

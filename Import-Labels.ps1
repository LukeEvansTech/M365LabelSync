#Requires -Version 7.0

<#
.SYNOPSIS
    Imports sensitivity labels and label policies into an M365 tenant from exported JSON files.

.PARAMETER UserPrincipalName
    UPN of a Compliance Administrator in the target tenant.

.PARAMETER InputDir
    Directory containing labels.json and label-policies.json. Defaults to .\export.

.PARAMETER SkipEncryption
    Skip encryption settings when creating labels (RMS keys/templates don't transfer between tenants).

.PARAMETER SkipPolicies
    Skip importing label policies.

.PARAMETER WhatIf
    Preview changes without creating anything in the target tenant.

.EXAMPLE
    .\Import-Labels.ps1 -UserPrincipalName admin@fabrikam.com -WhatIf
    .\Import-Labels.ps1 -UserPrincipalName admin@fabrikam.com -SkipEncryption
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,

    [string]$InputDir = (Join-Path $PSScriptRoot 'export'),

    [switch]$SkipEncryption,

    [switch]$SkipPolicies
)

$ErrorActionPreference = 'Stop'

# Import shared module
Import-Module (Join-Path $PSScriptRoot 'LabelHelpers.psm1') -Force

Write-Log "=== Sensitivity Label Import ===" -Level INFO
Write-Log "Input directory: $InputDir" -Level INFO
if ($WhatIfPreference) { Write-Log "*** WhatIf mode — no changes will be made ***" -Level WARN }
if ($SkipEncryption)   { Write-Log "Encryption settings will be skipped" -Level WARN }
if ($SkipPolicies)     { Write-Log "Policy import will be skipped" -Level WARN }

# Validate input files
$labelsPath   = Join-Path $InputDir 'labels.json'
$policiesPath = Join-Path $InputDir 'label-policies.json'

if (-not (Test-Path $labelsPath)) {
    Write-Log "Labels file not found: $labelsPath" -Level ERROR
    throw "Labels file not found: $labelsPath"
}

$labelData = Get-Content -Path $labelsPath -Raw | ConvertFrom-Json

# Handle single-object JSON (ConvertFrom-Json doesn't wrap single objects in an array)
if ($labelData -isnot [System.Collections.IEnumerable] -or $labelData -is [string]) {
    $labelData = @($labelData)
}

Write-Log "Loaded $($labelData.Count) label(s) from $labelsPath" -Level INFO

# Connect to target tenant
if (-not $WhatIfPreference) {
    Connect-ComplianceTenant -UserPrincipalName $UserPrincipalName
}

# Get existing labels in target tenant for idempotency
# Key by _LabelPath (Parent\Child) to handle duplicate DisplayNames across parents
$existingLabelsByPath = @{}
$targetLabelPathToGuid = @{}
$targetParentDnToGuid = @{}  # parent DisplayName -> GUID
if (-not $WhatIfPreference) {
    $targetLabels = Get-Label
    # First pass: index parents
    $guidToDn = @{}
    foreach ($lbl in $targetLabels) {
        $guidToDn[$lbl.Guid.ToString()] = $lbl.DisplayName
    }
    foreach ($lbl in $targetLabels) {
        $parentDn = $null
        if ($lbl.ParentId -and $lbl.ParentId -ne [Guid]::Empty) {
            $parentDn = $guidToDn[$lbl.ParentId.ToString()]
        }
        $path = if ($parentDn) { "$parentDn\$($lbl.DisplayName)" } else { $lbl.DisplayName }
        $existingLabelsByPath[$path] = $lbl
        $targetLabelPathToGuid[$path] = $lbl.Guid.ToString()
        if (-not $parentDn) {
            $targetParentDnToGuid[$lbl.DisplayName] = $lbl.Guid.ToString()
        }
    }
    Write-Log "Target tenant has $($existingLabelsByPath.Count) existing label(s)." -Level INFO
}

# Separate parents and children
$parentData = $labelData | Where-Object { -not $_.ParentLabelDisplayName }
$childData  = $labelData | Where-Object { $_.ParentLabelDisplayName }

# --- Phase 1: Parent Labels ---
Write-Log "`n--- Phase 1: Parent Labels ($($parentData.Count)) ---" -Level INFO

foreach ($label in $parentData) {
    $dn = $label.DisplayName
    $path = $label._LabelPath

    if ($existingLabelsByPath.ContainsKey($path)) {
        Write-Log "SKIP parent '$dn' — already exists in target tenant" -Level SKIP
        continue
    }

    $params = Build-NewLabelParameters -LabelData $label -SkipEncryption:$SkipEncryption

    if ($WhatIfPreference) {
        Write-Log "WHATIF: Would create parent label '$dn'" -Level INFO
        Write-Log "  Parameters: $($params.Keys -join ', ')" -Level INFO
        continue
    }

    Write-Log "Creating parent label '$dn'..." -Level INFO
    try {
        $newLabel = New-Label @params
        $targetLabelPathToGuid[$path] = $newLabel.Guid.ToString()
        $targetParentDnToGuid[$dn] = $newLabel.Guid.ToString()
        $existingLabelsByPath[$path] = $newLabel
        Write-Log "  Created '$dn' (GUID: $($newLabel.Guid))" -Level SUCCESS

        # Apply AdvancedSettings via Set-Label (not supported on New-Label)
        if ($label.AdvancedSettings -and $label.AdvancedSettings.Count -gt 0) {
            $advSettings = @{}
            foreach ($prop in $label.AdvancedSettings.PSObject.Properties) {
                $advSettings[$prop.Name] = $prop.Value
            }
            Set-Label -Identity $newLabel.Guid.ToString() -AdvancedSettings $advSettings
            Write-Log "  Applied $($advSettings.Count) advanced setting(s) to '$dn'" -Level SUCCESS
        }
    }
    catch {
        Write-Log "  FAILED to create parent label '$dn': $_" -Level ERROR
    }
}

# --- Phase 2: Sub-Labels ---
Write-Log "`n--- Phase 2: Sub-Labels ($($childData.Count)) ---" -Level INFO

foreach ($label in $childData) {
    $dn = $label.DisplayName
    $parentDn = $label.ParentLabelDisplayName
    $path = $label._LabelPath

    if ($existingLabelsByPath.ContainsKey($path)) {
        Write-Log "SKIP sub-label '$path' — already exists in target tenant" -Level SKIP
        continue
    }

    # Resolve parent GUID in target tenant
    $parentGuid = $targetParentDnToGuid[$parentDn]
    if (-not $parentGuid) {
        Write-Log "  SKIP sub-label '$path' — parent '$parentDn' not found in target tenant" -Level ERROR
        continue
    }

    $params = Build-NewLabelParameters -LabelData $label -SkipEncryption:$SkipEncryption -ParentId $parentGuid

    if ($WhatIfPreference) {
        Write-Log "WHATIF: Would create sub-label '$path'" -Level INFO
        Write-Log "  Parameters: $($params.Keys -join ', ')" -Level INFO
        continue
    }

    Write-Log "Creating sub-label '$path'..." -Level INFO
    try {
        $newLabel = New-Label @params
        $targetLabelPathToGuid[$path] = $newLabel.Guid.ToString()
        $existingLabelsByPath[$path] = $newLabel
        Write-Log "  Created '$path' (GUID: $($newLabel.Guid))" -Level SUCCESS

        # Apply AdvancedSettings via Set-Label
        if ($label.AdvancedSettings -and $label.AdvancedSettings.Count -gt 0) {
            $advSettings = @{}
            foreach ($prop in $label.AdvancedSettings.PSObject.Properties) {
                $advSettings[$prop.Name] = $prop.Value
            }
            Set-Label -Identity $newLabel.Guid.ToString() -AdvancedSettings $advSettings
            Write-Log "  Applied $($advSettings.Count) advanced setting(s) to '$path'" -Level SUCCESS
        }
    }
    catch {
        Write-Log "  FAILED to create sub-label '$path': $_" -Level ERROR
    }
}

# --- Phase 3: Policies ---
if ($SkipPolicies) {
    Write-Log "`n--- Phase 3: Policies — SKIPPED ---" -Level WARN
}
else {
    if (-not (Test-Path $policiesPath)) {
        Write-Log "Policies file not found: $policiesPath — skipping policy import" -Level WARN
    }
    else {
        $policyData = Get-Content -Path $policiesPath -Raw | ConvertFrom-Json
        if ($policyData -isnot [System.Collections.IEnumerable] -or $policyData -is [string]) {
            $policyData = @($policyData)
        }

        Write-Log "`n--- Phase 3: Policies ($($policyData.Count)) ---" -Level INFO

        # Get existing policies for idempotency
        $existingPolicies = @{}
        if (-not $WhatIfPreference) {
            $targetPolicies = Get-LabelPolicy
            foreach ($pol in $targetPolicies) {
                $existingPolicies[$pol.Name] = $pol
            }
        }

        foreach ($policy in $policyData) {
            $pName = $policy.Name

            if ($existingPolicies.ContainsKey($pName)) {
                Write-Log "SKIP policy '$pName' — already exists in target tenant" -Level SKIP
                continue
            }

            $params = Build-NewLabelPolicyParameters -PolicyData $policy -LabelPathToGuid $targetLabelPathToGuid

            if ($WhatIfPreference) {
                Write-Log "WHATIF: Would create policy '$pName'" -Level INFO
                Write-Log "  Labels: $($policy.LabelPaths -join ', ')" -Level INFO
                continue
            }

            Write-Log "Creating policy '$pName'..." -Level INFO
            try {
                $newPolicy = New-LabelPolicy @params
                Write-Log "  Created policy '$pName'" -Level SUCCESS

                # Apply AdvancedSettings via Set-LabelPolicy
                if ($policy.AdvancedSettings -and $policy.AdvancedSettings.Count -gt 0) {
                    $advSettings = @{}
                    foreach ($prop in $policy.AdvancedSettings.PSObject.Properties) {
                        $advSettings[$prop.Name] = $prop.Value
                    }
                    Set-LabelPolicy -Identity $pName -AdvancedSettings $advSettings
                    Write-Log "  Applied $($advSettings.Count) advanced setting(s) to policy '$pName'" -Level SUCCESS
                }
            }
            catch {
                Write-Log "  FAILED to create policy '$pName': $_" -Level ERROR
            }
        }
    }
}

# Save log
$logPath = Join-Path $InputDir 'import-log.txt'
Save-Log -Path $logPath

Write-Log "`nImport complete." -Level SUCCESS

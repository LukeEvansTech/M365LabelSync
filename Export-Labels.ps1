#Requires -Version 7.0

<#
.SYNOPSIS
    Exports sensitivity labels and label policies from an M365 tenant to JSON files.

.PARAMETER UserPrincipalName
    UPN of a Compliance Administrator in the source tenant.

.PARAMETER OutputDir
    Directory for exported JSON and log files. Defaults to .\export.

.EXAMPLE
    .\Export-Labels.ps1 -UserPrincipalName admin@contoso.com
    .\Export-Labels.ps1 -UserPrincipalName admin@contoso.com -OutputDir C:\migration\export
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,

    [string]$OutputDir = (Join-Path $PSScriptRoot 'export')
)

$ErrorActionPreference = 'Stop'

# Import shared module
Import-Module (Join-Path $PSScriptRoot 'LabelHelpers.psm1') -Force

Write-Log "=== Sensitivity Label Export ===" -Level INFO
Write-Log "Output directory: $OutputDir" -Level INFO

# Ensure output directory exists
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
    Write-Log "Created output directory: $OutputDir" -Level INFO
}

# Connect to source tenant
Connect-ComplianceTenant -UserPrincipalName $UserPrincipalName

# --- Export Labels ---
Write-Log "Retrieving labels..." -Level INFO
$allLabels = Get-Label -IncludeDetailedLabelActions

if (-not $allLabels -or $allLabels.Count -eq 0) {
    Write-Log "No labels found in tenant." -Level WARN
    $allLabels = @()
}

Write-Log "Found $($allLabels.Count) label(s)." -Level INFO

# Build GUID -> DisplayName lookup (needed for resolving ParentId)
$guidToDisplayName = @{}
foreach ($lbl in $allLabels) {
    $guidToDisplayName[$lbl.Guid.ToString()] = $lbl.DisplayName
}

# Build GUID -> LabelPath and Name -> LabelPath lookups (for policy resolution)
$guidToLabelPath = @{}
$nameToLabelPath = @{}
foreach ($lbl in $allLabels) {
    $parentDn = $null
    if ($lbl.ParentId -and $lbl.ParentId -ne [Guid]::Empty) {
        $parentDn = $guidToDisplayName[$lbl.ParentId.ToString()]
    }
    $path = if ($parentDn) { "$parentDn\$($lbl.DisplayName)" } else { $lbl.DisplayName }
    $guidToLabelPath[$lbl.Guid.ToString()] = $path
    $nameToLabelPath[$lbl.Name] = $path
}

# Separate parents and children, sort by priority
$parentLabels = $allLabels | Where-Object {
    -not $_.ParentId -or $_.ParentId -eq [Guid]::Empty
} | Sort-Object Priority

$childLabels = $allLabels | Where-Object {
    $_.ParentId -and $_.ParentId -ne [Guid]::Empty
} | Sort-Object Priority

Write-Log "Parents: $($parentLabels.Count), Sub-labels: $($childLabels.Count)" -Level INFO

# Convert to export objects (parents first, then children)
$exportLabels = @()

foreach ($lbl in $parentLabels) {
    $exported = $lbl | ConvertTo-LabelExportObject -GuidToDisplayName $guidToDisplayName
    $exportLabels += $exported

    if ($exported._HasEncryption) {
        Write-Log "  Label '$($exported.DisplayName)' has encryption — review before importing to another tenant" -Level WARN
    }
}

foreach ($lbl in $childLabels) {
    $exported = $lbl | ConvertTo-LabelExportObject -GuidToDisplayName $guidToDisplayName
    $exportLabels += $exported

    if ($exported._HasEncryption) {
        Write-Log "  Sub-label '$($exported.DisplayName)' (parent: $($exported.ParentLabelDisplayName)) has encryption" -Level WARN
    }
}

$labelsPath = Join-Path $OutputDir 'labels.json'
$exportLabels | ConvertTo-Json -Depth 10 | Out-File -FilePath $labelsPath -Encoding utf8
Write-Log "Exported $($exportLabels.Count) label(s) to $labelsPath" -Level SUCCESS

# --- Export Label Policies ---
Write-Log "Retrieving label policies..." -Level INFO
$allPolicies = Get-LabelPolicy

if (-not $allPolicies -or $allPolicies.Count -eq 0) {
    Write-Log "No label policies found in tenant." -Level WARN
    $allPolicies = @()
}

Write-Log "Found $($allPolicies.Count) policy/policies." -Level INFO

$exportPolicies = @()
foreach ($pol in $allPolicies) {
    $exported = $pol | ConvertTo-PolicyExportObject -GuidToLabelPath $guidToLabelPath -NameToLabelPath $nameToLabelPath
    $exportPolicies += $exported
}

$policiesPath = Join-Path $OutputDir 'label-policies.json'
$exportPolicies | ConvertTo-Json -Depth 10 | Out-File -FilePath $policiesPath -Encoding utf8
Write-Log "Exported $($exportPolicies.Count) policy/policies to $policiesPath" -Level SUCCESS

# Save log
$logPath = Join-Path $OutputDir 'export-log.txt'
Save-Log -Path $logPath

Write-Log "Export complete." -Level SUCCESS

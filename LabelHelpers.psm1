#Requires -Version 7.0

<#
.SYNOPSIS
    Shared helpers for sensitivity label export/import between M365 tenants.
#>

# Module-level log buffer
$script:LogEntries = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','SKIP')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"
    $script:LogEntries.Add($entry)

    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'SKIP'    { 'DarkGray' }
    }
    Write-Host $entry -ForegroundColor $color
}

function Save-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    $script:LogEntries | Out-File -FilePath $Path -Encoding utf8
    Write-Log "Log saved to $Path" -Level INFO
}

function Connect-ComplianceTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    Write-Log "Connecting to Security & Compliance as $UserPrincipalName..."

    try {
        Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ErrorAction Stop
        Write-Log "Connected successfully." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to connect: $_" -Level ERROR
        throw
    }

    # Verify connection by running a simple cmdlet
    try {
        $null = Get-Label -ErrorAction Stop | Select-Object -First 1
        Write-Log "Connection verified (Get-Label responded)." -Level SUCCESS
    }
    catch {
        Write-Log "Connection verification failed: $_" -Level ERROR
        throw "Connected but Get-Label failed. Ensure you have Compliance Administrator role."
    }
}

function ConvertTo-LabelExportObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]$Label,
        [hashtable]$GuidToDisplayName = @{}
    )

    process {
        $parentDisplayName = $null
        if ($Label.ParentId -and $Label.ParentId -ne [Guid]::Empty) {
            $parentGuid = $Label.ParentId.ToString()
            $parentDisplayName = $GuidToDisplayName[$parentGuid]
            if (-not $parentDisplayName) {
                Write-Log "Could not resolve ParentId $parentGuid for label '$($Label.DisplayName)'" -Level WARN
            }
        }

        # Compound key for cross-tenant matching (handles duplicate DisplayNames across parents)
        $labelPath = if ($parentDisplayName) { "$parentDisplayName\$($Label.DisplayName)" } else { $Label.DisplayName }

        $hasEncryption = [bool]($Label.EncryptionEnabled)

        $obj = [ordered]@{
            # Identity
            DisplayName              = $Label.DisplayName
            Name                     = $Label.Name
            Comment                  = $Label.Comment
            Tooltip                  = $Label.Tooltip
            Priority                 = $Label.Priority
            ContentType              = $Label.ContentType

            # Hierarchy
            ParentLabelDisplayName   = $parentDisplayName
            _LabelPath               = $labelPath

            # Content marking — Header
            HeaderEnabled            = $Label.HeaderEnabled
            HeaderText               = $Label.HeaderText
            HeaderFontName           = $Label.HeaderFontName
            HeaderFontSize           = $Label.HeaderFontSize
            HeaderFontColor          = $Label.HeaderFontColor
            HeaderAlignment          = $Label.HeaderAlignment

            # Content marking — Footer
            FooterEnabled            = $Label.FooterEnabled
            FooterText               = $Label.FooterText
            FooterFontName           = $Label.FooterFontName
            FooterFontSize           = $Label.FooterFontSize
            FooterFontColor          = $Label.FooterFontColor
            FooterAlignment          = $Label.FooterAlignment

            # Content marking — Watermark
            WatermarkEnabled         = $Label.WatermarkEnabled
            WatermarkText            = $Label.WatermarkText
            WatermarkFontName        = $Label.WatermarkFontName
            WatermarkFontSize        = $Label.WatermarkFontSize
            WatermarkFontColor       = $Label.WatermarkFontColor
            WatermarkLayout          = $Label.WatermarkLayout

            # Encryption
            _HasEncryption                    = $hasEncryption
            EncryptionEnabled                 = $Label.EncryptionEnabled
            EncryptionProtectionType          = $Label.EncryptionProtectionType
            EncryptionDoNotForward            = $Label.EncryptionDoNotForward
            EncryptionEncryptOnly             = $Label.EncryptionEncryptOnly
            EncryptionRightsDefinitions       = $Label.EncryptionRightsDefinitions
            EncryptionContentExpiredOnDateInDaysOrNever = $Label.EncryptionContentExpiredOnDateInDaysOrNever
            EncryptionOfflineAccessDays       = $Label.EncryptionOfflineAccessDays
            EncryptionPromptUser              = $Label.EncryptionPromptUser

            # Advanced settings
            AdvancedSettings         = $Label.AdvancedSettings

            # Metadata
            _SourceTenantLabelGuid   = $Label.Guid.ToString()
            _ExportedAt              = (Get-Date -Format 'o')
        }

        [PSCustomObject]$obj
    }
}

function ConvertTo-PolicyExportObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]$Policy,
        [hashtable]$GuidToLabelPath = @{},
        [hashtable]$NameToLabelPath = @{}
    )

    process {
        # Resolve label references to LabelPaths
        # Policy.Labels may contain GUIDs or Name values depending on label type
        $labelPaths = @()
        if ($Policy.Labels) {
            foreach ($ref in $Policy.Labels) {
                $refStr = $ref.ToString()
                $path = $GuidToLabelPath[$refStr]
                if (-not $path) {
                    $path = $NameToLabelPath[$refStr]
                }
                if ($path) {
                    $labelPaths += $path
                } else {
                    Write-Log "Policy '$($Policy.Name)': could not resolve label reference '$refStr'" -Level WARN
                    $labelPaths += "UNRESOLVED:$refStr"
                }
            }
        }

        $obj = [ordered]@{
            Name                     = $Policy.Name
            Comment                  = $Policy.Comment
            Enabled                  = $Policy.Enabled
            LabelPaths               = $labelPaths
            Priority                 = $Policy.Priority

            # Scoping — logged for manual remapping
            ExchangeLocation         = $Policy.ExchangeLocation
            ExchangeLocationException = $Policy.ExchangeLocationException
            ModernGroupLocation      = $Policy.ModernGroupLocation
            ModernGroupLocationException = $Policy.ModernGroupLocationException

            # Advanced policy settings
            AdvancedSettings         = $Policy.AdvancedSettings

            # Metadata
            _SourceTenantPolicyGuid  = $Policy.Guid.ToString()
            _ExportedAt              = (Get-Date -Format 'o')
        }

        [PSCustomObject]$obj
    }
}

function Build-NewLabelParameters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$LabelData,
        [switch]$SkipEncryption,
        [string]$ParentId
    )

    $params = @{
        DisplayName = $LabelData.DisplayName
        Name        = $LabelData.Name
        Tooltip     = $LabelData.Tooltip
    }

    if ($LabelData.Comment)     { $params['Comment']     = $LabelData.Comment }
    if ($LabelData.ContentType) { $params['ContentType'] = $LabelData.ContentType }
    if ($ParentId)              { $params['ParentId']     = $ParentId }

    # Content marking — Header
    if ($LabelData.HeaderEnabled) {
        $params['HeaderEnabled']   = $true
        if ($LabelData.HeaderText)      { $params['HeaderText']      = $LabelData.HeaderText }
        if ($LabelData.HeaderFontName)  { $params['HeaderFontName']  = $LabelData.HeaderFontName }
        if ($LabelData.HeaderFontSize)  { $params['HeaderFontSize']  = $LabelData.HeaderFontSize }
        if ($LabelData.HeaderFontColor) { $params['HeaderFontColor'] = $LabelData.HeaderFontColor }
        if ($LabelData.HeaderAlignment) { $params['HeaderAlignment'] = $LabelData.HeaderAlignment }
    }

    # Content marking — Footer
    if ($LabelData.FooterEnabled) {
        $params['FooterEnabled']   = $true
        if ($LabelData.FooterText)      { $params['FooterText']      = $LabelData.FooterText }
        if ($LabelData.FooterFontName)  { $params['FooterFontName']  = $LabelData.FooterFontName }
        if ($LabelData.FooterFontSize)  { $params['FooterFontSize']  = $LabelData.FooterFontSize }
        if ($LabelData.FooterFontColor) { $params['FooterFontColor'] = $LabelData.FooterFontColor }
        if ($LabelData.FooterAlignment) { $params['FooterAlignment'] = $LabelData.FooterAlignment }
    }

    # Content marking — Watermark
    if ($LabelData.WatermarkEnabled) {
        $params['WatermarkEnabled'] = $true
        if ($LabelData.WatermarkText)      { $params['WatermarkText']      = $LabelData.WatermarkText }
        if ($LabelData.WatermarkFontName)  { $params['WatermarkFontName']  = $LabelData.WatermarkFontName }
        if ($LabelData.WatermarkFontSize)  { $params['WatermarkFontSize']  = $LabelData.WatermarkFontSize }
        if ($LabelData.WatermarkFontColor) { $params['WatermarkFontColor'] = $LabelData.WatermarkFontColor }
        if ($LabelData.WatermarkLayout)    { $params['WatermarkLayout']    = $LabelData.WatermarkLayout }
    }

    # Encryption
    if ($LabelData._HasEncryption -and -not $SkipEncryption) {
        if ($LabelData.EncryptionEnabled)        { $params['EncryptionEnabled']        = $true }
        if ($LabelData.EncryptionProtectionType) { $params['EncryptionProtectionType'] = $LabelData.EncryptionProtectionType }
        if ($LabelData.EncryptionDoNotForward)   { $params['EncryptionDoNotForward']   = $LabelData.EncryptionDoNotForward }
        if ($LabelData.EncryptionEncryptOnly)    { $params['EncryptionEncryptOnly']    = $LabelData.EncryptionEncryptOnly }
        if ($LabelData.EncryptionRightsDefinitions) {
            $params['EncryptionRightsDefinitions'] = $LabelData.EncryptionRightsDefinitions
        }
        if ($LabelData.EncryptionContentExpiredOnDateInDaysOrNever) {
            $params['EncryptionContentExpiredOnDateInDaysOrNever'] = $LabelData.EncryptionContentExpiredOnDateInDaysOrNever
        }
        if ($null -ne $LabelData.EncryptionOfflineAccessDays) {
            $params['EncryptionOfflineAccessDays'] = $LabelData.EncryptionOfflineAccessDays
        }
        if ($null -ne $LabelData.EncryptionPromptUser) {
            $params['EncryptionPromptUser'] = $LabelData.EncryptionPromptUser
        }
    }
    elseif ($LabelData._HasEncryption -and $SkipEncryption) {
        Write-Log "  Skipping encryption for '$($LabelData.DisplayName)' (SkipEncryption flag set)" -Level WARN
    }

    $params
}

function Build-NewLabelPolicyParameters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$PolicyData,
        [hashtable]$LabelPathToGuid = @{}
    )

    # Resolve LabelPaths to target tenant GUIDs
    $resolvedLabels = @()
    foreach ($path in $PolicyData.LabelPaths) {
        if ($path -like 'UNRESOLVED:*') {
            Write-Log "  Policy '$($PolicyData.Name)': skipping unresolved label reference '$path'" -Level WARN
            continue
        }
        $guid = $LabelPathToGuid[$path]
        if ($guid) {
            $resolvedLabels += $guid
        } else {
            Write-Log "  Policy '$($PolicyData.Name)': label '$path' not found in target tenant, skipping" -Level WARN
        }
    }

    $params = @{
        Name   = $PolicyData.Name
        Labels = $resolvedLabels
    }

    if ($PolicyData.Comment) { $params['Comment'] = $PolicyData.Comment }

    # Log scoping info for manual remapping
    if ($PolicyData.ExchangeLocation -and $PolicyData.ExchangeLocation -ne 'All') {
        Write-Log "  Policy '$($PolicyData.Name)': ExchangeLocation scoping needs manual remapping: $($PolicyData.ExchangeLocation -join ', ')" -Level WARN
    }
    if ($PolicyData.ModernGroupLocation) {
        Write-Log "  Policy '$($PolicyData.Name)': ModernGroupLocation scoping needs manual remapping: $($PolicyData.ModernGroupLocation -join ', ')" -Level WARN
    }

    $params
}

Export-ModuleMember -Function @(
    'Write-Log'
    'Save-Log'
    'Connect-ComplianceTenant'
    'ConvertTo-LabelExportObject'
    'ConvertTo-PolicyExportObject'
    'Build-NewLabelParameters'
    'Build-NewLabelPolicyParameters'
)

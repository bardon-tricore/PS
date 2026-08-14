<#
.SYNOPSIS
    Datto RMM Component - Report the Microsoft 365 Apps (Click-to-Run) Update Channel.

.DESCRIPTION
    Determines the effective Office update channel by walking the update-related registry
    locations in order of authority, reports it in the component output, and writes it to
    a UDF.

    Authority order (first non-empty value wins), following Microsoft's published
    priority table for how Microsoft 365 Apps determines which update channel to apply:
        1. HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate  UpdatePath
        2. HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate  UpdateBranch
        3. HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate        UpdatePath
        4. HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate        UpdateBranch
        5. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 UpdateUrl
        -. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 UpdateChannel
        6. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 UnmanagedUpdateURL
        7. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 CDNBaseUrl

    UpdateUrl is the value the Office Deployment Tool writes, and is what the companion
    Set-OfficeUpdateChannel remediation component writes.

    UpdateChannel does NOT appear in Microsoft's priority table, but it is present and
    accurate on the large majority of real devices, so it is retained immediately below
    UpdateUrl: it is consulted only when UpdateUrl holds no value. Where both are set,
    UpdateUrl wins - which matters on a device that has just been remediated, because
    UpdateUrl reflects the new channel before UpdateChannel catches up.

    "UpdatePath" and the ClickToRun values hold a CDN URL, so the channel GUID is extracted
    from them. "UpdateBranch" holds a branch NAME (e.g. Current, MonthlyEnterprise) which is
    mapped to the matching GUID and reported under its friendly channel name.

    Note that "UpdatePath" may legitimately point at an on-premises share (UNC or local
    path) rather than the Microsoft CDN. In that case no GUID can be derived and the raw
    value is reported instead.

.INPUTS
    ENV:UDFNumber - UDF slot (1-300) to write the current channel to. Optional; if omitted
    or invalid the channel is still reported in the output and only the UDF write is
    skipped.

.OUTPUTS
    Datto RMM Result block:  Status=<current channel>

    UDF: HKLM:\SOFTWARE\CentraStage\Custom<UDFNumber> (REG_SZ) is set to the friendly
    channel name, or to one of "Office C2R not installed" / "Channel not set", or to the
    raw registry value when the channel cannot be resolved to a known channel.

.NOTES
    This component is report-only and always exits 0. Every outcome - including Office not
    being installed - is a valid observation rather than a failure, so nothing here raises
    a Datto alert. Filter or report on the UDF value to find devices of interest.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

#region ChannelMap

# Channel GUID -> friendly name
$ChannelMap = [ordered]@{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current Channel'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current Channel (Preview)'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise Channel'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise Channel'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise Channel (Preview)'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta Channel'
}

# "updatebranch" branch name -> channel GUID. Both the legacy branch names and the
# current-generation aliases are included, as either may be present depending on when
# and how the policy was authored.
$BranchNameMap = @{
    'current'              = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    'currentchannel'       = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    'monthly'              = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    'firstreleasecurrent'  = '64256afe-f5d9-4f86-8936-8840a6a4f5be'
    'currentpreview'       = '64256afe-f5d9-4f86-8936-8840a6a4f5be'
    'monthlypreview'       = '64256afe-f5d9-4f86-8936-8840a6a4f5be'
    'monthlyenterprise'    = '55336b82-a18d-4dd6-b5f6-9e5095c314a6'
    'deferred'             = '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
    'semiannual'           = '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
    'firstreleasedeferred' = 'b8f9b850-328d-4355-9145-c59439a0c4cf'
    'semiannualpreview'    = 'b8f9b850-328d-4355-9145-c59439a0c4cf'
    'targeted'             = 'b8f9b850-328d-4355-9145-c59439a0c4cf'
    'insiderfast'          = '5440fd1f-7ecb-4221-8110-145efaa6372f'
    'beta'                 = '5440fd1f-7ecb-4221-8110-145efaa6372f'
    'betachannel'          = '5440fd1f-7ecb-4221-8110-145efaa6372f'
    'dogfood'              = '5440fd1f-7ecb-4221-8110-145efaa6372f'
}

# Registry sources in descending order of authority.
$ChannelSources = @(
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatepath';         Kind = 'Url'    ; Label = '1. Cloud policy - UpdatePath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatebranch';       Kind = 'Branch' ; Label = '2. Cloud policy - UpdateBranch' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatepath';         Kind = 'Url'    ; Label = '3. GPO - UpdatePath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatebranch';       Kind = 'Branch' ; Label = '4. GPO - UpdateBranch' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'UpdateUrl';          Kind = 'Url'    ; Label = '5. ODT - UpdateUrl' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'UpdateChannel';      Kind = 'Url'    ; Label = '-. Observed - UpdateChannel' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'UnmanagedUpdateURL'; Kind = 'Url'    ; Label = '6. Unmanaged - UnmanagedUpdateURL' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'CDNBaseUrl';         Kind = 'Url'    ; Label = '7. Unmanaged - CDNBaseUrl' }
)

$C2RConfigPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$GuidPattern   = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

#endregion

#region Functions

function Resolve-ChannelGuid {
    <#  Extracts the lower-case channel GUID from a CDN URL, or returns $null. #>
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return $null }

    $GuidMatch = [regex]::Match($InputValue.Trim(), $GuidPattern)
    if ($GuidMatch.Success) { return $GuidMatch.Value.ToLowerInvariant() }

    return $null
}

function Resolve-BranchGuid {
    <#  Maps an "updatebranch" branch name to a channel GUID, or returns $null. #>
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return $null }

    $Normalised = ($InputValue -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    if ($BranchNameMap.ContainsKey($Normalised)) { return $BranchNameMap[$Normalised] }

    return $null
}

function Get-ChannelFriendlyName {
    param([string]$Guid)
    if ($Guid -and $ChannelMap.Contains($Guid)) { return $ChannelMap[$Guid] }
    return 'Unrecognised channel'
}

function Get-RegistryValue {
    <#  Returns the named value, or $null if the key or value is absent/empty. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $Key = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $Key) { return $null }

    $Value = [string]$Key.$Name
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    return $Value.Trim()
}

function Get-EffectiveChannel {
    <#
        Walks every source in authority order, records what each holds, and returns them
        in that order so the first entry is the effective one.
    #>
    $Found = @()

    foreach ($Source in $ChannelSources) {
        $Value = Get-RegistryValue -Path $Source.Path -Name $Source.Name
        if ($null -eq $Value) { continue }

        $Guid = if ($Source.Kind -eq 'Branch') {
            Resolve-BranchGuid -InputValue $Value
        }
        else {
            Resolve-ChannelGuid -InputValue $Value
        }

        $Found += [pscustomobject]@{
            Label = $Source.Label
            Path  = $Source.Path
            Name  = $Source.Name
            Kind  = $Source.Kind
            Value = $Value
            Guid  = $Guid
        }
    }

    return $Found
}

function Set-DattoUdf {
    <#  Writes a string to HKLM:\SOFTWARE\CentraStage as Custom<Number>. #>
    param(
        [Parameter(Mandatory)][int]$Number,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )

    $UdfPath = 'HKLM:\SOFTWARE\CentraStage'
    $UdfName = "Custom$Number"

    try {
        if (-not (Test-Path -LiteralPath $UdfPath)) {
            New-Item -Path $UdfPath -Force | Out-Null
        }
        New-ItemProperty -LiteralPath $UdfPath -Name $UdfName -Value $Value `
            -PropertyType String -Force -ErrorAction Stop | Out-Null
        Write-Host "UDF write        : $UdfName = $Value"
    }
    catch {
        Write-Host "WARNING: Failed to write $UdfName - $($_.Exception.Message)"
    }
}

function Write-DattoResult {
    <#  Flushes the UDF, emits the Result block and exits. Always exits 0. #>
    param(
        [Parameter(Mandatory)][string]$Status
    )

    # Flush the UDF on every exit path so a stale value is never left behind.
    if ($Script:UdfNumber) {
        Set-DattoUdf -Number $Script:UdfNumber -Value $Script:UdfValue
    }

    Write-Host '<-End Diagnostic->'
    Write-Host '<-Start Result->'
    Write-Host "Status=$Status"
    Write-Host '<-End Result->'
    exit 0
}

#endregion

#region Main

Write-Host '<-Start Diagnostic->'
Write-Host "Office Update Channel Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Device: $env:COMPUTERNAME"
Write-Host '--------------------------------------------------------------------'

#region UdfNumber

# Default the reported value; overwritten as soon as the real channel is known.
$Script:UdfValue  = 'Unknown'
$Script:UdfNumber = $null

$UdfRaw    = $env:UDFNumber
$ParsedUdf = 0

if ([string]::IsNullOrWhiteSpace($UdfRaw)) {
    Write-Host 'UDF              : UDFNumber variable not supplied - UDF write skipped'
}
elseif (-not [int]::TryParse($UdfRaw.Trim(), [ref]$ParsedUdf)) {
    Write-Host "WARNING: UDFNumber '$UdfRaw' is not a whole number - UDF write skipped"
}
elseif ($ParsedUdf -lt 1 -or $ParsedUdf -gt 300) {
    Write-Host "WARNING: UDFNumber $ParsedUdf is outside the valid range of 1-300 - UDF write skipped"
}
else {
    $Script:UdfNumber = $ParsedUdf
    Write-Host "UDF target       : Custom$($Script:UdfNumber)"
}

Write-Host ''

#endregion

#region ReadRegistry

# Office presence is judged on the ClickToRun key, not on the policy keys - a policy can
# be pushed to a device that has no Office installed at all.
if (-not (Test-Path -LiteralPath $C2RConfigPath)) {
    Write-Host 'Office Click-to-Run Configuration key not found.'
    Write-Host 'Office C2R (Microsoft 365 Apps) does not appear to be installed on this device.'
    $Script:UdfValue = 'Office C2R not installed'
    Write-DattoResult -Status 'Office C2R not installed'
}

$C2RConfig = Get-ItemProperty -LiteralPath $C2RConfigPath
Write-Host "Office version   : $(if ($C2RConfig.VersionToReport)   { $C2RConfig.VersionToReport }   else { '<unknown>' })"
Write-Host "Products         : $(if ($C2RConfig.ProductReleaseIds) { $C2RConfig.ProductReleaseIds } else { '<unknown>' })"
Write-Host ''

$FoundSources = Get-EffectiveChannel

Write-Host 'Update channel sources, in order of authority:'
foreach ($Source in $ChannelSources) {
    $Hit = $FoundSources | Where-Object { $_.Label -eq $Source.Label } | Select-Object -First 1
    if ($Hit) {
        $Resolved = if ($Hit.Guid) { Get-ChannelFriendlyName -Guid $Hit.Guid } else { 'no channel GUID derivable' }
        Write-Host "  [SET] $($Source.Label) = $($Hit.Value)  ->  $Resolved"
    }
    else {
        Write-Host "  [ - ] $($Source.Label)"
    }
}
Write-Host ''

if ($FoundSources.Count -eq 0) {
    Write-Host 'None of the update channel locations hold a value.'
    $Script:UdfValue = 'Channel not set'
    Write-DattoResult -Status 'Channel not set'
}

#endregion

#region Report

# Highest authority wins - Get-EffectiveChannel preserves the source order.
$Effective  = $FoundSources[0]
$ActualGuid = $Effective.Guid
$ActualName = if ($ActualGuid) { Get-ChannelFriendlyName -Guid $ActualGuid } else { 'Undetermined' }

Write-Host "Effective source : $($Effective.Label)"
Write-Host "Effective value  : $($Effective.Value)"
Write-Host "Current channel  : $ActualName"
Write-Host "Current GUID     : $(if ($ActualGuid) { $ActualGuid } else { '<none derivable>' })"

if ($FoundSources.Count -gt 1) {
    Write-Host ''
    Write-Host "NOTE: $($FoundSources.Count) sources hold a value. The one above takes precedence;"
    Write-Host '      lower-authority values are listed for reference only.'
}

Write-Host '--------------------------------------------------------------------'

if ($ActualGuid -and $ChannelMap.Contains($ActualGuid)) {
    # Known channel - report the friendly name.
    $Script:UdfValue = $ActualName
    Write-DattoResult -Status $ActualName
}
else {
    # Unrecognised GUID, or an on-premises update path with no GUID to derive. Report the
    # raw value so the device can still be identified from the UDF.
    Write-Host 'The effective value does not resolve to a known update channel.'
    Write-Host 'This is expected where updates are served from an on-premises share.'
    $Script:UdfValue = $Effective.Value
    Write-DattoResult -Status $Effective.Value
}

#endregion

#endregion
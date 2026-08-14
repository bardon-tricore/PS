<#
.SYNOPSIS
    Datto RMM Component - Audit Microsoft 365 Apps (Click-to-Run) Update Channel.

.DESCRIPTION
    Determines the effective Office update channel by walking the update-related registry
    locations in order of authority, and compares the winning value against the expected
    channel supplied in the Datto RMM variable "AuditUpdateChannel".

    Authority order (first non-empty value wins):
        1. HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate  updatepath
        2. HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate  updatebranch
        3. HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate        updatepath
        4. HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate        updatebranch
        5. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 UpdateChannel
        6. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 UnmanagedUpdateURL
        7. HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration                 CDNBaseUrl

    "updatepath" and the ClickToRun values hold a CDN URL, so the channel GUID is extracted
    from them. "updatebranch" holds a branch NAME (e.g. Current, MonthlyEnterprise) which is
    mapped to the matching GUID. Comparison is always done on the GUID, so http/https,
    trailing slashes and casing differences are all tolerated.

    Note that "updatepath" may legitimately point at an on-premises share (UNC or local
    path) rather than the Microsoft CDN. In that case no GUID can be derived, the channel
    is reported as undetermined, and the script exits gracefully rather than alerting.

.INPUTS
    ENV:AuditUpdateChannel - Expected channel as a full CDN URL, supplied by a Datto RMM
    selection variable, e.g. http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60

    ENV:UDFNumber - UDF slot (1-300) to write the current channel to. Optional; if omitted
    or invalid the audit still runs and only the UDF write is skipped.

.OUTPUTS
    Datto RMM Result block:  Status=<SUCCESS|FAILURE|...> : <detail>

    UDF: HKLM:\SOFTWARE\CentraStage\Custom<UDFNumber> (REG_SZ) is set to the friendly
    channel name, or to one of "Office C2R not installed" / "Channel not set", or to the
    raw registry value when the channel cannot be resolved to a known channel.

.NOTES
    Exit codes (Datto RMM only supports 0 and 1):
        1 = ONLY when the effective channel does not match the expected channel.
        0 = Everything else - a match, or any condition where the audit could not be
            performed (Office C2R not installed, no channel value present, channel could
            not be determined, or the AuditUpdateChannel variable was not supplied). These
            exit gracefully so they do not raise a false channel-drift alert; the Result
            line and the diagnostic output identify which condition occurred.
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
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatepath';         Kind = 'Url'    ; Label = 'Cloud policy - updatepath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatebranch';       Kind = 'Branch' ; Label = 'Cloud policy - updatebranch' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatepath';         Kind = 'Url'    ; Label = 'GPO - updatepath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatebranch';       Kind = 'Branch' ; Label = 'GPO - updatebranch' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'UpdateChannel';      Kind = 'Url'    ; Label = 'ClickToRun - UpdateChannel' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'UnmanagedUpdateURL'; Kind = 'Url'    ; Label = 'ClickToRun - UnmanagedUpdateURL' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration';                Name = 'CDNBaseUrl';         Kind = 'Url'    ; Label = 'ClickToRun - CDNBaseUrl' }
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
        Walks every source in authority order, records what each holds, and returns the
        highest-authority source that carries a value.
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
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][ValidateSet(0, 1)][int]$ExitCode
    )

    # Flush the UDF on every exit path so a stale value is never left behind.
    if ($Script:UdfNumber) {
        Set-DattoUdf -Number $Script:UdfNumber -Value $Script:UdfValue
    }

    Write-Host '<-End Diagnostic->'
    Write-Host '<-Start Result->'
    Write-Host "Status=$Status"
    Write-Host '<-End Result->'
    exit $ExitCode
}

#endregion

#region Main

Write-Host '<-Start Diagnostic->'
Write-Host "Office Update Channel Audit - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
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

#endregion

#region ExpectedChannel

$ExpectedRaw = $env:AuditUpdateChannel

if ([string]::IsNullOrWhiteSpace($ExpectedRaw)) {
    Write-Host 'ERROR: The component variable "AuditUpdateChannel" was not supplied.'
    Write-Host 'Exiting gracefully - check the component/site variable configuration.'
    # Nothing was learned about the device, so leave any existing UDF value intact.
    $Script:UdfNumber = $null
    Write-DattoResult -Status 'ERROR: AuditUpdateChannel variable not supplied - audit not performed' -ExitCode 0
}

$ExpectedGuid = Resolve-ChannelGuid -InputValue $ExpectedRaw
$ExpectedName = Get-ChannelFriendlyName -Guid $ExpectedGuid

Write-Host "Expected channel : $ExpectedName"
Write-Host "Expected GUID    : $ExpectedGuid"
Write-Host "Variable value   : $ExpectedRaw"
Write-Host ''

#endregion

#region ReadRegistry

# Office presence is judged on the ClickToRun key, not on the policy keys - a policy can
# be pushed to a device that has no Office installed at all.
if (-not (Test-Path -LiteralPath $C2RConfigPath)) {
    Write-Host 'Office Click-to-Run Configuration key not found.'
    Write-Host 'Office C2R (Microsoft 365 Apps) does not appear to be installed on this device.'
    Write-Host 'Exiting gracefully - there is no channel to audit on this device.'
    $Script:UdfValue = 'Office C2R not installed'
    Write-DattoResult -Status 'NOT APPLICABLE: Office Click-to-Run is not installed - no channel to audit' -ExitCode 0
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
    Write-Host 'Exiting gracefully - there is no channel value to compare.'
    $Script:UdfValue = 'Channel not set'
    Write-DattoResult -Status 'NOT APPLICABLE: No update channel value present - nothing to compare' -ExitCode 0
}

# Highest authority wins - Get-EffectiveChannel preserves the source order.
$Effective     = $FoundSources[0]
$ActualGuid    = $Effective.Guid
$ActualName    = if ($ActualGuid) { Get-ChannelFriendlyName -Guid $ActualGuid } else { 'Undetermined' }

Write-Host "Effective source : $($Effective.Label)"
Write-Host "Effective value  : $($Effective.Value)"
Write-Host "Current channel  : $ActualName"
Write-Host "Current GUID     : $(if ($ActualGuid) { $ActualGuid } else { '<none derivable>' })"

if ($FoundSources.Count -gt 1) {
    Write-Host ''
    Write-Host "NOTE: $($FoundSources.Count) sources hold a value. The one above takes precedence;"
    Write-Host '      lower-authority values are listed for reference only.'
}

# Report the friendly name where known, otherwise the raw value so it can be identified.
$Script:UdfValue = if ($ActualGuid -and $ChannelMap.Contains($ActualGuid)) { $ActualName } else { $Effective.Value }

#endregion

#region Compare

Write-Host '--------------------------------------------------------------------'

if (-not $ActualGuid) {
    # e.g. updatepath pointing at an on-premises share - a valid configuration, but the
    # channel cannot be determined from the registry alone.
    Write-Host 'The effective value does not resolve to a known update channel.'
    Write-Host 'This is expected where updates are served from an on-premises share.'
    Write-Host 'Exiting gracefully - unable to confirm the channel.'
    Write-DattoResult -Status "UNDETERMINED: Effective source $($Effective.Label) is set to '$($Effective.Value)' - channel could not be resolved" -ExitCode 0
}

if ($ActualGuid -eq $ExpectedGuid) {
    Write-Host "RESULT: MATCH - device is on $ExpectedName."
    Write-DattoResult -Status "SUCCESS: Office update channel is $ExpectedName ($ExpectedGuid)" -ExitCode 0
}
else {
    Write-Host "RESULT: MISMATCH - expected '$ExpectedName' but found '$ActualName'."
    Write-Host "        Set by: $($Effective.Label)"
    $Detail = "$ActualName ($ActualGuid) via $($Effective.Label)"
    Write-DattoResult -Status "FAILURE: Expected $ExpectedName but found $Detail" -ExitCode 1
}

#endregion

#endregion
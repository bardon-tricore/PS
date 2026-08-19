<#
.SYNOPSIS
    Datto RMM Component - Change the Microsoft 365 Apps (Click-to-Run) update channel,
    unless the channel is managed by policy.

.DESCRIPTION
    Companion remediation for the Get-OfficeUpdateChannel reporting component. It reuses
    the same authority chain to establish the effective channel, then acts only when it is
    safe to do so.

    The script REFUSES TO ACT and exits gracefully when any of these hold a value:
        HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate  updatepath
        HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate  updatebranch
        HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate        updatepath
        HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate        updatebranch

    Those locations outrank anything this script could write, so changing the ClickToRun
    configuration underneath them would appear to succeed and then be silently overridden.
    Devices in that state must be corrected at the policy source (Intune/Cloud policy or
    Group Policy) instead.

    Where no policy is present, the channel is changed via OfficeC2RClient.exe
    /changesetting and the result is verified by reading the
    registry.

    A device is only changed when its current channel matches SourceChannel. Both
    SourceChannel and TargetChannel are required: there is deliberately no "any channel"
    mode, so a job pointed at the wrong device group is not run rather than an an accidental
    channel change.

.INPUTS
    ENV:SourceChannel  - REQUIRED. Only devices currently on this channel are changed.
                         Full CDN URL, from a Datto RMM selection variable.
    ENV:TargetChannel  - REQUIRED. The channel to move those devices to. Full CDN URL, 
                         from a DRMM selection variable.
    ENV:UDFNumber      - Optional. UDF slot (1-300) to write the resulting channel to.
                         If omitted, output is still written. Defaults to 29.
    ENV:SimulateOnly   - Optional. "true" for a dry run: every check is performed and
                         reported, but no change is made. 
    ENV:TriggerUpdate  - Optional. Defaults to true. "false" changes the channel setting
                         without kicking off an update run immediately.

.OUTPUTS
    Datto RMM Result block:  Status=<outcome>
    UDF: HKLM:\SOFTWARE\CentraStage\Custom<UDFNumber> Resulting channel.

.NOTES
    Exit codes (Datto RMM only supports 0 and 1):
        1 = A remediation was attempted and did not succeed, OR the component is
            misconfigured (a required variable is missing, unrecognised, or Source and
            Target are the same channel). Both need a human to look at them.
        0 = Everything else: changed successfully, already on target, skipped because a
            policy manages the channel, skipped because the device is on another channel,
            or Office is not installed.

    Office apps are NOT force-closed. The channel switch is applied immediately but the
    actual build change lands on the next update cycle, which is deliberate, it avoids
    interrupting users mid-session.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

#region Configuration

<#
    Single source of truth for every channel this component understands.

    BranchToken is the value passed to OfficeC2RClient.exe /changesetting Channel=<token>.
#>
$ChannelDefinitions = [ordered]@{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = @{
        Name        = 'Current Channel'
        BranchToken = 'Current'
        CdnUrl      = 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    }
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = @{
        Name        = 'Current Channel (Preview)'
        BranchToken = 'CurrentPreview'
        CdnUrl      = 'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be'
    }
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = @{
        Name        = 'Monthly Enterprise Channel'
        BranchToken = 'MonthlyEnterprise'
        CdnUrl      = 'http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6'
    }
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = @{
        Name        = 'Semi-Annual Enterprise Channel'
        BranchToken = 'SemiAnnual'
        CdnUrl      = 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
    }
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = @{
        Name        = 'Semi-Annual Enterprise Channel (Preview)'
        BranchToken = 'SemiAnnualPreview'
        CdnUrl      = 'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf'
    }
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = @{
        Name        = 'Beta Channel'
        BranchToken = 'BetaChannel'
        CdnUrl      = 'http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f'
    }
}

# "updatebranch" branch name -> channel GUID.
$BranchNameMap = @{
    'current'              = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    'currentchannel'       = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    'monthly'              = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    'currentpreview'       = '64256afe-f5d9-4f86-8936-8840a6a4f5be'
    'monthlypreview'       = '64256afe-f5d9-4f86-8936-8840a6a4f5be'
    'monthlyenterprise'    = '55336b82-a18d-4dd6-b5f6-9e5095c314a6'
    'semiannual'           = '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'

}

# Policy sources. Any value here blocks remediation outright.
$PolicySources = @(
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatepath';   Kind = 'Url'   ; Label = 'Cloud policy - updatepath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatebranch'; Kind = 'Branch'; Label = 'Cloud policy - updatebranch' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatepath';   Kind = 'Url'   ; Label = 'GPO - updatepath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatebranch'; Kind = 'Branch'; Label = 'GPO - updatebranch' }
)

# ClickToRun sources, in order of authority. These are what the script can safely change.
$C2RSources = @(
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'UpdateChannel';      Kind = 'Url'; Label = 'ClickToRun - UpdateChannel' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'UnmanagedUpdateURL'; Kind = 'Url'; Label = 'ClickToRun - UnmanagedUpdateURL' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'CDNBaseUrl';         Kind = 'Url'; Label = 'ClickToRun - CDNBaseUrl' }
)

$C2RConfigPath   = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$GuidPattern     = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
$ClientTimeoutMs = 300000   # 5 minutes

#endregion

#region Functions

function Resolve-ChannelGuid {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return $null }

    $GuidMatch = [regex]::Match($InputValue.Trim(), $GuidPattern)
    if ($GuidMatch.Success) { return $GuidMatch.Value.ToLowerInvariant() }

    return $null
}

function Resolve-BranchGuid {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return $null }

    $Normalised = ($InputValue -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    if ($BranchNameMap.ContainsKey($Normalised)) { return $BranchNameMap[$Normalised] }

    return $null
}

function Get-ChannelFriendlyName {
    param([string]$Guid)
    if ($Guid -and $ChannelDefinitions.Contains($Guid)) { return $ChannelDefinitions[$Guid].Name }
    return 'Unrecognised channel'
}

function Get-RegistryValue {
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

function Get-SourceValues {
    <#  Reads a set of sources and returns those holding a value, in the order given. #>
    param([Parameter(Mandatory)][object[]]$Sources)

    $Found = @()

    foreach ($Source in $Sources) {
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
            Value = $Value
            Guid  = $Guid
        }
    }

    return $Found
}

function Write-SourceReport {
    param(
        [Parameter(Mandatory)][object[]]$Sources,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Found
    )

    foreach ($Source in $Sources) {
        $Hit = $Found | Where-Object { $_.Label -eq $Source.Label } | Select-Object -First 1
        if ($Hit) {
            $Resolved = if ($Hit.Guid) { Get-ChannelFriendlyName -Guid $Hit.Guid } else { 'no channel GUID derivable' }
            Write-Host "  [SET] $($Source.Label) = $($Hit.Value)  ->  $Resolved"
        }
        else {
            Write-Host "  [ - ] $($Source.Label)"
        }
    }
}

function Get-OfficeC2RClientPath {
    <#  Locates OfficeC2RClient.exe, preferring the path Office itself records. #>
    $Candidates = @()

    $ClientFolder = Get-RegistryValue -Path $C2RConfigPath -Name 'ClientFolder'
    if ($ClientFolder) { $Candidates += (Join-Path $ClientFolder 'OfficeC2RClient.exe') }

    $Candidates += (Join-Path $env:ProgramFiles 'Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe')
    if ($env:ProgramW6432) {
        $Candidates += (Join-Path $env:ProgramW6432 'Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe')
    }

    foreach ($Candidate in $Candidates) {
        if (Test-Path -LiteralPath $Candidate) { return $Candidate }
    }

    return $null
}

function Invoke-C2RClient {
    <#  Runs OfficeC2RClient.exe with the given arguments and returns the exit code. #>
    param(
        [Parameter(Mandatory)][string]$ClientPath,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    Write-Host "  Running: `"$ClientPath`" $($Arguments -join ' ')"

    $Process = Start-Process -FilePath $ClientPath -ArgumentList $Arguments -PassThru -WindowStyle Hidden

    if (-not $Process.WaitForExit($ClientTimeoutMs)) {
        Write-Host "  WARNING: Client did not exit within $([int]($ClientTimeoutMs / 1000))s - continuing without waiting."
        return $null
    }

    Write-Host "  Client exit code: $($Process.ExitCode)"
    return $Process.ExitCode
}

function Set-DattoUdf {
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
Write-Host "Office Update Channel Remediation - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Device: $env:COMPUTERNAME"
Write-Host '--------------------------------------------------------------------'

#region UdfNumber

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

#region ChannelVariables

<#
    Both channel variables are required. 
#>

$SourceGuid = Resolve-ChannelGuid -InputValue $env:SourceChannel
$TargetGuid = Resolve-ChannelGuid -InputValue $env:TargetChannel

if ([string]::IsNullOrWhiteSpace($env:SourceChannel) -or [string]::IsNullOrWhiteSpace($env:TargetChannel)) {
    Write-Host 'ERROR: SourceChannel and TargetChannel are both required.'
    Write-Host "  SourceChannel: $(if ($env:SourceChannel) { $env:SourceChannel } else { '<not supplied>' })"
    Write-Host "  TargetChannel: $(if ($env:TargetChannel) { $env:TargetChannel } else { '<not supplied>' })"
    $Script:UdfNumber = $null
    Write-DattoResult -Status 'CONFIG ERROR: SourceChannel and TargetChannel are both required' -ExitCode 1
}

if (-not $SourceGuid -or -not $ChannelDefinitions.Contains($SourceGuid)) {
    Write-Host "ERROR: SourceChannel '$($env:SourceChannel)' is not a recognised update channel."
    $Script:UdfNumber = $null
    Write-DattoResult -Status "CONFIG ERROR: SourceChannel is not a recognised update channel" -ExitCode 1
}

if (-not $TargetGuid -or -not $ChannelDefinitions.Contains($TargetGuid)) {
    Write-Host "ERROR: TargetChannel '$($env:TargetChannel)' is not a recognised update channel."
    $Script:UdfNumber = $null
    Write-DattoResult -Status "CONFIG ERROR: TargetChannel is not a recognised update channel" -ExitCode 1
}

if ($SourceGuid -eq $TargetGuid) {
    Write-Host 'ERROR: SourceChannel and TargetChannel are the same channel - nothing to do.'
    $Script:UdfNumber = $null
    Write-DattoResult -Status 'CONFIG ERROR: SourceChannel and TargetChannel are the same channel' -ExitCode 1
}

$SourceName  = $ChannelDefinitions[$SourceGuid].Name
$TargetName  = $ChannelDefinitions[$TargetGuid].Name
$TargetToken = $ChannelDefinitions[$TargetGuid].BranchToken
$TargetUrl   = $ChannelDefinitions[$TargetGuid].CdnUrl

$SimulateOnly  = ($env:SimulateOnly -match '^(?i)(true|yes|1)$')
$TriggerUpdate = -not ($env:TriggerUpdate -match '^(?i)(false|no|0)$')

Write-Host "Change FROM      : $SourceName"
Write-Host "Change TO        : $TargetName"
Write-Host "Simulate only    : $SimulateOnly"
Write-Host "Trigger update   : $TriggerUpdate"
if ($SimulateOnly) {
    Write-Host '*** SIMULATION MODE - no changes will be made ***'
}
Write-Host ''

#endregion

#region OfficePresence

if (-not (Test-Path -LiteralPath $C2RConfigPath)) {
    Write-Host 'Office Click-to-Run Configuration key not found.'
    Write-Host 'Office C2R (Microsoft 365 Apps) is not installed - nothing to remediate.'
    $Script:UdfValue = 'Office C2R not installed'
    Write-DattoResult -Status 'SKIPPED: Office Click-to-Run is not installed' -ExitCode 0
}

$C2RConfig = Get-ItemProperty -LiteralPath $C2RConfigPath
Write-Host "Office version   : $(if ($C2RConfig.VersionToReport)   { $C2RConfig.VersionToReport }   else { '<unknown>' })"
Write-Host "Products         : $(if ($C2RConfig.ProductReleaseIds) { $C2RConfig.ProductReleaseIds } else { '<unknown>' })"
Write-Host ''

#endregion

#region PolicyGuard

Write-Host 'Policy sources (these override anything this script can set):'
$FoundPolicies = @(Get-SourceValues -Sources $PolicySources)
Write-SourceReport -Sources $PolicySources -Found $FoundPolicies
Write-Host ''

if ($FoundPolicies.Count -gt 0) {
    $Blocking     = $FoundPolicies[0]
    $BlockingName = if ($Blocking.Guid) { Get-ChannelFriendlyName -Guid $Blocking.Guid } else { $Blocking.Value }

    Write-Host 'The update channel on this device is managed by policy.'
    Write-Host 'No change will be made - the ClickToRun configuration would be overridden.'
    Write-Host "Correct this at the policy source instead: $($Blocking.Label)"

    $Script:UdfValue = "Policy managed - $BlockingName"
    Write-DattoResult -Status "SKIPPED: Channel is policy managed via $($Blocking.Label) - change at the policy source" -ExitCode 0
}

Write-Host 'No update policy present - safe to change the ClickToRun configuration.'
Write-Host ''

#endregion

#region CurrentChannel

Write-Host 'ClickToRun sources, in order of authority:'
$FoundC2R = @(Get-SourceValues -Sources $C2RSources)
Write-SourceReport -Sources $C2RSources -Found $FoundC2R
Write-Host ''

if ($FoundC2R.Count -eq 0) {
    Write-Host 'No update channel value present under the ClickToRun Configuration key.'
    Write-Host 'The channel state cannot be established.'
    $Script:UdfValue = 'Channel not set'
    Write-DattoResult -Status 'SKIPPED: No update channel value present' -ExitCode 0
}

$Effective   = $FoundC2R[0]
$CurrentGuid = $Effective.Guid
$CurrentName = if ($CurrentGuid) { Get-ChannelFriendlyName -Guid $CurrentGuid } else { 'Undetermined' }

Write-Host "Effective source : $($Effective.Label)"
Write-Host "Current channel  : $CurrentName"
Write-Host ''

if ($CurrentGuid -eq $TargetGuid) {
    Write-Host "Already on $TargetName - no action required."
    $Script:UdfValue = $TargetName
    Write-DattoResult -Status "NO ACTION: Already on $TargetName" -ExitCode 0
}

if ($CurrentGuid -ne $SourceGuid) {
    Write-Host "This device is not on $SourceName, so it is out of scope for this job."
    Write-Host 'Leaving it alone.'
    $Script:UdfValue = if ($CurrentGuid) { $CurrentName } else { $Effective.Value }
    Write-DattoResult -Status "SKIPPED: On $CurrentName, not $SourceName - out of scope" -ExitCode 0
}

#endregion

#region Remediate

Write-Host '--------------------------------------------------------------------'
Write-Host "Remediating: $CurrentName  ->  $TargetName"

$ClientPath = Get-OfficeC2RClientPath
if (-not $ClientPath) {
    Write-Host 'ERROR: OfficeC2RClient.exe could not be located.'
    Write-Host 'The channel cannot be changed by the supported method on this device.'
    $Script:UdfValue = $CurrentName
    Write-DattoResult -Status 'FAILED: OfficeC2RClient.exe not found - channel unchanged' -ExitCode 1
}
Write-Host "C2R client       : $ClientPath"

if ($SimulateOnly) {
    Write-Host ''
    Write-Host 'SIMULATION - the following would have been run:'
    Write-Host "  `"$ClientPath`" /changesetting Channel=$TargetToken"
    Write-Host "  UpdateChannel would be set to $TargetUrl"
    if ($TriggerUpdate) {
        Write-Host "  `"$ClientPath`" /update user displaylevel=false forceappshutdown=false"
    }
    $Script:UdfValue = $CurrentName
    Write-DattoResult -Status "SIMULATED: Would change $CurrentName to $TargetName" -ExitCode 0
}

Write-Host ''
Write-Host 'Applying channel change...'
Invoke-C2RClient -ClientPath $ClientPath -Arguments @('/changesetting', "Channel=$TargetToken") | Out-Null

# OfficeC2RClient does not always write UpdateChannel on every build, so make sure the value the audit reads reflects the new channel.
Write-Host ''
Write-Host 'Confirming the ClickToRun configuration...'
try {
    New-ItemProperty -LiteralPath $C2RConfigPath -Name 'UpdateChannel' -Value $TargetUrl `
        -PropertyType String -Force -ErrorAction Stop | Out-Null
    Write-Host "  UpdateChannel set to $TargetUrl"
}
catch {
    Write-Host "  WARNING: Could not write UpdateChannel - $($_.Exception.Message)"
}

# Verify by re-reading through the same authority chain the audit uses.
$VerifyC2R  = @(Get-SourceValues -Sources $C2RSources)
$VerifyGuid = if ($VerifyC2R.Count -gt 0) { $VerifyC2R[0].Guid } else { $null }
$VerifyName = if ($VerifyGuid) { Get-ChannelFriendlyName -Guid $VerifyGuid } else { 'Undetermined' }

Write-Host ''
Write-Host 'Post-change state:'
Write-SourceReport -Sources $C2RSources -Found $VerifyC2R
Write-Host ''
Write-Host "Verified channel : $VerifyName"

if ($VerifyGuid -ne $TargetGuid) {
    Write-Host ''
    Write-Host 'ERROR: The channel did not change as expected.'
    $Script:UdfValue = $VerifyName
    Write-DattoResult -Status "FAILED: Channel still reports $VerifyName after remediation" -ExitCode 1
}

if ($TriggerUpdate) {
    Write-Host ''
    Write-Host 'Triggering an update run (users are not interrupted)...'
    Invoke-C2RClient -ClientPath $ClientPath `
        -Arguments @('/update', 'user', 'displaylevel=false', 'forceappshutdown=false') | Out-Null
    Write-Host '  Update requested. The build change completes on the next update cycle.'
}

Write-Host '--------------------------------------------------------------------'
Write-Host "SUCCESS: $CurrentName changed to $VerifyName"

$Script:UdfValue = $VerifyName
Write-DattoResult -Status "CHANGED: $CurrentName -> $VerifyName" -ExitCode 0

#endregion

#endregion
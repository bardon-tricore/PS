<#
.SYNOPSIS
    Datto RMM Component - Change the Microsoft 365 Apps (Click-to-Run) update channel by
    the ODT-equivalent method, unless the channel is managed by policy.

.DESCRIPTION
    Companion remediation for the Get-OfficeUpdateChannel reporting component.

    METHOD
    ------
    The channel is changed by writing UpdateUrl under the ClickToRun Configuration key -
    the same value the Office Deployment Tool injects when you run
    "setup.exe /configure config.xml" - and then running the "Office Automatic Updates 2.0"
    scheduled task, which is what actually detects the new setting and reassigns the
    channel.

    OfficeC2RClient.exe /changesetting is deliberately NOT used. It does not appear in
    Microsoft's list of supported channel change methods, and there are multiple reports of
    it having no effect on current builds.

    AUTHORITY (Microsoft's published priority order)
    -----------------------------------------------
        1. Cloud Update  UpdatePath          HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate
        2. Cloud Update  UpdateBranch        HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate
        3. Policy        UpdatePath          HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate
        4. Policy        UpdateBranch        HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate
        5. ODT           UpdateUrl           HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration   <-- written here
        6. Unmanaged     UnmanagedUpdateURL  HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
        7. Unmanaged     CDNBaseUrl          HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration

    Writing at position 5 leaves policy precedence intact: a real Group Policy or Intune
    profile still wins, as it should. The script refuses to act at all when positions 1-4
    hold a value, because the change would be silently overridden.

    TIMING - IMPORTANT
    ------------------
    A channel change can take up to 24 hours to apply, and Office only reports the new
    channel after a build from that channel installs. This component therefore CANNOT
    confirm the channel actually changed; it confirms only that the setting was written and
    the update task was started. Use the Get-OfficeUpdateChannel reporting component a day
    or two later to confirm the outcome across the fleet.

.INPUTS
    ENV:SourceChannel  - REQUIRED. Only devices currently on this channel are changed.
                         Full CDN URL, from a Datto RMM selection variable.
    ENV:TargetChannel  - REQUIRED. The channel to move those devices to. Full CDN URL.
    ENV:UDFNumber      - Optional. UDF slot (1-300) to write the outcome to.
    ENV:SimulateOnly   - Optional. "true" for a dry run: every check is performed and
                         reported, but nothing is written and no task is started.
    ENV:TriggerUpdate  - Optional. Defaults to true. Runs the "Office Automatic Updates
                         2.0" scheduled task immediately. Set "false" to leave it for the
                         task's own schedule.

.OUTPUTS
    Datto RMM Result block:  Status=<outcome>
    UDF: HKLM:\SOFTWARE\CentraStage\Custom<UDFNumber> (REG_SZ) - outcome.

.NOTES
    Exit codes (Datto RMM only supports 0 and 1):
        1 = The setting could not be written, the write did not read back, the update task
            is missing, the script hit an unexpected error, or the component is
            misconfigured. All need a human.
        0 = Everything else: setting written successfully, already on target, skipped
            because a policy manages the channel, skipped because the device is on another
            channel, or Office is not installed.

    Related: from the Version 2606 release (July 2026) Microsoft unified Semi-Annual
    Enterprise Channel and Monthly Enterprise Channel. Devices on SAEC already receive the
    MEC update experience. This script reports the installed build so you can see which
    devices are already past 2606.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

#region Configuration

<#
    Single source of truth for every channel this component understands.
    CdnUrl is what gets written to UpdateUrl.
#>
$ChannelDefinitions = [ordered]@{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = @{
        Name   = 'Current Channel'
        CdnUrl = 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    }
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = @{
        Name   = 'Current Channel (Preview)'
        CdnUrl = 'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be'
    }
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = @{
        Name   = 'Monthly Enterprise Channel'
        CdnUrl = 'http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6'
    }
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = @{
        Name   = 'Semi-Annual Enterprise Channel'
        CdnUrl = 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
    }
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = @{
        Name   = 'Semi-Annual Enterprise Channel (Preview)'
        CdnUrl = 'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf'
    }
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = @{
        Name   = 'Beta Channel'
        CdnUrl = 'http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f'
    }
}

# "UpdateBranch" branch name -> channel GUID, for reading the policy keys.
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

# Priority 1-4. Any value here blocks remediation outright.
$PolicySources = @(
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatepath';   Kind = 'Url'   ; Label = '1. Cloud policy - UpdatePath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\common\officeupdate'; Name = 'updatebranch'; Kind = 'Branch'; Label = '2. Cloud policy - UpdateBranch' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatepath';   Kind = 'Url'   ; Label = '3. GPO - UpdatePath' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate';       Name = 'updatebranch'; Kind = 'Branch'; Label = '4. GPO - UpdateBranch' }
)

<#
    Priority 5-7 per Microsoft's table, plus UpdateChannel.

    UpdateChannel does NOT appear in Microsoft's priority table, but it is present and
    accurate on the large majority of real devices, so it sits immediately below UpdateUrl
    and is consulted only when UpdateUrl holds no value. This ordering matches the
    Get-OfficeUpdateChannel reporting component, so the two never disagree about a device.
    UpdateChannel is only ever read, never written.
#>
$C2RSources = @(
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'UpdateUrl';          Kind = 'Url'; Label = '5. ODT - UpdateUrl' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'UpdateChannel';      Kind = 'Url'; Label = '-. Observed - UpdateChannel' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'UnmanagedUpdateURL'; Kind = 'Url'; Label = '6. Unmanaged - UnmanagedUpdateURL' }
    [pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'CDNBaseUrl';         Kind = 'Url'; Label = '7. Unmanaged - CDNBaseUrl' }
)

$C2RConfigPath  = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$WriteValueName = 'UpdateUrl'
$GuidPattern    = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# The task that detects the new setting and reassigns the channel.
$UpdateTaskPath = '\Microsoft\Office\'
$UpdateTaskName = 'Office Automatic Updates 2.0'

# Build at which SAEC/MEC unification took effect (Version 2606). Informational only.
$UnifiedBuild      = 20131
$UnifiedRevision   = 20000

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

function Test-UnifiedBuild {
    <#
        Reports whether the installed build is at or past Version 2606, from which SAEC
        devices already receive the Monthly Enterprise Channel experience. Informational.
    #>
    param([string]$VersionToReport)

    if ([string]::IsNullOrWhiteSpace($VersionToReport)) { return $null }

    $Parts = $VersionToReport.Split('.')
    if ($Parts.Count -lt 4) { return $null }

    $Build    = 0
    $Revision = 0
    if (-not [int]::TryParse($Parts[2], [ref]$Build))    { return $null }
    if (-not [int]::TryParse($Parts[3], [ref]$Revision)) { return $null }

    if ($Build -gt $UnifiedBuild) { return $true }
    if ($Build -eq $UnifiedBuild -and $Revision -gt $UnifiedRevision) { return $true }
    return $false
}

function Start-OfficeUpdateTask {
    <#
        Starts the Office Automatic Updates 2.0 task. Returns an object describing what
        was found and whether the start succeeded.
    #>
    $Task = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue

    if (-not $Task) {
        return [pscustomobject]@{ Exists = $false; Enabled = $false; Started = $false; Message = 'Task not found' }
    }

    $Enabled = ($Task.State -ne 'Disabled')

    try {
        Start-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction Stop
        return [pscustomobject]@{ Exists = $true; Enabled = $Enabled; Started = $true; Message = 'Started' }
    }
    catch {
        return [pscustomobject]@{ Exists = $true; Enabled = $Enabled; Started = $false; Message = $_.Exception.Message }
    }
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

$Script:UdfValue  = 'Unknown'
$Script:UdfNumber = $null

<#
    Everything below runs inside a try/catch. Without it an unexpected terminating error
    would kill the script before the Result block was emitted, leaving Datto with
    truncated output and no status at all.

    Write-DattoResult calls exit, which is not caught by catch - normal exits pass
    straight through this wrapper.
#>
try {

    #region UdfNumber

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
        Both channel variables are required. A configuration error is never treated as a
        device-state outcome: nothing is changed, the UDF is left untouched, and the
        component exits 1 so the misconfiguration is visible rather than reported as a
        set of green devices.
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
        Write-DattoResult -Status 'CONFIG ERROR: SourceChannel is not a recognised update channel' -ExitCode 1
    }

    if (-not $TargetGuid -or -not $ChannelDefinitions.Contains($TargetGuid)) {
        Write-Host "ERROR: TargetChannel '$($env:TargetChannel)' is not a recognised update channel."
        $Script:UdfNumber = $null
        Write-DattoResult -Status 'CONFIG ERROR: TargetChannel is not a recognised update channel' -ExitCode 1
    }

    if ($SourceGuid -eq $TargetGuid) {
        Write-Host 'ERROR: SourceChannel and TargetChannel are the same channel - nothing to do.'
        $Script:UdfNumber = $null
        Write-DattoResult -Status 'CONFIG ERROR: SourceChannel and TargetChannel are the same channel' -ExitCode 1
    }

    $SourceName = $ChannelDefinitions[$SourceGuid].Name
    $TargetName = $ChannelDefinitions[$TargetGuid].Name
    $TargetUrl  = $ChannelDefinitions[$TargetGuid].CdnUrl

    $SimulateOnly  = ($env:SimulateOnly -match '^(true|yes|1)$')
    $TriggerUpdate = -not ($env:TriggerUpdate -match '^(false|no|0)$')

    Write-Host "Change FROM      : $SourceName"
    Write-Host "Change TO        : $TargetName"
    Write-Host "Write            : $WriteValueName = $TargetUrl"
    Write-Host "Simulate only    : $SimulateOnly"
    Write-Host "Run update task  : $TriggerUpdate"
    if ($SimulateOnly) {
        Write-Host '*** SIMULATION MODE - nothing will be written and no task will be started ***'
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

    $C2RConfig      = Get-ItemProperty -LiteralPath $C2RConfigPath
    $InstalledBuild = [string]$C2RConfig.VersionToReport
    $IsUnified      = Test-UnifiedBuild -VersionToReport $InstalledBuild

    Write-Host "Office version   : $(if ($InstalledBuild) { $InstalledBuild } else { '<unknown>' })"
    Write-Host "Products         : $(if ($C2RConfig.ProductReleaseIds) { $C2RConfig.ProductReleaseIds } else { '<unknown>' })"

    if ($IsUnified -eq $true) {
        Write-Host 'Note             : Build is at or past Version 2606 - this device already'
        Write-Host '                   receives the Monthly Enterprise Channel update experience'
        Write-Host '                   regardless of how it is configured.'
    }
    Write-Host ''

    #endregion

    #region PolicyGuard

    Write-Host 'Policy sources, priority 1-4 (these override anything this script can set):'
    $FoundPolicies = @(Get-SourceValues -Sources $PolicySources)
    Write-SourceReport -Sources $PolicySources -Found $FoundPolicies
    Write-Host ''

    if ($FoundPolicies.Count -gt 0) {
        $Blocking     = $FoundPolicies[0]
        $BlockingName = if ($Blocking.Guid) { Get-ChannelFriendlyName -Guid $Blocking.Guid } else { $Blocking.Value }

        Write-Host 'The update channel on this device is managed by policy.'
        Write-Host "No change will be made - $WriteValueName sits at priority 5 and would be overridden."
        Write-Host "Correct this at the policy source instead: $($Blocking.Label)"

        $Script:UdfValue = "Policy managed - $BlockingName"
        Write-DattoResult -Status "SKIPPED: Channel is policy managed via $($Blocking.Label) - change at the policy source" -ExitCode 0
    }

    Write-Host "No update policy present - safe to write $WriteValueName."
    Write-Host ''

    #endregion

    #region CurrentChannel

    Write-Host 'ClickToRun sources, priority 5-7:'
    $FoundC2R = @(Get-SourceValues -Sources $C2RSources)
    Write-SourceReport -Sources $C2RSources -Found $FoundC2R
    Write-Host ''

    if ($FoundC2R.Count -eq 0) {
        Write-Host 'No update channel value present under the ClickToRun Configuration key.'
        Write-Host 'Leaving this device alone - the channel state cannot be established.'
        $Script:UdfValue = 'Channel not set'
        Write-DattoResult -Status 'SKIPPED: No update channel value present - state could not be established' -ExitCode 0
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

    if ($SimulateOnly) {
        Write-Host ''
        Write-Host 'SIMULATION - the following would have been done:'
        Write-Host "  Set $C2RConfigPath\$WriteValueName = $TargetUrl"
        if ($TriggerUpdate) {
            Write-Host "  Start scheduled task $UpdateTaskPath$UpdateTaskName"
        }
        $Script:UdfValue = $CurrentName
        Write-DattoResult -Status "SIMULATED: Would set $WriteValueName to $TargetName" -ExitCode 0
    }

    Write-Host ''
    Write-Host "Writing $WriteValueName..."
    try {
        New-ItemProperty -LiteralPath $C2RConfigPath -Name $WriteValueName -Value $TargetUrl `
            -PropertyType String -Force -ErrorAction Stop | Out-Null
        Write-Host "  $WriteValueName = $TargetUrl"
    }
    catch {
        Write-Host "  ERROR: Could not write $WriteValueName - $($_.Exception.Message)"
        $Script:UdfValue = $CurrentName
        Write-DattoResult -Status "FAILED: Could not write $WriteValueName - $($_.Exception.Message)" -ExitCode 1
    }

    <#
        Read the value back. This confirms the write landed and was not blocked or
        redirected - it does NOT confirm Office has accepted the channel, which can take
        up to 24 hours and is what the reporting component is for.
    #>
    $WrittenBack = Get-RegistryValue -Path $C2RConfigPath -Name $WriteValueName
    $WrittenGuid = Resolve-ChannelGuid -InputValue $WrittenBack

    if ($WrittenGuid -ne $TargetGuid) {
        Write-Host ''
        Write-Host "ERROR: $WriteValueName did not read back as expected."
        Write-Host "  Expected: $TargetUrl"
        Write-Host "  Found   : $(if ($WrittenBack) { $WrittenBack } else { '<empty>' })"
        $Script:UdfValue = $CurrentName
        Write-DattoResult -Status "FAILED: $WriteValueName did not read back as $TargetName" -ExitCode 1
    }
    Write-Host "  Read back OK - $WriteValueName resolves to $TargetName"

    #endregion

    #region UpdateTask

    <#
        Writing the setting is not enough on its own: the Office Automatic Updates 2.0
        task is what detects it and reassigns the channel. A missing task means the change
        will never apply, so that is treated as a failure.
    #>
    Write-Host ''
    Write-Host 'Office Automatic Updates 2.0 task...'

    $TaskResult = Start-OfficeUpdateTask

    if (-not $TaskResult.Exists) {
        Write-Host "  ERROR: Scheduled task '$UpdateTaskName' was not found."
        Write-Host '         The channel setting is written but will never be applied without it.'
        $Script:UdfValue = "$TargetName (pending - update task missing)"
        Write-DattoResult -Status "FAILED: $WriteValueName set to $TargetName but the Office Automatic Updates 2.0 task is missing" -ExitCode 1
    }

    if (-not $TaskResult.Enabled) {
        Write-Host '  WARNING: The task exists but is disabled. The channel change will not apply'
        Write-Host '           until it is enabled and allowed to run.'
    }

    if ($TriggerUpdate) {
        if ($TaskResult.Started) {
            Write-Host '  Task started.'
        }
        else {
            Write-Host "  WARNING: Could not start the task - $($TaskResult.Message)"
            Write-Host '           The setting is written; the task will pick it up on its own schedule.'
        }
    }
    else {
        Write-Host '  Not started (TriggerUpdate is false) - the task will pick the setting up on its own schedule.'
    }

    #endregion

    #region Outcome

    Write-Host ''
    Write-Host '--------------------------------------------------------------------'
    Write-Host "CONFIGURED: $WriteValueName now points at $TargetName."
    Write-Host ''
    Write-Host 'This device has NOT changed channel yet. Office applies the change on its own'
    Write-Host 'schedule - typically within 24 hours - and only reports the new channel once a'
    Write-Host 'build from it has installed. Re-run the reporting component in a day or two to'
    Write-Host 'confirm the outcome.'

    $Script:UdfValue = "$TargetName (pending)"
    Write-DattoResult -Status "CONFIGURED: $CurrentName -> $TargetName (applies within ~24h)" -ExitCode 0

    #endregion
}
catch {
    Write-Host ''
    Write-Host '===================================================================='
    Write-Host 'UNEXPECTED ERROR - the script stopped before completing.'
    Write-Host "  Message : $($_.Exception.Message)"
    Write-Host "  Type    : $($_.Exception.GetType().FullName)"
    Write-Host "  Location: line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
    Write-Host '===================================================================='

    # The device state is unknown at this point, so do not overwrite a good UDF value.
    $Script:UdfNumber = $null
    Write-DattoResult -Status "ERROR: Unexpected failure - $($_.Exception.Message)" -ExitCode 1
}

#endregion
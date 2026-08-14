<#
.SYNOPSIS
    Datto RMM Component - Audit Microsoft 365 Apps (Click-to-Run) Update Channel.

.DESCRIPTION
    Reads HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration and compares the
    configured update channel against the expected channel supplied in the Datto RMM
    site/component variable "AuditUpdateChannel".

    Comparison is done on the channel GUID only, so http/https, trailing slashes and
    casing differences between the variable and the registry are all tolerated.

.INPUTS
    ENV:AuditUpdateChannel - Expected channel as a full CDN URL, supplied by a Datto RMM
    selection variable, e.g. http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60

.OUTPUTS
    Datto RMM Result block:  Status=<SUCCESS|FAILURE|...> : <detail>

.NOTES
    Exit codes:
        0 = Channel matches the expected value
        1 = Any failure condition - mismatch, Office C2R not installed, no channel value
            present, or the AuditUpdateChannel variable was not supplied. The Result
            line identifies which.

#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

#region ChannelMap

# Channels selectable via the variable
$ChannelMap = [ordered]@{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current Channel'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current Channel (Preview)'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise Channel'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise Channel'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise Channel (Preview)'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta Channel'
}

$GuidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

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

function Get-ChannelFriendlyName {
    param([string]$Guid)
    if ($Guid -and $ChannelMap.Contains($Guid)) { return $ChannelMap[$Guid] }
    return 'Unrecognised channel'
}

function Get-C2RConfiguration {
    <#  Reads the ClickToRun Configuration key from the 64-bit registry view. #>
    $Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $BaseKey = $null
    $SubKey  = $null

    try {
        $BaseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Registry64
        )
        $SubKey = $BaseKey.OpenSubKey($Path)
        if ($null -ne $SubKey) {
            return [pscustomobject]@{
                KeyPath        = "HKLM:\$Path"
                UpdateChannel  = [string]$SubKey.GetValue('UpdateChannel')
                CDNBaseUrl     = [string]$SubKey.GetValue('CDNBaseUrl')
                UpdateUrl      = [string]$SubKey.GetValue('UpdateUrl')
                ProductVersion = [string]$SubKey.GetValue('VersionToReport')
                ProductIDs     = [string]$SubKey.GetValue('ProductReleaseIds')
            }
        }
    }
    catch {
        Write-Verbose "Failed reading the registry: $($_.Exception.Message)"
    }
    finally {
        if ($SubKey)  { $SubKey.Close() }
        if ($BaseKey) { $BaseKey.Close() }
    }

    return $null
}

function Get-UpdatePolicy {
    <#  Informational only - a GPO here overrides the Configuration key. #>
    $PolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'
    if (Test-Path -LiteralPath $PolicyPath) {
        $Policy = Get-ItemProperty -LiteralPath $PolicyPath -ErrorAction SilentlyContinue
        return [pscustomobject]@{
            UpdateBranch = [string]$Policy.updatebranch
            UpdatePath   = [string]$Policy.updatepath
            UpdateTarget = [string]$Policy.updatetargetversion
        }
    }
    return $null
}

function Write-DattoResult {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][ValidateSet(0, 1)][int]$ExitCode
    )
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

#region ExpectedChannel

$ExpectedRaw = $env:AuditUpdateChannel

if ([string]::IsNullOrWhiteSpace($ExpectedRaw)) {
    Write-Host 'ERROR: The component variable "AuditUpdateChannel" was not supplied.'
    Write-Host '<-End Diagnostic->'
    Write-DattoResult -Status 'ERROR: AuditUpdateChannel variable not supplied' -ExitCode 1
}

$ExpectedGuid = Resolve-ChannelGuid -InputValue $ExpectedRaw
$ExpectedName = Get-ChannelFriendlyName -Guid $ExpectedGuid

Write-Host "Expected channel : $ExpectedName"
Write-Host "Expected GUID    : $ExpectedGuid"
Write-Host "Variable value   : $ExpectedRaw"
Write-Host ''

#endregion

#region ReadRegistry

$Configuration = Get-C2RConfiguration

if ($null -eq $Configuration) {
    Write-Host 'Office Click-to-Run Configuration key not found.'
    Write-Host 'Office C2R (Microsoft 365 Apps) does not appear to be installed on this device.'
    Write-Host '<-End Diagnostic->'
    Write-DattoResult -Status 'NOT INSTALLED: Office Click-to-Run configuration key not found' -ExitCode 1
}

Write-Host "Registry key     : $($Configuration.KeyPath)"
Write-Host "UpdateChannel    : $(if ($Configuration.UpdateChannel)  { $Configuration.UpdateChannel }  else { '<not set>' })"
Write-Host "CDNBaseUrl       : $(if ($Configuration.CDNBaseUrl)     { $Configuration.CDNBaseUrl }     else { '<not set>' })"
Write-Host "UpdateUrl        : $(if ($Configuration.UpdateUrl)      { $Configuration.UpdateUrl }      else { '<not set>' })"
Write-Host "Office version   : $(if ($Configuration.ProductVersion) { $Configuration.ProductVersion } else { '<unknown>' })"
Write-Host "Products         : $(if ($Configuration.ProductIDs)     { $Configuration.ProductIDs }     else { '<unknown>' })"

# UpdateChannel is authoritative; fall back to CDNBaseUrl where it is absent.
$ActualRaw    = $null
$ActualSource = $null
foreach ($Candidate in @('UpdateChannel', 'CDNBaseUrl', 'UpdateUrl')) {
    if (-not [string]::IsNullOrWhiteSpace($Configuration.$Candidate)) {
        $ActualRaw    = $Configuration.$Candidate
        $ActualSource = $Candidate
        break
    }
}

if ([string]::IsNullOrWhiteSpace($ActualRaw)) {
    Write-Host ''
    Write-Host 'No UpdateChannel / CDNBaseUrl value present under the Configuration key.'
    Write-Host '<-End Diagnostic->'
    Write-DattoResult -Status 'NOT SET: No update channel value found in the ClickToRun Configuration key' -ExitCode 1
}

$ActualGuid = Resolve-ChannelGuid -InputValue $ActualRaw
$ActualName = if ($ActualGuid) { Get-ChannelFriendlyName -Guid $ActualGuid } else { 'Unrecognised channel' }

Write-Host ''
Write-Host "Value used       : $ActualSource"
Write-Host "Current channel  : $ActualName"
Write-Host "Current GUID     : $(if ($ActualGuid) { $ActualGuid } else { '<none found in value>' })"

#endregion

#region PolicyCheck

$UpdatePolicy = Get-UpdatePolicy
if ($UpdatePolicy -and ($UpdatePolicy.UpdateBranch -or $UpdatePolicy.UpdatePath)) {
    Write-Host ''
    Write-Host 'NOTE: An Office update GPO is present and may override the channel above:'
    if ($UpdatePolicy.UpdateBranch) { Write-Host "  updatebranch : $($UpdatePolicy.UpdateBranch)" }
    if ($UpdatePolicy.UpdatePath)   { Write-Host "  updatepath   : $($UpdatePolicy.UpdatePath)" }
    if ($UpdatePolicy.UpdateTarget) { Write-Host "  targetversion: $($UpdatePolicy.UpdateTarget)" }
}

#endregion

#region Compare

Write-Host '--------------------------------------------------------------------'

if ($ActualGuid -and ($ActualGuid -eq $ExpectedGuid)) {
    Write-Host "RESULT: MATCH - device is on $ExpectedName."
    Write-Host '<-End Diagnostic->'
    Write-DattoResult -Status "SUCCESS: Office update channel is $ExpectedName ($ExpectedGuid)" -ExitCode 0
}
else {
    Write-Host "RESULT: MISMATCH - expected '$ExpectedName' but found '$ActualName'."
    Write-Host '<-End Diagnostic->'
    $Detail = if ($ActualGuid) { "$ActualName ($ActualGuid)" } else { "unrecognised value '$ActualRaw'" }
    Write-DattoResult -Status "FAILURE: Expected $ExpectedName but found $Detail" -ExitCode 1
}

#endregion

#endregion
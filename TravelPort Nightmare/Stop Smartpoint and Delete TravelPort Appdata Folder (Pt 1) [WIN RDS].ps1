<#
    Travelport Smartpoint Reset - PART 1 of 2
    Stops Smartpoint for a specific user, then deletes their Roaming\Travelport folder.

    DATTO RMM COMPONENT VARIABLE
        Name : Action
        Type : Value (String)
        Input: target username WITHOUT domain, e.g.  fklein
               (a value like CTSCV\fklein is also accepted - the domain is stripped)

#>

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------- config
$Domain        = 'CTSCV'
$ProcessName   = 'Travelport.Smartpoint.App'
$ProfileRoot   = 'C:\Users'
$RelativePath  = 'AppData\Roaming\Travelport'
$KillTimeout   = 15   # seconds to wait for the process to actually exit
$SettleSeconds = 5    # grace period for file handles to release after exit

# ----------------------------------------------------------------------------- helpers
function Get-TargetProcess {
    # -IncludeUserName requires elevation. The DRMM agent runs as SYSTEM, so this is safe.
    param([string]$Name, [string]$Account)

    return @(Get-Process -Name $Name -IncludeUserName -ErrorAction SilentlyContinue |
             Where-Object { $_.UserName -ieq $Account })
}

# ----------------------------------------------------------------------------- input validation
$TargetUser = "$env:Action".Trim()

# Be forgiving if the tech pastes FQDN
if ($TargetUser -match '\\') { $TargetUser = $TargetUser.Split('\')[-1] }
if ($TargetUser -match '@')  { $TargetUser = $TargetUser.Split('@')[0]  }

if ([string]::IsNullOrWhiteSpace($TargetUser)) {
    Write-Output 'FAILED: No username was supplied. Nothing was stopped or deleted.'
    Write-Output '        Re-run and enter the username, for example: fklein'
    return
}

# Hard stop on anything that is not a plain username.
# A wildcard here would target every profile on the machine.
if ($TargetUser -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    Write-Output "FAILED: '$TargetUser' is not a valid username. Nothing was stopped or deleted."
    Write-Output '        Re-run using letters, digits, dot, dash or underscore only, for example: fklein'
    return
}

$Account     = "$Domain\$TargetUser"
$ProfilePath = Join-Path $ProfileRoot $TargetUser
$Folder      = Join-Path $ProfilePath $RelativePath

# Catches a mistyped username.
if (-not (Test-Path -LiteralPath $ProfilePath)) {
    Write-Output "FAILED: No profile for $TargetUser on this computer. Nothing was stopped or deleted."
    Write-Output '        Check the spelling of the username and that this is the correct device.'
    return
}

# ----------------------------------------------------------------------------- stop the process
$Processes = Get-TargetProcess -Name $ProcessName -Account $Account

if ($Processes.Count -gt 0) {
    foreach ($Process in $Processes) {
        try   { Stop-Process -Id $Process.Id -Force -ErrorAction Stop }
        catch { }   # verified below by re-checking for surviving instances
    }

    $Processes | Wait-Process -Timeout $KillTimeout -ErrorAction SilentlyContinue

    if ((Get-TargetProcess -Name $ProcessName -Account $Account).Count -gt 0) {
        Write-Output "FAILED: Smartpoint is still running for $TargetUser after $KillTimeout seconds."
        Write-Output '        The folder was NOT deleted, so nothing is half-removed.'
        Write-Output '        Have the user close Smartpoint, or end the task, then re-run.'
        return
    }

    Start-Sleep -Seconds $SettleSeconds
}

# ----------------------------------------------------------------------------- delete the folder
if (-not (Test-Path -LiteralPath $Folder)) {
    Write-Output "SUCCESS: Smartpoint is closed for $TargetUser and there was no Travelport folder to delete."
    Write-Output '         Part 2 can now be run.'
    return
}

try {
    Remove-Item -LiteralPath $Folder -Recurse -Force -ErrorAction Stop
    Write-Output "SUCCESS: Smartpoint was stopped and the Travelport folder deleted for $TargetUser."
    Write-Output '         Part 2 can now be run to restore from backup.'
}
catch {
    Write-Output "FAILED: The Travelport folder for $TargetUser could not be deleted."
    Write-Output '        It may be partly deleted, so do not leave the user here.'
    Write-Output '        Confirm Smartpoint is fully closed and re-run.'
    Write-Output "        Detail: $($_.Exception.Message.Trim())"
}
<#
    Travelport Smartpoint Reset - PART 1 of 2  (Stop, Backup, Delete)

    For the user named in ENV:Action this script will:
        1. Stop Smartpoint for that user
        2. Back up Roaming\Travelport to "Travelport BACKUP", but ONLY if no backup
           already exists. An existing backup is never replaced.
        3. Delete Roaming\Travelport

    The delete only happens once a usable backup is confirmed in place.
    Part 2 restores from "Travelport BACKUP".

    DATTO RMM COMPONENT VARIABLES
        Name : Action
        Type : Value (String)
        Input: target username WITHOUT domain, e.g.  fklein
               (a value like CTSCV\fklein is also accepted - the domain is stripped)

        Name : Verification
        Type : Value (String)
        Input: the tech must type   Permanently Delete Travelport Folder

#>

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------- config
$Domain             = 'CTSCV'
$ProcessName        = 'Travelport.Smartpoint.App'
$ProfileRoot        = 'C:\Users'
$SourcePath         = 'AppData\Roaming\Travelport'
$BackupPath         = 'AppData\Roaming\Travelport BACKUP'
$VerificationPhrase = 'Permanently Delete Travelport Folder'
$KillTimeout        = 15   # seconds to wait for the process to actually exit
$SettleSeconds      = 5    # grace period for file handles to release after exit

# ----------------------------------------------------------------------------- helpers
function Get-TargetProcess {
    # -IncludeUserName requires elevation. The DRMM agent runs as SYSTEM, so this is safe.
    param([string]$Name, [string]$Account)

    return @(Get-Process -Name $Name -IncludeUserName -ErrorAction SilentlyContinue |
             Where-Object { $_.UserName -ieq $Account })
}

function Get-FileCount {
    param([string]$Path)

    return @(Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue).Count
}

# ----------------------------------------------------------------------------- confirmation gate
# Trim the ends and collapse any repeated inner spaces, so a stray space does not
# reject a tech who typed the phrase correctly. Comparison is case-insensitive.
$Verification = ("$env:Verification".Trim() -replace '\s+', ' ')

if ($Verification -ne $VerificationPhrase) {
    Write-Output 'FAILED: Confirmation phrase missing or incorrect. Nothing was changed.'
    Write-Output '        Type the following into the Verification field, then re-run:'
    Write-Output "        $VerificationPhrase"
    return
}

# ----------------------------------------------------------------------------- input validation
$TargetUser = "$env:Action".Trim()

# Be forgiving if the tech pastes FQDN
if ($TargetUser -match '\\') { $TargetUser = $TargetUser.Split('\')[-1] }
if ($TargetUser -match '@')  { $TargetUser = $TargetUser.Split('@')[0]  }

if ([string]::IsNullOrWhiteSpace($TargetUser)) {
    Write-Output 'FAILED: No username was supplied. Nothing was changed.'
    Write-Output '        Re-run and enter the username, for example: fklein'
    return
}

# Hard stop on anything that is not a plain username.
# A wildcard here would target every profile on the machine.
if ($TargetUser -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    Write-Output "FAILED: '$TargetUser' is not a valid username. Nothing was changed."
    Write-Output '        Re-run using letters, digits, dot, dash or underscore only, for example: fklein'
    return
}

$Account     = "$Domain\$TargetUser"
$ProfilePath = Join-Path $ProfileRoot $TargetUser
$Source      = Join-Path $ProfilePath $SourcePath
$Backup      = Join-Path $ProfilePath $BackupPath

# Catches a mistyped username.
if (-not (Test-Path -LiteralPath $ProfilePath)) {
    Write-Output "FAILED: No profile for $TargetUser on this computer. Nothing was changed."
    Write-Output '        Check the spelling of the username and that this is the correct device.'
    return
}

$SourceExists = Test-Path -LiteralPath $Source
$BackupExists = Test-Path -LiteralPath $Backup

# ----------------------------------------------------------------------------- pre-flight checks
# Everything that can rule the job out is checked before Smartpoint is stopped, so the
# user is never interrupted for a run that was going to stop anyway.

if (-not $SourceExists) {
    if ($BackupExists) {
        Write-Output "SUCCESS: $TargetUser has no Travelport folder, so there was nothing to delete."
        Write-Output "         The existing backup was left untouched ($(Get-FileCount $Backup) file(s))."
        Write-Output '         Part 2 can be run to restore it.'
    }
    else {
        Write-Output "FAILED: $TargetUser has no Travelport folder and no backup. Nothing was changed."
        Write-Output '        There is nothing to delete and nothing for Part 2 to restore.'
        Write-Output '        Have the user launch Smartpoint once, then re-run.'
    }
    return
}

if ($BackupExists) {
    # A backup that exists but is empty would let the delete proceed and leave Part 2
    # with nothing to put back.
    $BackupCount = Get-FileCount $Backup

    if ($BackupCount -eq 0) {
        Write-Output "FAILED: The existing backup for $TargetUser is empty. Nothing was changed."
        Write-Output '        Deleting now would leave Part 2 with nothing to restore.'
        Write-Output "        Check or remove the empty folder manually, then re-run: $Backup"
        return
    }
}
else {
    # No backup yet, so one is taken below. An empty source would produce a backup that
    # looks valid and restores nothing.
    if ((Get-FileCount $Source) -eq 0) {
        Write-Output "FAILED: The Travelport folder for $TargetUser is empty. Nothing was changed."
        Write-Output '        Backing it up would give Part 2 nothing to restore.'
        Write-Output '        Check the folder manually before running this again.'
        return
    }
}

# ----------------------------------------------------------------------------- stop the process
# Done before the backup as well as the delete: Smartpoint is being closed either way,
# so copying afterwards avoids any chance of catching a file mid-write.
$Processes = Get-TargetProcess -Name $ProcessName -Account $Account

if ($Processes.Count -gt 0) {
    foreach ($Process in $Processes) {
        try   { Stop-Process -Id $Process.Id -Force -ErrorAction Stop }
        catch { }   # verified below by re-checking for surviving instances
    }

    $Processes | Wait-Process -Timeout $KillTimeout -ErrorAction SilentlyContinue

    if ((Get-TargetProcess -Name $ProcessName -Account $Account).Count -gt 0) {
        Write-Output "FAILED: Smartpoint is still running for $TargetUser after $KillTimeout seconds."
        Write-Output '        Nothing was backed up or deleted.'
        Write-Output '        Have the user close Smartpoint, or end the task, then re-run.'
        return
    }

    Start-Sleep -Seconds $SettleSeconds
}

# ----------------------------------------------------------------------------- backup
$SourceCount = Get-FileCount $Source

if ($BackupExists) {
    $Existing    = Get-Item -LiteralPath $Backup
    $BackupCount = Get-FileCount $Backup
    $BackupNote  = "Existing backup kept: $BackupCount file(s) from $($Existing.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))"
}
else {
    try {
        Copy-Item -LiteralPath $Source -Destination $Backup -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Output "FAILED: The backup for $TargetUser did not complete, so nothing was deleted."
        Write-Output '        A partial backup folder may have been left behind. Check it before re-running.'
        Write-Output "        Backup path: $Backup"
        Write-Output "        Detail: $($_.Exception.Message.Trim())"
        return
    }

    $BackupCount = Get-FileCount $Backup

    # The copy can skip a locked file without raising an error, so the result is counted
    # rather than assumed. Nothing is deleted unless the backup is complete.
    if ($BackupCount -lt $SourceCount) {
        Write-Output "FAILED: The backup for $TargetUser is incomplete, so nothing was deleted."
        Write-Output "        The folder holds $SourceCount file(s) but the backup has $BackupCount."
        Write-Output "        Check the backup manually before re-running: $Backup"
        return
    }

    $BackupNote = "New backup created: $BackupCount file(s)"
}

# ----------------------------------------------------------------------------- delete the folder
try {
    Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction Stop

    Write-Output "SUCCESS: Smartpoint stopped and the Travelport folder deleted for $TargetUser."
    Write-Output "         $BackupNote"
    Write-Output '         Part 2 can now be run to restore from backup.'
}
catch {
    Write-Output "FAILED: The Travelport folder for $TargetUser could not be deleted."
    Write-Output '        It may be partly deleted, so do not leave the user here.'
    Write-Output "        The backup is in place, so Part 2 can recover them: $BackupNote"
    Write-Output '        Confirm Smartpoint is fully closed and re-run.'
    Write-Output "        Detail: $($_.Exception.Message.Trim())"
}
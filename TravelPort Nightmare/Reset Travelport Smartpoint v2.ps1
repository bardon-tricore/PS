<#
    Travelport Smartpoint Reset - SINGLE RUN

    For the user named in ENV:Action this script will:
        1. Confirm the user is signed in to this host (required - Smartpoint must be
           relaunched inside their session)
        2. Stop Smartpoint
        3. Back up Roaming\Travelport in full to "Travelport BACKUP", ONLY if no backup
           exists. An existing backup is never replaced.
        4. Delete Roaming\Travelport
        5. Relaunch Smartpoint as the user so the folder is rebuilt, then wait
        6. Stop Smartpoint again
        7. Restore ONLY the folders listed in $RestoreFolders from the backup

    The whole folder is still backed up, but only the customisation folders are put
    back. Everything else is left as Smartpoint regenerated it. For each restored
    folder the rebuilt copy is deleted first, then replaced wholesale from the backup,
    so no regenerated files are left mixed in with restored ones.

    If any step after the delete fails, the backup is left intact and the Stop SmartPoint & Restore Folder (Complete)
    component can be run manually to recover the user. Which was once known as "part 2" due to this, keep it available.

    DATTO RMM COMPONENT VARIABLES
        Name : Action
        Type : Value (String)
        Input: target username WITHOUT domain, e.g.  fklein
               (a value like CTSCV\fklein is also accepted - the domain is stripped)

        Name : Verification
        Type : Value (String)
        Input: the tech must type   Permanently Delete Travelport Folder
               IMPORTANT - leave this variable's default value EMPTY in the component
               definition. If it is pre-filled, the confirmation is bypassed entirely.
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
$LaunchTimeout      = 20   # seconds to wait for Smartpoint to appear after relaunch
$ConfigWriteWait    = 30   # flat wait for Smartpoint to finish writing its config

# The only folders copied back from the backup. Wildcards are supported.

$RestoreFolders = @(
    '0370.00.QuickCommands',
    'SmartButtons'
)

# Used only if Smartpoint is not running when the script starts. When it IS running,
# the path is read from the live process, which is more reliable than hardcoding it.
$SmartpointExeFallback = 'C:\Program Files (x86)\Travelport\Smartpoint\Travelport.Smartpoint.App.exe'   # e.g. 'C:\Program Files (x86)\Travelport\Smartpoint\Travelport.Smartpoint.App.exe'

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

function Get-RestoreCandidates {
    # Returns the folders under $Root that match any entry in $Patterns.
    param([string]$Root, [string[]]$Patterns)

    $Folders = Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue
    if (-not $Folders) { return @() }

    return @($Folders | Where-Object {
        $Name = $_.Name
        ($Patterns | Where-Object { $Name -like $_ }).Count -gt 0
    })
}

function Stop-Smartpoint {
    # Returns $true once no instances remain for the account.
    param([string]$Name, [string]$Account, [int]$Timeout, [int]$Settle)

    $Running = Get-TargetProcess -Name $Name -Account $Account
    if ($Running.Count -eq 0) { return $true }

    foreach ($Process in $Running) {
        try   { Stop-Process -Id $Process.Id -Force -ErrorAction Stop }
        catch { }   # verified below by re-checking for surviving instances
    }

    $Running | Wait-Process -Timeout $Timeout -ErrorAction SilentlyContinue

    if ((Get-TargetProcess -Name $Name -Account $Account).Count -gt 0) { return $false }

    Start-Sleep -Seconds $Settle
    return $true
}

function Start-AsUser {
    # The agent runs as SYSTEM in session 0, so Start-Process cannot reach the user's
    # desktop. A transient scheduled task with an Interactive principal can.
    param([string]$ExePath, [string]$Account)

    $TaskName = "TravelportReset_$([guid]::NewGuid().ToString('N').Substring(0,8))"

    try {
        $TaskAction    = New-ScheduledTaskAction -Execute $ExePath
        $TaskPrincipal = New-ScheduledTaskPrincipal -UserId $Account -LogonType Interactive

        Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Principal $TaskPrincipal -Force | Out-Null
        Start-ScheduledTask -TaskName $TaskName
        Start-Sleep -Seconds 3
    }
    finally {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    }
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

$BackupExists = Test-Path -LiteralPath $Backup

# ----------------------------------------------------------------------------- pre-flight checks
# Everything that can rule the job out is checked before anything is stopped or changed.

if (-not (Test-Path -LiteralPath $Source)) {
    Write-Output "FAILED: $TargetUser has no Travelport folder. Nothing was changed."
    Write-Output '        There is nothing to reset. Have the user launch Smartpoint once, then re-run.'
    return
}

# Whichever folder the restore will draw from must actually contain the customisation
# folders. Without them the reset would wipe the user and put nothing back.
$CheckRoot   = if ($BackupExists) { $Backup } else { $Source }
$CheckLabel  = if ($BackupExists) { 'existing backup' } else { 'current Travelport folder' }
$Candidates  = Get-RestoreCandidates -Root $CheckRoot -Patterns $RestoreFolders

if ($Candidates.Count -eq 0) {
    Write-Output "FAILED: The $CheckLabel for $TargetUser holds none of the folders to be restored."
    Write-Output "        Looked in: $CheckRoot"
    Write-Output "        Looked for: $($RestoreFolders -join ', ')"
    Write-Output '        Resetting now would wipe the user and put nothing back.'
    Write-Output '        Check the folder names on the device. If Smartpoint has been upgraded, the'
    Write-Output '        version-numbered folder may have been renamed and $RestoreFolders needs updating.'
    return
}

# The user must be signed in: Smartpoint has to be relaunched inside their session
# for the folder to rebuild. explorer.exe runs per session, so it is a good signal.
if ((Get-TargetProcess -Name 'explorer' -Account $Account).Count -eq 0) {
    Write-Output "FAILED: $TargetUser is not signed in to this host. Nothing was changed."
    Write-Output '        Smartpoint has to be relaunched in their session to rebuild the folder.'
    Write-Output '        Have the user sign in, confirm this is the right session host, then re-run.'
    return
}

# Read the executable path from the live process where possible.
$Running       = Get-TargetProcess -Name $ProcessName -Account $Account
$SmartpointExe = if ($Running.Count -gt 0) { $Running[0].Path } else { $SmartpointExeFallback }

if ([string]::IsNullOrWhiteSpace($SmartpointExe) -or -not (Test-Path -LiteralPath $SmartpointExe)) {
    Write-Output 'FAILED: Could not determine where Smartpoint is installed. Nothing was changed.'
    Write-Output '        Smartpoint was not running, so its path could not be read from the process.'
    Write-Output '        Either have the user open Smartpoint and re-run, or set'
    Write-Output '        $SmartpointExeFallback at the top of this script to the full path of the exe.'
    return
}

# ----------------------------------------------------------------------------- 1. stop Smartpoint
if (-not (Stop-Smartpoint -Name $ProcessName -Account $Account -Timeout $KillTimeout -Settle $SettleSeconds)) {
    Write-Output "FAILED: Smartpoint is still running for $TargetUser after $KillTimeout seconds."
    Write-Output '        Nothing was backed up, deleted or changed.'
    Write-Output '        Have the user close Smartpoint, or end the task, then re-run.'
    return
}

# ----------------------------------------------------------------------------- 2. backup
$SourceCount = Get-FileCount $Source

if ($BackupExists) {
    $Existing    = Get-Item -LiteralPath $Backup
    $BackupCount = Get-FileCount $Backup
    $BackupNote  = "Existing backup used: $BackupCount file(s) from $($Existing.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))"
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

# Re-read against the backup now that it definitely exists.
$Candidates = Get-RestoreCandidates -Root $Backup -Patterns $RestoreFolders

if ($Candidates.Count -eq 0) {
    Write-Output "FAILED: The backup for $TargetUser holds none of the folders to be restored."
    Write-Output "        Nothing was deleted. Checked: $Backup"
    Write-Output "        Looked for: $($RestoreFolders -join ', ')"
    return
}

# ----------------------------------------------------------------------------- 3. delete
# Past this point the user's folder is gone. Every failure below reports that the
# backup is intact so the Stop SmartPoint & Restore Folder (Complete) component can recover them.
try {
    Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction Stop
}
catch {
    Write-Output "FAILED: The Travelport folder for $TargetUser could not be deleted."
    Write-Output '        It may be partly deleted, so do not leave the user here.'
    Write-Output "        $BackupNote"
    Write-Output '        Confirm Smartpoint is fully closed and re-run.'
    Write-Output "        Detail: $($_.Exception.Message.Trim())"
    return
}

# ----------------------------------------------------------------------------- 4. relaunch to rebuild
try {
    Start-AsUser -ExePath $SmartpointExe -Account $Account
}
catch {
    Write-Output "FAILED: Could not relaunch Smartpoint for $TargetUser. The folder is deleted."
    Write-Output "        $BackupNote"
    Write-Output '        RECOVER: have the user open Smartpoint themselves, then run the Stop SmartPoint'
    Write-Output '        & Restore Folder (Complete) component against this user.'
    Write-Output "        Detail: $($_.Exception.Message.Trim())"
    return
}

# Wait for the process to appear, then allow a flat period for it to write its config.
$Deadline = (Get-Date).AddSeconds($LaunchTimeout)
$Launched = $false

while ((Get-Date) -lt $Deadline) {
    if ((Get-TargetProcess -Name $ProcessName -Account $Account).Count -gt 0) { $Launched = $true; break }
    Start-Sleep -Seconds 2
}

if (-not $Launched) {
    Write-Output "FAILED: Smartpoint did not start for $TargetUser within $LaunchTimeout seconds."
    Write-Output '        The folder is deleted and was not rebuilt.'
    Write-Output "        $BackupNote"
    Write-Output '        RECOVER: have the user open Smartpoint themselves, then run the Stop SmartPoint'
    Write-Output '        & Restore (Complete) component against this user.'
    return
}

Start-Sleep -Seconds $ConfigWriteWait

if (-not (Test-Path -LiteralPath $Source)) {
    Write-Output "FAILED: Smartpoint started but did not rebuild the Travelport folder within $ConfigWriteWait seconds."
    Write-Output '        It may be waiting at a sign-in prompt in the user session.'
    Write-Output "        $BackupNote"
    Write-Output '        RECOVER: have the user sign in to Smartpoint, then run the Stop Smartpoint'
    Write-Output '        & Restore Folder (Complete) component against this user.'
    return
}

# ----------------------------------------------------------------------------- 5. stop again
if (-not (Stop-Smartpoint -Name $ProcessName -Account $Account -Timeout $KillTimeout -Settle $SettleSeconds)) {
    Write-Output "FAILED: Smartpoint would not close after rebuilding the folder for $TargetUser."
    Write-Output '        The folder was rebuilt but nothing was restored over it.'
    Write-Output "        $BackupNote"
    Write-Output '        RECOVER: have the user close Smartpoint, then run the Stop Smartpoint'
    Write-Output '        & Restore Folder (Complete) component against this user.'
    return
}

# ----------------------------------------------------------------------------- 6. restore the customisation folders
# Each target folder is replaced wholesale rather than merged, so no regenerated files
# are left mixed in with restored ones.
$Restored = @()

foreach ($Candidate in $Candidates) {
    $Destination = Join-Path $Source $Candidate.Name

    try {
        if (Test-Path -LiteralPath $Destination) {
            Remove-Item -LiteralPath $Destination -Recurse -Force -ErrorAction Stop
        }

        Copy-Item -LiteralPath $Candidate.FullName -Destination $Destination -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Output "FAILED: Could not restore '$($Candidate.Name)' for $TargetUser."
        Write-Output '        The folder was rebuilt but the restore is incomplete.'
        Write-Output "        $BackupNote"
        Write-Output '        RECOVER: run the Stop SmartPoint & Restore Folder (Complete) component against this user.'
        Write-Output "        Detail: $($_.Exception.Message.Trim())"
        return
    }

    $ExpectedCount = Get-FileCount $Candidate.FullName
    $ActualCount   = Get-FileCount $Destination

    # The copy can skip a locked file without raising an error, so the result is counted.
    if ($ActualCount -lt $ExpectedCount) {
        Write-Output "FAILED: '$($Candidate.Name)' was only partly restored for $TargetUser."
        Write-Output "        The backup holds $ExpectedCount file(s) but the restored folder has $ActualCount."
        Write-Output '        The copy reported no error. Check the folder before closing the ticket.'
        Write-Output "        $BackupNote"
        Write-Output '        RECOVER: run the Stop SmartPoint & Restore Folder (Complete) component against this user.'
        return
    }

    $Restored += "$($Candidate.Name) ($ActualCount file(s))"
}

# Report any configured folder that was not present in the backup. Not a failure - the
# user may simply never have had it - but the tech should know it was not restored.
$NotFound = @($RestoreFolders | Where-Object {
    $Pattern = $_
    ($Candidates | Where-Object { $_.Name -like $Pattern }).Count -eq 0
})

Write-Output "SUCCESS: Travelport folder rebuilt for $TargetUser and customisations restored."
Write-Output "         Restored: $($Restored -join ', ')"

if ($NotFound.Count -gt 0) {
    Write-Output "         Not in backup, so not restored: $($NotFound -join ', ')"
}

Write-Output "         $BackupNote"
Write-Output '         Smartpoint is closed. Have the user reopen it and confirm their'
Write-Output '         SmartButtons and quick commands.'
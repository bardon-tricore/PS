<#
    Travelport Smartpoint Reset - PART 2 of 2
    Stops Smartpoint for a specific user, then restores their Roaming\Travelport folder
    from "Travelport BACKUP".

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
$BackupPath    = 'AppData\Roaming\Travelport BACKUP'
$RestorePath   = 'AppData\Roaming\Travelport'
$KillTimeout   = 15   # seconds to wait for the process to actually exit
$SettleSeconds = 5    # grace period for file handles to release after exit

# ----------------------------------------------------------------------------- helpers
function Get-TargetProcess {
    param([string]$Name, [string]$Account)

    return @(Get-Process -Name $Name -IncludeUserName -ErrorAction SilentlyContinue |
             Where-Object { $_.UserName -ieq $Account })
}

Write-Output '=== Travelport Smartpoint Reset - Part 2: Stop and Restore ==='
Write-Output "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Output ''

# ----------------------------------------------------------------------------- input validation
$TargetUser = "$env:Action".Trim()
Write-Output "Raw value received in ENV:Action: '$TargetUser'"

#Be forgiving if the tech pastes FQDN
if ($TargetUser -match '\\') {
    $TargetUser = $TargetUser.Split('\')[-1]
    Write-Output "Domain prefix detected and removed. Using: '$TargetUser'"
}
if ($TargetUser -match '@') {
    $TargetUser = $TargetUser.Split('@')[0]
    Write-Output "UPN suffix detected and removed. Using: '$TargetUser'"
}

if ([string]::IsNullOrWhiteSpace($TargetUser)) {
    Write-Output ''
    Write-Output '*** SCRIPT FAILED ***'
    Write-Output 'Reason : The component variable "Action" was empty or contained only whitespace.'
    Write-Output '         Nothing was stopped and nothing was restored.'
    Write-Output 'Detail : This script needs a username to know whose backup folder to restore.'
    Write-Output 'Fix    : Re-run the component and enter the username, for example: fklein'
    return
}

if ($TargetUser -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    Write-Output ''
    Write-Output '*** SCRIPT FAILED ***'
    Write-Output "Reason : '$TargetUser' is not a valid username."
    Write-Output '         Only letters, digits, dot, dash and underscore are permitted, up to 64 characters.'
    Write-Output '         Nothing was stopped and nothing was restored.'
    Write-Output 'Detail : This value is used to build the source and destination folder paths.'
    Write-Output 'Fix    : Re-run the component with a plain username, for example: fklein'
    return
}

$Account = "$Domain\$TargetUser"
Write-Output "Target account: $Account"
Write-Output ''

# ----------------------------------------------------------------------------- resolve paths
$ProfilePath = Join-Path $ProfileRoot $TargetUser
$Source      = Join-Path $ProfilePath $BackupPath
$Destination = Join-Path $ProfilePath $RestorePath

Write-Output "Profile path: $ProfilePath"
Write-Output "Source      : $Source"
Write-Output "Destination : $Destination"
Write-Output ''

# Catches a mistyped username before anything else is attempted.
if (-not (Test-Path -LiteralPath $ProfilePath)) {
    Write-Output '*** SCRIPT FAILED ***'
    Write-Output "Reason : There is no user profile folder at $ProfilePath"
    Write-Output '         Nothing was stopped and nothing was restored.'
    Write-Output 'Detail : Either the username is misspelled, the user has never signed in to this'
    Write-Output '         computer, or the job was targeted at the wrong device.'
    Write-Output 'Fix    : Confirm the spelling of the username.'
    return
}

# Checked before stopping anything, so the user is not interrupted for a restore
# that could never have succeeded.
if (-not (Test-Path -LiteralPath $Source)) {
    Write-Output '*** SCRIPT FAILED ***'
    Write-Output "Reason : The backup folder does not exist for $Account."
    Write-Output "         Expected: $Source"
    Write-Output '         Nothing was stopped and nothing was restored.'
    Write-Output 'Detail : This was checked before stopping Smartpoint, so the user has not been interrupted.'
    Write-Output 'Fix    : Confirm a backup was taken for this user before Part 1 was run. If not,'
    Write-Output '         the folder will need to be rebuilt from another source.'
    return
}

$SourceCount = @(Get-ChildItem -LiteralPath $Source -Recurse -File -ErrorAction SilentlyContinue).Count
Write-Output "Backup folder found, containing $SourceCount file(s)."
Write-Output ''

# ----------------------------------------------------------------------------- stop the process
$Processes = Get-TargetProcess -Name $ProcessName -Account $Account

if ($Processes.Count -gt 0) {
    Write-Output "Found $($Processes.Count) running instance(s) of $ProcessName for $Account."

    foreach ($Process in $Processes) {
        Write-Output "Stopping $($Process.ProcessName) (PID $($Process.Id))"
        try {
            Stop-Process -Id $Process.Id -Force -ErrorAction Stop
        }
        catch {
            Write-Output "WARNING: Could not stop PID $($Process.Id)."
            Write-Output "         Reason: $($_.Exception.Message.Trim())"
        }
    }

    Write-Output "Waiting up to $KillTimeout seconds for the process to exit..."
    $Processes | Wait-Process -Timeout $KillTimeout -ErrorAction SilentlyContinue

    $Remaining = Get-TargetProcess -Name $ProcessName -Account $Account
    if ($Remaining.Count -gt 0) {
        Write-Output ''
        Write-Output '*** SCRIPT FAILED ***'
        Write-Output "Reason : $($Remaining.Count) instance(s) of $ProcessName are still running after"
        Write-Output "         $KillTimeout seconds, so the restore was deliberately NOT attempted."
        Write-Output '         Still running:'
        foreach ($Process in $Remaining) {
            Write-Output "           PID $($Process.Id)  $($Process.ProcessName)"
        }
        Write-Output 'Detail : Copying over files that Smartpoint holds open would produce a mix of'
        Write-Output '         old and new files, which looks like a successful restore but is not.'
        Write-Output '         Nothing was copied.'
        Write-Output 'Fix    : The process is most likely hung or protected. Have the user close'
        Write-Output '         Smartpoint manually, or end the task in Task Manager, then re-run.'
        return
    }

    Write-Output "Process stopped. Allowing $SettleSeconds seconds for file handles to release."
    Start-Sleep -Seconds $SettleSeconds
}
else {
    Write-Output "No instances of $ProcessName are running for $Account. Continuing to the restore step."
}

Write-Output ''

# ----------------------------------------------------------------------------- restore
if (-not (Test-Path -LiteralPath $Destination)) {
    Write-Output "Destination folder does not exist. Creating: $Destination"
    try {
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    }
    catch {
        Write-Output ''
        Write-Output '*** SCRIPT FAILED ***'
        Write-Output "Reason : The destination folder could not be created."
        Write-Output "         Path  : $Destination"
        Write-Output "         Error : $($_.Exception.Message.Trim())"
        Write-Output '         Nothing was copied.'
        Write-Output 'Fix    : Check permissions on the user profile and that the drive has free space.'
        return
    }
}

Write-Output 'Copying files...'
try {
    Copy-Item -Path (Join-Path $Source '*') -Destination $Destination -Recurse -Force -ErrorAction Stop

    $FileCount = @(Get-ChildItem -LiteralPath $Destination -Recurse -File -ErrorAction SilentlyContinue).Count
    Write-Output ''
    Write-Output '=== SUCCESS ==='
    Write-Output "The Travelport folder for $TargetUser was restored from backup."
    Write-Output "Source contained    : $SourceCount file(s)"
    Write-Output "Destination now has : $FileCount file(s)"
    if ($FileCount -lt $SourceCount) {
        Write-Output ''
        Write-Output 'WARNING: The destination has fewer files than the backup.'
    }
    Write-Output 'Have the user reopen Smartpoint and confirm their SmartButtons and custom scripts.'
}
catch {
    Write-Output ''
    Write-Output '*** SCRIPT FAILED ***'
    Write-Output "Reason : The copy did not complete."
    Write-Output "         Source     : $Source"
    Write-Output "         Destination: $Destination"
    Write-Output "         Error      : $($_.Exception.Message.Trim())"
    Write-Output 'Detail : The restore may be partial, so the destination folder could hold an'
    Write-Output '         incomplete set of files. Do not assume the user is back to normal.'
    Write-Output 'Fix    : Confirm Smartpoint is fully closed and the drive has free space, then'
    Write-Output '         re-run. The copy overwrites, so re-running is safe.'
}

Write-Output ''
Write-Output "Finished: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
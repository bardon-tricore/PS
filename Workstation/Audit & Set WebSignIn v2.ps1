<#
.SYNOPSIS
    Enables, Disables, or Audits Web Sign-In on a Windows endpoint, with a mandatory 
    expiration that automatically reverts the setting.

.DESCRIPTION
    Web Sign-In allows a user to authenticate to Windows using a web-based credential provider
    (Temporary Access Pass).

    When the Enable action runs, the script:
        1. Validates that Expiration is a whole number of hours between 1 and 12.
        2. Registers a scheduled task that will revert the setting once that window elapses.
        3. Sets EnableWebSignIn to 1.

    Re-running Enable overwrites the pending task, which restarts the expiration window.
    Running Disable reverts the value immediately and removes any pending task.

    Registry:
        Path:   HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Authentication
        DWORD:  EnableWebSignIn
        Value:  1 (Enabled) / 0 (Disabled)

    Available Actions:
        Enable  - Sets EnableWebSignIn to 1 and schedules the automatic revert.
        Disable - Sets EnableWebSignIn to 0 and removes any pending revert.
        Audit   - Reports the current value and any pending revert, without making changes.

.INPUTS
    Datto RMM Environment Variables:
        Action (Selection)     - Enable, Disable, or Audit
        Expiration (Selection) - 1 through 12. Hours that Web Sign-In remains enabled.
                                 Required by Enable. Ignored by Disable and Audit.

.OUTPUTS
    StdOut messages tagged with [INFO], [SUCCESS], [AUDIT], or [FAILED].

    Exit Code 0 - Action completed successfully.
    Exit Code 1 - Action failed, or an invalid Action or Expiration was supplied.

.NOTES
    Name:       Set-EnableWebSignIn.ps1
    Version:    3.0
    Date:       2026-08-19

    - Requires SYSTEM or local Administrator rights. Datto RMM runs as SYSTEM by default.
    - No files are written to the endpoint. The revert payload lives in the scheduled task's own
      argument string and is stored in the task XML as readable plain text, not encoded, so it can
      be reviewed by an administrator or an EDR product at any time.
    - A missing or out-of-range Expiration aborts the job before anything is
      written, so Web Sign-In is never left enabled without a scheduled revert. The scheduled task
      is also created before the registry value is set. If the task cannot be registered, the value
      is never enabled.
    - The revert task uses StartWhenAvailable, so an endpoint that is off or asleep at the
      expiration time runs the revert at its next opportunity rather than skipping it.
    - The revert deliberately unregisters the task only after a successful registry write. A failed
      revert therefore leaves the task in place, where the Audit action will surface it. An Audit
      reporting value '0' and no pending task is confirmation the revert completed.
    - The PolicyManager hive is normally owned by the MDM stack. On an Intune-enrolled device a
      policy sync can overwrite or remove these values, so prefer managing this through Intune
      where enrollment exists. This script is intended for non-enrolled or hybrid endpoints.
    - Web Sign-In requires Windows 10 1903 or later, or Windows 11.
    - Users must sign out or reboot before the credential provider appears on the lock screen, and
      an already-displayed lock screen may keep showing it until the logon UI is refreshed.

    Scheduled Task:  \DattoRMM\DisableWebSignIn
#>

#region Datto RMM Environment Variables

$Action     = "$ENV:Action".Trim()
$Expiration = "$ENV:Expiration".Trim()

#endregion Datto RMM Environment Variables

#region Script Variables

$RegPath      = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Authentication"
$ValueName    = "EnableWebSignIn"

$TaskName     = "DisableWebSignIn"
$TaskPath     = "\DattoRMM\"

$MinimumHours = 1
$MaximumHours = 12

#endregion Script Variables

#region Functions

function Remove-ExpirationTask {
    <#
        Removes the pending revert task. Safe to call when no task exists.
    #>

    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue

    if ($null -ne $ExistingTask) {
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
        Write-Output "[INFO] Removed pending expiration task '$TaskName'."
    }
}

function New-ExpirationTask {
    param (
        [Parameter(Mandatory = $true)]
        [int]$Hours
    )
    $RevertCommand =
        "if (-not (Test-Path '$RegPath')) { New-Item -Path '$RegPath' -Force | Out-Null }; " +
        "New-ItemProperty -Path '$RegPath' -Name '$ValueName' -PropertyType DWord -Value 0 -Force -ErrorAction Stop | Out-Null; " +
        "Unregister-ScheduledTask -TaskName '$TaskName' -TaskPath '$TaskPath' -Confirm:`$false"

    $TaskArgument = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command `"$RevertCommand`""

    $ExpirationTime = (Get-Date).AddHours($Hours)

    $TaskAction = New-ScheduledTaskAction `
        -Execute "$ENV:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -Argument $TaskArgument

    $TaskTrigger = New-ScheduledTaskTrigger -Once -At $ExpirationTime

    $TaskPrincipal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    $TaskSettings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

    Register-ScheduledTask `
        -TaskName $TaskName `
        -TaskPath $TaskPath `
        -Action $TaskAction `
        -Trigger $TaskTrigger `
        -Principal $TaskPrincipal `
        -Settings $TaskSettings `
        -Description "Reverts EnableWebSignIn to 0 when the approved window expires, then removes itself." `
        -Force | Out-Null

    return $ExpirationTime
}

#endregion Functions

#region Action

switch ($Action) {

    "Enable" {

        # Validate the expiration window before anything is written to the endpoint.
        $ExpirationHours = 0

        if (-not [int]::TryParse($Expiration, [ref]$ExpirationHours)) {
            Write-Output "[FAILED] Expiration must be a whole number of hours."
            Write-Output "[FAILED] Expected: $MinimumHours through $MaximumHours"
            Write-Output "[FAILED] Received: '$Expiration'"
            exit 1
        }

        if ($ExpirationHours -lt $MinimumHours -or $ExpirationHours -gt $MaximumHours) {
            Write-Output "[FAILED] Expiration is outside the permitted range."
            Write-Output "[FAILED] Expected: $MinimumHours through $MaximumHours"
            Write-Output "[FAILED] Received: '$Expiration'"
            exit 1
        }

        # Schedule the revert first so Web Sign-In is never enabled without an expiration.
        try {
            $ExpirationTime = New-ExpirationTask -Hours $ExpirationHours
            Write-Output "[INFO] Expiration task scheduled for '$ExpirationTime'."
        }
        catch {
            Write-Output "[FAILED] Unable to schedule the expiration task. '$ValueName' was not enabled."
            Write-Output "[FAILED] Task Name: $TaskPath$TaskName"
            Write-Output "[FAILED] Error: $($_.Exception.Message)"
            exit 1
        }

        try {
            if (-not (Test-Path $RegPath)) {
                Write-Output "[INFO] Registry path does not exist. Creating: $RegPath"
                New-Item -Path $RegPath -Force | Out-Null
            }

            New-ItemProperty `
                -Path $RegPath `
                -Name $ValueName `
                -PropertyType DWord `
                -Value 1 `
                -Force | Out-Null

            Write-Output "[SUCCESS] '$ValueName' configured with value '1' for $ExpirationHours hour(s)."
            Write-Output "[INFO] A sign-out or reboot is required before Web Sign-In appears on the lock screen."
            exit 0
        }
        catch {
            Write-Output "[FAILED] Unable to configure '$ValueName' with value '1'."
            Write-Output "[FAILED] Registry Path: $RegPath"
            Write-Output "[FAILED] Error: $($_.Exception.Message)"

            # Do not leave an orphaned revert task behind for a value that was never set.
            try {
                Remove-ExpirationTask
            }
            catch {
                Write-Output "[FAILED] Unable to remove the orphaned expiration task."
                Write-Output "[FAILED] Error: $($_.Exception.Message)"
            }

            exit 1
        }
    }

    "Disable" {
        try {
            if (-not (Test-Path $RegPath)) {
                Write-Output "[INFO] Registry path does not exist. Creating: $RegPath"
                New-Item -Path $RegPath -Force | Out-Null
            }

            New-ItemProperty `
                -Path $RegPath `
                -Name $ValueName `
                -PropertyType DWord `
                -Value 0 `
                -Force | Out-Null

            Write-Output "[SUCCESS] '$ValueName' configured with value '0'."

            Remove-ExpirationTask

            Write-Output "[INFO] A sign-out or reboot is required before the change takes effect."
            exit 0
        }
        catch {
            Write-Output "[FAILED] Unable to configure '$ValueName' with value '0'."
            Write-Output "[FAILED] Registry Path: $RegPath"
            Write-Output "[FAILED] Error: $($_.Exception.Message)"
            exit 1
        }
    }

    "Audit" {
        try {
            if (-not (Test-Path $RegPath)) {
                Write-Output "[AUDIT] Registry path does not exist: $RegPath"
            }
            else {
                $Property = Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction SilentlyContinue

                if ($null -eq $Property) {
                    Write-Output "[AUDIT] Registry value '$ValueName' does not exist."
                }
                else {
                    Write-Output "[AUDIT] Registry value '$ValueName' exists and is set to '$($Property.$ValueName)'."
                }
            }

            $ExistingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue

            if ($null -eq $ExistingTask) {
                Write-Output "[AUDIT] No expiration task is currently scheduled."
            }
            else {
                $TaskInfo = $ExistingTask | Get-ScheduledTaskInfo
                Write-Output "[AUDIT] Expiration task is scheduled to run at '$($TaskInfo.NextRunTime)'."
            }

            exit 0
        }
        catch {
            Write-Output "[FAILED] Unable to audit configuration."
            Write-Output "[FAILED] Registry Path: $RegPath"
            Write-Output "[FAILED] Error: $($_.Exception.Message)"
            exit 1
        }
    }

    default {
        Write-Output "[FAILED] Invalid Action specified."
        Write-Output "[FAILED] Expected: Enable, Disable, or Audit"
        Write-Output "[FAILED] Received: '$Action'"
        exit 1
    }
}

#endregion Action
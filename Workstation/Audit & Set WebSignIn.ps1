<#
.SYNOPSIS
    Enables, Disables, or Audits Web Sign-In on a Windows endpoint via the PolicyManager registry hive.

.DESCRIPTION
    Web Sign-In allows a user to authenticate to Windows using TAP

    This script creates the "Authentication" key beneath the device PolicyManager hive if it does
    not already exist, then writes the EnableWebSignIn DWORD value.

    Path:   HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Authentication
    DWORD:  EnableWebSignIn
    Value:  1 (Enabled) / 0 (Disabled)

    Available Actions:
        Enable  - Creates the key if missing and sets EnableWebSignIn to 1.
        Disable - Creates the key if missing and sets EnableWebSignIn to 0.
        Audit   - Reports the current configuration without making any changes.

.INPUTS
    Datto RMM Environment Variable:
        Action (Selection / String) - Enable, Disable, or Audit

.OUTPUTS
    StdOut messages tagged with [INFO], [SUCCESS], [AUDIT], or [FAILED].

    Exit Code 0 - Action completed successfully.
    Exit Code 1 - Action failed or an invalid Action was supplied.

.NOTES
    Name:       Set EnableWebSignIn.ps1
    Version:    1.0
    Date:       2026-08-17

    - The PolicyManager hive is normally owned by the MDM stack. On an Intune-enrolled device a
      policy sync can overwrite or remove these values.
    - Users must sign out or reboot before the credential provider appears on the lock screen.
#>

#region Datto RMM Environment Variables

$Action = "$ENV:Action".Trim()

#endregion Datto RMM Environment Variables

#region Script Variables

$RegPath   = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Authentication"
$ValueName = "EnableWebSignIn"

#endregion Script Variables

#region Action

switch ($Action) {

    "Enable" {
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

            Write-Output "[SUCCESS] '$ValueName' configured with value '1'."
            Write-Output "[INFO] A sign-out or reboot is required before Web Sign-In appears on the lock screen."
            exit 0
        }
        catch {
            Write-Output "[FAILED] Unable to configure '$ValueName' with value '1'."
            Write-Output "[FAILED] Registry Path: $RegPath"
            Write-Output "[FAILED] Error: $($_.Exception.Message)"
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
                exit 0
            }

            $Property = Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction SilentlyContinue

            if ($null -eq $Property) {
                Write-Output "[AUDIT] Registry value '$ValueName' does not exist."
            }
            else {
                Write-Output "[AUDIT] Registry value '$ValueName' exists and is set to '$($Property.$ValueName)'."
            }

            exit 0
        }
        catch {
            Write-Output "[FAILED] Unable to audit registry configuration."
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
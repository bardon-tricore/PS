# Check for and remove Dell SupportAssist Remediation 5.5.16.0
try {
    $RemediationFound = $false

    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* ,
                     HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -eq "Dell SupportAssist Remediation" -and
        $_.DisplayVersion -eq "5.5.16.0"
    } |
    ForEach-Object {
        $RemediationFound = $true

        Write-Output "Dell SupportAssist Remediation 5.5.16.0 found. Starting uninstall."

        if ($_.UninstallString) {
            Start-Process -FilePath "cmd.exe" `
                -ArgumentList "/c $($_.UninstallString) /quiet /norestart" `
                -Wait `
                -NoNewWindow `
                -ErrorAction Stop

            Write-Output "Dell SupportAssist Remediation removed successfully."
        }
        else {
            Write-Output "Dell SupportAssist Remediation uninstall string not found."
        }
    }

    if (-not $RemediationFound) {
        Write-Output "Dell SupportAssist Remediation 5.5.16.0 not installed."
    }
}
catch {
    Write-Output "ERROR removing Dell SupportAssist Remediation: $($_.Exception.Message)"
}

# Check for and remove Dell SupportAssist OS Recovery Plugin for Dell Update 5.5.16.0
try {
    $RecoveryPluginFound = $false

    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* ,
                     HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -eq "Dell SupportAssist OS Recovery Plugin for Dell Update" -and
        $_.DisplayVersion -eq "5.5.16.0"
    } |
    ForEach-Object {
        $RecoveryPluginFound = $true

        Write-Output "Dell SupportAssist OS Recovery Plugin for Dell Update 5.5.16.0 found. Starting uninstall."

        if ($_.UninstallString) {
            Start-Process -FilePath "cmd.exe" `
                -ArgumentList "/c $($_.UninstallString) /quiet /norestart" `
                -Wait `
                -NoNewWindow `
                -ErrorAction Stop

            Write-Output "Dell SupportAssist OS Recovery Plugin removed successfully."
        }
        else {
            Write-Output "Dell SupportAssist OS Recovery Plugin uninstall string not found."
        }
    }

    if (-not $RecoveryPluginFound) {
        Write-Output "Dell SupportAssist OS Recovery Plugin for Dell Update 5.5.16.0 not installed."
    }
}
catch {
    Write-Output "ERROR removing Dell SupportAssist OS Recovery Plugin: $($_.Exception.Message)"
}
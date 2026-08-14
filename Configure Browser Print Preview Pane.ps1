##This script assumes Chrome is installed for all users (which it is on all TriCore Immy Bot'd devices)
##This script creates 2 DWORDs with the registry (see below) forcing the browsers to print using the system dialog rather than their own Preview Panes

##HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge\UseSystemPrintDialog
##HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\DisablePrintPreview

$ErrorActionPreference = "Stop"
$Failed = $false

# DRMM inputs
$BrowserSelect = "$ENV:BrowserSelect"
$Action = "$ENV:Action"

# CHROME
if ($BrowserSelect -in @("Chrome", "All")) {

    $ChromePath    = "HKLM:\SOFTWARE\Policies\Google\Chrome"
    $ChromeValue   = "DisablePrintPreview"
    $ChromeDesired = 1

    try {
        if ($Action -eq "Audit") {
            if (-not (Test-Path $ChromePath)) {
                Write-Output "[AUDIT] Chrome policy path does not exist."
            }
            else {
                $reg = Get-ItemProperty -Path $ChromePath -ErrorAction SilentlyContinue
                $CurrentValue = if ($reg) { $reg.$ChromeValue } else { $null }

                if ($null -eq $CurrentValue) {
                    Write-Output "[AUDIT] Chrome - $ChromeValue does not exist."
                }
                else {
                    Write-Output "[AUDIT] Chrome - $ChromeValue = $CurrentValue"
                }
            }
        }

        elseif ($Action -eq "Apply") {
            if (-not (Test-Path $ChromePath)) {
                New-Item -Path $ChromePath -Force | Out-Null
                Write-Output "[CREATED] Registry path: $ChromePath"
            }

            $reg = Get-ItemProperty -Path $ChromePath -ErrorAction SilentlyContinue
            $CurrentValue = if ($reg) { $reg.$ChromeValue } else { $null }

            if ($null -eq $CurrentValue) {
                New-ItemProperty -Path $ChromePath -Name $ChromeValue -Value $ChromeDesired -PropertyType DWord -Force | Out-Null
                Write-Output "[CREATED] Chrome - $ChromeValue = $ChromeDesired"
            }
            elseif ($CurrentValue -eq $ChromeDesired) {
                Write-Output "[EXISTS] Chrome - $ChromeValue already set to $ChromeDesired"
            }
            else {
                Set-ItemProperty -Path $ChromePath -Name $ChromeValue -Value $ChromeDesired
                Write-Output "[UPDATED] Chrome - $ChromeValue changed from $CurrentValue to $ChromeDesired"
            }
        }

        elseif ($Action -eq "Remove") {
            if (-not (Test-Path $ChromePath)) {
                Write-Output "[SKIPPED] Chrome registry path does not exist."
            }
            else {
                $reg = Get-ItemProperty -Path $ChromePath -ErrorAction SilentlyContinue
                $CurrentValue = if ($reg) { $reg.$ChromeValue } else { $null }

                if ($null -eq $CurrentValue) {
                    Write-Output "[SKIPPED] Chrome - $ChromeValue does not exist."
                }
                else {
                    Remove-ItemProperty -Path $ChromePath -Name $ChromeValue -ErrorAction SilentlyContinue
                    Write-Output "[REMOVED] Chrome - $ChromeValue"
                }
            }
        }
    }
    catch {
        Write-Warning "[FAILED] Chrome policy configuration failed. $($_.Exception.Message)"
        $Failed = $true
    }
}

# EDGE
if ($BrowserSelect -in @("Edge", "All")) {

    $EdgePath    = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
    $EdgeValue   = "UseSystemPrintDialog"
    $EdgeDesired = 1

    try {
        if ($Action -eq "Audit") {
            if (-not (Test-Path $EdgePath)) {
                Write-Output "[AUDIT] Edge policy path does not exist."
            }
            else {
                $reg = Get-ItemProperty -Path $EdgePath -ErrorAction SilentlyContinue
                $CurrentValue = if ($reg) { $reg.$EdgeValue } else { $null }

                if ($null -eq $CurrentValue) {
                    Write-Output "[AUDIT] Edge - $EdgeValue does not exist."
                }
                else {
                    Write-Output "[AUDIT] Edge - $EdgeValue = $CurrentValue"
                }
            }
        }

        elseif ($Action -eq "Apply") {
            if (-not (Test-Path $EdgePath)) {
                New-Item -Path $EdgePath -Force | Out-Null
                Write-Output "[CREATED] Registry path: $EdgePath"
            }

            $reg = Get-ItemProperty -Path $EdgePath -ErrorAction SilentlyContinue
            $CurrentValue = if ($reg) { $reg.$EdgeValue } else { $null }

            if ($null -eq $CurrentValue) {
                New-ItemProperty -Path $EdgePath -Name $EdgeValue -Value $EdgeDesired -PropertyType DWord -Force | Out-Null
                Write-Output "[CREATED] Edge - $EdgeValue = $EdgeDesired"
            }
            elseif ($CurrentValue -eq $EdgeDesired) {
                Write-Output "[EXISTS] Edge - $EdgeValue already set to $EdgeDesired"
            }
            else {
                Set-ItemProperty -Path $EdgePath -Name $EdgeValue -Value $EdgeDesired
                Write-Output "[UPDATED] Edge - $EdgeValue changed from $CurrentValue to $EdgeDesired"
            }
        }

        elseif ($Action -eq "Remove") {
            if (-not (Test-Path $EdgePath)) {
                Write-Output "[SKIPPED] Edge registry path does not exist."
            }
            else {
                $reg = Get-ItemProperty -Path $EdgePath -ErrorAction SilentlyContinue
                $CurrentValue = if ($reg) { $reg.$EdgeValue } else { $null }

                if ($null -eq $CurrentValue) {
                    Write-Output "[SKIPPED] Edge - $EdgeValue does not exist."
                }
                else {
                    Remove-ItemProperty -Path $EdgePath -Name $EdgeValue -ErrorAction SilentlyContinue
                    Write-Output "[REMOVED] Edge - $EdgeValue"
                }
            }
        }
    }
    catch {
        Write-Warning "[FAILED] Edge policy configuration failed. $($_.Exception.Message)"
        $Failed = $true
    }
}

# SUMMARY
if ($Failed) {
    Write-Error "Script completed with one or more errors."
    exit 1
}
else {
    Write-Output "Script completed successfully."
    exit 0
}
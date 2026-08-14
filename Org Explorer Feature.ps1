##Uses a registry key to Review or Adjust the "Org Explorer Feature" within Outlook

##Enable, Disable, and Audit actions are available. 

##Path: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Search Show
##DWORD: DisableOrgExplorerSearch
##Value: 1/0

# Datto RMM Enviroment Variable
$Action = "$ENV:Action"

$RegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Search Show"
$ValueName = "DisableOrgExplorerSearch"

#Action
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
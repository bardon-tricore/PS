# Check if Remote Desktop (RDP) is enabled and store result in Datto RMM UDF

try {

    $RDPStatus = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction Stop

    if ($RDPStatus.fDenyTSConnections -eq 0) {

        $RDPStatusUDFOutput = "RDP Enabled"
        Write-Output "RDP is ENABLED on this machine."

    }
    elseif ($RDPStatus.fDenyTSConnections -eq 1) {

        $RDPStatusUDFOutput = "RDP Disabled"
        Write-Output "RDP is DISABLED on this machine."

    }
    else {

        $RDPStatusUDFOutput = "Unknown"
        Write-Output "Unknown RDP status detected."

    }

}
catch {

    $RDPStatusUDFOutput = "Error"
    Write-Output "Unable to determine RDP status. Error: $($_.Exception.Message)"

}

# Write result to Datto RMM UDF
New-ItemProperty HKLM:\SOFTWARE\CentraStage -Name "custom$ENV:UDFField" -PropertyType string -Value "$RDPStatusUDFOutput" -Force | Out-Null


### Below was removed from the deployed script

# Disable Remote Desktop
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 1

# Disable Remote Desktop Firewall Rules
Disable-NetFirewallRule -DisplayGroup "Remote Desktop"

# Verify RDP Status
if ((Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server").fDenyTSConnections -eq 1) {
    Write-Host "RDP has been successfully disabled."
}
else {
    Write-Host "Failed to disable RDP."
}
}
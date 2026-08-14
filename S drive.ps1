$DriveLetter = "S"
$SharePath = "\\192.168.253.224\BND Sharepoint"

Write-Host "Mapping drive $DriveLetter`: to $SharePath..."

# Remove any existing 
if (Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue) {
    Write-Host "Existing drive mapping found. Removing..."
    Remove-PSDrive -Name $DriveLetter -Force
}

# Nuke existing
cmd.exe /c "net use $DriveLetter`: /delete /y" | Out-Null

# Create a persistent mapping
New-PSDrive -Name $DriveLetter `
            -PSProvider FileSystem `
            -Root $SharePath `
            -Persist

Write-Host "Drive $DriveLetter`: mapped successfully."
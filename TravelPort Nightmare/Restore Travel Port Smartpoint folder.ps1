#Restore backup files for Fran (SmartButtons, Custom Actions/Scripts)

$Source = "C:\Users\fklein\AppData\Roaming\Travelport BACKUP"
$Destination = "C:\Users\fklein\AppData\Roaming\Travelport"

if (-not (Test-Path $Source)) {
    Write-Error "Source folder does not exist: $Source"
    exit 1
}

# Create destination if it doesn't exist
if (-not (Test-Path $Destination)) {
    New-Item -ItemType Directory -Path $Destination | Out-Null
}

Write-Host "Copying files..."

Copy-Item -Path "$Source\*" `
          -Destination $Destination `
          -Recurse `
          -Force

Write-Host "Copy completed successfully."
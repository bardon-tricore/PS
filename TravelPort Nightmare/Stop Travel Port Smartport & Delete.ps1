# Stop Travelport Smartpoint for specific user
$User = "CTSCV\fklein"
$ProcessName = "Travelport.Smartpoint.App"

$Processes = Get-Process -Name $ProcessName -IncludeUserName -ErrorAction SilentlyContinue |
    Where-Object { $_.UserName -eq $User }

if ($Processes) {
    foreach ($Process in $Processes) {
        Write-Host "Stopping $($Process.ProcessName) (PID $($Process.Id)) for $($Process.UserName)"
        Stop-Process -Id $Process.Id -Force
    }

    Write-Host "Completed stopping Smartpoint."
}
else {
    Write-Host "No instances of $ProcessName found running for $User."
}

# Wait for Smartpoint to fully release files
Write-Host "Waiting 15 seconds for Smartpoint to close..."
Start-Sleep -Seconds 15

# Delete Travelport folder
$Folder = "C:\Users\fklein\AppData\Roaming\Travelport"

if (Test-Path $Folder) {
    Write-Host "Deleting $Folder..."
    Remove-Item -Path $Folder -Recurse -Force
    Write-Host "Folder deleted successfully."
}
else {
    Write-Host "Folder does not exist."
}

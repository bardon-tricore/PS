$LibraryId = "c9f9fa27-5ec5-46cf-914f-5d61ebf672c4"

Write-Host "Starting SharePoint library sync for Library ID: $LibraryId"

Invoke-ImmyCommand -RunAsUser {
    # Note: These inner hosts may not stream back to the ImmyBot console, but they won't break anything.
    Write-Host "Running in user context: $env:USERNAME"
    
    $SyncUrl = "odopen://sync?libraryId=$using:LibraryId"
    Write-Host "Launching sync URL: $SyncUrl"

    Start-Process $SyncUrl -UseShellExecute

    Write-Host "Sync request submitted to OneDrive client."
}

Write-Host "SharePoint library sync command completed."


### ABOVE FAILED
# Define your specific target target IDs
$SiteId     = "XXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
$WebId      = "XXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
$LibraryId  = "c9f9fa27-5ec5-46cf-914f-5d61ebf672c4" # (This is your listId)
$WebUrl     = "https://yourtenant.sharepoint.com/sites/YourSite"

Write-Host "Starting SharePoint library sync configuration..."

Invoke-ImmyCommand -RunAsUser {
    # Dynamically grab the active logged-in user's email 
    # (Since ImmyBot runs this block as the user, we can grab their environment UPN or email)
    $UserEmail = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.EmailAddress
    if (-not $UserEmail) {
        # Fallback to a common Entra ID environment variable if domain isn't fully synced locally
        $UserEmail = $env:UPN
    }

    # Construct the full, structurally complete Microsoft sync URL
    $SyncUrl = "odopen://sync/?siteId=$using:SiteId&webId=$using:WebId&listId=$using:LibraryId&webUrl=$using:WebUrl&userEmail=$UserEmail"

    Write-Host "Launching fully-formed sync URL for $UserEmail"
    Start-Process $SyncUrl -UseShellExecute
}

Write-Host "SharePoint library sync command completed."
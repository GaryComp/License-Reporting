$userId = "70ee6e97-4f36-4270-9242-59ba9e1a09f2"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

# Step 1: Try to force-provision by calling the drive endpoint
Write-Host "[1] Attempting to provision OneDrive..." -ForegroundColor Cyan
try {
    $drive = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId/drive" -Headers $headers
    Write-Host "Drive provisioned!" -ForegroundColor Green
    Write-Host "Drive ID  : $($drive.id)"
    Write-Host "Web URL   : $($drive.webUrl)"
} catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Status code: $statusCode"
    Write-Warning "Response: $_"
}

# Step 2: Check if a OneDrive site exists for this user via SharePoint admin
Write-Host "`n[2] Looking up OneDrive site via SharePoint..." -ForegroundColor Cyan
try {
    $sites = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=sky" -Headers $headers
    $sites.value | Select-Object displayName, webUrl, id | Format-Table -AutoSize
} catch {
    Write-Warning "Site search failed: $_"
}

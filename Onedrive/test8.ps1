$userId = "70ee6e97-4f36-4270-9242-59ba9e1a09f2"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

# Try 1: User drive via object ID with explicit host header
Write-Host "[1] Trying user drive with x-anchormailbox header..." -ForegroundColor Cyan
$headers2 = @{
    Authorization = "Bearer $($tokenJson.accessToken)"
    "x-anchormailbox" = "sky@basecampresorts.com"
}
try {
    $drive = Invoke-RestMethod -Uri ("https://graph.microsoft.com/v1.0/users/" + $userId + "/drive") -Headers $headers2
    Write-Host "Drive ID: $($drive.id)" -ForegroundColor Green
    Write-Host "Web URL : $($drive.webUrl)" -ForegroundColor Green
} catch {
    Write-Warning "Try 1 failed: $_"
}

# Try 2: Access via SharePoint REST API directly
Write-Host "`n[2] Trying SharePoint REST API directly..." -ForegroundColor Cyan
$spToken = az account get-access-token --resource "https://basecampresorts931-my.sharepoint.com" | ConvertFrom-Json
$spHeaders = @{ Authorization = "Bearer $($spToken.accessToken)" }
try {
    $sp = Invoke-RestMethod -Uri "https://basecampresorts931-my.sharepoint.com/personal/sky_basecampresorts_com/_api/web" -Headers $spHeaders
    Write-Host "SharePoint site accessible: $($sp.Url)" -ForegroundColor Green
} catch {
    Write-Warning "Try 2 failed: $_"
}

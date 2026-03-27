$TenantName = "basecampresorts931"
$siteId = "basecampresorts931-my.sharepoint.com,3426d4c0-7ba7-48e2-a7e3-3ae1b1b55505,34b7dd76-4b30-4d55-94ad-e959d02e6fe0"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

$url1 = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/drives"
Write-Host "Calling: $url1" -ForegroundColor Yellow

$raw = Invoke-RestMethod -Uri $url1 -Headers $headers
Write-Host "Raw response:" -ForegroundColor Cyan
$raw | ConvertTo-Json -Depth 5

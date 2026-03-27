$TenantName = "basecampresorts931"
$siteId = "basecampresorts931-my.sharepoint.com,3426d4c0-7ba7-48e2-a7e3-3ae1b1b55505,34b7dd76-4b30-4d55-94ad-e959d02e6fe0"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

$url1 = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/drives"
Write-Host "[1] Listing drives on site..." -ForegroundColor Cyan
$drives = Invoke-RestMethod -Uri $url1 -Headers $headers
$drives.value | Select-Object id, name, driveType, webUrl | Format-Table -AutoSize

Write-Host "[2] Listing root folders..." -ForegroundColor Cyan
$driveId = $drives.value[0].id
$url2 = "https://graph.microsoft.com/v1.0/drives/" + $driveId + "/root/children"
$root = Invoke-RestMethod -Uri $url2 -Headers $headers
$root.value | Select-Object name, @{n="Type";e={if($_.folder){"Folder"}else{"File"}}}, @{n="SizeGB";e={[math]::Round($_.size/1GB,2)}} | Format-Table -AutoSize

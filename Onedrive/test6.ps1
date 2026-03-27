$TenantName = "basecampresorts931"
$siteId = "basecampresorts931-my.sharepoint.com,3426d4c0-7ba7-48e2-a7e3-3ae1b1b55505,34b7dd76-4b30-4d55-94ad-e959d02e6fe0"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

$url1 = "https://graph.microsoft.com/v1.0/sites/" + $siteId + "/drives"
$drives = Invoke-RestMethod -Uri $url1 -Headers $headers

# Print all drives so we can see what we got
Write-Host "Available drives:" -ForegroundColor Cyan
$drives.value | ForEach-Object {
    Write-Host "  Name: $($_.name) | ID: $($_.id) | Type: $($_.driveType)"
}

# Use the first drive ID explicitly
$driveId = $drives.value[0].id
Write-Host "`nUsing Drive ID: $driveId" -ForegroundColor Yellow

$url2 = "https://graph.microsoft.com/v1.0/drives/" + $driveId + "/root/children"
Write-Host "Calling: $url2" -ForegroundColor Yellow

$root = Invoke-RestMethod -Uri $url2 -Headers $headers
$root.value | Select-Object name, @{n="Type";e={if($_.folder){"Folder"}else{"File"}}}, @{n="SizeGB";e={[math]::Round($_.size/1GB,2)}} | Format-Table -AutoSize

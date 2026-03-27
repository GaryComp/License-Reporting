$userId = "70ee6e97-4f36-4270-9242-59ba9e1a09f2"
$TenantName = "basecampresorts931"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

# Try with explicit site URL via SharePoint host
$siteUrl = "https://$TenantName-my.sharepoint.com/personal/sky_basecampresorts_com"
Write-Host "Trying site: $siteUrl" -ForegroundColor Cyan

$site = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/${TenantName}-my.sharepoint.com:/personal/sky_basecampresorts_com" -Headers $headers
Write-Host "Site ID: $($site.id)" -ForegroundColor Green

# Now try the drive via site ID
$drive = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive/root/children" -Headers $headers
$drive.value | Where-Object { $_.folder } | Select-Object name, id | Format-Table -AutoSize

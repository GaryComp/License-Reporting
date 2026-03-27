$userId = "70ee6e97-4f36-4270-9242-59ba9e1a09f2"

$tokenJson = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($tokenJson.accessToken)" }

# Check who WE are authenticated as
Write-Host "[1] Checking authenticated identity..." -ForegroundColor Cyan
$me = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers $headers
Write-Host "Logged in as : $($me.userPrincipalName)"
Write-Host "Object ID    : $($me.id)"

# Check our assigned roles
Write-Host "`n[2] Checking directory roles..." -ForegroundColor Cyan
$roles = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me/memberOf" -Headers $headers
$roles.value | Select-Object displayName, '@odata.type' | Format-Table -AutoSize

# Try accessing drive with beta endpoint (sometimes works when v1.0 doesn't)
Write-Host "`n[3] Trying beta endpoint..." -ForegroundColor Cyan
try {
    $drive = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/users/$userId/drive/root/children" -Headers $headers
    $drive.value | Where-Object { $_.folder } | Select-Object name, id | Format-Table -AutoSize
} catch {
    Write-Warning "Beta endpoint failed: $_"
}

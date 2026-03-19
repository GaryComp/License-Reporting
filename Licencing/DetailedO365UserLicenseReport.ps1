Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserNamesFile,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force


# ------------------------- Export User License Details -------------------------
Function Export-UserLicenseDetails {
    param (
        [Parameter(Mandatory=$true)][object]$User,
        [string]$ExportCSV,
        [object]$SkuMapping,
        [object]$ServiceMapping
    )

    $UPN = $User.UserPrincipalName
    $DisplayName = $User.DisplayName
    $Country = if ($User.Country) { $User.Country } else { "-" }

    Write-Progress -Activity "Exported user count: $Global:LicensedUserCount" -Status "Processing: $UPN"

    $SKUs = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction SilentlyContinue
    if (!$SKUs) { return }

    foreach ($Sku in $SKUs) {
        $SkuId = $Sku.SkuPartNumber.Trim()
        $FriendlyLicense = $SkuId
        if ($SkuMapping) {
            $match = $SkuMapping | Where-Object { $_.SkuPartNumber -and ($_.SkuPartNumber.Trim().ToLower() -eq $SkuId.ToLower()) }
            if ($match) { $FriendlyLicense = $match.Product_Display_Name }
        }

        foreach ($Service in $Sku.ServicePlans) {
            $ServiceName = $Service.ServicePlanName.Trim()
            $FriendlyService = $ServiceName
            if ($ServiceMapping) {
                $match = $ServiceMapping | Where-Object { $_.Service_Plan_Name -and ($_.Service_Plan_Name.Trim().ToLower() -eq $ServiceName.ToLower()) }
                if ($match) { $FriendlyService = $match.ServicePlanDisplayName }
            }

            [PSCustomObject]@{
                DisplayName                = $DisplayName
                UserPrincipalName          = $UPN
                Country                    = $Country
                LicenseSkuPartNumber       = $SkuId
                LicenseFriendlyName        = $FriendlyLicense
                ServicePlanName            = $ServiceName
                ServicePlanFriendlyName    = $FriendlyService
                ProvisioningStatus         = $Service.ProvisioningStatus
            } | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
        }
    }
}

# ------------------------- Close Connection -------------------------
Function Close-Connection {
    Disconnect-MgGraph | Out-Null
    Exit
}

# ------------------------- Main -------------------------
Function main {
    Write-Host "`nNote: For best results, run this in a fresh PowerShell window." -ForegroundColor Yellow

    $timestamp = (Get-Date -Format "yyyy-MMM-dd-ddd hh-mm tt").ToString()
    $ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
    if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
    $ExportCSV = Join-Path $ExportsDir "DetailedO365UserLicenseReport_$timestamp.csv"

    # Try to load mapping file (optional)
    $SkuMapping = $null
    $ServiceMapping = $null
    $skuCsvPath = Join-Path $PSScriptRoot '..' 'Supporting_Files' 'Product names and service plan identifiers for licensing.csv'
    if (Test-Path $skuCsvPath) {
        try {
            $csvData = Import-Csv -Path $skuCsvPath
            if ($csvData | Get-Member -Name "SkuPartNumber" -ErrorAction SilentlyContinue) {
                $SkuMapping = $csvData | Where-Object { $_.SkuPartNumber -and $_.Product_Display_Name }
            }
            if ($csvData | Get-Member -Name "Service_Plan_Name" -ErrorAction SilentlyContinue) {
                $ServiceMapping = $csvData | Where-Object { $_.Service_Plan_Name -and $_.ServicePlanDisplayName }
            }
        } catch {
            Write-Host "Warning: Failed to parse mapping file, continuing without friendly names." -ForegroundColor Yellow
        }
    }

    $Global:LicensedUserCount = 0

    if ($UserNamesFile) {
        $UserNames = Import-Csv -Header "UserPrincipalName" $UserNamesFile
        foreach ($item in $UserNames) {
            $user = Get-MgUser -UserId $item.UserPrincipalName -ErrorAction SilentlyContinue
            if ($user) {
                $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
                if ($licenseDetails) {
                    Export-UserLicenseDetails -User $user -ExportCSV $ExportCSV -SkuMapping $SkuMapping -ServiceMapping $ServiceMapping
                    $Global:LicensedUserCount++
                }
            }
        }
    } else {
        Get-MgUser -All | ForEach-Object {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $_.Id -ErrorAction SilentlyContinue
            if ($licenseDetails) {
                Export-UserLicenseDetails -User $_ -ExportCSV $ExportCSV -SkuMapping $SkuMapping -ServiceMapping $ServiceMapping
                $Global:LicensedUserCount++
            }
        }
    }

    Write-Progress -Activity "Processing users" -Completed
    if (Test-Path -Path $ExportCSV) {
        Write-Host "`nDetailed report available at: $ExportCSV" -ForegroundColor Cyan
        Write-Host "$Global:LicensedUserCount users processed." -ForegroundColor Green
    } else {
        Write-Host "No data found." -ForegroundColor Yellow
    }
    Close-Connection
}

Connect-M365Services -Services "Graph" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
main

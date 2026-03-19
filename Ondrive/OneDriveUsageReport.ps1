<#
=============================================================================================
Name:           OneDrive for Business Usage Report
Version:        2.0
Website:        ClutchSolutions.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
    """_summary_
    """1. Reports OneDrive for Business storage consumption for all licensed users.
2. Includes user details: city, country, department, job title.
3. Calculates quota used (GB), storage used (GB), and percentage used.
4. Exports report results to CSV in the shared Exports folder.
5. Supports certificate-based (unattended), credential-based, and interactive authentication.
6. Automatically handles tenants with obscured report data, restoring the setting after the run.
7. Uses the shared M365AuthModule for consistent authentication across all scripts.

~~~~~~~~~~~~~~~~~~
Required Graph Permissions (Application or Delegated):
~~~~~~~~~~~~~~~~~~
  - User.Read.All
  - Reports.Read.All
  - ReportSettings.ReadWrite.All   (only required if tenant has obscured report names)

~~~~~~~~~~~
Change Log:
~~~~~~~~~~~
V1.0 (21-Feb-2024) - Original script (github.com/12Knocksinna/Office365itpros)
V2.0 (17-Mar-2026) - Refactored for Clutch Solutions: shared auth module, Exports folder,
                      Param block, division-by-zero guard, path hardening.
============================================================================================
#>
Param(
    [string]$TenantId             = "",
    [string]$ClientId             = "",
    [string]$CertificateThumbprint = ""
)

# Import shared auth helper module
Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

# Connect to Microsoft Graph
Connect-M365Services -Services Graph `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint `
    -GraphScopes @("User.Read.All", "Reports.Read.All", "ReportSettings.ReadWrite.All")

# Ensure Exports folder exists
$ExportsDir = Join-Path $PSScriptRoot ".." "Exports"
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }

$CSVOutputFile = Join-Path $ExportsDir "OneDriveUsageReport_$((Get-Date -Format 'yyyy-MMM-dd-ddd_HH-mm-ss').ToString()).csv"
$TempExportFile = Join-Path $env:TEMP "OD4B_TempExport_$([System.IO.Path]::GetRandomFileName()).csv"

# Check if the tenant has obscured real names for reports
# Requires ReportSettings.ReadWrite.All — gracefully skip if permission is absent
$ObscureFlag = $false
try {
    If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $true) {
        Write-Host "Unhiding obscured report data for the script to run..." -ForegroundColor Yellow
        Update-MgAdminReportSetting -BodyParameter @{ displayConcealedNames = $false }
        $ObscureFlag = $true
    }
} catch {
    Write-Host "Unable to check report concealment setting (ReportSettings.ReadWrite.All permission may be missing). Continuing..." -ForegroundColor Yellow
}

try {
    # Get user account information into a hash table for lookup
    Write-Host "Finding user account information..." -ForegroundColor Cyan
    [array]$Users = Get-MgUser -All `
        -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
        -ConsistencyLevel Eventual -CountVariable UserCount `
        -Sort "displayName" `
        -Property Id, displayName, userPrincipalName, city, country, department, jobTitle

    $UserHash = @{}
    foreach ($User in $Users) {
        $UserHash[$User.userPrincipalName] = $User
    }

    # Fetch OneDrive usage report via Graph API
    Write-Host "Finding OneDrive sites..." -ForegroundColor Cyan
    $Uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempExportFile
    $ProgressPreference = 'Continue'

    [array]$ODFBSites = Import-Csv $TempExportFile | Sort-Object "User display name"

    if (-not $ODFBSites) {
        Write-Host "No OneDrive sites found." -ForegroundColor Yellow
        exit 1
    }

    # Calculate total storage used
    $TotalODFBGBUsed = [Math]::Round(($ODFBSites."Storage Used (Byte)" | Measure-Object -Sum).Sum / 1GB, 2)

    # Build report
    $Report = [System.Collections.Generic.List[Object]]::new()
    foreach ($Site in $ODFBSites) {
        $UserData    = $UserHash[$Site."Owner Principal name"]
        $AllocBytes  = [double]$Site."Storage Allocated (Byte)"
        $UsedBytes   = [double]$Site."Storage Used (Byte)"
        $PercentUsed = if ($AllocBytes -gt 0) { [Math]::Round(($UsedBytes / $AllocBytes * 100), 4) } else { 0 }

        $Report.Add([PSCustomObject]@{
            Owner         = $Site."Owner display name"
            UPN           = $Site."Owner Principal name"
            City          = $UserData.city
            Country       = $UserData.country
            Department    = $UserData.department
            "Job Title"   = $UserData.jobTitle
            QuotaGB       = [Math]::Round($AllocBytes / 1GB, 2)
            UsedGB        = [Math]::Round($UsedBytes / 1GB, 4)
            PercentUsed   = $PercentUsed
        })
    }

    $Report | Export-Csv -NoTypeInformation -Path $CSVOutputFile
    $Report | Sort-Object UsedGB -Descending | Out-GridView -Title "OneDrive Usage Report"
    Write-Host ("Current OneDrive for Business storage consumption: {0} GB. Report saved to: {1}" -f $TotalODFBGBUsed, $CSVOutputFile) -ForegroundColor Green
}
finally {
    # Restore tenant report obscure data setting if it was changed
    if ($ObscureFlag) {
        Write-Host "Restoring tenant data concealment setting to True..." -ForegroundColor Yellow
        Update-MgAdminReportSetting -BodyParameter @{ displayConcealedNames = $true }
    }
    # Clean up temp file
    if (Test-Path $TempExportFile) { Remove-Item $TempExportFile -Force }
}

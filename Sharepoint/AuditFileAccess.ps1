<#
=============================================================================================
Name:           Audit File Access in SharePoint Online
Version:        2.0
Website:        ClutchSolutions.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Retrieves file access audit logs using modern authentication.
2. Supports MFA-enabled accounts.
3. Retrieves file access audit log for 180 days by default.
4. Allows generating a file access audit report for a custom period.
5. Helps find recently accessed files (e.g., files accessed in the last 30 days).
6. Identifies files accessed by external/guest users.
7. Helps monitor all files accessed by a specific user.
8. Tracks SharePoint Online file accesses only.
9. Tracks OneDrive file accesses only.
10. Finds files accessed by a specific person in a specific period.
11. Exports report results to CSV in the shared Exports folder.

~~~~~~~~~~~~~~~~~~
Required Permissions:
~~~~~~~~~~~~~~~~~~
  - Exchange Online: View-Only Audit Logs (or Audit Logs) role

~~~~~~~~~~~
Change Log:
~~~~~~~~~~~
V1.0 (Original) - Audit File Access report (source: github.com/admindroid-community)
V2.0 (18-Mar-2026) - Refactored for Clutch Solutions: shared auth module, Exports folder,
                     null-comparison fixes, style alignment, bug fixes.
============================================================================================
#>
Param (
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [int]$RecentlyAccessedFiles_In_Days,
    [Switch]$SharePointOnlineOnly,
    [Switch]$OneDriveOnly,
    [Switch]$FileAccessedByExternalUsersOnly,
    [string]$AccessedBy,
    [string]$TenantId              = "",
    [string]$ClientId              = "",
    [string]$CertificateThumbprint = "",
    [string]$UserName              = "",
    [SecureString]$Password        = $null
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

Connect-M365Services -Services ExchangeOnline `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint `
    -UserName $UserName `
    -Password $Password

$ExportsDir = Join-Path $PSScriptRoot ".." "Exports"
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }

$OutputCSV   = Join-Path $ExportsDir "Audit_File_Access_Report_$((Get-Date -Format 'yyyy-MMM-dd-ddd_HH-mm-ss').ToString()).csv"
$MaxStartDate = ((Get-Date).AddDays(-179)).Date

if ($RecentlyAccessedFiles_In_Days -ne 0) {
    $StartDate = ((Get-Date).AddDays(-$RecentlyAccessedFiles_In_Days)).Date
    $EndDate   = (Get-Date).Date
}

# Interactive prompt when no date range was supplied via parameters
if (($null -eq $StartDate) -and ($null -eq $EndDate) -and ($RecentlyAccessedFiles_In_Days -eq 0)) {
    Write-Host "`nHow far back should the audit log search go?" -ForegroundColor Cyan
    Write-Host "  Enter a number of days (1-180), or press Enter to use the full 180-day window." -ForegroundColor Yellow
    Write-Host "  Warning- The full 180-day window. will generate a large report. ( a medium sized platform will be + 150 - 250 MB )" -ForegroundColor Red
    Write-Host "  Warning- We strongly recommend that you begin with 30 days)" -ForegroundColor Red
    $DaysInput = (Read-Host "Days to review").Trim()
    if ($DaysInput -match '^\d+$') {
        $Days = [int]$DaysInput
        if ($Days -lt 1 -or $Days -gt 180) {
            Write-Host "Value must be between 1 and 180. Defaulting to 180 days." -ForegroundColor Yellow
            $Days = 180
        }
        $StartDate = ((Get-Date).AddDays(-$Days)).Date
        $EndDate   = (Get-Date).Date
    } else {
        # Empty input or non-numeric — fall through to the full 180-day default below
        Write-Host "Using full 180-day window." -ForegroundColor Yellow
    }
}

# Default to past 180 days if no date range provided
if (($null -eq $StartDate) -and ($null -eq $EndDate)) {
    $EndDate   = (Get-Date).Date
    $StartDate = $MaxStartDate
}

# Prompt for and validate start date
While ($true) {
    if ($null -eq $StartDate) {
        $StartDate = Read-Host "Enter start date for report generation (e.g. 12/15/2023)"
    }
    try {
        $Date = [DateTime]$StartDate
        if ($Date -ge $MaxStartDate) {
            break
        } else {
            Write-Host "`nAudit can be retrieved only for the past 180 days. Please select a date after $MaxStartDate" -ForegroundColor Red
            return
        }
    } catch {
        Write-Host "`nNot a valid date." -ForegroundColor Red
    }
}

# Prompt for and validate end date
While ($true) {
    if ($null -eq $EndDate) {
        $EndDate = Read-Host "Enter end date for report generation (e.g. 12/15/2023)"
    }
    try {
        $Date = [DateTime]$EndDate
        if ($EndDate -lt $StartDate) {
            Write-Host "End date should be later than start date." -ForegroundColor Red
            return
        }
        break
    } catch {
        Write-Host "`nNot a valid date." -ForegroundColor Red
    }
}

$IntervalTimeInMinutes = 1440
$CurrentStart          = $StartDate
$CurrentEnd            = $CurrentStart.AddMinutes($IntervalTimeInMinutes)

if ($CurrentEnd -gt $EndDate)  { $CurrentEnd = $EndDate }

if ($CurrentStart -eq $CurrentEnd) {
    Write-Host "Start and end time are the same. Please enter a different time range." -ForegroundColor Red
    Exit
}

$i                   = 0
$OutputEvents        = 0
$CurrentResultCount  = 0
$Operations          = "FileAccessed"

if ($FileAccessedByExternalUsersOnly.IsPresent) {
    $UserIds = "*#EXT*"
} elseif (-not [string]::IsNullOrEmpty($AccessedBy)) {
    $UserIds = $AccessedBy
} else {
    $UserIds = "*"
}

Write-Host "`nRetrieving file access audit log from $StartDate to $EndDate..." -ForegroundColor Yellow

try {
    while ($true) {
        $Results     = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd `
                           -Operations $Operations -UserIds $UserIds `
                           -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
        $ResultCount = ($Results | Measure-Object).Count

        foreach ($Result in $Results) {
            $i++
            $PrintFlag   = $true
            $MoreInfo    = $Result.auditdata
            $AuditData   = $Result.auditdata | ConvertFrom-Json
            $ActivityTime = (Get-Date($AuditData.CreationTime)).ToLocalTime()
            $UserID       = $AuditData.userId
            $AccessedFile = $AuditData.SourceFileName
            $FileExtension = $AuditData.SourceFileExtension
            $SiteURL      = $AuditData.SiteURL
            $Workload     = $AuditData.Workload

            if ($SharePointOnlineOnly.IsPresent -and ($Workload -ne "SharePoint")) { $PrintFlag = $false }
            if ($OneDriveOnly.IsPresent -and ($Workload -ne "OneDrive"))           { $PrintFlag = $false }

            if ($PrintFlag) {
                $OutputEvents++
                [PSCustomObject]@{
                    'Accessed Time' = $ActivityTime
                    'Accessed By'   = $UserID
                    'Accessed File' = $AccessedFile
                    'Site URL'      = $SiteURL
                    'File Extension' = $FileExtension
                    'Workload'      = $Workload
                    'More Info'     = $MoreInfo
                } | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
            }
        }

        Write-Progress -Activity "Retrieving file access audit data from $StartDate to $EndDate..." `
            -Status "Processed audit record count: $i"

        $CurrentResultCount = $CurrentResultCount + $ResultCount

        if ($CurrentResultCount -ge 50000) {
            Write-Host "Retrieved max records for current range. Proceeding may cause data loss." -ForegroundColor Red
            $Confirm = Read-Host "`nAre you sure you want to continue? [Y] Yes [N] No"
            if ($Confirm -match "[Y]") {
                Write-Host "Proceeding audit log collection with possible data loss." -ForegroundColor Yellow
                [DateTime]$CurrentStart    = $CurrentEnd
                [DateTime]$CurrentEnd      = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
                $CurrentResultCount        = 0
                if ($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
            } else {
                Write-Host "Please rerun the script with a reduced time interval." -ForegroundColor Red
                Exit
            }
        }

        if ($ResultCount -lt 5000) {
            if ($CurrentEnd -eq $EndDate) { break }
            $CurrentStart = $CurrentEnd
            if ($CurrentStart -gt (Get-Date)) { break }
            $CurrentEnd         = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
            $CurrentResultCount = 0
            if ($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
        }

        $ResultCount = 0
    }
    Write-Progress -Activity "Retrieving file access audit data..." -Completed

    if ($OutputEvents -eq 0) {
        Write-Host "`nNo records found." -ForegroundColor Yellow
    } else {
        Write-Host "`nThe output file contains $OutputEvents audit records." -ForegroundColor Green
        if (Test-Path $OutputCSV) {
            Write-Host "Report saved to: $OutputCSV" -ForegroundColor Green
        }
    }
} finally {
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
}

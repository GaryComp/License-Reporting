<#
=============================================================================================
Name:           Teams Activity Summary Report
Version:        2.0
Website:        ClutchSolutions.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Generates Teams activity reports for periods of 7, 30, 90, or 180 days.
2. Retrieves Teams usage data for a specific date within the last 28 days.
3. Identifies inactive Teams based on customizable inactivity thresholds.
4. Supports certificate-based (unattended) and interactive authentication.
5. The script is scheduler-friendly for automated reporting.
6. Exports results to timestamped CSV files in the shared Exports folder.

~~~~~~~~~~~~~~~~~~
Required Graph Permissions (Application or Delegated):
~~~~~~~~~~~~~~~~~~
  - Reports.Read.All

~~~~~~~~~~~
Change Log:
~~~~~~~~~~~
V1.0 (Original) - Export Teams Usage Report (source: github.com/admindroid-community)
V2.0 (18-Mar-2026) - Refactored for Clutch Solutions: shared auth module, Exports folder,
                     temp file cleanup, progress suppression, style alignment.
============================================================================================
#>
Param (
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period,
    [string]$ReportDate,
    [int]$InactiveDays,
    [string]$TenantId              = "",
    [string]$ClientId              = "",
    [string]$CertificateThumbprint = ""
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

Function Export-TeamsActivityCsv {
    param (
        [array]$CsvData,
        [string]$OutputPath
    )
    if ($CsvData.Count -eq 0) { return }

    $CsvData = $CsvData | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_.'Last Activity Date')) { $_.'Last Activity Date' = 'Never Active' }
        $_
    }

    $CsvData | Select-Object "Team Name", "Team Id", "Team Type", "Is Deleted", "Last Activity Date",
        "Active Users", @{ Name = "Active External Users"; Expression = { $_."Active External Users" } },
        @{ Name = "Active Guests"; Expression = { $_."Guests" } }, "Active Channels", "Active Shared Channels",
        "Post Messages", "Urgent Messages", "Mentions", "Channel Messages", "Reply Messages",
        "Reactions", "Meetings Organized" | Export-Csv -Path $OutputPath -NoTypeInformation
}

Connect-M365Services -Services Graph `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint `
    -GraphScopes @("Reports.Read.All")

$ExportsDir = Join-Path $PSScriptRoot ".." "Exports"
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }

$Timestamp    = Get-Date -Format 'yyyy-MMM-dd-ddd_HH-mm-ss'
$TempFilePath = Join-Path $env:TEMP "Teams_Activity_Temp_$([System.IO.Path]::GetRandomFileName()).csv"
$CsvFilePath  = Join-Path $ExportsDir "Teams_Activity_Summary_Report_$Timestamp.csv"

if (-not ($Period -or $ReportDate -or $InactiveDays)) {
    Write-Host "`nWe can perform below operations." -ForegroundColor Cyan
    Write-Host "      1. Audit Teams activity for a period of time" -ForegroundColor Yellow
    Write-Host "      2. Get Teams activity for a specific day" -ForegroundColor Yellow
    Write-Host "      3. Find Inactive Teams" -ForegroundColor Yellow
    Write-Host "      4. Exit" -ForegroundColor Yellow
    [int]$Action = Read-Host "`nPlease choose the action to continue"
} else {
    if ($Period) { $Action = 1 } elseif ($ReportDate) { $Action = 2 } elseif ($InactiveDays) { $Action = 3 }
}

try {
    Switch ($Action) {
        1 {
            $validPeriods = @("D7", "D30", "D90", "D180")
            if (-not $Period) {
                Write-Host "`nAvailable periods: $($validPeriods -join ', ')"
                $Period = (Read-Host "Enter your preferred period (e.g., D30)").Trim().ToUpper()
            }
            if ($validPeriods -contains $Period) {
                $ProgressPreference = 'SilentlyContinue'
                Get-MgReportTeamActivityDetail -Period $Period -OutFile $TempFilePath
                $ProgressPreference = 'Continue'
                $csvdata = Import-Csv -Path $TempFilePath
                Export-TeamsActivityCsv -CsvData $csvdata -OutputPath $CsvFilePath
            } else {
                Write-Host "Invalid period entered." -ForegroundColor Red
                Exit
            }
        }

        2 {
            if (-not $ReportDate) {
                $ReportDate = Read-Host "`nEnter a date starting from $((Get-Date).AddDays(-28).ToString('yyyy-MM-dd'))"
            }
            try {
                $ProgressPreference = 'SilentlyContinue'
                Get-MgReportTeamActivityDetail -Date $ReportDate -OutFile $TempFilePath -ErrorAction Stop
                $ProgressPreference = 'Continue'
                $csvdata = Import-Csv -Path $TempFilePath
                Export-TeamsActivityCsv -CsvData $csvdata -OutputPath $CsvFilePath
            } catch {
                Write-Host "Error retrieving Teams activity for '$ReportDate': $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        3 {
            if (-not $InactiveDays) {
                $InactiveDays = Read-Host "`nEnter number of inactive days"
            }
            $ProgressPreference = 'SilentlyContinue'
            Get-MgReportTeamActivityDetail -Period 'D180' -OutFile $TempFilePath
            $ProgressPreference = 'Continue'
            $csvdata     = Import-Csv -Path $TempFilePath
            $cutoffDate  = (Get-Date).AddDays(-[int]$InactiveDays)
            $inactiveData = $csvdata | Where-Object {
                if ($_.'Is Deleted' -eq $true) { return $false }
                if (-not [string]::IsNullOrWhiteSpace($_.'Last Activity Date')) {
                    try {
                        $lastActivity = Get-Date $_.'Last Activity Date' -ErrorAction Stop
                        return ($lastActivity -lt $cutoffDate)
                    } catch { return $true }
                }
                return $true
            }
            Export-TeamsActivityCsv -CsvData $inactiveData -OutputPath $CsvFilePath
        }

        4 { Exit }

        default {
            Write-Host "`nInvalid choice. Please select a valid action." -ForegroundColor Red
            Exit
        }
    }

    if ((Test-Path $CsvFilePath) -and ($null -ne (Get-Content $CsvFilePath | Where-Object { $_ -match '\S' }))) {
        Write-Host "`nReport saved to: $CsvFilePath" -ForegroundColor Green
        
    } else {
        Write-Host "`nNo records found." -ForegroundColor Yellow
    }
} finally {
    Disconnect-MgGraph | Out-Null
    if (Test-Path $TempFilePath) { Remove-Item $TempFilePath -Force }
}

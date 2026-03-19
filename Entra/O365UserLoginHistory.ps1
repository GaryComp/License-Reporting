<#
=============================================================================================
Name:           Office 365 User Login History Report
Description:    This script retrieves the login history of Office 365 users and exports the report to CSV. 
                The report includes details such as login time, user name, IP address, operation, result status, and workload.  
Script Highlights: 
~~~~~~~~~~~~~~~~~

1.The script uses modern authentication to connect to Exchange Online.
2.Allows you to filter the result based on successful and failed logon attempts. 
3.The exported report has IP addresses from where your office 365 users are login. 
4.This script can be executed with MFA enabled account. 
5.You can export the report to choose either “All Office 365 users’ login attempts” or “Specific Office user’s logon attempts”. 
6.By using advanced filtering options, you can export “Office 365 users Sign-in report” and “Suspicious login report”. 
7.Exports report result to CSV. 
8.Automatically installs the EXO V2 module (if not installed already) upon your confirmation. 
9.This script is scheduler friendly. I.e., credentials can be passed as a parameter instead of saving inside the script. 
10.Our Logon history report tracks login events in AzureActiveDirectory (UserLoggedIn, UserLoginFailed), ExchangeOnline (MailboxLogin) and MicrosoftTeams (TeamsSessionStarted). 
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Success,
    [switch]$Failed,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$AdminName,
    [SecureString]$Password
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

#Getting StartDate and EndDate for Audit log
if ((($null -eq $StartDate) -and ($null -ne $EndDate)) -or (($null -ne $StartDate) -and ($null -eq $EndDate)))
{
 Write-Host `nPlease enter both StartDate and EndDate for Audit log collection -ForegroundColor Red
 exit
}   
elseif(($null -eq $StartDate) -and ($null -eq $EndDate))
{
 $StartDate=(((Get-Date).AddDays(-90))).Date
 $EndDate=Get-Date
}
else
{
 $StartDate=[DateTime]$StartDate
 $EndDate=[DateTime]$EndDate
 if($StartDate -lt ((Get-Date).AddDays(-90)))
 { 
  Write-Host `nAudit log can be retrieved only for past 90 days. Please select a date after (Get-Date).AddDays(-90) -ForegroundColor Red
  Exit
 }
 if($EndDate -lt ($StartDate))
 {
  Write-Host `nEnd time should be later than start time -ForegroundColor Red
  Exit
 }
}

Connect-M365Services -Services "ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $AdminName -Password $Password

$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$OutputCSV = Join-Path $ExportsDir "UserLoginHistoryReport_$((Get-Date -format 'yyyy-MMM-dd-ddd hh-mm tt').ToString()).csv"
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Filter for successful login attempts
if($success.IsPresent)
{
 $Operation="UserLoggedIn,TeamsSessionStarted,MailboxLogin,SignInEvent"
}
#Filter for successful login attempts
elseif($Failed.IsPresent)
{
 $Operation="UserLoginFailed"
}
else
{
 $Operation="UserLoggedIn,UserLoginFailed,TeamsSessionStarted,MailboxLogin,SignInEvent"
}

#Check whether CurrentEnd exceeds EndDate(checks for 1st iteration)
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = 0
$CurrentResultCount=0
Write-Host `nRetrieving audit log from $StartDate to $EndDate... -ForegroundColor Yellow

while($true)
{ 
 #Write-Host Retrieving audit log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }

 #Getting audit log for specific user(s) for a given time range
 if($UserName -ne "")
 {
  $Results=Search-UnifiedAuditLog -UserIds $UserName -StartDate $CurrentStart -EndDate $CurrentEnd -operations $Operation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 -WarningAction SilentlyContinue
 }

 #Getting audit log for all users for a given time range
 else
 {
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 -WarningAction SilentlyContinue
 }
 $ResultsCount=($Results|Measure-Object).count
 $AllAuditData=@()
 $AllAudits=
 foreach($Result in $Results)
 {
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $AuditData.CreationTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()
  $AllAudits=@{'Login Time'=$AuditData.CreationTime;'User Name'=$AuditData.UserId;'IP Address'=$AuditData.ClientIP;'Operation'=$AuditData.Operation;'Result Status'=$AuditData.ResultStatus;'Workload'=$AuditData.Workload}
  $AllAuditData= New-Object PSObject -Property $AllAudits
  $AllAuditData | Sort-Object 'Login Time','User Name' | Select-Object 'Login Time','User Name','IP Address',Operation,'Result Status',Workload | Export-Csv $OutputCSV -NoTypeInformation -Append
 }
 
 #$CurrentResult += $Results
 $currentResultCount=$CurrentResultCount+$ResultsCount
 $AggregateResults +=$ResultsCount
 $TotalMinutes   = ($EndDate - $StartDate).TotalMinutes
 $ElapsedMinutes = ($CurrentStart - $StartDate).TotalMinutes
 $PercentComplete = [Math]::Min(100, [Math]::Round(($ElapsedMinutes / $TotalMinutes) * 100))
 Write-Progress -Activity "Retrieving audit log" `
     -Status "Window: $CurrentStart  →  $CurrentEnd  |  Records so far: $AggregateResults" `
     -PercentComplete $PercentComplete
 if(($CurrentResultCount -eq 50000) -or ($ResultsCount -lt 5000))
 {
  if($CurrentResultCount -eq 50000)
  {
   Write-Host Retrieved max record for the current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
   $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
   if($Confirm -notmatch "[Y]")
   {
    Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
    Exit
   }
   else
   {
    Write-Host Proceeding audit log collection with data loss
   }
  } 
  #Check for last iteration
  if(($CurrentEnd -eq $EndDate))
  {
   break
  }
  [DateTime]$CurrentStart=$CurrentEnd
  #Break loop if start date exceeds current date(There will be no data)
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }
  
  $CurrentResultCount=0
 }
}
Write-Progress -Activity "Retrieving audit log" -Completed

#Open output file after execution
If($AggregateResults -eq 0)
{
 Write-Host "No records found for the given criteria." -ForegroundColor Yellow
}
else
{
 Write-Host "`nThe output file contains $AggregateResults audit records."
 Write-Host " The output file available in: " -NoNewline -ForegroundColor Yellow
 Write-Host $OutputCSV
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue

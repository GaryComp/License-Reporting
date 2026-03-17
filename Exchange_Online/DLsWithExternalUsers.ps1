<#
=============================================================================================
Name:           Get Distribution Lists with External Users in Microsoft 365  
Version:        1.0
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Generates a list of distribution groups with external users in Microsoft 365.  
2. Excludes external mail contacts by default, with an option to include them if required. 
3. The script automatically verifies and installs the Exchange PowerShell module (if not installed already) upon your confirmation. 
4. Exports report results to CSV. 
5. The script supports Certificate-based authentication (CBA).  
6. The script is schedular friendly. �

For detailed Script execution: https://o365reports.com/2024/12/17/get-distribution-lists-with-external-users-in-microsoft-365/


============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
	[string]$UserName=$Null,
    [string]$Password=$Null,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
	[switch]$IncludeContacts
	)
Import-Module "$PSScriptRoot\M365AuthModule.psm1" -Force
Connect-M365Services -Services "ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password

$Location=Get-Location
$OutputCsv="$Location\DLs_with_ExternalUsers_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$DLWithExternalUsersCount=0
$DLCount=0
    
#Processing all DLs
Get-DistributionGroup -ResultSize Unlimited | ForEach{
 $DLCount++
 Write-Progress -Activity "Finding Distribution List with ExternalUsers:" -Status "Processed DL Count: $DLCount" -CurrentOperation "Currently Processing DL Name: $_ "
 $DLName =$_.Name
 $DLEmailAddress=$_.PrimarySMTPAddress
 If($IncludeContacts.IsPresent)
 {
  $ExternalUsers=(Get-DistributionGroupMember $_.Name  |where{$_.RecipientType -eq "MailContact" -or $_.RecipientType -eq "MailUser" })
 }
 Else
 {
  $ExternalUsers=(Get-DistributionGroupMember $_.Name  |where{ $_.RecipientType -eq "MailUser" })	
 }

 #Processing DLs with external users
 $ExternalUsersCount= ($ExternalUsers | Measure-Object).count
 If($ExternalUsersCount -ne '0')
 {
  $EmailAddress=$ExternalUsers.ExternalEmailAddress | foreach {
    $_.split(":")[1]}
  $EmailAddress= $EmailAddress -join (",")
  $ExternalUserName=$ExternalUsers.Name -join ','
  $ExternalUserDisplayName=$ExternalUsers.DisplayName -join ','
  $Result=New-Object PsObject -Property @{'DL Name'=$DLName;'DL Email Address'=$DLEmailAddress;'No of External Users in DL'=$ExternalUsersCount;'External Users Name'=$ExternalUserName;'External Users Display Name'=$ExternalUserDisplayName;'Exterenal Users Email Address'=$EmailAddress}
  $Result|Select-Object 'DL Name','DL Email Address','No of External Users in DL','External Users Name','External Users Display Name','Exterenal Users Email Address'|Export-CSV -Path $OutputCsv -NoTypeInformation -Append
  $DLWithExternalUsersCount++
 }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false


Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n
 

if((Test-Path -Path $OutputCsv) -eq "True") 
{
 Write-Host $DLWithExternalUsersCount DLs contain external user as members.
 Write-Host `nDetailed report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $OutputCSV
 $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
 If ($UserInput -eq 6)   
 {   
   Invoke-Item "$OutputCSV"   
 } 
}
else
{
 Write-Host No DLs found with the specific criteria.
}
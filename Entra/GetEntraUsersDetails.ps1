<#
=============================================================================================
-----------------
Script Highlights
-----------------
1. The script automatically verifies and installs the Microsoft Graph PowerShell SDK module (if not installed already) upon your confirmation.
2. Exports all users from Microsoft Entra.
3. Allows filtering and exporting users that match the selected filters.
    -> Guest users
    -> Sign-in enabled users
    -> Sign-in blocked users
    -> License assigned users
    -> Users without any license
    -> Users without a manager
4. Identifies recently created users in Microsoft Entra (e.g., within the last n days).
5. Exports the report result to CSV.
6. This script can be scheduled to run automatically.
============================================================================================
\#>
Param
(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [int]$RecentlyCreatedUsers,
    [Switch]$GuestUsersOnly,
    [Switch]$EnabledUsersOnly,
    [Switch]$DisabledUsersOnly,
    [Switch]$LicensedUsersOnly,
    [Switch]$UnlicensedUsersOnly,
    [Switch]$UnmanagedUsers

)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force
Connect-M365Services -Services "Graph" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint


$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$ExportCSV = Join-Path $ExportsDir "EntraUsers_Report_$((Get-Date -format 'yyyy-MMM-dd-ddd hh-mm-ss tt').ToString()).csv"
$Count=0
$PrintedUsers=0
$RequiredProperties=@('UserPrincipalName','LastPasswordChangeDateTime','AccountEnabled','Country','Department','Jobtitle','SigninActivity','DisplayName','UserType','CreatedDateTime')

Write-Host Generating Entra users report...
Get-MgUser -All -Property $RequiredProperties  | ForEach-Object {
 $Print=1
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.DisplayName
 $Count++
 Write-Progress -Activity "`n     Processed users: $Count - $UPN "
 $LastPwdSet=$_.LastPasswordChangeDateTime
 $AccountEnabled=$_.AccountEnabled
 if($AccountEnabled -eq $true)
 {
  $SigninStatus="Allowed"
 }
 else
 {
  $SigninStatus="Denied"
 }

 $SKUs = (Get-MgUserLicenseDetail -UserId $UPN).SkuPartNumber
 $Sku= $SKUs -join ","
 $Department=$_.Department
 $JobTitle=$_.JobTitle
 $LastSigninTime=($_.SignInActivity).LastSignInDateTime
 $LastNonInteractiveSignIn=($_.SignInActivity).LastNonInteractiveSignInDateTime
 $Manager=(Get-MgUserManager -UserId $UPN -ErrorAction SilentlyContinue)
 $ManagerDetails=$Manager.AdditionalProperties
 $ManagerName=$ManagerDetails.userPrincipalName
 $Country= $_.Country
 $CreationTime=$_.CreatedDateTime
 $CreatedSince=(New-TimeSpan -Start $CreationTime).Days
 $UserType=$_.UserType

 #Filter for guest users
 if($GuestUsersOnly.IsPresent -and ($UserType -ne "Guest"))
 { 
  $Print=0
 }
 #Filter for recently created users
 if(($RecentlyCreatedUsers -ne "") -and ($CreatedSince -gt $RecentlyCreatedUsers))
 { 
  $Print=0
 }
 #Filter for sign-in allowed users
 if($EnabledUsersOnly.IsPresent -and ($AccountEnabled -eq $false))
 {
  $Print=0
 }
 #Filter for sign-in disabled users
 if($DisabledUsersOnly.IsPresent -and ($AccountEnabled -eq $true))
 {
  $Print=0
 }
 #Filter for licensed users
 if(($LicensedUsersOnly.IsPresent) -and ($Sku.Length -eq 0))
 {
  $Print=0
 }
 #Filter for unlicensed users
 if(($UnlicensedUsersOnly.IsPresent) -and ($Sku.Length -ne 0))
 {
  $Print=0
 }
 #Filter for users withour manager
 if(($UnmanagedUsers.IsPresent) -and ($Manager -ne $null))
 {
  $Print=0
 }
 
 #Export users based on the given criteria
 if($Print -eq 1)
 {
  $PrintedUsers++
  $Result=[PSCustomObject]@{'Name'=$UPN;'Display Name'=$DisplayName;'User Type'=$UserType;'Sign-in Status'=$SigninStatus;'License'=$SKU;'Department'=$Department;'Job Title'=$JobTitle;'Country'=$Country;'Manager'=$ManagerName;'Pwd Last Change Date'=$LastPwdSet;'Last Signin Date'=$LastSigninTime;'Last Non-interactive Signin Date'=$LastNonInteractiveSignIn;'Creation Time'=$CreationTime}
  $Result | Export-Csv -Path $ExportCSV -Notype -Append
 }
}
Write-Progress -Activity "Processing users" -Completed

Disconnect-MgGraph | Out-Null
  
#Open Output file after execution
 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `The exported report contains $PrintedUsers users.
  Write-Host `nEntra users report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
 }
 else
 {
  Write-Host No users found for the given criteria.
 }
<#
=============================================================================================
Name:           Get all enterprise apps and their owners 
Script Highlights:  
~~~~~~~~~~~~~~~~~
1. The script exports all enterprise apps along with its owners in Microsoft Entra.  
2. Generates report for sign-in enabled applications alone. 
3. Exports report for sign-in disabled applications only. 
4. Filters applications that are hidden from all users except assigned users. 
5. Provides the list of applications that are visible to all users in the organization. 
6. Lists applications that are accessible to all users in the organization.  
7. Identifies applications that can be accessed only by assigned users. 
8. Fetches the list of ownerless applications in Microsoft Entra. 
9. Assists in filtering home tenant applications only. 
10. Exports applications from external tenants only. 
11. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation.  
12. Exports the report result to CSV. 
13. The script can be sheduled to run automatically.  
============================================================================================
#>
Param
(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$SigninEnabledAppsOnly,
    [Switch]$SigninDisabledAppsOnly,
    [Switch]$HiddenApps,
    [Switch]$VisibleToAllUsers,
    [Switch]$AccessScopeToAllUsers,
    [Switch]$RoleAssignmentRequiredApps,
    [Switch]$OwnerlessApps,
    [Switch]$HomeTenantAppsOnly,
    [Switch]$ExternalTenantAppsOnly
)

# Import shared auth helper module
Import-Module "${PSScriptRoot}\..\M365AuthModule.psm1" -Force

# Connect to Graph (will use config file / params for cert auth or interactive)
Connect-M365Services -Services Graph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -GraphScopes @('Application.Read.All')

# Ensure output folder exists
$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }

$ExportCSV = Join-Path $ExportsDir "EnterpriseApps_and_their_Owners_Report_$((Get-Date -format 'yyyy-MMM-dd-ddd hh-mm-ss tt').ToString()).csv"
$PrintedCount=0
$Count=0
$TenantGUID= (Get-MgOrganization).Id


$RequiredProperties=@('DisplayName','AccountEnabled','Id','Tags','AppRoleAssignmentRequired','ServicePrincipalType','AdditionalProperties','AppDisplayName','AppOwnerOrganizationId','createdDateTime')
Get-MgServicePrincipal -All -Property $RequiredProperties | ForEach-Object {
 $Print=1
 $Count++
 $EnterpriseAppName=$_.DisplayName
 Write-Progress -Activity "`n     Processed enterprise apps: $Count - $EnterpriseAppName "
 $UserSigninStatus=$_.AccountEnabled
 $Id=$_.Id
 $Tags=$_.Tags
 if($Tags -contains "HideApp")
 {
  $UserVisibility="Hidden"
 }
 else
 {
  $UserVisibility="Visible"
 }
 $IsRoleAssignmentRequired=$_.AppRoleAssignmentRequired
 if($IsRoleAssignmentRequired -eq $true)
 {
  $AccessScope="Only assigned users can access"
 }
 else
 {
  $AccessScope="All users can access"
 }
 $parsedDate = $_.AdditionalProperties.createdDateTime -as [DateTime]
 $CreationTime = if ($parsedDate) { $parsedDate.ToLocalTime() } else { '-' }
 $ServicePrincipalType=$_.ServicePrincipalType
 $AppRegistrationName=$_.AppDisplayName
 $AppOwnerOrgId=$_.AppOwnerOrganizationId
 if($AppOwnerOrgId -eq $TenantGUID)
 {
  $AppOrigin="Home tenant"
 }
 else
 {
  $AppOrigin="External tenant"
 }
 $Owners=(Get-MgServicePrincipalOwner -ServicePrincipalId $Id).AdditionalProperties.userPrincipalName
 $Owners=$Owners -join ","
 if($owners -eq "")
 {
  $Owners="-"
 }

 #Filtering the result
 if(($SigninEnabledAppsOnly.IsPresent) -and ($UserSigninStatus -eq $false))
 {
  $Print=0
 }
 elseif(($SigninDisabledAppsOnly.IsPresent) -and ($UserSigninStatus -eq $true))
 {
  $Print=0
 }
 if(($HiddenApps.IsPresent) -and ($UserVisibility -eq "Visible"))
 {
  $Print=0
 }
 elseif(($VisibleToAllUsers.IsPresent) -and ($UserVisibility -eq "Hidden"))
 {
  $Print=0
 }
 if(($AccessScopeToAllUsers.IsPresent) -and ($AccessScope -eq "Only assigned users can access"))
 {
  $Print=0
 }
 elseif(($RoleAssignmentRequiredApps.IsPresent) -and ($AccessScope -eq "All users can access"))
 {
  $Print=0
 }
 if(($OwnerlessApps.IsPresent) -and ($Owners -ne "-"))
 {
  $Print=0
 }
 if(($HomeTenantAppsOnly.IsPresent) -and ($AppOrigin -eq "External tenant"))
 {
  $Print=0
 }
 elseif(($ExternalTenantAppsOnly.IsPresent) -and ($AppOrigin -eq "Home tenant"))
 {
  $Print=0
 }

 if($Print -eq 1)
   {
   $PrintedCount++
   $ExportResult=[PSCustomObject]@{'Enterprise App Name'=$EnterpriseAppName;'App Id'=$Id;'App Owners'=$Owners;'App Creation Time'=$CreationTime;'User Signin Allowed'=$UserSigninStatus;'User Visibility'=$UserVisibility;'Role Assignment Required'=$AccessScope;'Service Principal Type'=$ServicePrincipalType;'App Registration Name'=$AppRegistrationName;'App Origin'=$AppOrigin;'App Org Id'=$AppOwnerOrgId}
   $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
  }
}
Write-Progress -Activity "Processed enterprise apps" -Completed
Disconnect-MgGraph | Out-Null

#Open output file after execution
 If($PrintedCount -eq 0)
 {
  Write-Host No data found for the given criteria
 
 }
 else
 {
  Write-Host `nThe script processed $Count enterprise apps and the output file contains $PrintedCount records.
  if((Test-Path -Path $ExportCSV) -eq "True") 
  {

   Write-Host `n The Output file available in: -NoNewline -ForegroundColor Yellow
   Write-Host $ExportCSV 
  
 }
}

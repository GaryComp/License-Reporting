<#
=============================================================================================
Name:           Find Unused Licenses in Microsoft 365 Using PowerShell
Description:    This script exports a report of unused Microsoft 365 licenses by identifying inactive users through their last successful sign-in activity.
Version:        1.0
website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~

1. Retrieves unused licenses based on users' last successful sign-in time.
2. Lists licenses assigned to sign-in disabled users.
3. Identifies licenses assigned to never logged-in user accounts.
4. Filters unused licenses by type, such as paid, free, or trial.
5. Fetches inactive licenses assigned to external accounts.
6. Identifies unused specific licenses, such as Power BI Pro.
7. Automatically verifies and installs the Microsoft Graph PowerShell Module (if not already installed) upon your confirmation.
8. Supports Certificate-based Authentication (CBA) too.
9. The script is scheduler-friendly.


For detailed script execution: https://o365reports.com/2025/09/02/find-unused-licenses-in-microsoft-365-using-powershell/ 

============================================================================================
#>
Param(
    [int]$InactiveDays,
    [Nullable[int]]$LicenseCount = $null,
    [string]$ImportCSVPath,
    [switch]$ReturnNeverLoggedInUser,
    [ValidateSet("InternalUser", "ExternalUser")]
    [string]$UserType,
    [ValidateSet("EnabledUser", "DisabledUser")]
    [string]$UserState,
    [ValidateSet( "Paid", "Trial", "Free")]
    [string]$LicenseType,
    [string[]]$LicensePlanList,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

if (-not $InactiveDays -and -not $ReturnNeverLoggedInUser) {
    do {
        $InactiveDays = Read-Host "`nEnter the number of inactive days"
        if ($InactiveDays -notmatch '^\d+$') {
            Write-Host "Please enter a valid number." -ForegroundColor Red
        }
    } while ($InactiveDays -notmatch '^\d+$')
    $InactiveDays = [int]$InactiveDays
}

Connect-M365Services -Services "Graph" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -GraphScopes "User.Read.All","Group.Read.All","Organization.Read.All","AuditLog.Read.All"

$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$ExportCSV = Join-Path $ExportsDir "UnusedM365LicensesByLastSignIn_$((Get-Date -format 'yyyy-MMM-dd-ddd hh-mm-ss tt').ToString()).csv"
$ExportResult = $null
$Count = 0
$PrintedUsers = 0

if( $ImportCSVPath -ne "" -and !(Test-Path $ImportCSVPath)){
    Write-Host "File not found: $($ImportCSVPath)" -ForegroundColor Red
    Exit
}

#Get Licenses
$FriendlyNameHash = @{}
$LicenseFriendlyNamePath = Join-Path $PSScriptRoot '..' 'Supporting_Files' 'LicenseFriendlyName.txt'
if (Test-Path $LicenseFriendlyNamePath) {
    $loaded = Get-Content -Raw -Path $LicenseFriendlyNamePath -ErrorAction SilentlyContinue | ConvertFrom-StringData
    if ($loaded) { $FriendlyNameHash = $loaded }
} else {
    Write-Host "Warning: LicenseFriendlyName.txt not found. License friendly names will not be resolved." -ForegroundColor Yellow
}
$LicenseMap = @{}
if($ImportCSVPath -ne ""){
    $LicenseNames = Import-Csv -Header "SkuPartNumber" -path $ImportCSVPath | ForEach-Object { $_ }
    foreach ($License in $LicenseNames) {
        $SkuPartNumber = $License.SkuPartNumber
        if ($FriendlyNameHash.ContainsKey($SkuPartNumber)) {
            $LicenseMap[$SkuPartNumber] = $FriendlyNameHash[$SkuPartNumber]
        }
    }
}
else{
    $LicenseMap = $FriendlyNameHash
} 

#Get License type details
$LicenseSkuIdMap = @{}
$LifeCycleDateInfo = Get-MgDirectorySubscription -All 
$LifeCycleDateInfo | ForEach-Object{
    $LicenseSkuIdMap[$_.SkuId] = $_.SkuPartNumber
}

#Retrieve Users
Write-Host "`nRetrieving inactive users with assigned licenses..."
$RequiredProperties = @('UserPrincipalName','DisplayName','SignInActivity','UserType','CreatedDateTime','AccountEnabled', 'LicenseAssignmentStates', 'Department','JobTitle')  
Get-MgUser -All -Property $RequiredProperties | Select-Object $RequiredProperties | ForEach-Object{
    $Count++
    $UPN = $_.UserPrincipalName
    Write-Progress -Activity "        Processing  user: $($count) $($UPN)" 
    $DisplayName = $_.DisplayName
    $UserCategory = $_.UserType
    $LastSuccessfulSigninDate = $_.SignInActivity.LastSuccessfulSignInDateTime
    $LastInteractiveSignIn = $_.SignInActivity.LastSignInDateTime
    $LastNon_InterativeSignIn = $_.SignInActivity.LastNonInteractiveSignInDateTime
    $CreatedDate = $_.CreatedDateTime
    $AccountEnabled = $_.AccountEnabled
    $Department = if($null -eq $_.Department) {" -"} else{$_.Department}
    $JobTitle = if($null -eq $_.JobTitle) {" -"} else{$_.JobTitle}
    $TotalLicenses = 0
    $LicenseStates = $_.LicenseAssignmentStates 
    $Print = 1
    
    #Calculate Inactive users days
    if($null -eq $LastSuccessfulSigninDate){
        $LastSuccessfulSigninDate = "Never Logged In"
        $InactiveUserDays = "-"
    }else{
        $InactiveUserDays = (New-TimeSpan -Start $LastSuccessfulSigninDate).Days
    }

    if($null -eq $LastInteractiveSignIn){
        $LastInteractiveSignIn = "Never Logged In"
    }

    if($null -eq $LastNon_InterativeSignIn){
        $LastNon_InterativeSignIn = "Never Logged In"
    }

    #Get account status
    if($AccountEnabled -eq $true){
        $AccountStats = "Enabled"
    }
    else{
        $AccountStats = "Disabled"
    }
    
    #Inactive days based on last successful signins filter
    if ($ReturnNeverLoggedInUser.IsPresent -and ($LastInteractiveSignIn -ne "Never Logged In" -or $LastNon_InterativeSignIn -ne "Never Logged In")) {
        $Print = 0
    }
    elseif (-not $ReturnNeverLoggedInUser.IsPresent) {
        if ($LastSuccessfulSigninDate -eq "Never Logged In") {
            $Print = 0
        }
        # Filter by inactive days
        elseif (($InactiveDays -ne 0) -and ($InactiveDays -ge $InactiveUserDays)) {
            $Print = 0
        }
    }
    
    #Filter for internal users only
    if(($UserType -eq "InternalUser") -and ($UserCategory -eq "Guest")){
        $Print = 0
    }

    #Filter for external users only
    if(($UserType -eq "ExternalUser") -and ($UserCategory -ne "Guest")){
        $Print = 0
    }

    #Signin allowed Users
    if(($UserState -eq "EnabledUser") -and ($AccountStats -eq 'Disabled')){
        $Print = 0
    }

    #Signin disabled Users
    if(($UserState -eq "DisabledUser") -and ($AccountStats -eq 'Enabled')){
        $Print = 0
    }
    
    #Licensed users only
    $LicensePartNumbers = @()
    $Groups = @()
    $GroupLicense = @()
    $DirectLicense = @()
    if($LicenseStates.Count -ne 0){
        foreach($State in $LicenseStates){
            if($State){
                $Flag = 1
                $LicensePartNumber = ""
                $LicenseName = ""
                if($LicenseSkuIdMap.ContainsKey($State.SkuId)){
                    $LicensePartNumber = $LicenseSkuIdMap[$State.SkuId]
                    $LicenseName = $LicenseMap[$LicensePartNumber]
                    $MoreSkuDetails = $LifeCycleDateInfo | Where-Object {$_.skuId -eq $State.SkuId}
                    $ExpiryDate = $MoreSkuDetails.nextLifeCycleDateTime
                    #Filter SkuPartNumber
                    if($LicensePlanList){
                        if($LicensePlanList -notcontains $LicensePartNumber){
                            $Flag = 0
                        }
                    }
                    
                    #Filter Free Licensed User
                    if($LicenseType -eq "Free"){
                        if($null -ne $ExpiryDate){
                            $Flag = 0
                        }
                    }

                    #Filter Trial Licensed User
                    if($LicenseType -eq "Trial"){
                        if(-not $MoreSkuDetails.isTrial){
                            $Flag = 0
                        }
                    }

                    #Filter Paid Licensed User
                    if($LicenseType -eq "Paid"){
                        if(($null -eq $ExpiryDate) -or ($MoreSkuDetails.isTrial)){
                            $Flag = 0
                        }
                    }

                    if($Flag -eq 1){
                        if($LicenseName){
                            $LicensePartNumbers += $LicensePartNumber
                            if($null -ne $State.AssignedByGroup){
                                $Groups += (Get-MgGroup -GroupId $State.AssignedByGroup -ErrorAction SilentlyContinue).DisplayName
                                $GroupLicense += $LicenseName
                            }
                            else{
                                $DirectLicense += $LicenseName
                            }
                        }
                    }
                }
            }
        }

        if(($DirectLicense.Count -ne 0) -or ($GroupLicense.Count -ne 0)) {
            $LicensePlans = $LicensePartNumbers -join ", "
            $TotalLicenses = $DirectLicense.Count + $GroupLicense.Count
            $GroupNames = if($Groups.Count -ne 0) {$Groups -join ","} else {"- "}
            $GroupLicenseNames = if($GroupLicense.Count -ne 0) { $GroupLicense -join ","} else {"- "}
            $DirectLicenseNames = if($DirectLicense.Count -ne 0) { $DirectLicense -join ","} else{"- "}
        }
        else{
            $Print = 0
        }
    }
    else{
        $Print = 0;
    }
    
    #LicenseCount above users only
    if($null -ne $LicenseCount){
        if($LicenseCount -gt $TotalLicenses){
            $Print = 0
        }
    }

    #Export users to output file
    if($Print -eq 1 ){
        $PrintedUsers++
        $ExportResult = [PSCustomObject]@{ 'Display Name' = $DisplayName; 'UPN' = $UPN; 'User Type' = $UserCategory; 'Account Status' = $AccountStats; 'License Plans' = $LicensePlans; 'Directly Assigned Licenses' = $DirectLicenseNames; 'Licenses Assigned via Groups' =$GroupLicenseNames; 'Assigned via (Group Names)' = $GroupNames; 'License Count' = $TotalLicenses;'Last Successful SignIn Date '= $LastSuccessfulSigninDate; 'Inactive Days' = $InactiveUserDays; 'Last Interactive SignIn Date' = $LastInteractiveSignIn; 'Last Non-Interactive SignIn Date' = $LastNon_InterativeSignIn;'Creation Date' = $CreatedDate; 'Department' = $Department;'Job Title' = $JobTitle;}
        $ExportResult | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
    }
}

Write-Progress -Activity "Processing users" -Completed
Disconnect-MgGraph | Out-Null

Write-Host "`nScript completed. Processed $Count user(s) — $PrintedUsers matched the criteria."

if($PrintedUsers -eq 0)
{
    Write-Host "No users found matching the given criteria." -ForegroundColor Yellow
}
else
{
    Write-Host "The generated report is available in: " -NoNewline -ForegroundColor Yellow
    Write-Host $ExportCSV
}
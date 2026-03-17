<#
=============================================================================================
Name:           Export Office 365 SPO External Users Report
Description:    This script exports Office 365 SPO external users to CSV
Version:        1.0
Website:        o365reports.com

Script Highlights:   
~~~~~~~~~~~~~~~~~
1. Generates 3 different SharePoint Online external user reports. 
2. Automatically installs the SharePoint Management Shell module upon your confirmation when it is not available in your system.  
3. Shows list of all external users in SharePoint Online in the tenant.   
4. You can get SharePoint sites’ external users separately. 
5. Allows retrieving external user accounts added recently. 
6. Supports both MFA and Non-MFA accounts.     
7. Exports the report in CSV format.   
8. Scheduler-friendly. You can automate the report generation upon passing credentials as parameters. 

For detailed Script execution: http://o365reports.com/2021/08/03/get-all-external-users-in-sharepoint-online-powershell
============================================================================================
#>

param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [Int] $GuestsCreatedWithin_Days,
    [Switch] $SiteWiseGuest,
    [Parameter(Mandatory = $True)]
    [string] $HostName = $null

)

#StartFromLine: 153

Import-Module "$PSScriptRoot\M365AuthModule.psm1" -Force
$AdminUrl = "https://$HostName-admin.sharepoint.com/"
Connect-M365Services -Services "SharePoint" -AdminUrl $AdminUrl -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password

#This function checks the user choice and get the guest user data
Function FindGuestUsers {
    $AllGuestUserData = @()
    $global:ExportCSVFileName = "SPOExternalUsersReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 

    #Checks the SPO Sites and lists the guest users
    if ($SiteWiseGuest.IsPresent) {
        $GuestDataAvailable = $false
        Get-SPOSite | foreach-object {
            $CurrSite = $_
            Write-Progress "Finding the guest users in the site: $($CurrSite.Url)" "Processing the sites with guest users..."
            #Fiters the sites with guest users
            Get-SPOUser -Site $CurrSite.Url | where-object { $_.LoginName -like "*#ext#*" -or $_.LoginName -like "urn:spo:guest#*"} | foreach-object {
                    $global:ExportedGuestUser = $global:ExportedGuestUser + 1    
                    $CurrGuestData = $_
                    ExportGuestsAndSitesData
                }
        }
        if ($global:ExportedGuestUser -eq 0) {
            Write-Host `n"No SharePoint Online guests in any SharePoint Online sites in your tenant" -ForegroundColor Magenta
        }
    }
   
    #Checks the guest user acount creation within the mentioned days and retrieves it
    elseif ($GuestsCreatedWithin_Days -gt 0) {
        $AccountCreationDate = (Get-date).AddDays(-$GuestsCreatedWithin_Days).Date
        for (($i = 0), ($errVar = @()); (($errVar.Count) -eq 0); $i += 50) {
        Get-SPOExternalUser -Position $i -PageSize 50 -ErrorAction SilentlyContinue -ErrorVariable errVar | where-object { $_.WhenCreated -ge $AccountCreationDate } | foreach-object {
                $global:ExportedGuestUser = $global:ExportedGuestUser + 1   
                $CurrGuestData = $_
                ExportGuestUserDetails
            }
        }
        if ($global:ExportedGuestUser -eq 0) {
            Write-host "No SharePoint Online guests created in last $GuestsCreatedWithin_Days days" -ForegroundColor Magenta
        }
    }

    #Returns all SPO guest users in your tenant
    else {
        for (($i = 0), ($errVar = @()); (($errVar.Count) -eq 0); $i += 50) {
            Get-SPOExternalUser -Position $i -PageSize 50 -ErrorAction SilentlyContinue -ErrorVariable errVar | ForEach-Object {
                $global:ExportedGuestUser = $global:ExportedGuestUser + 1   
                $CurrGuestData = $_
                ExportGuestUserDetails
            }
        }
        if ($global:ExportedGuestUser -eq 0) {
            Write-Host "No SharePoint Online guest users found in your tenant." -ForegroundColor Magenta
        }
    }
}

#Saves site-wise guest user data
Function ExportGuestsAndSitesData {
    $SiteName = $CurrSite.Title
    $SiteUrl = $CurrSite.Url
    $GuestDisplayName = $CurrGuestData.DisplayName
    $GuestEmailAddress = $CurrGuestData.LoginName
    if($GuestEmailAddress -like "*ext*"){
    $GuestDomain = ($CurrGuestData.LoginName).split("_#") | Select-Object -Index 1
    }
    else{
    $GuestDomain = ($CurrGuestData.LoginName).split("@") | Select-Object -Index 1
    }
    
    $ExportResult = @{'Guest User' = $GuestDisplayName; 'Email Address' = $GuestEmailAddress; 'Site Name' = $SiteName; 'Site Url' = $SiteUrl; 'Guest Domain' = $GuestDomain }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Guest User', 'Email Address', 'Site Name', 'Site Url', 'Guest Domain' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
      
}

#Saves guest user data
Function ExportGuestUserDetails {
    $GuestDisplayName = $CurrGuestData.DisplayName
    $GuestEmailAddress = $CurrGuestData.Email
    $GuestInviteAcceptedAs = $CurrGuestData.AcceptedAs
    $CreationDate = ($CurrGuestData.WhenCreated).ToString().split(" ") | Select-Object -Index 0
    $GuestDomain = ($CurrGuestData.Email).split("@") | Select-Object -Index 1
    
    Write-Progress "Retrieving the Guest User: $GuestDisplayName" "Processed Guest Users Count: $global:ExportedGuestUser"
   
    #Exports the guest user data to the csv file format

    $ExportResult = @{'Guest User' = $GuestDisplayName; 'Email Address' = $GuestEmailAddress; 'Invitation Accepted via' = $GuestInviteAcceptedAs; 'Created On' = $CreationDate; 'Guest Domain' = $GuestDomain }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Guest User', 'Email Address', 'Created On', 'Invitation Accepted via', 'Guest Domain' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
    
}

#Execution starts here.
$global:ExportedGuestUser = 0
FindGuestUsers

if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    #Open file after code execution finishes  
    Write-Host " The output file available in:" -NoNewline -ForegroundColor Yellow; Write-Host .\$global:ExportCSVFileName `n
    write-host "Exported $global:ExportedGuestUser records to CSV." `n 
    Write-host "Disconnected SPOService Session Successfully" `n 
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green  
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
    Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    If ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
} 
Disconnect-SPOService

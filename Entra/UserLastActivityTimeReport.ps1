<#
=============================================================================================
Name:           Export Office 365 users real last activity time report
Version:        3.0
Website:        o365reports.com
Script by:      O365Reports Team

Script Highlights :
~~~~~~~~~~~~~~~~~

1.	Reports the user’s activity time based on the user’s last action time(LastUserActionTime). 
2.	Exports result to CSV file. 
3.	Result can be filtered based on inactive days. 
4.	You can filter the result based on user/mailbox type. 
5.	Result can be filtered to list never logged in mailboxes alone. 
6.	You can filter the result based on licensed user.
7.	Shows result with the user’s administrative roles in the Office 365. 
8.	The assigned licenses column will show you the user-friendly-name like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’. 
9.	The script can be executed with MFA enabled account. 
10.	The script is scheduler friendly. i.e., credentials can be passed as a parameter instead of saving inside the script. 


For detailed script execution:  https://o365reports.com/2019/06/18/export-office-365-users-real-last-logon-time-report-csv/#
============================================================================================
#>
#If you connect via Certificate based authentication, then your application required "Directory.Read.All" application permission, assign exchange administrator role and  Exchange.ManageAsApp permission to your application.
#Accept input parameter
Param
(
    [string]$MBNamesFile,
    [int]$InactiveDays,
    [switch]$UserMailboxOnly,
    [switch]$LicensedUserOnly,
    [switch]$ReturnNeverLoggedInMBOnly,
    [switch]$FriendlyTime,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)
Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force
Function Get_LastLogonTime
{
    $MailboxStatistics = Get-MailboxStatistics -Identity $UPN
    $LastActionTime = $MailboxStatistics.LastUserActionTime
    $PercentComplete=($MBUserCount/($Mailboxes.Count))*100
    Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount out of $($Mailboxes.Count)"`n"  Currently Processing: $DisplayName"  -PercentComplete $PercentComplete
    $Script:MBUserCount++ 
 
    #Retrieve lastlogon time and then calculate Inactive days 
    if($null -eq $LastActionTime)
    { 
        $LastActionTime = "Never Logged In" 
        $InactiveDaysOfUser = "-" 
    } 
    else
    { 
        $InactiveDaysOfUser = (New-TimeSpan -Start $LastActionTime).Days
        #Convert Last Action Time to Friendly Time
        if($friendlyTime.IsPresent) 
        {
            $FriendlyLastActionTime = ConvertTo-HumanDate ($LastActionTime)
            $friendlyLastActionTime = "("+$FriendlyLastActionTime+")"
            $LastActionTime = "$LastActionTime $FriendlyLastActionTime" 
        }
    }
    #Get licenses assigned to mailboxes 
    $Licenses = (Get-MgBetaUserLicenseDetail -UserId $UPN -ErrorAction SilentlyContinue).SkuPartNumber 
    $AssignedLicense = @()
    if($Licenses.Count -eq 0) 
    { 
        $AssignedLicense = "No License Assigned" 
    }  
    #Convert license plan to friendly name 
    else
    {
        foreach($License in $Licenses) 
        {
            $EasyName = $FriendlyNameHash[$License]  
            if(!($EasyName))  
            {
                $NamePrint = $License
            }  
            else  
            {
                $NamePrint = $EasyName
            } 
            $AssignedLicense += $NamePrint
        }
        $AssignedLicense = @($AssignedLicense) -join ','
    }
    #Inactive days based filter 
    if($InactiveDaysOfUser -ne "-")
    { 
        if(($InactiveDays -ne "") -and ([int]$InactiveDays -gt $InactiveDaysOfUser)) 
        { 
            return
        }
    } 

    #Filter result based on user mailbox 
    if(($UserMailboxOnly.IsPresent) -and ($MBType -ne "UserMailbox"))
    { 
        return
    } 

    #Never Logged In user
    if(($ReturnNeverLoggedInMBOnly.IsPresent) -and ($LastActionTime -ne "Never Logged In"))
    {
        return
    }

    #Filter result based on license status
    if(($LicensedUserOnly.IsPresent) -and ($AssignedLicense -eq "No License Assigned"))
    {
        return
    }
    #Get admin roles assigned to user 
    $RoleList=Get-MgBetaUserTransitiveMemberOf -UserId $UPN|Select-Object -ExpandProperty AdditionalProperties
    $RoleList = $RoleList | Where-Object {$_.'@odata.type' -eq '#microsoft.graph.directoryRole'}
    $Roles = @($RoleList.displayName) -join ','
    if($RoleList.count -eq 0)
    {
        $Roles = "No roles"
    }

    #Export result to CSV file 
    $Result = [PSCustomObject] @{'UserPrincipalName'=$UPN;'DisplayName'=$DisplayName;'LastUserActionTime'=$LastActionTime;'CreationTime'=$CreationTime;'InactiveDays'=$InactiveDaysOfUser;'MailboxType'=$MBType; 'AssignedLicenses'=$AssignedLicense;'Roles'=$Roles} 
    $Result | Export-Csv -Path $ExportCSV -Notype -Append
}
Function CloseConnection
{
    Disconnect-MgGraph | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

Connect-M365Services -Services "Graph","ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
Write-Host "`nNote: If you encounter module related conflicts, run the script in a fresh PowerShell window.`n" -ForegroundColor Yellow

#Friendly DateTime conversion
if($FriendlyTime.IsPresent)
{
    if(((Get-Module -Name PowerShellHumanizer -ListAvailable).Count) -eq 0)
    {
        Write-Host Installing PowerShellHumanizer for Friendly DateTime conversion 
        Install-Module -Name PowerShellHumanizer
    }
}
$Result = ""  
$MBUserCount = 1 

#Get friendly name of license plan from external file 
$FriendlyNameHash = Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction SilentlyContinue -ErrorVariable LicenseFileError | ConvertFrom-StringData
if($null -ne $LicenseFileError)
{
    Write-Host $LicenseFileError -ForegroundColor Red
    CloseConnection
}

#Set output file 
$Path = (Get-Location).Path
$ExportCSV = "$Path\LastAccessTimeReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"

#Check for input file
if([string]$MBNamesFile -ne "") 
{ 
    #We have an input file, read it into memory 
    $Mailboxes = @()
    try{
        $Mailboxes = Import-Csv -Header "MBIdentity" $MBNamesFile
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        CloseConnection
    }
    Foreach($item in $Mailboxes)
    {
        $MBDetails = Get-Mailbox -Identity $item.MBIdentity
        $DisplayName = $MBDetails.DisplayName 
        $UPN = $MBDetails.UserPrincipalName 
        $CreationTime = $MBDetails.WhenCreated
        $MBType = $MBDetails.RecipientTypeDetails
        Get_LastLogonTime    
    }
}

#Get all mailboxes from Office 365
else
{
    $MailBoxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.DisplayName -notlike "Discovery Search Mailbox"}
    ForEach($Mail in $MailBoxes) {
        $DisplayName=$Mail.DisplayName  
        $UPN = $Mail.UserPrincipalName 
        $CreationTime = $Mail.WhenCreated
        $MBType = $Mail.RecipientTypeDetails
        Get_LastLogonTime
    } 
}

#Open output file after execution 
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host "Detailed report available in:" -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
    $Prompt = New-Object -ComObject wscript.shell  
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)  
    if ($UserInput -eq 6)  
    {  
        Invoke-Item "$ExportCSV"  
    } 
}
else
{
    Write-Host "No mailbox found" -ForegroundColor Red
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
CloseConnection

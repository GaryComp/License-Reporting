<#
=============================================================================================
Name:           Identify Non-Compliant Shared Mailboxes in Microsoft 365  
Version:        1.0
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Generates all non-compliant shared mailboxes in Microsoft 365.  
2. Exports report results to CSV file. 
3. The script automatically verifies and installs the MS Graph PowerShell SDK and Exchange Online PowerShell modules (if they are not already installed) upon your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. The script supports Certificate-based authentication (CBA). 
6. The script is scheduler-friendly.   

For detailed Script execution: https://o365reports.com/2024/12/10/identify-non-compliant-shared-mailboxes-in-microsoft-365/


============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [string]$TenantId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

Import-Module "$PSScriptRoot\M365AuthModule.psm1" -Force
Connect-M365Services -Services "Graph","ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password -GraphScopes "User.Read.All","AuditLog.Read.All"


$ProgressIndex = 0
$OutputCSV = "$(Get-Location)\NonCompliant_Shared_Mailboxes_$((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')).csv"
$NonCompliantCount = 0


$ExchangeOnlineServicePlans = @(
    "efb87545-963c-4e0d-99df-69c6916d9eb0", # EXCHANGE_S_ENTERPRISE(EXO Plan2)
    "9aaf7827-d63c-4b61-89c3-182f06f82e5c"  # EXCHANGE_S_STANDARD  (EXO Plan1)
)

# Retrieve all shared mailboxes
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $MailboxName = $_.DisplayName
    $ProgressIndex++
    Write-Progress -Activity "`n     Processed shared mailbox count: $ProgressIndex"`n"  Checking $MailboxName mailbox"
    # Retrieve additional details using Microsoft Graph
    $User = Get-MgUser -UserId $_.ExternalDirectoryObjectId -Property AccountEnabled, DisplayName, UserPrincipalName, SignInActivity | 
            Select-Object DisplayName, UserPrincipalName, AccountEnabled, @{Name="LastSignInTime";Expression={$_.SignInActivity.LastSignInDateTime}}
    
    if ($User.AccountEnabled -eq $true) {
        $LicenseDetails = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName
        $HasExchangeOnline = $false
        foreach ($License in $LicenseDetails) {
            foreach ($ServicePlan in $License.ServicePlans) {
                if ($ExchangeOnlineServicePlans -contains $ServicePlan.ServicePlanId) {
                    $HasExchangeOnline = $true
                    break
                }
            }
            if ($HasExchangeOnline) { break }
        }

        # Check for non-compliance
        if (-not $HasExchangeOnline) {
            $NonCompliantCount++

            # Gather all necessary properties
            $NonCompliantDetails = [PSCustomObject]@{
                "Shared Mailbox Name"    = $User.DisplayName
                "Primary SMTP Address"   = $_.PrimarySmtpAddress
                "Sign-In Enabled"        = "Enabled"
                "Exchange License"       = "No"
                "Last Sign-In Time"      = if ($User.LastSignInTime -eq $null) { "Never logged-in" } else { $User.LastSignInTime }
                "Creation Time"          = $_.WhenCreated
            }

            # Append to CSV
            $NonCompliantDetails | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
        }
    }
}    


Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

# Disconnect sessions
#Disconnect-MgGraph | Out-Null
#Disconnect-ExchangeOnline -Confirm:$false | Out-Null


# Prompt to open CSV if non-compliant mailboxes exist
if(Test-Path -Path $OutputCSV) {   
    Write-Host  " Found $NonCompliantCount non-compliant shared mailboxes." -ForegroundColor Yellow;
    Write-Host  " Output CSV file saved to: " -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV"  
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6) {
        Invoke-Item "$OutputCSV"
    }
}
else {
    Write-Host `n"No non-compliant shared mailboxes found." 
}

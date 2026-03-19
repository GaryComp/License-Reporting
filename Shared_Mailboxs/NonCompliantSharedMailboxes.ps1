<#
=============================================================================================
Name:           Identify Non-Compliant Shared Mailboxes in Microsoft 365  
Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Generates all non-compliant shared mailboxes in Microsoft 365.  
2. Exports report results to CSV file. 
3. The script automatically verifies and installs the MS Graph PowerShell SDK and Exchange Online PowerShell modules (if they are not already installed) upon your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. The script supports Certificate-based authentication (CBA). 
6. The script is scheduler-friendly.   

============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [string]$TenantId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [SecureString]$Password
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force
Connect-M365Services -Services "Graph","ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password -GraphScopes "User.Read.All","AuditLog.Read.All"


$ProgressIndex = 0
$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$OutputCSV = Join-Path $ExportsDir "NonCompliant_Shared_Mailboxes_$((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')).csv"
$NonCompliantCount = 0


$ExchangeOnlineServicePlans = @(
    "efb87545-963c-4e0d-99df-69c6916d9eb0", # EXCHANGE_S_ENTERPRISE(EXO Plan2)
    "9aaf7827-d63c-4b61-89c3-182f06f82e5c"  # EXCHANGE_S_STANDARD  (EXO Plan1)
)

# Retrieve all shared mailboxes
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $MailboxName = $_.DisplayName
    $ProgressIndex++
    Write-Progress -Activity "Processing shared mailboxes" -Status "Processed: $ProgressIndex | Checking: $MailboxName"
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
                "Last Sign-In Time"      = if ($null -eq $User.LastSignInTime) { "Never logged-in" } else { $User.LastSignInTime }
                "Creation Time"          = $_.WhenCreated
            }

            # Append to CSV
            $NonCompliantDetails | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
        }
    }
}
Write-Progress -Activity "Processing shared mailboxes" -Completed

# Disconnect sessions
Disconnect-MgGraph | Out-Null
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

if ($NonCompliantCount -eq 0) {
    Write-Host "`nNo non-compliant shared mailboxes found." -ForegroundColor Yellow
} else {
    Write-Host "`nFound $NonCompliantCount non-compliant shared mailboxes."
    Write-Host "Report available at: " -NoNewline -ForegroundColor Yellow
    Write-Host $OutputCSV
}

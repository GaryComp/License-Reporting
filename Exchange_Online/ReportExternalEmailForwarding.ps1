<#-------------------------------------------------------------------------------------------------------------------------------------------------------------
Name: External Email Forwarding Report for Exchange Online
~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange PowerShell module (if not installed already) upon your confirmation.
2. Exports the 'External email forwarding report' for all mailboxes into a CSV file.
3. Exports the 'Inbox rules with external forwarding report' into a CSV file.
4. Allows checking external email forwarding for specific mailboxes.
5. The script can be executed with an MFA-enabled account too.
6. The script supports Certificate-based authentication (CBA).
--------------------------------------------------------------------------------------------------------------------------------------------------------------#>

param (
    [string] $CertificateThumbprint,
    [string] $ClientId,
    [string] $TenantId,
    [string] $UserName,
    [SecureString] $Password,
    [Switch] $ExcludeGuests,
    [Switch] $ExcludeInternalGuests,
    [String] $MailboxNames
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

Function SplittingLegacyValue{
    param($value)
    $SplittedValue = (($value).split(':') | Select-Object -Index 1)
    $FinalValue = (($SplittedValue).split(']') | Select-Object -Index 0)
    return $FinalValue
}

Function SplittingQuotes{
    param($value)
    $SplitValue = (($value).split('"') | Select-Object -Index 1)
    return $SplitValue
}

Function FindForwardingActionIsExternal{
    param($Action)
    if($Action.contains("[SMTP:")){
        $SplitQuoteValue = SplittingQuotes -value $Action
        return $SplitQuoteValue
    }
    elseif($Action.contains("[EX:")){
        $SplittedLegacy = SplittingLegacyValue -value $Action
        if($ExcludeInternalGuests){
            if($global:InternalGuest.$SplittedLegacy){ return }
        }
        if($global:GuestUsers.$SplittedLegacy){
            if($ExcludeGuests){ return }
            else{
                $SplitQuoteValue = SplittingQuotes -value $Action
                return $SplitQuoteValue
            }
        }
        if($global:MailUsers.$SplittedLegacy){
            $SplitQuoteValue = SplittingQuotes -value $Action
            return $SplitQuoteValue
        }
        if($global:Contacts.$SplittedLegacy){
            $SplitQuoteValue = SplittingQuotes -value $Action
            return $SplitQuoteValue
        }
    }
}

Function GetInboxRule{
    param($Mailbox)
    Write-Progress -Activity "Getting inbox rules with external forwarding for: $($Mailbox)"
    Get-InboxRule -Mailbox $Mailbox | Where-Object {
        ($_.ForwardAsAttachmentTo -ne $Empty -or $_.ForwardTo -ne $Empty -or $_.RedirectTo -ne $Empty) -and ($_.Enabled -eq $True)
    } | ForEach-Object{
        $ForwardTo = @()
        $ForwardAsAttachmentTo = @()
        $RedirectTo = @()
        if($_.ForwardTo){
            ForEach($Forward in $_.ForwardTo){
                $IsExternal = FindForwardingActionIsExternal -Action $Forward
                $ForwardTo = $ForwardTo + $IsExternal
            }
        }
        if($_.RedirectTo){
            ForEach($Redirect in $_.RedirectTo){
                $IsExternal = FindForwardingActionIsExternal -Action $Redirect
                $RedirectTo = $RedirectTo + $IsExternal
            }
        }
        if($_.ForwardAsAttachmentTo){
            ForEach($ForwardAsAttach in $_.ForwardAsAttachmentTo){
                $IsExternal = FindForwardingActionIsExternal -Action $ForwardAsAttach
                $ForwardAsAttachmentTo = $ForwardAsAttachmentTo + $IsExternal
            }
        }
        if(($ForwardTo.count -gt 0) -or ($ForwardAsAttachmentTo.count -gt 0) -or ($RedirectTo.count -gt 0)){
            $ExportResult = @{
                'Mailbox Name'            = $_.MailboxOwnerId
                'User Principal Name'     = $Mailbox
                'Inbox Rule Name'         = $_.Name
                'Rule Identity'           = $_.Identity
                'Forward To'              = $ForwardTo -join ","
                'Forward As Attachment To'= $ForwardAsAttachmentTo -join ","
                'Redirect To'             = $RedirectTo -join ","
            }
            $ExportResults = New-Object PSObject -Property $ExportResult
            $ExportResults | Select-Object 'Mailbox Name','User Principal Name','Inbox Rule Name','Rule Identity','Forward To','Forward As Attachment To','Redirect To' |
                Export-Csv -Path $global:ExportInboxRule -NoType -Append -Force
        }
    }
    Write-Progress -Activity "Getting inbox rules with external forwarding for: $($Mailbox)" -Completed
}

Function CheckEmailForwardingForExternal{
    param($Mailbox)
    Write-Progress -Activity "Checking external forwarding for: $($Mailbox.DisplayName)"
    $ForwardingAddress         = '-'
    $ForwardingSMTPAddress     = '-'
    $ExternalForwardingAddress = '-'
    $ExternalForwardingSMTPAddress = '-'
    $ExternalAddressFound      = $True
    $ExternalSMTPAddressFound  = $True

    if($Mailbox.ForwardingAddress){
        $ExternalAddressFound = $False
        $ForwardingAddress = $Mailbox.ForwardingAddress
        if($global:Contacts.$ForwardingAddress){
            $ExternalForwardingAddress = $ForwardingAddress
            $ExternalAddressFound = $True
        }
    }
    if($Mailbox.ForwardingSMTPAddress){
        $ExternalSMTPAddressFound = $False
        $ForwardingSMTPAddress = (($Mailbox.ForwardingSMTPAddress).split(":") | Select-Object -Index 1)
        $checkDomainIsInternal = (($Mailbox.ForwardingSMTPAddress).split("@") | Select-Object -Index 1)
        if(!$global:Domain.$checkDomainIsInternal){
            $ExternalForwardingSMTPAddress = $ForwardingSMTPAddress
            $ExternalSMTPAddressFound = $True
        }
    }
    if($ExcludeInternalGuests){
        if(!$ExternalAddressFound -and $global:InternalGuest.$ForwardingAddress)     { $ExternalAddressFound     = $True }
        if(!$ExternalSMTPAddressFound -and $global:InternalGuest.$ForwardingSMTPAddress) { $ExternalSMTPAddressFound = $True }
    }
    if(!$ExternalAddressFound){
        if($global:GuestUsers.$ForwardingAddress){
            if($ExcludeGuests){ $ExternalAddressFound = $True }
            else{ $ExternalForwardingAddress = $ForwardingAddress; $ExternalAddressFound = $True }
        }
        elseif($global:MailUsers.$ForwardingAddress){
            $ExternalForwardingAddress = $ForwardingAddress; $ExternalAddressFound = $True
        }
    }
    if(!$ExternalSMTPAddressFound){
        if($global:GuestUsers.$ForwardingSMTPAddress){
            if($ExcludeGuests){ $ExternalSMTPAddressFound = $True }
            else{ $ExternalForwardingSMTPAddress = $ForwardingSMTPAddress; $ExternalSMTPAddressFound = $True }
        }
        elseif($global:MailUsers.$ForwardingSMTPAddress){
            $ExternalForwardingSMTPAddress = $ForwardingSMTPAddress; $ExternalSMTPAddressFound = $True
        }
    }

    if(($ExternalForwardingAddress -ne "-") -or ($ExternalForwardingSMTPAddress -ne "-")){
        $ExportResult = @{
            'Display Name'                 = $Mailbox.DisplayName
            'User Principal Name'          = $Mailbox.UserPrincipalName
            'Forwarding Address'           = $ExternalForwardingAddress
            'Forwarding SMTP Address'      = $ExternalForwardingSMTPAddress
            'Deliver To Mailbox and Forward' = $Mailbox.DeliverToMailboxAndForward
        }
        $ExportResults = New-Object PSObject -Property $ExportResult
        $ExportResults | Select-Object 'Display Name','User Principal Name','Forwarding Address','Forwarding SMTP Address','Deliver To Mailbox and Forward' |
            Export-Csv -Path $global:ExportEmailForwarding -NoType -Append -Force
    }
    Write-Progress -Activity "Checking external forwarding for: $($Mailbox.DisplayName)" -Completed
}

Function GetMailbox{
    if($MailboxNames){
        Import-Csv $MailboxNames | ForEach-Object {
            $Mailbox = Get-EXOMailbox $($_.'User Principal Name') -Properties ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward |
                Where-Object {($null -ne $_.ForwardingAddress) -or ($null -ne $_.ForwardingSMTPAddress)}
            if($Mailbox){ CheckEmailForwardingForExternal -Mailbox $Mailbox }
            GetInboxRule -Mailbox $_.'User Principal Name'
        }
    }
    else{
        Get-EXOMailbox -ResultSize Unlimited -Properties ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward | ForEach-Object {
            if(($null -ne $_.ForwardingAddress) -or ($null -ne $_.ForwardingSMTPAddress)){
                CheckEmailForwardingForExternal -Mailbox $_
            }
            GetInboxRule -Mailbox $_.UserPrincipalName
        }
    }
}

Function ShowOutputFiles{
    $emailExists = Test-Path -Path $global:ExportEmailForwarding
    $ruleExists  = Test-Path -Path $global:ExportInboxRule

    if(-not $emailExists -and -not $ruleExists){
        Write-Host "`nNo external forwarding found for the given input." -ForegroundColor Green
        return
    }
    Write-Host "`nReport(s) saved to the Exports folder:" -ForegroundColor Yellow
    if($emailExists){ Write-Host "  $global:ExportEmailForwarding" -ForegroundColor Cyan }
    if($ruleExists) { Write-Host "  $global:ExportInboxRule"       -ForegroundColor Cyan }
}

#................................................Execution starts here.....................................................

if($MailboxNames -and -not (Test-Path $MailboxNames -PathType Leaf)){
    Write-Host "Error: The specified CSV file does not exist or is not accessible." -ForegroundColor Red
    Exit
}

Connect-M365Services -Services "ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password

Write-Host "`nFetching mailboxes with external email forwarding..."

$global:Domain = @{}
Get-AcceptedDomain | ForEach-Object{ $global:Domain[$_.DomainName] = $_.DomainName }

$global:GuestUsers = @{}
Get-User -ResultSize Unlimited | Where-Object {$_.UserType -eq 'Guest'} | Select-Object Identity,LegacyExchangeDN,UserPrincipalName | ForEach-Object{
    $global:GuestUsers[$_.Identity]          = $_.Identity
    $global:GuestUsers[$_.LegacyExchangeDN]  = $_.LegacyExchangeDN
    $global:GuestUsers[$_.UserPrincipalName] = $_.UserPrincipalName
}

if($ExcludeInternalGuests){
    $global:InternalGuest = @{}
    Get-User -ResultSize Unlimited | Where-Object {$_.UserPersona -eq 'InternalGuest'} | Select-Object Identity,LegacyExchangeDN,UserPrincipalName | ForEach-Object{
        $global:InternalGuest[$_.Identity]          = $_.Identity
        $global:InternalGuest[$_.LegacyExchangeDN]  = $_.LegacyExchangeDN
        $global:InternalGuest[$_.UserPrincipalName] = $_.UserPrincipalName
    }
}

$global:Contacts = @{}
Get-MailContact -ResultSize Unlimited | Select-Object Identity,LegacyExchangeDN | ForEach-Object{
    $global:Contacts[$_.Identity]         = $_.Identity
    $global:Contacts[$_.LegacyExchangeDN] = $_.LegacyExchangeDN
}

$global:MailUsers = @{}
Get-MailUser -ResultSize Unlimited | Select-Object Identity,LegacyExchangeDN,UserPrincipalName | ForEach-Object{
    $global:MailUsers[$_.Identity]          = $_.Identity
    $global:MailUsers[$_.LegacyExchangeDN]  = $_.LegacyExchangeDN
    $global:MailUsers[$_.UserPrincipalName] = $_.UserPrincipalName
}

$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$global:ExportEmailForwarding = Join-Path $ExportsDir ("ExternalEmailForwardingReport_" + (Get-Date -Format "yyyy-MMM-dd-ddd_hh-mm-ss_tt") + ".csv")
$global:ExportInboxRule       = Join-Path $ExportsDir ("InboxRulesWithExternalForwardingReport_" + (Get-Date -Format "yyyy-MMM-dd-ddd_hh-mm-ss_tt") + ".csv")

GetMailbox
ShowOutputFiles

Disconnect-ExchangeOnline -Confirm:$false

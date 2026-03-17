<#
=============================================================================================
Name:           Find shared mailboxes with licenses in Office 365
Description:    This script exports licensed shared mailboxes
Updated March 16-2026
CLutchSolutions.com

============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

Import-Module "$PSScriptRoot\M365AuthModule.psm1" -Force
Connect-M365Services -Services "Graph","ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password -GraphScopes "User.Read.All","Organization.Read.All"


# Build SKU GUID -> friendly name lookup table from subscribed SKUs
Write-Host "`nBuilding license SKU lookup table..." -ForegroundColor Cyan
$SkuTable = @{}
Get-MgSubscribedSku | ForEach-Object { $SkuTable[$_.SkuId] = $_.SkuPartNumber }

$ExportCSV = ".\LicensedSharedMailboxesReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$Count = 0

Write-Host "Retrieving licensed shared mailboxes..." -ForegroundColor Cyan

# Get all shared mailboxes from Exchange Online, then check Graph for licenses
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $UPN = $_.UserPrincipalName
    $Name = $_.DisplayName

    # Retrieve license assignments from Graph
    $UserLicenses = (Get-MgUser -UserId $UPN -Property AssignedLicenses).AssignedLicenses

    # Skip if no licenses assigned
    if ($UserLicenses.Count -eq 0) { return }

    $Count++
    Write-Progress -Activity "Found $Count licensed shared mailboxes" -Status "Currently processing: $Name"

    $LitigationHoldEnabled = $_.LitigationHoldEnabled
    $InPlaceHoldEnabled = if ($_.InPlaceHolds -ne $null -and $_.InPlaceHolds.Count -gt 0) { "True" } else { "False" }

    $MailboxItemSize = (Get-MailboxStatistics -Identity $UPN).TotalItemSize.Value
    $MailboxItemSizeParts = $MailboxItemSize.ToString().Split("()")
    $MBSize = $MailboxItemSizeParts | Select-Object -Index 0
    $MBSizeInBytes = $MailboxItemSizeParts | Select-Object -Index 1

    # Resolve SKU GUIDs to friendly part numbers
    $AssignedLicenses = ($UserLicenses | ForEach-Object {
        if ($SkuTable.ContainsKey($_.SkuId)) { $SkuTable[$_.SkuId] } else { $_.SkuId }
    }) -join ","

    [PSCustomObject]@{
        'Name'                    = $Name
        'UPN'                     = $UPN
        'Shared MB Size'          = $MBSize
        'MB Size (Bytes)'         = $MBSizeInBytes
        'Litigation Hold Enabled' = $LitigationHoldEnabled
        'In-place Archive Enabled'= $InPlaceHoldEnabled
        'Assigned Licenses'       = $AssignedLicenses
    } | Export-Csv $ExportCSV -NoTypeInformation -Append
}

# Open output file after execution
If ($Count -eq 0) {
    Write-Host "No shared mailbox found with a license."
} else {
    Write-Host "`nThe output file contains $Count licensed shared mailboxes."
    if (Test-Path -Path $ExportCSV) {
        Write-Host "`nThe output file is available at:" -NoNewline -ForegroundColor Yellow
        Write-Host " $ExportCSV"
        $Prompt = New-Object -ComObject wscript.shell
        If ($UserInput -eq 6) {
            Invoke-Item "$ExportCSV"
        }
    }
}

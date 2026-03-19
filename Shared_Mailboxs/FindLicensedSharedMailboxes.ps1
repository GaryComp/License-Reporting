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
    [SecureString]$Password
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force
Connect-M365Services -Services "Graph","ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password -GraphScopes "User.Read.All","Organization.Read.All"


# Build SKU GUID -> friendly name lookup table from subscribed SKUs
Write-Host "`nBuilding license SKU lookup table..." -ForegroundColor Cyan
$SkuTable = @{}
Get-MgSubscribedSku | ForEach-Object { $SkuTable[$_.SkuId] = $_.SkuPartNumber }

$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$ExportCSV = Join-Path $ExportsDir "LicensedSharedMailboxesReport_$((Get-Date -format 'yyyy-MMM-dd-ddd hh-mm tt').ToString()).csv"
$Count = 0

Write-Host "Retrieving licensed shared mailboxes..." -ForegroundColor Cyan

# Get all shared mailboxes from Exchange Online, then check Graph for licenses
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $UPN = $_.UserPrincipalName
    $Name = $_.DisplayName

    # Retrieve license assignments from Graph
    $UserLicenses = (Get-MgUser -UserId $UPN -Property AssignedLicenses -ErrorAction SilentlyContinue).AssignedLicenses

    # Skip if no licenses assigned
    if ($UserLicenses.Count -eq 0) { return }

    $Count++
    Write-Progress -Activity "Found $Count licensed shared mailboxes" -Status "Currently processing: $Name"

    $LitigationHoldEnabled = $_.LitigationHoldEnabled
    $InPlaceHoldEnabled = if ($null -ne $_.InPlaceHolds -and $_.InPlaceHolds.Count -gt 0) { "True" } else { "False" }

    $MBSize = "-"
    $MBSizeInBytes = "-"
    $MailboxStats = Get-MailboxStatistics -Identity $UPN -ErrorAction SilentlyContinue
    if ($null -ne $MailboxStats -and $null -ne $MailboxStats.TotalItemSize) {
        $MailboxItemSizeParts = $MailboxStats.TotalItemSize.Value.ToString().Split("()")
        $MBSize = $MailboxItemSizeParts | Select-Object -Index 0
        $MBSizeInBytes = $MailboxItemSizeParts | Select-Object -Index 1
    }

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
Write-Progress -Activity "Processing shared mailboxes" -Completed

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

If ($Count -eq 0) {
    Write-Host "No shared mailbox found with a license." -ForegroundColor Yellow
} else {
    Write-Host "`nThe output file contains $Count licensed shared mailboxes."
    Write-Host "Report available at: " -NoNewline -ForegroundColor Yellow
    Write-Host $ExportCSV
}

#Version 2.1
Param(
    [switch]$Trial,
    [switch]$Free,
    [switch]$Purchased,
    [switch]$Expired,
    [switch]$Active,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force
Connect-M365Services -Services "Graph" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

# ------------------------- Output Paths -------------------------
$TimeStamp = Get-Date -Format "yyyy-MMM-dd-ddd__hh-mm_tt"
$ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
$ExportCSV = Join-Path $ExportsDir "LicenseExpiryReport_$TimeStamp.csv"

# ------------------------- Load Friendly Names -------------------------
$FriendlyNameHash = @{}
$skuCsvPath = Join-Path $PSScriptRoot '..' 'Supporting_Files' 'Product names and service plan identifiers for licensing.csv'
if (Test-Path $skuCsvPath) {
    Import-Csv $skuCsvPath | ForEach-Object {
        if ($_.SkuPartNumber -and $_.Product_Display_Name) {
            $FriendlyNameHash[$_.SkuPartNumber.Trim()] = $_.Product_Display_Name.Trim()
        }
    }
}
else {
    Write-Host "Warning: Product names CSV not found. Friendly names will not be resolved." -ForegroundColor Yellow
}

# ------------------------- Determine Filters -------------------------
$ShowAll = -not ($Trial -or $Free -or $Purchased -or $Expired -or $Active)

# ------------------------- Retrieve License Info -------------------------
Write-Host "`nRetrieving subscribed SKUs..." -ForegroundColor Cyan
$Skus = Get-MgSubscribedSku -All

# Lifecycle info (try v1.0, fallback to beta)
$lifecycleUriV1   = "https://graph.microsoft.com/v1.0/directory/subscriptions"
$lifecycleUriBeta = "https://graph.microsoft.com/beta/directory/subscriptions"
$lifecycleInfo = $null

try {
    $lifecycleInfo = (Invoke-MgGraphRequest -Uri $lifecycleUriV1 -Method GET -ErrorAction Stop).value
}
catch {
    Write-Host "Warning: v1.0 lifecycle endpoint failed, falling back to beta endpoint" -ForegroundColor Yellow
    $lifecycleInfo = (Invoke-MgGraphRequest -Uri $lifecycleUriBeta -Method GET -ErrorAction Stop).value
}

# ------------------------- Process Results -------------------------
$Results = @()
foreach ($Sku in $Skus) {
    $SkuId = $Sku.SkuId
    $SkuPartNumber = $Sku.SkuPartNumber

    if ($FriendlyNameHash.ContainsKey($SkuPartNumber)) {
        $FriendlyName = $FriendlyNameHash[$SkuPartNumber]
    } else {
        $FriendlyName = $SkuPartNumber
    }

    $Consumed = if ($Sku.ConsumedUnits) { [int]$Sku.ConsumedUnits } else { 0 }

    $Lifecycle = $lifecycleInfo | Where-Object { $_.skuId -eq $SkuId }
    if (-not $Lifecycle) { continue }

    $Created    = ($Lifecycle.createdDateTime | Select-Object -First 1)
    $Status     = ($Lifecycle.status | Select-Object -First 1)
    $Total      = ($Lifecycle.totalLicenses | Select-Object -First 1)
    $ExpiryDate = ($Lifecycle.nextLifecycleDateTime | Select-Object -First 1)

    # Ensure numeric subtraction works
    $Remaining = 0
    if ($Total -and $Consumed -is [int]) {
        $Remaining = $Total - $Consumed
    }

    # Subscription type
    if ($SkuPartNumber -like "*Free*" -and -not $ExpiryDate) {
        $Type = "Free"
    }
    elseif ($Lifecycle.isTrial) {
        $Type = "Trial"
    }
    else {
        $Type = "Purchased"
    }

    # Subscribed date
    if ($Created) {
        try {
            $SubscribedDate = [datetime]$Created
            $SubscribedAgo = (New-TimeSpan -Start $SubscribedDate -End (Get-Date)).Days
            if ($SubscribedAgo -eq 0) {
                $SubscribedFriendly = "Today"
            } else {
                $SubscribedFriendly = "$SubscribedAgo days ago"
            }
            $SubscribedString = "$SubscribedDate ($SubscribedFriendly)"
        }
        catch {
            $SubscribedString = "Invalid date"
        }
    }
    else {
        $SubscribedString = "Unknown"
    }

    # Expiry date
    if ($ExpiryDate) {
        try {
            $ExpiryDateTime = [datetime]$ExpiryDate
            $DaysRemaining = (New-TimeSpan -Start (Get-Date) -End $ExpiryDateTime).Days
            switch -Regex ($Status) {
                "^Enabled$"   { $ExpiryNote = "Will expire in $DaysRemaining days" }
                "^Warning$"   { $ExpiryNote = "Expired. Will suspend in $DaysRemaining days" }
                "^Suspended$" { $ExpiryNote = "Expired. Will delete in $DaysRemaining days" }
                "^LockedOut$" { $ExpiryNote = "Subscription is locked. Contact Microsoft." }
                default       { $ExpiryNote = "Unknown status" }
            }
        }
        catch {
            $ExpiryNote = "Invalid expiry date"
            $ExpiryDateTime = "-"
        }
    }
    else {
        $ExpiryNote = "Never Expires"
        $ExpiryDateTime = "-"
    }

    # Apply filters
    $Include = $false
    if ($ShowAll) { $Include = $true }
    if ($Trial -and $Type -eq "Trial") { $Include = $true }
    if ($Free -and $Type -eq "Free") { $Include = $true }
    if ($Purchased -and $Type -eq "Purchased") { $Include = $true }
    if ($Expired -and $Status -ne "Enabled") { $Include = $true }
    if ($Active -and $Status -eq "Enabled") { $Include = $true }

    if ($Include) {
        $Results += New-Object PSObject -Property @{
            "Subscription Name"                             = $SkuPartNumber
            "Friendly Subscription Name"                    = $FriendlyName
            "Subscribed Date"                               = $SubscribedString
            "Total Units"                                   = $Total
            "Consumed Units"                                = $Consumed
            "Remaining Units"                               = $Remaining
            "Subscription Type"                             = $Type
            "License Expiry Date / Next Lifecycle Activity" = $ExpiryDateTime
            "Friendly Expiry Date"                          = $ExpiryNote
            "Status"                                        = $Status
            "SKU Id"                                        = $SkuId
        }
    }
}

# ------------------------- Export -------------------------
if ($Results.Count -gt 0) {
    $Results | Export-Csv -Path $ExportCSV -NoTypeInformation
    Write-Host "`nReport saved to: " -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" -ForegroundColor Cyan
    Write-Host "$($Results.Count) subscriptions included.`n"
}
else {
    Write-Host "No subscriptions matched the given filters." -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null

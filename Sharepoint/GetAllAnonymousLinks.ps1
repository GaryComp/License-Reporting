<#
=============================================================================================
Name:           Get All Anonymous Links in SharePoint Online
Version:        2.0
Website:        ClutchSolutions.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Exports all anonymous links in your SharePoint Online environment.
2. Identifies files and folders with only active anonymous links.
3. Lists files and folders that have only expired anonymous links.
4. Generates a report for never-expiring anonymous links.
5. Exports a report for anonymous links set with expiration.
6. Allows export of files and folders with soon-to-expire links (e.g., 30 days, 90 days).
7. Supports certificate-based (unattended) and interactive authentication.
8. Exports report results to CSV in the shared Exports folder.

~~~~~~~~~~~~~~~~~~
Note:
~~~~~~~~~~~~~~~~~~
The app registration used for certificate-based authentication must be granted Application
permissions for "Files.Read.All" and "Sites.Read.All". Without these you will receive:
"Get-PnPFileSharingLink: Either scp or roles claim need to be present in the token."

~~~~~~~~~~~~~~~~~~
Required Graph/PnP Permissions (Application):
~~~~~~~~~~~~~~~~~~
  - Sites.Read.All
  - Files.Read.All

~~~~~~~~~~~
Change Log:
~~~~~~~~~~~
V1.0 (Original) - Get All Anonymous Links report (source: github.com/admindroid-community)
V2.0 (18-Mar-2026) - Refactored for Clutch Solutions: shared auth module, Exports folder,
                     null-comparison fixes, dead code removal, style alignment.
============================================================================================
#>
Param (
    [string]$TenantId,
    [string]$TenantName            = "",
    [string]$ClientId              = "",
    [string]$CertificateThumbprint = "",
    [string]$UserName              = "",
    [SecureString]$Password        = $null,
    [string]$ImportCsv             = "",
    [Switch]$ActiveLinks,
    [Switch]$ExpiredLinks,
    [Switch]$LinksWithExpiration,
    [Switch]$NeverExpiresLinks,
    [int]$SoonToExpireInDays
)

Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force

Function Connect-SiteModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    Connect-M365Services -Services PnP -SiteUrl $Url `
        -TenantId $TenantId -ClientId $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -UserName $UserName -Password $Password
}

Function Get-SharedLinks {
    try {
        $ExcludedLists = @("Form Templates", "Style Library", "Site Assets", "Site Pages",
            "Preservation Hold Library", "Pages", "Images",
            "Site Collection Documents", "Site Collection Images")

        $DocumentLibraries = Get-PnPList | Where-Object {
            $_.Hidden -eq $false -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"
        }

        foreach ($List in $DocumentLibraries) {
            $ListItems = Get-PnPListItem -List $List -PageSize 2000
            foreach ($Item in $ListItems) {
                $FileName    = $Item.FieldValues.FileLeafRef
                $ObjectType  = $Item.FileSystemObjectType
                Write-Progress -Activity "Site: $($Site.Title)" -Status "Processing: $FileName"

                $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments
                if (-not $HasUniquePermissions) { continue }

                $FileUrl = $Item.FieldValues.FileRef
                if ($ObjectType -eq "File") {
                    $FileSharingLinks = Get-PnPFileSharingLink -Identity $FileUrl
                } elseif ($ObjectType -eq "Folder") {
                    $FileSharingLinks = Get-PnPFolderSharingLink -Folder $FileUrl
                } else {
                    continue
                }

                foreach ($FileSharingLink in $FileSharingLinks) {
                    $Link = $FileSharingLink.Link
                    if ($Link.Scope -ne "Anonymous") { continue }

                    $Permission      = $Link.Type
                    $SharedLink      = $Link.WebUrl
                    $PasswordProtected = $FileSharingLink.HasPassword
                    $BlockDownload   = $Link.PreventsDownload
                    $RoleList        = $FileSharingLink.Roles -join ","
                    $ExpirationDate  = $FileSharingLink.ExpirationDateTime
                    $CurrentDateTime = (Get-Date).Date

                    if ($null -ne $ExpirationDate) {
                        $ExpiryDate  = ([DateTime]$ExpirationDate).ToLocalTime()
                        $ExpiryDays  = (New-TimeSpan -Start $CurrentDateTime -End $ExpiryDate).Days
                        if ($ExpiryDate -lt $CurrentDateTime) {
                            $LinkStatus          = "Expired"
                            $FriendlyExpiryTime  = "Expired $($ExpiryDays * -1) days ago"
                        } else {
                            $LinkStatus         = "Active"
                            $FriendlyExpiryTime = "Expires in $ExpiryDays days"
                        }
                    } else {
                        $LinkStatus         = "Active"
                        $ExpiryDays         = "-"
                        $ExpiryDate         = "-"
                        $FriendlyExpiryTime = "Never Expires"
                    }

                    # Apply filters
                    if ($ActiveLinks.IsPresent       -and $LinkStatus -ne "Active")          { continue }
                    if ($ExpiredLinks.IsPresent       -and $LinkStatus -ne "Expired")         { continue }
                    if ($LinksWithExpiration.IsPresent -and $null -eq $ExpirationDate)        { continue }
                    if ($NeverExpiresLinks.IsPresent  -and $FriendlyExpiryTime -ne "Never Expires") { continue }
                    if ($SoonToExpireInDays -ne 0 -and (($null -eq $ExpirationDate) -or ($SoonToExpireInDays -lt $ExpiryDays) -or ($ExpiryDays -lt 0))) { continue }

                    [PSCustomObject]@{
                        "Site Name"            = $Site.Title
                        "Library"              = $List.Title
                        "Object Type"          = $ObjectType
                        "File/Folder Name"     = $FileName
                        "File/Folder URL"      = $FileUrl
                        "Access Type"          = $Permission
                        "Roles"                = $RoleList
                        "File Type"            = $Item.FieldValues.File_x0020_Type
                        "Link Status"          = $LinkStatus
                        "Link Expiry Date"     = $ExpiryDate
                        "Days Since/To Expiry" = $ExpiryDays
                        "Friendly Expiry Time" = $FriendlyExpiryTime
                        "Password Protected"   = $PasswordProtected
                        "Block Download"       = $BlockDownload
                        "Shared Link"          = $SharedLink
                    } | Export-Csv -Path $ReportOutput -NoTypeInformation -Append

                    $Global:ItemCount++
                }
            }
        }
    } catch {
        Write-Host "$($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$ExportsDir = Join-Path $PSScriptRoot ".." "Exports"
if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }

$ReportOutput      = Join-Path $ExportsDir "AnonymousLink_Report_$((Get-Date -Format 'yyyy-MMM-dd-ddd_HH-mm-ss').ToString()).csv"
$Global:ItemCount  = 0

# Resolve TenantName from config if not supplied as a parameter
if ([string]::IsNullOrEmpty($TenantName)) {
    $configPath = Join-Path $PSScriptRoot ".." "M365AuthConfig.psd1"
    if (Test-Path $configPath) {
        $authConfig = Import-PowerShellDataFile -Path $configPath
        if ($authConfig.AdminUrl) {
            # Extract from "https://tenantname-admin.sharepoint.com"
            $TenantName = $authConfig.AdminUrl -replace 'https://', '' -replace '-admin\.sharepoint\.com.*', ''
        }
    }
}
if ([string]::IsNullOrEmpty($TenantName)) {
    $TenantName = Read-Host "Enter your tenant name (e.g., 'contoso' for 'contoso.onmicrosoft.com')"
}

if (-not [string]::IsNullOrEmpty($ImportCsv)) {
    $SiteCollections = Import-Csv -Path $ImportCsv
    foreach ($Site in $SiteCollections) {
        Connect-SiteModule -Url $Site.SiteUrl
        try {
            $Site = Get-PnPWeb
            Get-SharedLinks
        } catch {
            Write-Host "$($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    Disconnect-PnPOnline -WarningAction SilentlyContinue
} else {
    Connect-SiteModule -Url "https://$TenantName-admin.sharepoint.com"
    $SiteCollections = Get-PnPTenantSite | Where-Object {
        $_.Template -notin @("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0",
                              "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
    }
    foreach ($Site in $SiteCollections) {
        Connect-SiteModule -Url $Site.Url
        try {
            Get-SharedLinks
        } catch {
            Write-Host "$($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    Disconnect-PnPOnline -WarningAction SilentlyContinue
}

Write-Progress -Activity "Processing anonymous links..." -Completed

if (Test-Path $ReportOutput) {
    Write-Host "`nThe output file contains $($Global:ItemCount) anonymous links." -ForegroundColor Green
    Write-Host "Report saved to: $ReportOutput" -ForegroundColor Green
} else {
    Write-Host "`nNo records found." -ForegroundColor Yellow
}

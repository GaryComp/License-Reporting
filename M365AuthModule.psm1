<#
.SYNOPSIS
    Shared Microsoft 365 authentication module for Clutch Solutions management scripts.

.DESCRIPTION
    Provides a single Connect-M365Services function that handles module installation checks
    and authentication for Microsoft Graph, Exchange Online, SharePoint Online, and PnP Online.
    Supports certificate-based (unattended), credential-based, and interactive authentication.

.NOTES
    Requires a single App Registration configured per AppRegistration-Setup.md
    Updated: March 2026 - ClutchSolutions.com
#>

function Connect-M365Services {
    <#
    .SYNOPSIS
        Connects to one or more Microsoft 365 services using a unified authentication path.

    .PARAMETER Services
        One or more services to connect to: Graph, ExchangeOnline, SharePoint, PnP

    .PARAMETER TenantId
        Azure AD Tenant ID (GUID or primary domain). Required for certificate-based auth.

    .PARAMETER ClientId
        App Registration Application (Client) ID. Required for certificate-based auth.

    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate uploaded to the App Registration.

    .PARAMETER UserName
        UPN for credential-based auth (legacy, avoid where possible).

    .PARAMETER Password
        Plain-text password for credential-based auth (legacy, avoid where possible).

    .PARAMETER AdminUrl
        SharePoint admin URL, e.g. https://contoso-admin.sharepoint.com  Required for SharePoint service.

    .PARAMETER SiteUrl
        Site collection URL for PnP connections, e.g. https://contoso.sharepoint.com/sites/MySite

    .PARAMETER GraphScopes
        Graph permission scopes used during interactive auth. Defaults cover all scripts in this repo.

    .EXAMPLE
        # Certificate-based (unattended)
        Connect-M365Services -Services "Graph","ExchangeOnline" `
            -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

    .EXAMPLE
        # Interactive browser auth
        Connect-M365Services -Services "ExchangeOnline"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Graph", "ExchangeOnline", "SharePoint", "PnP")]
        [string[]]$Services,

        [string]$TenantId             = "",
        [string]$ClientId             = "",
        [string]$CertificateThumbprint = "",
        [string]$UserName             = "",
        [string]$Password             = "",
        [string]$AdminUrl             = "",
        [string]$SiteUrl              = "",
        [string[]]$GraphScopes        = @(
            "User.Read.All",
            "Directory.Read.All",
            "AuditLog.Read.All",
            "Organization.Read.All"
        )
    )

    $IsCertAuth = ($TenantId -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
    $IsCredAuth = ($UserName -ne "" -and $Password -ne "")

    $authMode = if ($IsCertAuth) { "Certificate" } elseif ($IsCredAuth) { "Credential" } else { "Interactive" }
    Write-Host "Authentication mode: $authMode" -ForegroundColor Cyan

    # ── Microsoft Graph ──────────────────────────────────────────────────────
    if ("Graph" -in $Services) {
        _Assert-Module -Name "Microsoft.Graph.Authentication"

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                Connect-MgGraph -TenantId $TenantId `
                                -ClientId $ClientId `
                                -CertificateThumbprint $CertificateThumbprint `
                                -NoWelcome -ErrorAction Stop
            } elseif ($IsCredAuth) {
                $Credential = _New-Credential -UserName $UserName -Password $Password
                Connect-MgGraph -Credential $Credential -NoWelcome -ErrorAction Stop
            } else {
                Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
            }
        } catch {
            Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }

        if ($null -eq (Get-MgContext)) {
            throw "Microsoft Graph connection could not be verified."
        }
        Write-Host "Microsoft Graph connected successfully." -ForegroundColor Green
    }

    # ── Exchange Online ───────────────────────────────────────────────────────
    if ("ExchangeOnline" -in $Services) {
        _Assert-Module -Name "ExchangeOnlineManagement"

        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                Connect-ExchangeOnline -AppId $ClientId `
                                       -CertificateThumbprint $CertificateThumbprint `
                                       -Organization $TenantId `
                                       -ShowBanner:$false -ErrorAction Stop
            } elseif ($IsCredAuth) {
                $Credential = _New-Credential -UserName $UserName -Password $Password
                Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
        } catch {
            Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
        Write-Host "Exchange Online connected successfully." -ForegroundColor Green
    }

    # ── SharePoint Online ─────────────────────────────────────────────────────
    if ("SharePoint" -in $Services) {
        _Assert-Module -Name "Microsoft.Online.SharePoint.PowerShell"

        if ($AdminUrl -eq "") {
            throw "AdminUrl is required for SharePoint connections (e.g. https://contoso-admin.sharepoint.com)."
        }

        Write-Host "Connecting to SharePoint Online ($AdminUrl)..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                Connect-SPOService -Url $AdminUrl `
                                   -ClientId $ClientId `
                                   -Thumbprint $CertificateThumbprint `
                                   -Tenant $TenantId -ErrorAction Stop | Out-Null
            } elseif ($IsCredAuth) {
                $Credential = _New-Credential -UserName $UserName -Password $Password
                Connect-SPOService -Url $AdminUrl -Credential $Credential -ErrorAction Stop | Out-Null
            } else {
                Connect-SPOService -Url $AdminUrl -ErrorAction Stop | Out-Null
            }
        } catch {
            Write-Host "Failed to connect to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
        Write-Host "SharePoint Online connected successfully." -ForegroundColor Green
    }

    # ── PnP Online ────────────────────────────────────────────────────────────
    if ("PnP" -in $Services) {
        _Assert-Module -Name "PnP.PowerShell" -MinimumVersion "1.12.0"

        if ($SiteUrl -eq "") {
            throw "SiteUrl is required for PnP connections."
        }

        Write-Host "Connecting to PnP Online ($SiteUrl)..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                Connect-PnPOnline -Url $SiteUrl `
                                  -ClientId $ClientId `
                                  -Thumbprint $CertificateThumbprint `
                                  -Tenant $TenantId -ErrorAction Stop
            } elseif ($IsCredAuth) {
                $Credential = _New-Credential -UserName $UserName -Password $Password
                Connect-PnPOnline -Url $SiteUrl -Credential $Credential `
                                  -ClientId $ClientId -ErrorAction Stop
            } else {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive -ErrorAction Stop
            }
        } catch {
            Write-Host "Failed to connect to PnP Online: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
        Write-Host "PnP Online connected successfully." -ForegroundColor Green
    }
}

# ── Private helpers ───────────────────────────────────────────────────────────

function _Assert-Module {
    param(
        [string]$Name,
        [string]$MinimumVersion = ""
    )
    $params = @{ Name = $Name; ErrorAction = "SilentlyContinue" }
    if ($MinimumVersion -ne "") { $params.MinimumVersion = $MinimumVersion }

    $module = Get-InstalledModule @params
    if ($null -eq $module) {
        Write-Host "$Name module is not available." -ForegroundColor Yellow
        $confirm = Read-Host "Install $Name now? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Write-Host "Installing $Name..."
            Install-Module -Name $Name -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        } else {
            throw "$Name module is required but not installed. Run: Install-Module $Name"
        }
    }
}

function _New-Credential {
    param([string]$UserName, [string]$Password)
    $secure = ConvertTo-SecureString -AsPlainText $Password -Force
    return New-Object System.Management.Automation.PSCredential($UserName, $secure)
}

Export-ModuleMember -Function Connect-M365Services

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
        SecureString password for credential-based auth (legacy, avoid where possible).

    .PARAMETER AdminUrl
        SharePoint admin URL, e.g. https://contoso-admin.sharepoint.com  Required for SharePoint service.

    .PARAMETER SiteUrl
        Site collection URL for PnP connections, e.g. https://contoso.sharepoint.com/sites/MySite

    .PARAMETER GraphScopes
        Graph permission scopes used during interactive auth. Defaults cover all scripts in this repo. Can be overridden in the config file.

    .PARAMETER ConfigPath
        Optional path to an auth config file (PowerShell data file / .psd1). When omitted, the module looks for a file named
        'M365AuthConfig.psd1' next to this module.

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
        [SecureString]$Password        = $null,
        [string]$AdminUrl             = "",
        [string]$SiteUrl              = "",
        [string]$ConfigPath           = "",
        [string[]]$GraphScopes        = $null
    )

    # Load auth config from a shared file if available (allows centralizing TenantId/ClientId/etc.)
    if ($ConfigPath -eq "") {
        $defaultConfig = Join-Path $PSScriptRoot "M365AuthConfig.psd1"
        if (Test-Path $defaultConfig) { $ConfigPath = $defaultConfig }
    }

    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        try {
            $config = Import-PowerShellDataFile -Path $ConfigPath
        } catch {
            throw "Failed to load auth config from '$ConfigPath': $($_.Exception.Message)"
        }

        if (-not $TenantId) { $TenantId = $config.TenantId }
        if (-not $ClientId) { $ClientId = $config.ClientId }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = $config.CertificateThumbprint }
        if (-not $UserName) { $UserName = $config.UserName }
        if (-not $Password -and $config.Password) { $Password = ConvertTo-SecureString -AsPlainText $config.Password -Force }
        if (-not $AdminUrl) { $AdminUrl = $config.AdminUrl }
        if (-not $SiteUrl) { $SiteUrl = $config.SiteUrl }

        if (-not $GraphScopes -or $GraphScopes.Count -eq 0) {
            if ($config.GraphScopes) { $GraphScopes = $config.GraphScopes }
        }
    }

    # Default Graph scopes (if still missing)
    if (-not $GraphScopes -or $GraphScopes.Count -eq 0) {
        $GraphScopes = @(
            "User.Read.All",
            "Directory.Read.All",
            "AuditLog.Read.All",
            "Organization.Read.All"
        )
    }

    $IsCertAuth = ($TenantId -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
    $IsCredAuth = ($UserName -ne "" -and $null -ne $Password)

    $authMode = if ($IsCertAuth) { "Certificate" } elseif ($IsCredAuth) { "Credential" } else { "Interactive" }
    Write-Host "Authentication mode: $authMode" -ForegroundColor Cyan

    # ── Microsoft Graph ──────────────────────────────────────────────────────
    if ("Graph" -in $Services) {
        Confirm-M365Module -Name "Microsoft.Graph.Authentication"

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                Connect-MgGraph -TenantId $TenantId `
                                -ClientId $ClientId `
                                -CertificateThumbprint $CertificateThumbprint `
                                -NoWelcome -ErrorAction Stop
            } elseif ($IsCredAuth) {
                $Credential = New-M365Credential -UserName $UserName -Password $Password
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
        Confirm-M365Module -Name "ExchangeOnlineManagement"

        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                Connect-ExchangeOnline -AppId $ClientId `
                                       -CertificateThumbprint $CertificateThumbprint `
                                       -Organization $TenantId `
                                       -ShowBanner:$false -ErrorAction Stop
            } elseif ($IsCredAuth) {
                $Credential = New-M365Credential -UserName $UserName -Password $Password
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
        Confirm-M365Module -Name "Microsoft.Online.SharePoint.PowerShell"

        if ($AdminUrl -eq "") {
            throw "AdminUrl is required for SharePoint connections (e.g. https://contoso-admin.sharepoint.com)."
        }

        Write-Host "Connecting to SharePoint Online ($AdminUrl)..." -ForegroundColor Cyan

        try {
            if ($IsCertAuth) {
                $SpoCert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                if (-not $SpoCert) {
                    $SpoCert = Get-Item "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                }
                if (-not $SpoCert) {
                    throw "Certificate with thumbprint '$CertificateThumbprint' was not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My."
                }
                Connect-SPOService -Url $AdminUrl `
                                   -ClientId $ClientId `
                                   -Certificate $SpoCert `
                                   -Tenant $TenantId -ErrorAction Stop | Out-Null
            } elseif ($IsCredAuth) {
                $Credential = New-M365Credential -UserName $UserName -Password $Password
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
        Confirm-M365Module -Name "PnP.PowerShell" -MinimumVersion "1.12.0"

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
                $Credential = New-M365Credential -UserName $UserName -Password $Password
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

function Confirm-M365Module {
    param(
        [string]$Name,
        [string]$MinimumVersion = ""
    )
    # Microsoft.Online.SharePoint.PowerShell is .NET Framework-based and must exist in the
    # Windows PowerShell module path (not PS7 paths) for -UseWindowsPowerShell to load it.
    # Check and install via powershell.exe (WinPS 5.1) to ensure it lands in the right location.
    if ($Name -eq "Microsoft.Online.SharePoint.PowerShell") {
        $availableInWinPS = & powershell.exe -NonInteractive -Command `
            "if (Get-Module -ListAvailable -Name '$Name') { 'true' } else { 'false' }"
        if ($availableInWinPS -notmatch 'true') {
            Write-Host "$Name module is not available in Windows PowerShell." -ForegroundColor Yellow
            $confirm = Read-Host "Install $Name now? [Y] Yes [N] No"
            if ($confirm -notmatch "[yY]") {
                throw "$Name is required. Run in Windows PowerShell: Install-Module $Name -Scope CurrentUser"
            }
            Write-Host "Installing $Name into Windows PowerShell (CurrentUser scope)..." -ForegroundColor Cyan
            & powershell.exe -Command `
                "Install-Module -Name '$Name' -Repository PSGallery -AllowClobber -Force -Scope CurrentUser -Confirm:`$false"
            # Verify it landed in the WinPS path
            $verifyInstall = & powershell.exe -NonInteractive -Command `
                "if (Get-Module -ListAvailable -Name '$Name') { 'true' } else { 'false' }"
            if ($verifyInstall -notmatch 'true') {
                throw "Installation could not be verified. Try manually in Windows PowerShell: Install-Module $Name -Scope CurrentUser"
            }
        }
        Import-Module -Name $Name -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
        return
    }

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

    Import-Module -Name $Name -DisableNameChecking -ErrorAction Stop
}

function New-M365Credential {
    param([string]$UserName, [SecureString]$Password)
    return New-Object System.Management.Automation.PSCredential($UserName, $Password)
}

Export-ModuleMember -Function Connect-M365Services

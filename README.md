# App Registration Setup Guide

## Overview

All PowerShell scripts in this repository authenticate through a **single Azure AD App Registration**. This eliminates the need for per-script credentials or multiple registrations. Authentication is handled by `M365AuthModule.psm1`, which supports certificate-based (unattended/scheduled), credential-based, and interactive modes.

---

## Step 1: Create the App Registration

1. Sign in to the [Azure Portal](https://portal.azure.com) as a Global Administrator.
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**.
3. Fill in:
   - **Name**: `ClutchSolutions-M365Scripts` (or your preferred name)
   - **Supported account types**: Accounts in this organizational directory only (single tenant)
   - **Redirect URI**: Leave blank
4. Click **Register**.
5. Copy and save the **Application (client) ID** and **Directory (tenant) ID** — you will need both.

---

## Step 2: API Permissions

In your new App Registration, go to **API permissions** > **Add a permission**.

### Microsoft Graph (Application permissions)

| Permission | Purpose |
|---|---|
| `User.Read.All` | Read all user profiles |
| `Directory.Read.All` | Read directory data and roles |
| `AuditLog.Read.All` | Read sign-in and audit logs |
| `Organization.Read.All` | Read organization/subscription info |
| `Application.Read.All` | Read enterprise apps and service principals (GetEnterpriseAppsReport) |
| `Reports.Read.All` | Read Teams and other usage reports (TeamsActivitySummaryReport) |
| `Sites.Read.All` | Read SharePoint sites (PnP anonymous links) |
| `Sites.FullControl.All` | Have full control of all site collections |
| `Files.Read.All` | Read files in SharePoint/OneDrive (PnP anonymous links) |
| `ReportSettings.ReadWrite.All` | Read and write all admin report settings |

### Exchange Online (Application permissions)
### to Add this,   Search Under "APIs my organization uses": In the API Permissions blade, click "Add a permission" -> "APIs my organization uses" -> Type "Office 365 Exchange Online" --> thenm select Application Permissions --> and add below:
| Permission | Purpose |
|---|---|
| `Exchange.ManageAsApp` | Connect to Exchange Online as an application |



### SharePoint (Application permissions)

| Permission | Purpose |
|---|---|
| `Sites.FullControl.All` | Full control of all site collections (SPO external users, PnP) |

---

## Step 3: Grant Admin Consent

After adding all permissions:

1. In the **API permissions** blade, click **Grant admin consent for \<your tenant\>**.
2. Confirm by clicking **Yes**.
3. All permissions should show a green checkmark under **Status**.

---

## Step 4: Assign Exchange and Sharepont Administrator Role

Certificate-based authentication to Exchange Online requires the App Registration to hold the **Exchange Administrator** directory role.

1. Navigate to **Microsoft Entra** > **Roles and administrators**.
2. Search for and click **Exchange administrator**.
3. Click **Add assignments** > search for your App Registration name > **Add**.
4. Repeat for Sharepoint Administrator as well
---

## Step 5: Create and Install the Certificate Locally

> **Why local?** The certificate's **private key** must live in your Windows certificate store for PowerShell to sign authentication tokens. Certificates created inside the Azure Portal or Cloud Shell do not install the private key on your machine — only the public key ends up in Azure, leaving the scripts unable to authenticate.

**Prerequisites:** Open a **local** PowerShell session (not Cloud Shell) and `cd` to the repo root before pasting.

```powershell
cd "Path_To_\O365-Reporting"
```

Then paste this entire block at once:
```powershell
# From Here 
$repoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# 1 — Create the certificate and install the private key into the local store
$cert = New-SelfSignedCertificate `
    -Subject           "CN=ClutchSolutions-M365Scripts" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy   Exportable `
    -KeySpec           Signature `
    -NotAfter          (Get-Date).AddYears(2) `
    -HashAlgorithm     SHA256
Write-Host "Certificate created." -ForegroundColor Green
Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Yellow

# 2 — Export the public key (.cer) — upload this to Azure in Step 6
$cerPath = Join-Path $repoRoot "ClutchSolutions-M365Scripts.cer"
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Write-Host "Public key exported:  $cerPath" -ForegroundColor Cyan

# 3 — Export an encrypted private key backup (.pfx) — needed on any other machine
$pfxPath = Join-Path $repoRoot "ClutchSolutions-M365Scripts.pfx"
$pfxPassword = Read-Host "Enter a password to protect the .pfx backup" -AsSecureString
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword | Out-Null
Write-Host "Private key backup:   $pfxPath" -ForegroundColor Cyan

# 4 — Update M365AuthConfig.psd1 with the new thumbprint automatically
$configPath = Join-Path $repoRoot "M365AuthConfig.psd1"
if (Test-Path $configPath) {
    (Get-Content $configPath) `
        -replace 'CertificateThumbprint\s*=\s*".*"', "CertificateThumbprint = `"$($cert.Thumbprint)`"" |
        Set-Content $configPath
    Write-Host "M365AuthConfig.psd1 updated with new thumbprint." -ForegroundColor Green
} else {
    Write-Host "M365AuthConfig.psd1 not found at $configPath" -ForegroundColor Red
    Write-Host "Manually set: CertificateThumbprint = `"$($cert.Thumbprint)`"" -ForegroundColor Yellow
}

Write-Host "`nDone. Upload $cerPath to your App Registration (Step 6)." -ForegroundColor Green
# To Here

> **Keep the `.pfx` file secure.** Anyone with the file and password can authenticate as your App Registration. Do not commit it to source control.

---

## Step 6: Upload the Public Key to the App Registration

1. In the [Azure Portal](https://portal.azure.com), open your App Registration.
2. Go to **Certificates & secrets** > **Certificates** tab.
3. Click **Upload certificate**.
4. Browse to `ClutchSolutions-M365Scripts.cer` exported in Step 5.
5. Add an optional description (e.g. `ClutchSolutions local cert`) and click **Add**.
6. Confirm the thumbprint shown in Azure matches the one printed by the PowerShell block above.

> **If you need to run the scripts on a second machine**, copy the `.pfx` file to that machine and import it:
> ```powershell
> $pfxPassword = Read-Host "Enter the .pfx password" -AsSecureString
> Import-PfxCertificate -FilePath ".\ClutchSolutions-M365Scripts.pfx" `
>     -CertStoreLocation "Cert:\CurrentUser\My" -Password $pfxPassword
> ```
> No Azure changes are needed — the same thumbprint is already registered.

---

## Step 7: SharePoint App-Only Access (for SPO and PnP Scripts)

Scripts that use SharePoint Online or PnP Online also require SharePoint app-only access to be granted via PnP PowerShell. Run this once per tenant:

```powershell
# Install PnP if not already present
Install-Module PnP.PowerShell -Scope CurrentUser -Force
```

> **Note for SPOExternalUsersReport.ps1:** The legacy SharePoint Management Shell module must be installed into Windows PowerShell (not PS7). Run the following from an **elevated Windows PowerShell 5.1** prompt:
> ```powershell
> Install-Module Microsoft.Online.SharePoint.PowerShell -Scope AllUsers -Force
> ```

```powershell
# Register app-only access (opens a browser for consent)
Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName "ClutchSolutions-M365Scripts" `
    -Tenant "contoso.onmicrosoft.com" `
    -Interactive
```

Alternatively, the `Sites.FullControl.All` Graph application permission granted in Step 3 covers PnP certificate-based access when using `Connect-PnPOnline` with `-Thumbprint`.

---

## Step 8: Record Your Values

After completing the steps above, record these three values — they are passed as parameters to every script:

| Parameter | Where to find it |
|---|---|
| `$TenantId` | Azure AD > Overview > **Directory (tenant) ID** |
| `$ClientId` | App Registration > Overview > **Application (client) ID** |
| `$CertificateThumbprint` | Output of `New-SelfSignedCertificate` in Step 5, or the Certificates blade in Step 6 |

---

## Step 9: Running the Scripts

All scripts read authentication credentials automatically from `M365AuthConfig.psd1` in the repo root. No parameters are required for certificate-based (unattended) authentication once the config file is populated. All CSV output is saved to the `Exports\` folder at the repo root.

To override the config or run interactively, pass parameters explicitly. If all three cert parameters are omitted, the scripts fall back to interactive browser login.

All scripts must be run from within their own subfolder, or called by full path.

### Licencing\Licence_report.ps1

```powershell
.\Licencing\Licence_report.ps1
# Optional: -UserNamesFile "users.csv"
```

### Licencing\LicenseExpiryDateReport.ps1

```powershell
.\Licencing\LicenseExpiryDateReport.ps1
# Optional filters: -Trial -Free -Purchased -Expired -Active
```

### Shared_Mailboxs\FindLicensedSharedMailboxes.ps1

```powershell
.\Shared_Mailboxs\FindLicensedSharedMailboxes.ps1
```

### Shared_Mailboxs\NonCompliantSharedMailboxes.ps1

```powershell
.\Shared_Mailboxs\NonCompliantSharedMailboxes.ps1
```

### Sharepoint\AuditFileAccess.ps1

```powershell
.\Sharepoint\AuditFileAccess.ps1
# Optional: -StartDate "2026-01-01" -EndDate "2026-03-16" -SharePointOnlineOnly -OneDriveOnly
```

### Exchange_Online\BlockExternalEmailForwarding.ps1

```powershell
.\Exchange_Online\BlockExternalEmailForwarding.ps1
```

### Exchange_Online\DLsWithExternalUsers.ps1

```powershell
.\Exchange_Online\DLsWithExternalUsers.ps1
```

### Exchange_Online\GetMailboxFolderStatisticsReport.ps1

```powershell
.\Exchange_Online\GetMailboxFolderStatisticsReport.ps1
```

### Entra\O365UserLoginHistory.ps1

```powershell
.\Entra\O365UserLoginHistory.ps1
```

### Sharepoint\SPOExternalUsersReport.ps1

```powershell
.\Sharepoint\SPOExternalUsersReport.ps1 -HostName "contoso"
```

### Sharepoint\GetAllAnonymousLinks.ps1

```powershell
.\Sharepoint\GetAllAnonymousLinks.ps1 -TenantName "contoso"
```

### Ondrive\OneDriveUsageReport.ps1

```powershell
.\Ondrive\OneDriveUsageReport.ps1 -HostName "contoso"
```

### Entra\UserLastActivityTimeReport.ps1

```powershell
.\Entra\UserLastActivityTimeReport.ps1
```

### Entra\GetEnterpriseAppsReport.ps1

```powershell
.\Entra\GetEnterpriseAppsReport.ps1
```

### Entra\GetEntraUsersDetails (1).ps1

```powershell
& ".\Entra\GetEntraUsersDetails (1).ps1"
```

### Teams\TeamsActivitySummaryReport.ps1

```powershell
.\Teams\TeamsActivitySummaryReport.ps1 -Period D30
# Periods: D7, D30, D90, D180 | Or: -ReportDate "2026-03-01" | -InactiveDays 90
```

### FindUnusedM365Licenses\FindUnusedM365Licenses.ps1

```powershell
.\FindUnusedM365Licenses\FindUnusedM365Licenses.ps1 -InactiveDays 90
```

---

## Troubleshooting

**"Insufficient privileges to complete the operation"**
- Ensure admin consent was granted (Step 3).
- Verify the Exchange Administrator role is assigned to the App Registration (Step 4).

**"Certificate with thumbprint '...' was not found in certificate store or has expired"**
- The certificate private key is not installed on this machine. This happens when the cert was created in the Azure Portal or Cloud Shell instead of locally.
- Check your local store: `Get-ChildItem Cert:\CurrentUser\My | Select Thumbprint, Subject, NotAfter`
- If you have the `.pfx` backup from Step 5, import it: `Import-PfxCertificate -FilePath ".\ClutchSolutions-M365Scripts.pfx" -CertStoreLocation "Cert:\CurrentUser\My" -Password (Read-Host -AsSecureString)`
- If the private key is lost, repeat Steps 5 and 6. The Step 5 script updates `M365AuthConfig.psd1` automatically.

**"AADSTS700027: The certificate with identifier used to sign the client assertion is not registered on application"**
- The `.cer` public key in Azure does not match the thumbprint in `M365AuthConfig.psd1`. Re-upload the `.cer` exported in Step 5 to the App Registration Certificates blade.

**"Get-PnPFileSharingLink: Either scp or roles claim need to be present in the token"**
- Ensure `Sites.Read.All` and `Files.Read.All` Graph application permissions are granted and admin consent was given.

**"Connect-MgGraph: No account found in the token cache"**
- Run `Disconnect-MgGraph` then retry. If persistent, open a fresh PowerShell window to avoid module DLL conflicts.

**Module not found errors**
- The `M365AuthModule.psm1` will prompt to install missing modules automatically. Ensure you have internet access and that `PSGallery` is a trusted repository (`Set-PSRepository -Name PSGallery -InstallationPolicy Trusted`).

**Certificate expired**
- Repeat Steps 5 and 6 to generate and upload a new certificate. The PowerShell block in Step 5 updates `M365AuthConfig.psd1` automatically.

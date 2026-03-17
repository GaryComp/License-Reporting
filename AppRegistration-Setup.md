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
| `Sites.Read.All` | Read SharePoint sites (PnP anonymous links) |
| `Files.Read.All` | Read files in SharePoint/OneDrive (PnP anonymous links) |

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

## Step 4: Assign Exchange Administrator Role

Certificate-based authentication to Exchange Online requires the App Registration to hold the **Exchange Administrator** directory role.

1. Navigate to **Microsoft Entra** > **Roles and administrators**.
2. Search for and click **Exchange administrator**.
3. Click **Add assignments** > search for your App Registration name > **Add**.

---

## Step 5: Create a Self-Signed Certificate

Run the following in PowerShell to create a certificate valid for 2 years:

```powershell
$certParams = @{
    Subject           = "CN=ClutchSolutions-M365Scripts"
    CertStoreLocation = "Cert:\CurrentUser\My"
    KeyExportPolicy   = "Exportable"
    KeySpec           = "Signature"
    NotAfter          = (Get-Date).AddYears(2)
    HashAlgorithm     = "SHA256"
}
$cert = New-SelfSignedCertificate @certParams

# Export public key (.cer) for upload to Azure
Export-Certificate -Cert $cert -FilePath ".\ClutchSolutions-M365Scripts.cer"

# Note the thumbprint — you will use this as $CertificateThumbprint
Write-Host "Thumbprint: $($cert.Thumbprint)"
```

---

## Step 6: Upload Certificate to App Registration

1. In your App Registration, go to **Certificates & secrets** > **Certificates** tab.
2. Click **Upload certificate**.
3. Browse to the `.cer` file exported in Step 5.
4. Click **Add**.
5. The certificate thumbprint displayed in Azure should match the one from PowerShell.

---

## Step 7: SharePoint App-Only Access (for SPO and PnP Scripts)

Scripts that use SharePoint Online or PnP Online also require SharePoint app-only access to be granted via PnP PowerShell. Run this once per tenant:

```powershell
# Install PnP if not already present
Install-Module PnP.PowerShell -Scope CurrentUser -Force

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

All scripts accept the same three core parameters for certificate-based (unattended) authentication. If omitted, the scripts fall back to interactive login.

### Licence_report.ps1

```powershell
.\Licence_report.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### LicenseExpiryDateReport.ps1

```powershell
.\LicenseExpiryDateReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
# Optional filters: -Trial -Free -Purchased -Expired -Active
```

### FindLicensedSharedMailboxes.ps1

```powershell
.\FindLicensedSharedMailboxes.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### NonCompliantSharedMailboxes.ps1

```powershell
.\NonCompliantSharedMailboxes.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### AuditFileAccess.ps1

```powershell
.\AuditFileAccess.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
# Optional: -StartDate "2026-01-01" -EndDate "2026-03-16" -SharePointOnlineOnly -OneDriveOnly
```

### BlockExternalEmailForwarding.ps1

```powershell
.\BlockExternalEmailForwarding.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### DLsWithExternalUsers.ps1

```powershell
.\DLsWithExternalUsers.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### GetMailboxFolderStatisticsReport.ps1

```powershell
.\GetMailboxFolderStatisticsReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### O365UserLoginHistory.ps1

```powershell
.\O365UserLoginHistory.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### SPOExternalUsersReport.ps1

```powershell
.\SPOExternalUsersReport.ps1 -HostName "contoso" -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### GetAllAnonymousLinks.ps1

```powershell
.\GetAllAnonymousLinks.ps1 -TenantId "your-tenant-id" -TenantName "contoso" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

### Office365UserLastActivityTime\UserLastActivityTimeReport.ps1

```powershell
.\Office365UserLastActivityTime\UserLastActivityTimeReport.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "THUMBPRINT"
```

---

## Troubleshooting

**"Insufficient privileges to complete the operation"**
- Ensure admin consent was granted (Step 3).
- Verify the Exchange Administrator role is assigned to the App Registration (Step 4).

**"AADSTS700027: The certificate with identifier used to sign the client assertion is not registered on application"**
- The certificate thumbprint does not match the certificate uploaded in Azure. Re-upload the `.cer` file or regenerate the certificate.

**"Get-PnPFileSharingLink: Either scp or roles claim need to be present in the token"**
- Ensure `Sites.Read.All` and `Files.Read.All` Graph application permissions are granted and admin consent was given.

**"Connect-MgGraph: No account found in the token cache"**
- Run `Disconnect-MgGraph` then retry. If persistent, open a fresh PowerShell window to avoid module DLL conflicts.

**Module not found errors**
- The `M365AuthModule.psm1` will prompt to install missing modules automatically. Ensure you have internet access and that `PSGallery` is a trusted repository (`Set-PSRepository -Name PSGallery -InstallationPolicy Trusted`).

**Certificate expired**
- Repeat Steps 5 and 6 to generate and upload a new certificate, then update `$CertificateThumbprint` in your scheduled tasks or scripts.

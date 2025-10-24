# Quick Reference Guide

This quick reference provides essential commands and information for SharePoint Embedded ISV setup.

## Essential Information to Collect

### From ISV (Your Organization)
```
Application (Client) ID: _______________
Directory (Tenant) ID:   _______________
Certificate Thumbprint:  _______________
Certificate Path:        _______________
Certificate Password:    _______________
```

### From Customer (Consuming Tenant)
```
Tenant ID/Domain:        _______________
Admin Email:             _______________
SharePoint Admin URL:    _______________
```

## Quick Setup Commands

### 1. Install Prerequisites

```powershell
# Install required PowerShell modules
Install-Module -Name PnP.PowerShell -Force -AllowClobber
Install-Module -Name Microsoft.Graph -Force -AllowClobber
```

### 2. Generate Certificate (ISV - One Time)

```powershell
# Generate self-signed certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=SPEmbeddedApp" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export certificate
$certPassword = ConvertTo-SecureString -String "YOUR_SECURE_PASSWORD_HERE" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath ".\SPEmbeddedApp.pfx" -Password $certPassword
Export-Certificate -Cert $cert -FilePath ".\SPEmbeddedApp.cer"

Write-Host "Thumbprint: $($cert.Thumbprint)"
```

### 3. Generate Admin Consent URL (ISV)

```powershell
# Your application details
$clientId = "your-app-id"
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

# Generate consent URL
$consentUrl = "https://login.microsoftonline.com/organizations/v2.0/adminconsent" +
    "?client_id=$clientId" +
    "&redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUri) +
    "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/.default")

Write-Host $consentUrl
```

### 4. Register Owning App (After Admin Consent)

```powershell
.\scripts\Register-SPEOwningApp.ps1 `
    -OwningAppId "your-app-id" `
    -CertificatePath ".\certificate.pfx" `
    -CertificatePassword "your-password" `
    -ConsumingTenantId "customer.onmicrosoft.com" `
    -ConsumingTenantAdminUrl "https://customer-admin.sharepoint.com" `
    -ContainerTypeDisplayName "My Container Type"
```

### 5. Create Container

```powershell
.\scripts\New-SPEContainer.ps1 `
    -ContainerTypeId "your-container-type-id" `
    -DisplayName "Project Files" `
    -OwningAppId "your-app-id" `
    -CertificatePath ".\certificate.pfx" `
    -CertificatePassword "your-password" `
    -ConsumingTenantId "customer.onmicrosoft.com"
```

## Common Verification Commands

### Check if App Exists in Tenant

```powershell
Connect-MgGraph -Scopes "Application.Read.All"
$sp = Get-MgServicePrincipal -Filter "appId eq 'your-app-id'"
$sp | Format-List DisplayName, AppId, Id
Disconnect-MgGraph
```

### List Container Types

```powershell
Connect-PnPOnline -Url "https://tenant-admin.sharepoint.com" -Interactive
Get-PnPContainerType | Format-Table ContainerTypeId, DisplayName, OwningAppId
Disconnect-PnPOnline
```

### List Containers

```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com" -Interactive
Get-PnPContainer | Format-Table ContainerId, DisplayName, ContainerTypeId
Disconnect-PnPOnline
```

### Check Container Permissions

```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com" -Interactive
Get-PnPContainerPermission -ContainerId "container-id"
Disconnect-PnPOnline
```

## API Permissions Required

### Microsoft Graph
- `Sites.FullControl.All` (Application)
- `Files.ReadWrite.All` (Application)
- `Files.Read.All` (Application) - Optional

### SharePoint
- `Sites.FullControl.All` (Application)
- `TermStore.ReadWrite.All` (Application)

## Troubleshooting Quick Checks

### Check Certificate

```powershell
# List certificates
Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Subject -like "*SPEmbedded*"}

# Check certificate expiry
$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq "your-thumbprint"}
Write-Host "Expires: $($cert.NotAfter)"
Write-Host "Valid: $($cert.NotAfter -gt (Get-Date))"
```

### Test Connection

```powershell
# Test SharePoint connection
Connect-PnPOnline -Url "https://tenant.sharepoint.com" `
    -ClientId "your-app-id" `
    -Tenant "tenant.onmicrosoft.com" `
    -Thumbprint "your-thumbprint"

$context = Get-PnPContext
Write-Host "Connected to: $($context.Url)"
Disconnect-PnPOnline
```

### Check Admin Consent

```powershell
Connect-MgGraph -Scopes "Application.Read.All"
$sp = Get-MgServicePrincipal -Filter "appId eq 'your-app-id'"

if ($sp) {
    Write-Host "✓ App found - consent granted" -ForegroundColor Green
    $oauth2 = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $sp.Id
    $appRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id
    Write-Host "OAuth2 Permissions: $($oauth2.Count)"
    Write-Host "App Permissions: $($appRoles.Count)"
} else {
    Write-Host "✗ App not found - consent not granted" -ForegroundColor Red
}
Disconnect-MgGraph
```

## Common Error Codes

| Error Code | What It Means | Quick Fix |
|------------|---------------|-----------|
| AADSTS65001 | Consent not granted | Run admin consent process |
| AADSTS70001 | App not found | Check client ID is correct |
| AADSTS700016 | App not in directory | Grant admin consent in customer tenant |
| AADSTS700027 | Invalid certificate | Verify certificate and thumbprint match |

## Useful URLs

### Azure Portal Pages
- App Registrations: `https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps`
- Enterprise Applications: `https://portal.azure.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/~/AppAppsPreview`

### SharePoint Admin Centers
- `https://{tenant}-admin.sharepoint.com`
- `https://{tenant}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx`

### Microsoft 365 Admin
- Service Health: `https://admin.microsoft.com/Adminportal/Home#/servicehealth`
- App Permissions: `https://admin.microsoft.com/Adminportal/Home#/Settings/IntegratedApps`

## Script Parameters Reference

### Register-SPEOwningApp.ps1

| Parameter | Required | Description |
|-----------|----------|-------------|
| OwningAppId | Yes | Application (Client) ID |
| CertificatePath | Yes* | Path to PFX file |
| CertificatePassword | Yes* | PFX password |
| CertificateThumbprint | Yes* | Cert thumbprint (alternative to path) |
| ConsumingTenantId | Yes | Customer tenant ID or domain |
| ConsumingTenantAdminUrl | Yes | SharePoint admin URL |
| ContainerTypeDisplayName | No | Display name for container type |
| ContainerTypeDescription | No | Description for container type |

*Either CertificatePath + Password OR CertificateThumbprint required

### New-SPEContainer.ps1

| Parameter | Required | Description |
|-----------|----------|-------------|
| ContainerTypeId | Yes | Container type ID from registration |
| DisplayName | Yes | Container display name |
| OwningAppId | Yes | Application (Client) ID |
| CertificatePath | Yes* | Path to PFX file |
| CertificatePassword | Yes* | PFX password |
| CertificateThumbprint | Yes* | Cert thumbprint (alternative) |
| ConsumingTenantId | Yes | Customer tenant ID or domain |
| Description | No | Container description |
| SiteUrl | No | SharePoint site URL |
| SetPermissions | No | Switch to set initial permissions |
| Owners | No | Comma-separated owner UPNs |
| Members | No | Comma-separated member UPNs |
| Readers | No | Comma-separated reader UPNs |

*Either CertificatePath + Password OR CertificateThumbprint required

## Step-by-Step Checklist

### ISV Setup (One Time)
- [ ] Create Azure App Registration
- [ ] Add required API permissions
- [ ] Grant admin consent in your tenant
- [ ] Generate certificate
- [ ] Upload certificate to Azure
- [ ] Record Application ID and Certificate Thumbprint
- [ ] Generate admin consent URL

### Per Customer Setup
- [ ] Share admin consent URL with customer admin
- [ ] Verify customer admin granted consent
- [ ] Run Register-SPEOwningApp.ps1
- [ ] Save Container Type ID
- [ ] Create initial container(s) using New-SPEContainer.ps1
- [ ] Test container access from application
- [ ] Provide container information to customer

## Best Practices

1. **Certificate Management**
   - Use certificates, not client secrets
   - Set 2-year expiration
   - Implement rotation before expiry
   - Store PFX files securely
   - Use Azure Key Vault for production

2. **Permissions**
   - Request minimum permissions needed
   - Document why each permission is required
   - Review permissions annually
   - Remove unused permissions

3. **Security**
   - Never commit certificates to source control
   - Use strong passwords for PFX files
   - Rotate credentials regularly
   - Monitor application access logs
   - Implement least privilege access

4. **Documentation**
   - Maintain list of customer deployments
   - Document container type IDs per customer
   - Track certificate expiration dates
   - Keep customer contact information updated

5. **Testing**
   - Test in a dev tenant first
   - Verify all operations before customer deployment
   - Have rollback plan
   - Document known issues
   - Test certificate renewal process

## Support Resources

- **Documentation**: See `docs/` folder for detailed guides
- **Scripts**: See `scripts/` folder for PowerShell automation
- **Microsoft Docs**: https://learn.microsoft.com/sharepoint/dev/embedded/
- **Community**: https://techcommunity.microsoft.com/

## Next Steps

1. Review the full documentation in `docs/azure-app-registration.md`
2. Follow the admin consent guide in `docs/admin-consent.md`
3. Use the troubleshooting guide when issues occur: `docs/troubleshooting.md`
4. Keep this quick reference handy for day-to-day operations

# Troubleshooting Guide for SharePoint Embedded ISV Setup

This guide provides solutions to common issues encountered when setting up SharePoint Embedded applications in customer tenants.

## Table of Contents

1. [Azure App Registration Issues](#azure-app-registration-issues)
2. [Certificate Authentication Issues](#certificate-authentication-issues)
3. [Admin Consent Issues](#admin-consent-issues)
4. [Owning App Registration Issues](#owning-app-registration-issues)
5. [Container Creation Issues](#container-creation-issues)
6. [Permission Issues](#permission-issues)
7. [PowerShell Script Issues](#powershell-script-issues)
8. [General Debugging Tips](#general-debugging-tips)

## Azure App Registration Issues

### Issue: Can't create app registration

**Error**: "Insufficient privileges to complete the operation"

**Cause**: Your account lacks the required permissions.

**Solution**:
```
- Request Global Administrator or Application Administrator role
- Or request Cloud Application Administrator role
- Contact your Azure AD administrator if you don't have these roles
```

### Issue: Multi-tenant option not available

**Cause**: Your organization's settings may restrict multi-tenant app creation.

**Solution**:
1. Check Azure AD settings: **Enterprise applications** > **User settings**
2. Verify "Users can consent to apps accessing company data on their behalf" is enabled
3. Contact your Azure AD administrator to modify tenant settings

### Issue: Can't add API permissions

**Error**: "Permission not found" or permissions list is empty

**Solution**:
- Refresh the Azure Portal page
- Try a different browser or clear cache
- Ensure the API (Microsoft Graph, SharePoint) is available in your tenant
- Wait a few minutes and try again (sometimes there's a replication delay)

## Certificate Authentication Issues

### Issue: Certificate upload fails

**Error**: "Invalid certificate format"

**Cause**: Uploading the wrong file format or file is corrupted.

**Solution**:
```powershell
# Verify you're uploading the .cer (public key) file, not .pfx
# Re-export the certificate if needed

# PowerShell - Re-export public key
$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq "your-thumbprint"}
Export-Certificate -Cert $cert -FilePath ".\certificate.cer"

# OpenSSL - Re-export public key
openssl x509 -in certificate.crt -out certificate.cer -outform DER
```

### Issue: Certificate authentication fails

**Error**: "AADSTS700027: Client assertion contains an invalid signature"

**Cause**: 
- Certificate mismatch between uploaded cert and the one being used
- Certificate has expired
- Wrong certificate thumbprint

**Solution**:
```powershell
# Verify certificate details
$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Subject -like "*SPEmbedded*"}

Write-Host "Certificate Details:"
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host "Subject: $($cert.Subject)"
Write-Host "Expiry: $($cert.NotAfter)"
Write-Host "Valid: $($cert.NotAfter -gt (Get-Date))"

# Check if thumbprint matches what's in Azure
# Azure Portal > App Registration > Certificates & secrets
```

### Issue: Can't find certificate in store

**Error**: "Cannot find certificate with thumbprint"

**Cause**: Certificate not in the expected certificate store.

**Solution**:
```powershell
# Search all certificate stores
Get-ChildItem -Path Cert:\CurrentUser\My
Get-ChildItem -Path Cert:\LocalMachine\My

# If certificate is in PFX file, import it
$certPassword = ConvertTo-SecureString -String "your-password" -Force -AsPlainText
Import-PfxCertificate -FilePath ".\certificate.pfx" -CertStoreLocation Cert:\CurrentUser\My -Password $certPassword
```

### Issue: Certificate expired

**Error**: Certificate validation failed

**Solution**:
```powershell
# Generate new certificate and upload to Azure
$cert = New-SelfSignedCertificate `
    -Subject "CN=SPEmbeddedApp" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export and upload the new certificate
Export-Certificate -Cert $cert -FilePath ".\NewCert.cer"

# Then upload to Azure Portal and update your scripts with new thumbprint
```

## Admin Consent Issues

### Issue: Admin consent URL doesn't work

**Error**: "AADSTS500011: The resource principal named https://graph.microsoft.com was not found"

**Cause**: URL is malformed or parameters are incorrect.

**Solution**:
```powershell
# Regenerate URL with proper encoding
$clientId = "your-app-id"
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

$consentUrl = "https://login.microsoftonline.com/organizations/v2.0/adminconsent" +
    "?client_id=$clientId" +
    "&redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUri) +
    "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/.default")

Write-Host $consentUrl
```

### Issue: "Need admin approval" appears after consent

**Cause**: User-level consent is showing instead of admin consent.

**Solution**:
- Ensure the admin uses the admin consent URL, not a regular sign-in URL
- The URL must include `/adminconsent` endpoint
- Admin must have Global Administrator role

### Issue: Consent shows but doesn't persist

**Cause**: Browser issues or session problems.

**Solution**:
- Clear browser cache and cookies
- Try in an incognito/private browsing window
- Try a different browser
- Ensure the admin completes the entire flow without closing the window

## Owning App Registration Issues

### Issue: Owning app registration fails

**Error**: "The specified container type does not exist"

**Cause**: Container type ID is incorrect or not created yet.

**Solution**:
```powershell
# Verify container type exists
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive

# List all container types
$containerTypes = Get-PnPContainerType
$containerTypes | Format-Table ContainerTypeId, DisplayName, OwningAppId

# If not found, create the container type first
```

### Issue: "Application not found in tenant"

**Cause**: Admin consent wasn't granted in the consuming tenant.

**Solution**:
1. Verify admin consent was completed
2. Check in Azure AD > Enterprise Applications
3. Re-run the admin consent process if needed
4. Wait 5-10 minutes for replication

### Issue: Permission denied during registration

**Error**: "Access denied. You do not have permission to perform this action"

**Cause**: Insufficient permissions in the consuming tenant.

**Solution**:
- Ensure you have SharePoint Administrator role or higher
- Check if the customer has restricted SharePoint Embedded features
- Verify the owning app has necessary permissions in the consuming tenant

## Container Creation Issues

### Issue: Container creation fails

**Error**: "The container type is not registered in the tenant"

**Cause**: Owning app not registered in the consuming tenant.

**Solution**:
```powershell
# Verify owning app registration
Connect-PnPOnline -Url "https://tenant-admin.sharepoint.com" -Interactive

# Check if container type is registered
$containerType = Get-PnPContainerType -ContainerTypeId "your-container-type-id"

if ($containerType) {
    Write-Host "Container type registered" -ForegroundColor Green
} else {
    Write-Host "Container type NOT registered. Run Register-SPEOwningApp.ps1 first" -ForegroundColor Red
}
```

### Issue: "Unauthorized" when creating container

**Cause**: Authentication token doesn't have required permissions.

**Solution**:
```powershell
# Verify you're authenticated with the owning app identity
Connect-PnPOnline -Url "https://tenant.sharepoint.com" `
    -ClientId "your-app-id" `
    -CertificatePath ".\certificate.pfx" `
    -CertificatePassword (ConvertTo-SecureString "password" -AsPlainText -Force) `
    -Tenant "tenant.onmicrosoft.com"

# Test connection
$context = Get-PnPContext
Write-Host "Connected as: $($context.Url)"
```

### Issue: Container appears but is not accessible

**Cause**: Permissions not properly configured on the container.

**Solution**:
```powershell
# Set container permissions explicitly
Set-PnPContainerPermission `
    -ContainerId "container-id" `
    -UserPrincipalName "user@domain.com" `
    -Role "Owner"
```

## Permission Issues

### Issue: "Insufficient privileges" when calling Graph API

**Cause**: Required permissions not granted or not effective yet.

**Solution**:
```powershell
# 1. Verify permissions in Azure Portal
# App Registration > API Permissions > Check all have green checkmark

# 2. Re-grant admin consent if needed
# Click "Grant admin consent for [Organization]"

# 3. Wait 5-10 minutes for permissions to propagate

# 4. Test with Microsoft Graph Explorer
# https://developer.microsoft.com/graph/graph-explorer
```

### Issue: "Access denied" to SharePoint sites

**Cause**: SharePoint permissions not granted or not effective.

**Solution**:
```powershell
# Grant SharePoint admin consent
$tenantUrl = "https://tenant-admin.sharepoint.com"
Connect-PnPOnline -Url $tenantUrl -Interactive

# Verify app has necessary permissions
# Azure Portal > SharePoint API Permissions > Sites.FullControl.All should be granted
```

### Issue: "Sites.Selected" permission issues

**Cause**: Sites.Selected requires explicit site-level permissions.

**Solution**:
```powershell
# Grant permission to specific site
$siteUrl = "https://tenant.sharepoint.com/sites/sitename"
$appId = "your-app-id"

Connect-PnPOnline -Url $siteUrl -Interactive

# Grant full control to the site
Grant-PnPAzureADAppSitePermission `
    -AppId $appId `
    -DisplayName "Your App Name" `
    -Permissions FullControl
```

## PowerShell Script Issues

### Issue: "Module not found" error

**Error**: "The specified module 'PnP.PowerShell' was not loaded"

**Solution**:
```powershell
# Install required modules
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser

# Verify installation
Get-Module -ListAvailable -Name PnP.PowerShell
Get-Module -ListAvailable -Name Microsoft.Graph

# Import modules
Import-Module PnP.PowerShell
Import-Module Microsoft.Graph
```

### Issue: Execution policy prevents script running

**Error**: "File cannot be loaded because running scripts is disabled"

**Solution**:
```powershell
# Check current execution policy
Get-ExecutionPolicy

# Set execution policy (run as Administrator)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Or bypass for current session only
PowerShell.exe -ExecutionPolicy Bypass -File .\script.ps1
```

### Issue: Script fails with certificate password error

**Error**: "The specified network password is not correct"

**Cause**: Certificate password is incorrect or not provided.

**Solution**:
```powershell
# Create secure password correctly
$certPassword = ConvertTo-SecureString -String "YourActualPassword" -Force -AsPlainText

# Or prompt for password
$certPassword = Read-Host -AsSecureString -Prompt "Enter certificate password"

# Test certificate can be loaded
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import(".\certificate.pfx", $certPassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet)
```

### Issue: "Connect-PnPOnline" fails

**Error**: Various connection errors

**Solution**:
```powershell
# Try interactive login first to verify connectivity
Connect-PnPOnline -Url "https://tenant.sharepoint.com" -Interactive

# If that works, then try certificate-based auth
Connect-PnPOnline -Url "https://tenant.sharepoint.com" `
    -ClientId "your-app-id" `
    -Tenant "tenant.onmicrosoft.com" `
    -CertificatePath ".\certificate.pfx" `
    -CertificatePassword (ConvertTo-SecureString "password" -AsPlainText -Force)

# Enable verbose logging to see details
Connect-PnPOnline -Url "https://tenant.sharepoint.com" -Interactive -Verbose
```

## General Debugging Tips

### Enable Verbose Logging

```powershell
# PowerShell verbose output
$VerbosePreference = "Continue"

# Or run commands with -Verbose flag
Connect-PnPOnline -Url "https://tenant.sharepoint.com" -Interactive -Verbose
```

### Capture Detailed Errors

```powershell
# Capture full error details
try {
    # Your command here
    Connect-PnPOnline -Url "https://tenant.sharepoint.com" -Interactive
} catch {
    Write-Host "Error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host "Stack Trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace
    Write-Host "Error Details:" -ForegroundColor Yellow
    Write-Host ($_ | Format-List -Force | Out-String)
}
```

### Check Service Health

```powershell
# Check Microsoft 365 service health
# Visit: https://admin.microsoft.com/Adminportal/Home#/servicehealth

# Or use PowerShell
Connect-MgGraph -Scopes "ServiceHealth.Read.All"
Get-MgServiceAnnouncementIssue | Where-Object {$_.Service -eq "SharePoint"}
```

### Verify Tenant Configuration

```powershell
# Check SharePoint Online tenant settings
Connect-PnPOnline -Url "https://tenant-admin.sharepoint.com" -Interactive

# Get tenant properties
$tenant = Get-PnPTenant
$tenant | Format-List

# Check if SharePoint Embedded is enabled
# (There may not be a specific setting, but this shows general tenant config)
```

### Test with Microsoft Graph Explorer

For permission and API issues:
1. Go to [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
2. Sign in with the admin account
3. Try the same API calls your script is making
4. Compare results and permissions

### Use Fiddler or Network Tracing

For authentication issues:
1. Install Fiddler or use browser DevTools
2. Capture network traffic during authentication
3. Look for failed requests and error responses
4. Check token claims and permissions

### Common Error Codes

| Error Code | Meaning | Solution |
|------------|---------|----------|
| AADSTS50001 | Resource not found | Check resource URL and app registration |
| AADSTS65001 | Consent not granted | Run admin consent process |
| AADSTS70001 | Application not found | Verify client ID is correct |
| AADSTS700016 | Application not in directory | App not in target tenant, need consent |
| AADSTS700027 | Invalid certificate signature | Check certificate and thumbprint |
| AADSTS90002 | Tenant not found | Verify tenant ID is correct |

## Getting Help

If you're still experiencing issues:

1. **Check Microsoft Documentation**:
   - [SharePoint Embedded Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview)
   - [Azure AD Error Codes](https://learn.microsoft.com/en-us/azure/active-directory/develop/reference-aadsts-error-codes)
   - [Microsoft Graph Errors](https://learn.microsoft.com/en-us/graph/errors)

2. **Community Support**:
   - [Microsoft Tech Community](https://techcommunity.microsoft.com/)
   - [Stack Overflow](https://stackoverflow.com/questions/tagged/sharepoint-embedded)
   - [SharePoint Dev Community](https://aka.ms/sppnp)

3. **Microsoft Support**:
   - Open a support ticket through Azure Portal
   - For SharePoint issues: Microsoft 365 Admin Center > Support

4. **Enable Logging**:
   - Enable diagnostic logging in your scripts
   - Collect error messages and stack traces
   - Note the exact steps that lead to the error

## Quick Checklist

When debugging, verify:

- [ ] All required PowerShell modules are installed and up to date
- [ ] App registration exists and is configured as multi-tenant
- [ ] All required API permissions are added and granted
- [ ] Admin consent was granted in the consuming tenant
- [ ] Certificate is valid and not expired
- [ ] Certificate thumbprint matches what's in Azure
- [ ] Service principal exists in the consuming tenant
- [ ] User has appropriate admin roles (Global Admin, SharePoint Admin, etc.)
- [ ] Tenant allows the required operations
- [ ] No service outages or known issues
- [ ] URLs and IDs are correct (no typos)
- [ ] Sufficient wait time for replication (5-10 minutes after changes)

## Additional Resources

- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [Microsoft Graph PowerShell Documentation](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [Azure AD App Registration Troubleshooting](https://learn.microsoft.com/en-us/azure/active-directory/develop/howto-troubleshoot-app-registration)
- [SharePoint Embedded Known Issues](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/known-issues)

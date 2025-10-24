# Complete Example Walkthrough

This document provides a complete, real-world example of setting up a SharePoint Embedded application for a customer tenant. Follow along to understand the entire process from start to finish.

## Scenario

**ISV Company**: Contoso Software Inc.  
**Application**: Contoso Document Manager  
**Customer**: Fabrikam Corporation  

**Goal**: Set up Contoso Document Manager in Fabrikam's tenant so their users can leverage SharePoint Embedded containers for document management.

## Prerequisites Completed

Before starting this walkthrough, ensure:
- ✅ PowerShell 7.x installed
- ✅ PnP.PowerShell module installed
- ✅ Microsoft.Graph module installed
- ✅ Azure subscription with appropriate permissions
- ✅ Access to both ISV and customer tenants

## Phase 1: ISV Setup (One-Time Configuration)

### Step 1.1: Create Azure App Registration

**Location**: Contoso's Azure Portal  
**Performed by**: Contoso IT Administrator

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations** > **New registration**

**Configuration**:
```
Name: Contoso Document Manager - SPE
Supported account types: Accounts in any organizational directory (Multitenant)
Redirect URI: https://login.microsoftonline.com/common/oauth2/nativeclient
```

**Result**:
```
Application (Client) ID: 12345678-90ab-cdef-1234-567890abcdef
Directory (Tenant) ID:   abcdef12-3456-7890-abcd-ef1234567890
```

### Step 1.2: Add API Permissions

Still in Azure Portal, add the following permissions:

**Microsoft Graph - Application Permissions**:
- ✅ Sites.FullControl.All
- ✅ Files.ReadWrite.All
- ✅ Files.Read.All

**SharePoint - Application Permissions**:
- ✅ Sites.FullControl.All
- ✅ TermStore.ReadWrite.All

**Grant admin consent for Contoso**:
- Click "Grant admin consent for Contoso Software"
- Confirm the action

### Step 1.3: Generate Certificate

Open PowerShell as Administrator:

```powershell
# Generate certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=ContosoDocumentManager" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export PFX (with private key)
$certPassword = ConvertTo-SecureString -String "C0nt0s0Secure!2024" -Force -AsPlainText
Export-PfxCertificate `
    -Cert $cert `
    -FilePath "C:\Contoso\ContosoDocMgr.pfx" `
    -Password $certPassword

# Export CER (public key for Azure)
Export-Certificate `
    -Cert $cert `
    -FilePath "C:\Contoso\ContosoDocMgr.cer"

Write-Host "Certificate Details:"
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host "Expires: $($cert.NotAfter)"
```

**Result**:
```
Thumbprint: A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0
Expires: 12/25/2026 10:30:00 AM
```

### Step 1.4: Upload Certificate to Azure

1. In Azure Portal, go to your app registration
2. Select **Certificates & secrets** > **Certificates** tab
3. Click **Upload certificate**
4. Select `ContosoDocMgr.cer`
5. Add description: "Production Certificate - Expires Dec 2026"
6. Click **Add**

Verify the thumbprint matches: `A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0`

### Step 1.5: Record Important Information

Create a secure document with this information:

```
=== Contoso Document Manager - Azure App Details ===

Application (Client) ID: 12345678-90ab-cdef-1234-567890abcdef
Directory (Tenant) ID:   abcdef12-3456-7890-abcd-ef1234567890
Object ID:               fedcba09-8765-4321-fedc-ba0987654321

Certificate Information:
  Thumbprint:            A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0
  PFX Path:              C:\Contoso\ContosoDocMgr.pfx
  Password:              [Stored in Azure Key Vault]
  Expires:               December 25, 2026

API Permissions:
  ✓ Sites.FullControl.All (Microsoft Graph)
  ✓ Files.ReadWrite.All (Microsoft Graph)
  ✓ Sites.FullControl.All (SharePoint)
  ✓ TermStore.ReadWrite.All (SharePoint)
  ✓ Admin Consent Granted
```

## Phase 2: Customer Onboarding (Per Customer)

### Step 2.1: Generate Admin Consent URL

**Performed by**: Contoso Support Team

```powershell
# Generate admin consent URL for Fabrikam
$clientId = "12345678-90ab-cdef-1234-567890abcdef"
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
$state = [guid]::NewGuid().ToString()

$consentUrl = "https://login.microsoftonline.com/organizations/v2.0/adminconsent" +
    "?client_id=$clientId" +
    "&redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUri) +
    "&state=$state" +
    "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/.default")

Write-Host "Admin Consent URL for Fabrikam:"
Write-Host $consentUrl
Write-Host "`nState Value (save for verification): $state"
```

**Result**:
```
https://login.microsoftonline.com/organizations/v2.0/adminconsent?client_id=12345678-90ab-cdef-1234-567890abcdef&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient&state=a1b2c3d4-e5f6-7890-abcd-ef1234567890&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default

State Value: a1b2c3d4-e5f6-7890-abcd-ef1234567890
```

### Step 2.2: Send Email to Customer

**From**: Contoso Support (support@contoso.com)  
**To**: Fabrikam IT Admin (admin@fabrikam.com)  
**Subject**: Admin Consent Required - Contoso Document Manager Setup

```
Dear Fabrikam IT Team,

Thank you for choosing Contoso Document Manager! To complete your setup, we need 
admin consent from your Microsoft 365 tenant.

What This Does:
- Authorizes Contoso Document Manager to create and manage SharePoint Embedded 
  containers in your tenant
- Grants secure access to document storage and management APIs
- Enables your users to leverage advanced document collaboration features

Who Can Grant Consent:
- Global Administrator
- Privileged Role Administrator

Please follow these steps:

1. Click this admin consent URL (or have your Global Administrator click it):
   [Insert the generated consent URL here]

2. Sign in with your admin account (@fabrikam.com)

3. Review the requested permissions:
   - Full control of all site collections
   - Read and write files
   - Manage term store

4. Click "Accept" to grant consent

5. Reply to this email once complete

Security Information:
- Application Name: Contoso Document Manager - SPE
- Publisher: Contoso Software Inc.
- Application ID: 12345678-90ab-cdef-1234-567890abcdef
- Privacy Policy: https://contoso.com/privacy
- Terms of Service: https://contoso.com/terms

Questions? Contact our support team at support@contoso.com

Best regards,
Contoso Support Team
```

### Step 2.3: Customer Grants Admin Consent

**Performed by**: Fabrikam Global Administrator

1. Admin receives email from Contoso
2. Admin clicks the consent URL
3. Admin signs in with admin@fabrikam.com
4. Reviews consent screen showing:
   - Application: Contoso Document Manager - SPE
   - Publisher: Contoso Software Inc.
   - Permissions requested (list shown)
5. Clicks "Accept"
6. Gets redirected to success page
7. Replies to Contoso email: "Admin consent granted"

**Verification** (Optional - performed by Fabrikam Admin):

```powershell
Connect-MgGraph -Scopes "Application.Read.All"

$appId = "12345678-90ab-cdef-1234-567890abcdef"
$sp = Get-MgServicePrincipal -Filter "appId eq '$appId'"

if ($sp) {
    Write-Host "✓ Contoso Document Manager found in tenant" -ForegroundColor Green
    Write-Host "Display Name: $($sp.DisplayName)"
    Write-Host "App ID: $($sp.AppId)"
} else {
    Write-Host "✗ Application not found - consent may have failed" -ForegroundColor Red
}

Disconnect-MgGraph
```

### Step 2.4: Register Owning App in Fabrikam Tenant

**Performed by**: Contoso Support Team (with Fabrikam's information)

**Fabrikam Information Provided**:
```
Tenant Domain: fabrikam.onmicrosoft.com
Tenant ID: 98765432-1fed-cba9-8765-432198765432
SharePoint Admin URL: https://fabrikam-admin.sharepoint.com
```

**Run Registration Script**:

```powershell
# Navigate to scripts directory
cd C:\Contoso\SPE-Setup-ISV\scripts

# Run registration script
.\Register-SPEOwningApp.ps1 `
    -OwningAppId "12345678-90ab-cdef-1234-567890abcdef" `
    -CertificatePath "C:\Contoso\ContosoDocMgr.pfx" `
    -CertificatePassword "C0nt0s0Secure!2024" `
    -ConsumingTenantId "fabrikam.onmicrosoft.com" `
    -ConsumingTenantAdminUrl "https://fabrikam-admin.sharepoint.com" `
    -ContainerTypeDisplayName "Contoso Document Manager Containers"
```

**Script Output**:
```
======================================================================
  SharePoint Embedded Owning App Registration
======================================================================

[Step 1/4] Checking prerequisites...
✓ Module 'PnP.PowerShell' is installed
✓ Module 'Microsoft.Graph.Authentication' is installed

[Step 2/4] Verifying owning app...
Connecting to Microsoft Graph...
✓ Owning app found in tenant: Contoso Document Manager - SPE
  App ID: 12345678-90ab-cdef-1234-567890abcdef
  Object ID: fedcba09-8765-4321-fedc-ba0987654321
✓ Owning app verification complete

[Step 3/4] Connecting to SharePoint...
Connecting to SharePoint: https://fabrikam-admin.sharepoint.com
✓ Successfully connected to SharePoint

[Step 4/4] Registering container type...
Creating new container type...
✓ Container type registered successfully!
  Container Type ID: b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
  Display Name: Contoso Document Manager Containers
  Owning App ID: 12345678-90ab-cdef-1234-567890abcdef

======================================================================
  REGISTRATION SUMMARY
======================================================================

Owning App Information:
  Application ID:    12345678-90ab-cdef-1234-567890abcdef
  Consuming Tenant:  fabrikam.onmicrosoft.com

Container Type Information:
  Container Type ID: b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
  Display Name:      Contoso Document Manager Containers
  Description:       SharePoint Embedded container type
  Owning App ID:     12345678-90ab-cdef-1234-567890abcdef

Next Steps:
  1. Save the Container Type ID for creating containers
  2. Use the New-SPEContainer.ps1 script to create containers
  3. Test container creation and access

======================================================================

✓ Registration completed successfully!
```

**Record Container Type ID**:
```
Container Type ID for Fabrikam: b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
```

## Phase 3: Container Creation and Testing

### Step 3.1: Create First Container for Fabrikam

**Performed by**: Contoso Support Team or Fabrikam Admin

```powershell
# Navigate to scripts directory
cd C:\Contoso\SPE-Setup-ISV\scripts

# Create container for Fabrikam's HR department
.\New-SPEContainer.ps1 `
    -ContainerTypeId "b!ISV.default|12345678-90ab-cdef-1234-567890abcdef" `
    -DisplayName "Fabrikam HR Documents" `
    -Description "Human Resources department document repository" `
    -OwningAppId "12345678-90ab-cdef-1234-567890abcdef" `
    -CertificatePath "C:\Contoso\ContosoDocMgr.pfx" `
    -CertificatePassword "C0nt0s0Secure!2024" `
    -ConsumingTenantId "fabrikam.onmicrosoft.com" `
    -SetPermissions `
    -Owners "hr-admin@fabrikam.com" `
    -Members "hr-team@fabrikam.com"
```

**Script Output**:
```
======================================================================
  SharePoint Embedded Container Creation
======================================================================

[Step 1/4] Checking prerequisites...
✓ Module 'PnP.PowerShell' is installed

[Step 2/4] Connecting to SharePoint...
Connecting to SharePoint: https://fabrikam.sharepoint.com
✓ Successfully connected to SharePoint

[Step 3/4] Verifying container type...
Verifying container type exists...
✓ Container type found
  Container Type ID: b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
  Display Name: Contoso Document Manager Containers
  Owning App ID: 12345678-90ab-cdef-1234-567890abcdef

[Step 4/4] Creating container...

Creating SharePoint Embedded container...
  Display Name: Fabrikam HR Documents
✓ Container created successfully!
  Container ID: b!HzFmW9X2qEqAZV9nEXL8_1t1J_O3fqZLgY_qvZ5kD1E2F3A4B5C
  Display Name: Fabrikam HR Documents
  Container Type ID: b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
  Created: 2024-10-24T12:45:30Z

Setting container permissions...
  ✓ Added owner: hr-admin@fabrikam.com
  ✓ Added member: hr-team@fabrikam.com
✓ Permissions configured

======================================================================
  CONTAINER CREATED SUCCESSFULLY
======================================================================

Container Information:
  Container ID:       b!HzFmW9X2qEqAZV9nEXL8_1t1J_O3fqZLgY_qvZ5kD1E2F3A4B5C
  Display Name:       Fabrikam HR Documents
  Description:        Human Resources department document repository
  Container Type ID:  b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
  Created:            2024-10-24T12:45:30Z

Using the Container:
  1. Use the Container ID in your application
  2. Access via Microsoft Graph API:
     GET https://graph.microsoft.com/v1.0/storage/fileStorage/containers/b!HzFmW9X2qEqAZV9nEXL8_1t1J_O3fqZLgY_qvZ5kD1E2F3A4B5C
  3. Manage permissions using Set-PnPContainerPermission cmdlet

Next Steps:
  - Upload files to the container using your application
  - Set additional permissions as needed
  - Test accessing the container from your application

======================================================================

✓ Container creation completed successfully!
```

### Step 3.2: Create Additional Containers

Create containers for other Fabrikam departments:

```powershell
# Finance Department
.\New-SPEContainer.ps1 `
    -ContainerTypeId "b!ISV.default|12345678-90ab-cdef-1234-567890abcdef" `
    -DisplayName "Fabrikam Finance Documents" `
    -Description "Finance department document repository" `
    -OwningAppId "12345678-90ab-cdef-1234-567890abcdef" `
    -CertificateThumbprint "A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0" `
    -ConsumingTenantId "fabrikam.onmicrosoft.com" `
    -SetPermissions `
    -Owners "finance-admin@fabrikam.com"

# Engineering Department
.\New-SPEContainer.ps1 `
    -ContainerTypeId "b!ISV.default|12345678-90ab-cdef-1234-567890abcdef" `
    -DisplayName "Fabrikam Engineering Documents" `
    -Description "Engineering department document repository" `
    -OwningAppId "12345678-90ab-cdef-1234-567890abcdef" `
    -CertificateThumbprint "A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0" `
    -ConsumingTenantId "fabrikam.onmicrosoft.com" `
    -SetPermissions `
    -Owners "eng-admin@fabrikam.com" `
    -Members "engineers@fabrikam.com"
```

### Step 3.3: Test Container Access

Verify containers can be accessed via Microsoft Graph:

```powershell
# Connect to Microsoft Graph as the application
Connect-MgGraph `
    -ClientId "12345678-90ab-cdef-1234-567890abcdef" `
    -TenantId "fabrikam.onmicrosoft.com" `
    -CertificateThumbprint "A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0"

# List all containers
$containerId = "b!HzFmW9X2qEqAZV9nEXL8_1t1J_O3fqZLgY_qvZ5kD1E2F3A4B5C"

# Get container details
$container = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/$containerId"

Write-Host "Container Details:"
Write-Host "  ID: $($container.id)"
Write-Host "  Display Name: $($container.displayName)"
Write-Host "  Description: $($container.description)"
Write-Host "  Created: $($container.createdDateTime)"

# Test file upload (from application)
# This would be done through your Contoso Document Manager application
# using the container ID

Disconnect-MgGraph
```

## Phase 4: Deployment Summary

### Customer Information Record

Create a deployment record for Fabrikam:

```
=== Fabrikam Corporation - Deployment Summary ===

Customer Information:
  Organization:          Fabrikam Corporation
  Tenant Domain:         fabrikam.onmicrosoft.com
  Tenant ID:             98765432-1fed-cba9-8765-432198765432
  SharePoint Admin URL:  https://fabrikam-admin.sharepoint.com
  Contact Email:         admin@fabrikam.com
  Contact Name:          Alex Johnson
  Deployment Date:       October 24, 2024

Container Type:
  Container Type ID:     b!ISV.default|12345678-90ab-cdef-1234-567890abcdef
  Display Name:          Contoso Document Manager Containers

Containers Created:
  1. HR Documents
     Container ID: b!HzFmW9X2qEqAZV9nEXL8_1t1J_O3fqZLgY_qvZ5kD1E2F3A4B5C
     Owners: hr-admin@fabrikam.com
     
  2. Finance Documents
     Container ID: b!GhIjKlMnOpQrStUvWxYz0123456789AbCdEfGhIjKlMnOpQr
     Owners: finance-admin@fabrikam.com
     
  3. Engineering Documents
     Container ID: b!StUvWxYz0123456789AbCdEfGhIjKlMnOpQrStUvWxYz0123
     Owners: eng-admin@fabrikam.com

Status: ✓ Active
Admin Consent: ✓ Granted
App Registered: ✓ Yes
Containers: 3
Last Updated: October 24, 2024
```

### Handoff to Customer

Send completion email:

```
Dear Fabrikam Team,

Your Contoso Document Manager setup is complete! Here's what was configured:

✓ Admin consent granted
✓ Owning application registered
✓ Container type created
✓ 3 department containers created and configured

Container IDs for your applications:
- HR Documents: b!HzFmW9X2qEqAZV9nEXL8_1t1J_O3fqZLgY_qvZ5kD1E2F3A4B5C
- Finance Documents: b!GhIjKlMnOpQrStUvWxYz0123456789AbCdEfGhIjKlMnOpQr
- Engineering Documents: b!StUvWxYz0123456789AbCdEfGhIjKlMnOpQrStUvWxYz0123

Next Steps:
1. Your users can now log into Contoso Document Manager
2. Department admins have been granted ownership of respective containers
3. Users can start uploading and managing documents

Support:
- Documentation: https://docs.contoso.com
- Support Email: support@contoso.com
- Support Portal: https://support.contoso.com

Thank you for choosing Contoso Document Manager!

Best regards,
Contoso Support Team
```

## Success Criteria

✅ Azure App Registration created in ISV tenant  
✅ API permissions configured and admin consent granted  
✅ Certificate generated and uploaded  
✅ Admin consent granted by customer  
✅ Owning app registered in customer tenant  
✅ Container type created  
✅ Containers created and accessible  
✅ Permissions configured correctly  
✅ Testing completed successfully  
✅ Documentation provided to customer  

## Ongoing Maintenance

### Monthly Tasks
- Monitor certificate expiration dates
- Review application usage logs
- Check for permission changes or revocations
- Verify container access

### Quarterly Tasks
- Review and optimize permissions
- Update documentation
- Check for SharePoint Embedded updates
- Customer satisfaction check

### Annual Tasks
- Rotate certificates before expiration
- Conduct security audit
- Review all customer deployments
- Update processes based on learnings

## Conclusion

This completes the full walkthrough for setting up SharePoint Embedded for a customer. The process can be repeated for each new customer with their specific tenant information.

**Key Takeaways**:
1. ISV setup is done once and reused for all customers
2. Customer onboarding requires admin consent first
3. Scripts automate the registration and container creation
4. Proper documentation ensures smooth deployments
5. Maintain records of all customer deployments

For questions or issues, refer to the troubleshooting guide at `docs/troubleshooting.md`.

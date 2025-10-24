# Admin Consent Process for SharePoint Embedded

This guide explains how to obtain admin consent from a consuming tenant administrator for your SharePoint Embedded application.

## Table of Contents

1. [Overview](#overview)
2. [Generate Admin Consent URL](#generate-admin-consent-url)
3. [Share with Customer Admin](#share-with-customer-admin)
4. [Admin Consent Steps (Customer)](#admin-consent-steps-customer)
5. [Verify Consent](#verify-consent)
6. [Troubleshooting](#troubleshooting)

## Overview

Admin consent is a critical step that allows your multi-tenant application to access resources in a customer's tenant. When a customer tenant administrator grants consent:

- Your application is authorized to access the customer's SharePoint and Microsoft Graph resources
- The permissions you requested during app registration are granted
- Users in the customer tenant can use your application without individual consent prompts

**Important**: Admin consent must be granted by someone with one of these roles in the customer tenant:
- Global Administrator
- Privileged Role Administrator
- Cloud Application Administrator (for some permissions)

## Generate Admin Consent URL

### Step 1: Construct the Admin Consent URL

The admin consent URL follows this format:

```
https://login.microsoftonline.com/{tenant}/v2.0/adminconsent
  ?client_id={client_id}
  &redirect_uri={redirect_uri}
  &state={state}
  &scope={scope}
```

### Step 2: Fill in the Parameters

Replace the placeholders with your application's values:

- `{tenant}`: Use `organizations` for multi-tenant apps, or the specific tenant ID
- `{client_id}`: Your Application (client) ID from Azure App Registration
- `{redirect_uri}`: A redirect URI registered in your app (URL-encoded)
- `{state}`: (Optional) A value to maintain state between request and callback
- `{scope}`: (Optional) Space-separated list of permissions, or use `.default`

### Example Admin Consent URL

```
https://login.microsoftonline.com/organizations/v2.0/adminconsent?client_id=12345678-1234-1234-1234-123456789abc&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient&state=12345&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default
```

### Step 3: Generate Using PowerShell

You can also generate the URL using PowerShell:

```powershell
# Your application details
$clientId = "your-app-id-here"
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
$state = [guid]::NewGuid().ToString()

# Construct the URL
$consentUrl = "https://login.microsoftonline.com/organizations/v2.0/adminconsent" +
    "?client_id=$clientId" +
    "&redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUri) +
    "&state=$state" +
    "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/.default")

Write-Host "Admin Consent URL:" -ForegroundColor Green
Write-Host $consentUrl
Write-Host ""
Write-Host "Share this URL with the customer tenant administrator." -ForegroundColor Yellow
Write-Host "State value (for verification): $state" -ForegroundColor Cyan
```

## Share with Customer Admin

### Step 1: Prepare Communication

When sharing the admin consent URL with your customer, provide clear instructions and context:

**Sample Email Template**:

```
Subject: Admin Consent Required for [Your Application Name]

Dear [Customer Name],

To complete the setup of [Your Application Name] in your Microsoft 365 tenant, 
we need admin consent from a Global Administrator in your organization.

What this does:
- Authorizes our application to access SharePoint and Microsoft Graph resources in your tenant
- Enables our application to create and manage SharePoint Embedded containers
- Grants the following permissions: [list key permissions]

Who can grant consent:
- Global Administrator
- Privileged Role Administrator

Please follow these steps:
1. Click the admin consent URL below (or have your Global Administrator click it)
2. Sign in with your admin account
3. Review the requested permissions
4. Click "Accept" to grant consent

Admin Consent URL:
[Insert your generated consent URL here]

What to expect:
- You'll see a consent screen showing our application name and requested permissions
- After accepting, you'll be redirected to a confirmation page
- The entire process takes less than 2 minutes

Security Notes:
- Review all requested permissions before accepting
- Ensure the application name matches: [Your Application Name]
- The publisher should show: [Your Organization Name]
- You can revoke consent at any time from your Azure AD portal

If you have any questions or concerns about the requested permissions, please don't 
hesitate to contact us.

Best regards,
[Your Name]
[Your Company]
```

### Step 2: Provide Supporting Documentation

Include links to:
- This documentation
- Your privacy policy
- Your terms of service
- SharePoint Embedded documentation
- List of requested permissions and why they're needed

## Admin Consent Steps (Customer)

These are the steps the customer administrator will follow:

### Step 1: Access the Consent URL

1. The administrator receives the admin consent URL from you (the ISV)
2. They click the link or copy it into a browser
3. They are redirected to Microsoft's login page

### Step 2: Sign In

1. The administrator signs in with their Global Administrator account
2. Multi-factor authentication may be required (if enabled)
3. They are redirected to the consent screen

### Step 3: Review Permissions

The consent screen displays:

- **Application Name**: Your application's name from Azure AD
- **Publisher**: Your verified publisher domain (if verified)
- **Permissions Requested**: List of all API permissions
- **Description**: What each permission allows the app to do

**Example permissions displayed**:
```
This app would like to:
✓ Read and write files in all site collections
✓ Have full control of all site collections
✓ Read and write managed metadata
✓ Read all users' full profiles
```

### Step 4: Grant Consent

1. The administrator reviews all requested permissions carefully
2. If acceptable, they click **Accept**
3. If not acceptable, they can click **Cancel** and contact you for clarification

### Step 5: Confirmation

After accepting:
1. The administrator is redirected to the redirect URI
2. The URL will include the `state` parameter (if provided) and a `tenant` parameter
3. They may see a success message or be redirected to your application

Example redirect:
```
https://login.microsoftonline.com/common/oauth2/nativeclient
  ?tenant=customer-tenant-id
  &state=12345
  &admin_consent=True
```

## Verify Consent

### Method 1: Check in Customer's Azure Portal

The customer administrator can verify consent was granted:

1. Sign in to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **Enterprise applications**
3. Search for your application name
4. Select your application
5. Click on **Permissions**
6. Verify all permissions show "Granted for [Organization]"

### Method 2: Using PowerShell

The customer can verify using PowerShell:

```powershell
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All"

# Find the service principal for your app
$clientId = "your-app-id"
$sp = Get-MgServicePrincipal -Filter "appId eq '$clientId'"

if ($sp) {
    Write-Host "Application found in tenant: $($sp.DisplayName)" -ForegroundColor Green
    Write-Host "App ID: $($sp.AppId)"
    Write-Host "Object ID: $($sp.Id)"
    
    # Get granted permissions
    $oauth2Permissions = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $sp.Id
    
    if ($oauth2Permissions) {
        Write-Host "OAuth2 Permissions Granted: Yes" -ForegroundColor Green
    }
    
    # Get app role assignments
    $appRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id
    
    if ($appRoles) {
        Write-Host "Application Permissions Granted: Yes" -ForegroundColor Green
        Write-Host "Number of permissions: $($appRoles.Count)"
    }
} else {
    Write-Host "Application not found. Consent may not have been granted." -ForegroundColor Red
}

Disconnect-MgGraph
```

### Method 3: Test API Access

After consent is granted, you can test by attempting to acquire a token:

```powershell
# This script tests if consent was granted by attempting to get an access token

$tenantId = "customer-tenant-id"
$clientId = "your-app-id"
$certThumbprint = "your-cert-thumbprint"

try {
    # Connect using certificate
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certThumbprint
    
    Write-Host "Successfully connected! Admin consent is working." -ForegroundColor Green
    
    # Test a simple Graph call
    $context = Get-MgContext
    Write-Host "Connected as: $($context.AppName)"
    Write-Host "Tenant: $($context.TenantId)"
    
    Disconnect-MgGraph
} catch {
    Write-Host "Failed to connect. Possible issues:" -ForegroundColor Red
    Write-Host "- Admin consent not granted"
    Write-Host "- Certificate not valid"
    Write-Host "- Incorrect tenant ID"
    Write-Host "Error: $($_.Exception.Message)"
}
```

## Troubleshooting

### Issue: "AADSTS65001: The user or administrator has not consented"

**Cause**: Admin consent has not been granted yet.

**Solution**: 
- Ensure the admin clicked the consent URL
- Verify the admin completed the consent process
- Check if consent was accidentally denied

### Issue: "AADSTS50020: User account from identity provider does not exist in tenant"

**Cause**: The user signing in doesn't belong to the target tenant.

**Solution**:
- Ensure the administrator signs in with an account from THEIR tenant
- Don't use your (ISV) credentials
- Use an account with Global Administrator role in the customer tenant

### Issue: "AADSTS700016: Application not found in the directory"

**Cause**: The application doesn't exist or the client ID is incorrect.

**Solution**:
- Verify the client ID in the consent URL
- Ensure the app registration exists in your Azure AD
- Check that the app is configured as multi-tenant

### Issue: Consent screen shows "Unverified" publisher

**Cause**: Your app registration doesn't have a verified publisher domain.

**Solution**:
- This is a warning but doesn't prevent consent
- Consider verifying your publisher domain in Azure AD for production apps
- Provide additional documentation to reassure customers

### Issue: Some permissions not showing as granted

**Cause**: Certain permissions may require additional admin roles or may not be available.

**Solution**:
- Verify the admin has sufficient privileges
- Some permissions require Global Administrator specifically
- Check if the customer's license includes the required services

### Issue: Redirect fails after consent

**Cause**: The redirect URI may not be registered in your app.

**Solution**:
- Ensure the redirect URI in the consent URL matches one registered in your app
- URL-encode the redirect URI properly
- Use `https://login.microsoftonline.com/common/oauth2/nativeclient` as a safe default

## Consent Revocation

Customers can revoke consent at any time:

1. Sign in to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **Enterprise applications**
3. Find and select your application
4. Click **Permissions**
5. Click **Revoke permissions** or remove specific grants

After revocation:
- Your application will no longer be able to access the tenant's resources
- Users will be prompted to consent again
- Existing containers and data remain but become inaccessible

## Security Best Practices

### For ISVs:
- Request only the minimum permissions needed
- Clearly document why each permission is required
- Implement proper error handling for consent failures
- Use certificate-based authentication, not client secrets
- Regularly rotate certificates

### For Customer Admins:
- Review all requested permissions carefully
- Verify the publisher's identity
- Check the redirect URI is legitimate
- Document which applications have been granted consent
- Regularly audit enterprise applications
- Revoke consent for unused applications

## Next Steps

After admin consent is granted:

1. **Verify consent** using one of the methods above
2. **Register the owning app** in the consuming tenant using the `Register-SPEOwningApp.ps1` script
3. **Create containers** using the `New-SPEContainer.ps1` script
4. **Test application access** to ensure everything works

## Additional Resources

- [Admin Consent Workflow](https://learn.microsoft.com/en-us/azure/active-directory/manage-apps/configure-admin-consent-workflow)
- [Application Consent Experience](https://learn.microsoft.com/en-us/azure/active-directory/develop/application-consent-experience)
- [Permissions and Consent](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent)
- [Publisher Verification](https://learn.microsoft.com/en-us/azure/active-directory/develop/publisher-verification-overview)

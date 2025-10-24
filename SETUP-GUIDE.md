# SharePoint Embedded Token Generator - Setup Guide

## Overview

This guide provides complete instructions for ISVs to set up certificate-based authentication and generate tokens for registering SharePoint Embedded owning applications in consuming (customer) tenants.

## What This Solution Provides

The `New-SharePointEmbeddedToken.ps1` script automates the following tasks:

1. **Certificate Management**
   - Creates new self-signed certificates
   - Uses existing certificates from the certificate store
   - Exports certificates for distribution and backup

2. **Token Generation**
   - Builds JWT (JSON Web Token) assertions using certificate signing
   - Requests OAuth 2.0 access tokens from Azure AD
   - Returns tokens ready for use in API calls

3. **User Guidance**
   - Comprehensive help documentation
   - Clear error messages and troubleshooting tips
   - Next-step instructions after token generation

## Prerequisites Checklist

Before using this script, ensure you have:

- [ ] PowerShell 5.1 or later installed
- [ ] Administrator rights on your machine (for certificate creation)
- [ ] Azure AD tenant access (both owning and consuming tenants)
- [ ] An Azure AD app registration for your SharePoint Embedded application
- [ ] Internet connectivity to reach Azure AD endpoints

## Complete Setup Process

### Phase 1: Owning Tenant Setup (ISV Environment)

#### Step 1: Create Azure AD App Registration

1. Sign in to [Azure Portal](https://portal.azure.com) with your owning tenant credentials
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter details:
   - **Name**: Your SharePoint Embedded App Name
   - **Supported account types**: Multitenant (for ISV scenarios)
5. Click **Register**
6. Copy the **Application (client) ID** - you'll need this later

#### Step 2: Generate Certificate

Run the script to create a certificate:

```powershell
.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "your-owning-tenant.onmicrosoft.com" `
    -ClientId "your-app-client-id" `
    -ExportCertificate `
    -CertificatePath "C:\SPE-Certs"
```

This creates:
- A certificate in your certificate store
- A `.pfx` file (certificate with private key)
- A `.cer` file (public certificate only)

#### Step 3: Upload Certificate to Azure AD

1. In your app registration, go to **Certificates & secrets**
2. Click **Upload certificate**
3. Select the `.cer` file created in Step 2
4. Click **Add**
5. Note the certificate thumbprint displayed

#### Step 4: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions**
5. Add the following permissions:
   - `Sites.FullControl.All`
   - `Files.ReadWrite.All`
   - (Add other permissions as required by your app)
6. Click **Grant admin consent** (requires admin privileges)

### Phase 2: Consuming Tenant Registration (Customer Environment)

#### Step 5: Share Required Information with Customer

Provide the customer with:
- Your Application (Client) ID
- Instructions for granting consent
- Required API permissions list
- This setup guide

#### Step 6: Generate Token for Customer Tenant

Once the customer has granted consent, generate a token:

```powershell
.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "customer-tenant.onmicrosoft.com" `
    -ClientId "your-app-client-id" `
    -CertificateThumbprint "your-certificate-thumbprint"
```

#### Step 7: Use Token to Register SharePoint Embedded Container

Use the generated token in your API calls:

```powershell
# Capture the token
$token = .\New-SharePointEmbeddedToken.ps1 `
    -TenantId "customer-tenant.onmicrosoft.com" `
    -ClientId "your-app-client-id" `
    -CertificateThumbprint "your-certificate-thumbprint"

# Prepare headers
$headers = @{
    "Authorization" = "Bearer $($token.access_token)"
    "Content-Type" = "application/json"
}

# Make API calls to register your container
# Example: Register container type
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/..." `
    -Headers $headers -Method Post -Body $jsonBody
```

## Common Scenarios

### Scenario 1: Initial Development and Testing

```powershell
# Create certificate and generate token for your dev tenant
.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "dev-tenant.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -ExportCertificate

# Save the thumbprint for future use
```

### Scenario 2: Customer Onboarding

```powershell
# Generate token for new customer
.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "newcustomer.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateThumbprint "ABC123..."
```

### Scenario 3: Automated Deployment Pipeline

```powershell
# In your CI/CD pipeline
$token = .\New-SharePointEmbeddedToken.ps1 `
    -TenantId $env:CUSTOMER_TENANT_ID `
    -ClientId $env:APP_CLIENT_ID `
    -CertificateThumbprint $env:CERT_THUMBPRINT

# Use $token.access_token in subsequent deployment steps
```

## Security Considerations

### Certificate Protection

1. **Private Keys**: Never share `.pfx` files or private keys
2. **Storage**: Store certificates in:
   - Windows Certificate Store (recommended for production)
   - Azure Key Vault (best for cloud deployments)
   - Hardware Security Modules (HSMs) for highest security

3. **Access Control**: 
   - Limit who can access certificate files
   - Use ACLs to protect certificate store access
   - Audit certificate access regularly

### Token Handling

1. **Transmission**: Always use HTTPS for API calls
2. **Storage**: 
   - Never log tokens in plain text
   - Don't commit tokens to source control
   - Clear tokens from memory after use
3. **Expiration**: Tokens typically expire after 3600 seconds (1 hour)
4. **Refresh**: Generate new tokens when needed; don't reuse expired tokens

### Certificate Lifecycle

1. **Rotation**: Rotate certificates every 12-24 months
2. **Monitoring**: Set alerts for approaching expiration dates
3. **Revocation**: Have a process to revoke compromised certificates
4. **Backup**: Maintain secure backups of certificates

## Troubleshooting Guide

### Error: "Certificate not found"

**Cause**: Certificate isn't in the certificate store or thumbprint is incorrect

**Solutions**:
- Verify thumbprint has no extra spaces or special characters
- Check both Current User and Local Machine stores
- Recreate the certificate if necessary

### Error: "Failed to get access token"

**Cause**: Authentication failed at Azure AD

**Solutions**:
- Ensure certificate is uploaded to Azure AD app registration
- Verify Client ID and Tenant ID are correct
- Check that API permissions are granted with admin consent
- Verify certificate hasn't expired

### Error: "Access Denied"

**Cause**: Insufficient permissions

**Solutions**:
- Grant admin consent for required API permissions
- Verify app is registered in the target tenant
- Check that app registration supports multitenant scenarios

### Error: "Certificate does not have private key"

**Cause**: Certificate in store doesn't include private key

**Solutions**:
- Re-import certificate with private key
- Use the `.pfx` file with password when importing
- Recreate certificate using the script with `-ExportCertificate` flag

## Best Practices

### For Development

1. Use separate certificates for development and production
2. Test token generation in development environment first
3. Implement error handling in your integration code
4. Log non-sensitive information for troubleshooting

### For Production

1. Use production-grade certificate storage (Azure Key Vault, HSM)
2. Implement token caching to reduce API calls
3. Set up monitoring and alerting for token failures
4. Document your certificate rotation process
5. Maintain audit logs of token generation

### For Multi-Tenant Scenarios

1. Maintain configuration per customer tenant
2. Automate token generation for deployments
3. Implement tenant-specific error handling
4. Document customer-specific configurations

## Additional Resources

- [Microsoft Identity Platform documentation](https://docs.microsoft.com/azure/active-directory/develop/)
- [SharePoint Embedded documentation](https://docs.microsoft.com/sharepoint/dev/embedded/)
- [OAuth 2.0 Client Credentials Flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)
- [Certificate-based authentication](https://docs.microsoft.com/azure/active-directory/develop/active-directory-certificate-credentials)

## Support

For issues or questions:
1. Check the troubleshooting guide above
2. Review the example usage file (`Example-Usage.ps1`)
3. Consult Microsoft documentation
4. Contact your SharePoint Embedded support team

## Version History

- **v1.0** - Initial release with certificate management and token generation

---

**Note**: This script and documentation are provided as-is for ISVs implementing SharePoint Embedded solutions. Always follow your organization's security policies and Microsoft's best practices when handling certificates and tokens.

# Azure App Registration for SharePoint Embedded

This guide provides detailed step-by-step instructions for creating and configuring an Azure App Registration that an ISV can use to register their SharePoint Embedded application on a consuming tenant.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Create the App Registration](#create-the-app-registration)
3. [Configure API Permissions](#configure-api-permissions)
4. [Create and Configure Certificate Authentication](#create-and-configure-certificate-authentication)
5. [Configure Application Settings](#configure-application-settings)
6. [Record Important Values](#record-important-values)

## Prerequisites

Before you begin, ensure you have:

- Access to Azure Portal (https://portal.azure.com)
- Global Administrator or Application Administrator role in Azure AD
- OpenSSL or PowerShell for certificate generation
- Understanding of OAuth 2.0 and Azure AD concepts

## Create the App Registration

### Step 1: Navigate to Azure Portal

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** (or **Microsoft Entra ID**)
3. Select **App registrations** from the left navigation menu
4. Click **+ New registration**

### Step 2: Configure Basic Settings

1. **Name**: Enter a descriptive name for your application
   - Example: `Contoso SharePoint Embedded App`
   - This name will be visible to users when they grant consent

2. **Supported account types**: Select the appropriate option
   - **Recommended**: "Accounts in any organizational directory (Any Azure AD directory - Multitenant)"
   - This allows your application to work across multiple customer tenants

3. **Redirect URI**: (Optional at this stage, can be configured later)
   - Platform: Web
   - URI: `https://login.microsoftonline.com/common/oauth2/nativeclient` (or your application's callback URL)

4. Click **Register**

### Step 3: Note the Application Details

After registration, you'll be taken to the app's overview page. Record these important values:

- **Application (client) ID**: This is your `OwningAppId`
- **Directory (tenant) ID**: Your owning tenant ID
- **Object ID**: The application's object ID in Azure AD

## Configure API Permissions

SharePoint Embedded applications require specific Microsoft Graph and SharePoint permissions.

### Step 1: Add Microsoft Graph Permissions

1. From your app registration page, select **API permissions** from the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** (not Delegated)
5. Add the following permissions:

   **Required Permissions**:
   - `Sites.FullControl.All` - Full control of all site collections
   - `Files.ReadWrite.All` - Read and write files in all site collections
   - `Files.Read.All` - Read files in all site collections (if read-only access is needed)

   **Additional Recommended Permissions**:
   - `Sites.Selected` - Access to selected site collections (for more granular control)
   - `User.Read.All` - Read all users' full profiles

6. Click **Add permissions**

### Step 2: Add SharePoint Permissions

1. Click **+ Add a permission** again
2. Select **SharePoint**
3. Choose **Application permissions**
4. Add the following permissions:

   **Required Permissions**:
   - `Sites.FullControl.All` - Have full control of all site collections
   - `TermStore.ReadWrite.All` - Read and write managed metadata

5. Click **Add permissions**

### Step 3: Grant Admin Consent (for your owning tenant)

⚠️ **Important**: This step grants consent for your app in YOUR tenant. Customer tenants will need to grant consent separately.

1. After adding all permissions, click **Grant admin consent for [Your Organization]**
2. Confirm by clicking **Yes**
3. Verify that all permissions show a green checkmark under the "Status" column

## Create and Configure Certificate Authentication

Certificate-based authentication is more secure than client secrets and is recommended for production applications.

### Step 1: Generate a Certificate

#### Option A: Using PowerShell (Windows)

```powershell
# Generate a self-signed certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=SPEmbeddedApp" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export the certificate with private key (PFX)
$certPassword = ConvertTo-SecureString -String "YourStrongPassword123!" -Force -AsPlainText
Export-PfxCertificate `
    -Cert $cert `
    -FilePath ".\SPEmbeddedApp.pfx" `
    -Password $certPassword

# Export the public key (CER) for uploading to Azure
Export-Certificate `
    -Cert $cert `
    -FilePath ".\SPEmbeddedApp.cer"

Write-Host "Certificate Thumbprint: $($cert.Thumbprint)"
```

#### Option B: Using OpenSSL (Cross-platform)

```bash
# Generate private key
openssl genrsa -out SPEmbeddedApp.key 2048

# Generate certificate signing request
openssl req -new -key SPEmbeddedApp.key -out SPEmbeddedApp.csr

# Generate self-signed certificate (valid for 2 years)
openssl x509 -req -days 730 -in SPEmbeddedApp.csr -signkey SPEmbeddedApp.key -out SPEmbeddedApp.crt

# Convert to PFX format
openssl pkcs12 -export -out SPEmbeddedApp.pfx -inkey SPEmbeddedApp.key -in SPEmbeddedApp.crt

# Extract public key for Azure upload
openssl x509 -in SPEmbeddedApp.crt -out SPEmbeddedApp.cer -outform DER
```

⚠️ **Security Best Practices**:
- Store the PFX file securely (it contains the private key)
- Use a strong password for the PFX file
- Consider using Azure Key Vault for production certificates
- Set appropriate expiration dates and implement certificate rotation

### Step 2: Upload Certificate to Azure App Registration

1. In your app registration, select **Certificates & secrets** from the left menu
2. Click the **Certificates** tab
3. Click **Upload certificate**
4. Select the `.cer` file (public key) you generated
5. Add a description (e.g., "Production Certificate - Expires 2025")
6. Click **Add**
7. Note the certificate **Thumbprint** - you'll need this for authentication

### Step 3: (Optional) Add Client Secret as Fallback

While certificates are recommended, you may want to add a client secret for testing:

1. Click the **Client secrets** tab
2. Click **+ New client secret**
3. Add a description (e.g., "Development Secret")
4. Select an expiration period (maximum 24 months)
5. Click **Add**
6. **Important**: Copy the secret value immediately - it won't be shown again!

## Configure Application Settings

### Step 1: Configure Authentication Settings

1. Select **Authentication** from the left menu
2. Under **Platform configurations**, verify or add your redirect URIs
3. Under **Implicit grant and hybrid flows**, ensure settings match your application needs:
   - For SharePoint Embedded, typically no checkboxes are needed
4. Under **Advanced settings**:
   - **Allow public client flows**: No (for server-side applications)
   - **Enable the following mobile and desktop flows**: No

### Step 2: Configure Token Configuration (Optional)

1. Select **Token configuration** from the left menu
2. Add optional claims if needed for your application:
   - Click **+ Add optional claim**
   - Select token type (ID, Access, or SAML)
   - Choose claims like `email`, `family_name`, `given_name`, etc.

### Step 3: Branding Configuration

1. Select **Branding & properties** from the left menu
2. Configure the following for better user experience:

   - **Name**: User-facing application name
   - **Logo**: Upload a logo (240x240 pixels recommended)
   - **Home page URL**: Your application's home page
   - **Terms of service URL**: Link to your terms of service
   - **Privacy statement URL**: Link to your privacy policy
   - **Publisher domain**: Your verified domain

3. Click **Save**

## Record Important Values

Before proceeding, ensure you have recorded all these values securely:

### Application Information
```
Application (client) ID: ________________________________
Directory (tenant) ID:   ________________________________
Object ID:               ________________________________
```

### Certificate Information
```
Certificate Thumbprint:  ________________________________
Certificate Path:        ________________________________
Certificate Password:    ________________________________
Certificate Expiry Date: ________________________________
```

### API Permissions Granted
- [ ] Sites.FullControl.All (Microsoft Graph)
- [ ] Files.ReadWrite.All (Microsoft Graph)
- [ ] Sites.FullControl.All (SharePoint)
- [ ] TermStore.ReadWrite.All (SharePoint)
- [ ] Admin consent granted

### Application URLs (if configured)
```
Home Page URL:           ________________________________
Redirect URI:            ________________________________
```

## Next Steps

After completing the Azure App Registration:

1. **Share consent URL with customers**: See [Admin Consent Process](admin-consent.md)
2. **Register owning app in consuming tenant**: Use the `Register-SPEOwningApp.ps1` script
3. **Create containers**: Use the `New-SPEContainer.ps1` script

## Verification Checklist

Before moving to the next step, verify:

- [ ] App registration created successfully
- [ ] All required API permissions added
- [ ] Admin consent granted (for your owning tenant)
- [ ] Certificate created and uploaded
- [ ] Certificate thumbprint recorded
- [ ] PFX file stored securely
- [ ] Application ID and Tenant ID recorded
- [ ] Branding configured (optional but recommended)

## Common Issues

### Issue: "Insufficient privileges" error when granting admin consent
**Solution**: Ensure you have Global Administrator or Privileged Role Administrator role.

### Issue: Certificate upload fails
**Solution**: Ensure you're uploading the `.cer` file (public key), not the `.pfx` file (private key).

### Issue: Can't find SharePoint permissions
**Solution**: Make sure you're selecting the "SharePoint" API, not "Microsoft Graph" when adding SharePoint permissions.

## Additional Resources

- [Azure AD App Registration Documentation](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Certificate Credentials](https://learn.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials)
- [SharePoint Embedded Overview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/overview)
- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)

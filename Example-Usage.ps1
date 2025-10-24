<#
.SYNOPSIS
    Example usage scenarios for the SharePoint Embedded Token Generator script.

.DESCRIPTION
    This file contains example scenarios and usage patterns for generating
    certificate-based tokens for SharePoint Embedded applications.
    
    DO NOT RUN THIS ENTIRE FILE - Copy individual examples to use them.
#>

# ============================================================================
# Example 1: First-Time Setup - Create Certificate and Generate Token
# ============================================================================
# This example creates a new certificate, generates a token, and exports
# the certificate for future use.

.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -ExportCertificate `
    -CertificatePath "C:\SPE-Certs"

# After running this:
# 1. Upload the .cer file to your Azure AD app registration
# 2. Save the thumbprint displayed in the output
# 3. Use the thumbprint for future token generations


# ============================================================================
# Example 2: Generate Token Using Existing Certificate
# ============================================================================
# Use this when you've already created a certificate and uploaded it to Azure AD

.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "customer.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF01234567"


# ============================================================================
# Example 3: Create Certificate with Custom Subject
# ============================================================================
# Use a custom certificate subject name

.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateSubject "CN=ContosoSPEApp" `
    -ExportCertificate


# ============================================================================
# Example 4: Capture Token for Programmatic Use
# ============================================================================
# Generate a token and use it in subsequent API calls

$tokenResponse = .\New-SharePointEmbeddedToken.ps1 `
    -TenantId "customer.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF01234567"

# Use the token in API calls
$headers = @{
    "Authorization" = "Bearer $($tokenResponse.access_token)"
    "Content-Type"  = "application/json"
}

# Example: List SharePoint sites
$sitesResponse = Invoke-RestMethod `
    -Uri "https://graph.microsoft.com/v1.0/sites" `
    -Headers $headers `
    -Method Get

Write-Host "Found $($sitesResponse.value.Count) sites"


# ============================================================================
# Example 5: Export Certificate with Password Protection
# ============================================================================
# Create and export a certificate with password protection

$certPassword = ConvertTo-SecureString -String "YourSecurePassword123!" -AsPlainText -Force

.\New-SharePointEmbeddedToken.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificatePassword $certPassword `
    -ExportCertificate `
    -CertificatePath "C:\SPE-Certs"


# ============================================================================
# Example 6: Complete Registration Workflow
# ============================================================================
# Full workflow to register an app in a consuming tenant

# Step 1: Generate token
Write-Host "Step 1: Generating access token..." -ForegroundColor Cyan
$token = .\New-SharePointEmbeddedToken.ps1 `
    -TenantId "customer.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF01234567"

# Step 2: Prepare headers for API calls
$headers = @{
    "Authorization" = "Bearer $($token.access_token)"
    "Content-Type"  = "application/json"
}

# Step 3: Register SharePoint Embedded container type (example)
Write-Host "Step 2: Registering SharePoint Embedded container..." -ForegroundColor Cyan
$containerBody = @{
    displayName = "My SharePoint Embedded Container"
    description = "Container for ISV application"
} | ConvertTo-Json

# Note: Adjust the API endpoint according to SharePoint Embedded documentation
# $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/..." `
#     -Headers $headers -Method Post -Body $containerBody

Write-Host "Registration complete!" -ForegroundColor Green


# ============================================================================
# Example 7: Token Refresh Workflow
# ============================================================================
# Generate a new token when the current one expires

function Get-SPEToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertThumbprint
    )
    
    $token = .\New-SharePointEmbeddedToken.ps1 `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -CertificateThumbprint $CertThumbprint
    
    return $token
}

# Use in a loop or scheduled task
$token = Get-SPEToken -TenantId "customer.onmicrosoft.com" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertThumbprint "A1B2C3D4E5F67890123456789ABCDEF01234567"

# Token is valid for the duration specified in expires_in (typically 3600 seconds)
Write-Host "Token expires in $($token.expires_in) seconds"


# ============================================================================
# Example 8: Error Handling
# ============================================================================
# Implement proper error handling when generating tokens

try {
    $token = .\New-SharePointEmbeddedToken.ps1 `
        -TenantId "customer.onmicrosoft.com" `
        -ClientId "12345678-1234-1234-1234-123456789012" `
        -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF01234567" `
        -ErrorAction Stop
    
    Write-Host "Token generated successfully" -ForegroundColor Green
    
    # Use token for API calls
    $headers = @{
        "Authorization" = "Bearer $($token.access_token)"
    }
    
    # Your API calls here...
}
catch {
    Write-Host "Failed to generate token: $($_.Exception.Message)" -ForegroundColor Red
    
    # Implement fallback or retry logic
    Write-Host "Please verify:"
    Write-Host "  1. Certificate is uploaded to Azure AD"
    Write-Host "  2. Client ID and Tenant ID are correct"
    Write-Host "  3. App has required permissions"
}


# ============================================================================
# Tips and Best Practices
# ============================================================================

<#
1. Certificate Management:
   - Keep certificate thumbprints in a secure configuration file
   - Rotate certificates regularly (every 12-24 months)
   - Store .pfx files securely (Azure Key Vault recommended)

2. Token Handling:
   - Cache tokens until they expire
   - Implement token refresh logic
   - Never log or expose tokens in plain text

3. Security:
   - Use strong passwords for certificate export
   - Limit access to certificate files
   - Use HTTPS for all API communications
   - Implement proper error handling to avoid leaking sensitive information

4. Multi-Tenant Scenarios:
   - Maintain separate certificates per tenant if required
   - Document tenant-specific configurations
   - Implement tenant validation logic

5. Automation:
   - Integrate token generation in CI/CD pipelines
   - Use scheduled tasks for token rotation
   - Implement monitoring and alerting for token failures
#>

<#
.SYNOPSIS
    Creates a certificate-based SharePoint Token for registering SharePoint Embedded owning app in a consuming tenant.

.DESCRIPTION
    This script generates an access token using certificate-based authentication for SharePoint Embedded applications.
    It can create a new self-signed certificate or use an existing certificate to authenticate and obtain a token
    that can be used to register a SharePoint Embedded owning app in a consuming tenant.

.PARAMETER TenantId
    The Azure AD Tenant ID where the app will be registered (the consuming tenant).

.PARAMETER ClientId
    The Application (Client) ID of the SharePoint Embedded owning app.

.PARAMETER CertificateThumbprint
    The thumbprint of an existing certificate to use for authentication.
    If not provided, a new self-signed certificate will be created.

.PARAMETER CertificateSubject
    The subject name for the certificate. Default is "CN=SharePointEmbeddedApp".
    Only used when creating a new certificate.

.PARAMETER CertificatePassword
    The password to protect the certificate's private key.
    Only used when creating a new certificate.

.PARAMETER ExportCertificate
    If specified, exports the certificate to a PFX file for backup or distribution.

.PARAMETER CertificatePath
    The path where the certificate will be exported. Default is current directory.
    Only used when ExportCertificate is specified.

.EXAMPLE
    .\New-SharePointEmbeddedToken.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "12345678-1234-1234-1234-123456789012"
    
    Creates a new self-signed certificate and generates a token for the specified tenant and app.

.EXAMPLE
    .\New-SharePointEmbeddedToken.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "ABC123..."
    
    Uses an existing certificate to generate a token.

.EXAMPLE
    .\New-SharePointEmbeddedToken.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "12345678-1234-1234-1234-123456789012" -ExportCertificate -CertificatePath "C:\Certs"
    
    Creates a new certificate, generates a token, and exports the certificate to the specified path.

.NOTES
    Author: SharePoint Embedded ISV Setup
    Version: 1.0
    Requirements: 
    - PowerShell 5.1 or later
    - Admin rights to create certificates in the certificate store
    - Internet connectivity to reach Azure AD endpoints
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "The Azure AD Tenant ID (e.g., contoso.onmicrosoft.com or GUID)")]
    [string]$TenantId,

    [Parameter(Mandatory = $true, HelpMessage = "The Application (Client) ID of the SharePoint Embedded app")]
    [string]$ClientId,

    [Parameter(Mandatory = $false, HelpMessage = "The thumbprint of an existing certificate")]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false, HelpMessage = "The subject name for the certificate")]
    [string]$CertificateSubject = "CN=SharePointEmbeddedApp",

    [Parameter(Mandatory = $false, HelpMessage = "Password to protect the certificate's private key")]
    [SecureString]$CertificatePassword,

    [Parameter(Mandatory = $false, HelpMessage = "Export the certificate to a PFX file")]
    [switch]$ExportCertificate,

    [Parameter(Mandatory = $false, HelpMessage = "Path where the certificate will be exported")]
    [string]$CertificatePath = (Get-Location).Path
)

#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Function to create a JWT token
function New-JwtToken {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    try {
        # Token endpoint
        $authority = "https://login.microsoftonline.com/$TenantId"
        
        # Create JWT header
        $header = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [Convert]::ToBase64String($Certificate.GetCertHash()) -replace '\+', '-' -replace '/', '_' -replace '='
        }

        # Create JWT payload
        $now = [Math]::Floor([decimal](Get-Date(Get-Date).ToUniversalTime() -UFormat "%s"))
        $exp = $now + 3600  # Token valid for 1 hour

        $payload = @{
            aud = "$authority/oauth2/v2.0/token"
            exp = $exp
            iss = $ClientId
            jti = [Guid]::NewGuid().ToString()
            nbf = $now
            sub = $ClientId
        }

        # Encode header and payload
        $headerJson = $header | ConvertTo-Json -Compress
        $payloadJson = $payload | ConvertTo-Json -Compress
        
        $headerEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerJson)) -replace '\+', '-' -replace '/', '_' -replace '='
        $payloadEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson)) -replace '\+', '-' -replace '/', '_' -replace '='

        # Create signature
        $signatureInput = "$headerEncoded.$payloadEncoded"
        $signatureInputBytes = [System.Text.Encoding]::UTF8.GetBytes($signatureInput)
        
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
        $signatureBytes = $rsa.SignData($signatureInputBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $signatureEncoded = [Convert]::ToBase64String($signatureBytes) -replace '\+', '-' -replace '/', '_' -replace '='

        # Combine to create JWT
        $jwt = "$headerEncoded.$payloadEncoded.$signatureEncoded"
        
        return $jwt
    }
    catch {
        Write-Error "Failed to create JWT token: $_"
        throw
    }
}

# Function to get access token
function Get-AccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientAssertion
    )

    try {
        $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        $body = @{
            client_id             = $ClientId
            client_assertion      = $ClientAssertion
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
        }

        Write-Verbose "Requesting access token from $tokenEndpoint"
        $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
        
        return $response
    }
    catch {
        Write-Error "Failed to get access token: $_"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Error "Response: $responseBody"
        }
        throw
    }
}

# Main script logic
try {
    Write-Host "SharePoint Embedded Token Generator" -ForegroundColor Cyan
    Write-Host "=====================================" -ForegroundColor Cyan
    Write-Host ""

    # Get or create certificate
    $certificate = $null
    
    if ($CertificateThumbprint) {
        Write-Host "Looking for existing certificate with thumbprint: $CertificateThumbprint" -ForegroundColor Yellow
        $certificate = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -Recurse | 
                       Where-Object { $_.Thumbprint -eq $CertificateThumbprint } | 
                       Select-Object -First 1
        
        if (-not $certificate) {
            throw "Certificate with thumbprint '$CertificateThumbprint' not found in certificate store."
        }
        
        Write-Host "Certificate found: $($certificate.Subject)" -ForegroundColor Green
    }
    else {
        Write-Host "Creating new self-signed certificate..." -ForegroundColor Yellow
        
        # Create certificate parameters
        $certParams = @{
            Subject           = $CertificateSubject
            CertStoreLocation = "Cert:\CurrentUser\My"
            KeyExportPolicy   = "Exportable"
            KeySpec           = "Signature"
            KeyLength         = 2048
            KeyAlgorithm      = "RSA"
            HashAlgorithm     = "SHA256"
            NotAfter          = (Get-Date).AddYears(2)
            Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"
        }
        
        $certificate = New-SelfSignedCertificate @certParams
        
        Write-Host "Certificate created successfully!" -ForegroundColor Green
        Write-Host "  Thumbprint: $($certificate.Thumbprint)" -ForegroundColor Gray
        Write-Host "  Subject: $($certificate.Subject)" -ForegroundColor Gray
        Write-Host "  Valid Until: $($certificate.NotAfter)" -ForegroundColor Gray
        
        # Export certificate if requested
        if ($ExportCertificate) {
            if (-not $CertificatePassword) {
                $CertificatePassword = Read-Host -Prompt "Enter password for certificate export" -AsSecureString
            }
            
            $certFileName = "SharePointEmbedded_$($certificate.Thumbprint).pfx"
            $certFilePath = Join-Path -Path $CertificatePath -ChildPath $certFileName
            
            Write-Host "Exporting certificate to: $certFilePath" -ForegroundColor Yellow
            Export-PfxCertificate -Cert $certificate -FilePath $certFilePath -Password $CertificatePassword | Out-Null
            Write-Host "Certificate exported successfully!" -ForegroundColor Green
            
            # Also export public key for upload to Azure AD
            $cerFileName = "SharePointEmbedded_$($certificate.Thumbprint).cer"
            $cerFilePath = Join-Path -Path $CertificatePath -ChildPath $cerFileName
            Export-Certificate -Cert $certificate -FilePath $cerFilePath | Out-Null
            Write-Host "Public certificate exported to: $cerFilePath" -ForegroundColor Green
            Write-Host "Upload this .cer file to your Azure AD app registration." -ForegroundColor Cyan
        }
    }

    # Verify certificate has private key
    if (-not $certificate.HasPrivateKey) {
        throw "Certificate does not have a private key. Cannot sign JWT token."
    }

    Write-Host ""
    Write-Host "Generating JWT assertion..." -ForegroundColor Yellow
    $jwt = New-JwtToken -ClientId $ClientId -TenantId $TenantId -Certificate $certificate
    
    Write-Host "Requesting access token..." -ForegroundColor Yellow
    $tokenResponse = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientAssertion $jwt
    
    Write-Host ""
    Write-Host "Success! Access token obtained." -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Token Details:" -ForegroundColor Cyan
    Write-Host "  Token Type: $($tokenResponse.token_type)" -ForegroundColor Gray
    Write-Host "  Expires In: $($tokenResponse.expires_in) seconds" -ForegroundColor Gray
    Write-Host "  Scope: $($tokenResponse.scope)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Access Token:" -ForegroundColor Cyan
    Write-Host $tokenResponse.access_token -ForegroundColor White
    Write-Host ""
    
    # Provide usage instructions
    Write-Host "Next Steps:" -ForegroundColor Cyan
    Write-Host "1. Copy the access token above" -ForegroundColor White
    Write-Host "2. Use it in your API calls to register the SharePoint Embedded app" -ForegroundColor White
    Write-Host "3. Include it in the Authorization header as: Bearer <token>" -ForegroundColor White
    
    if (-not $CertificateThumbprint) {
        Write-Host ""
        Write-Host "Important: Certificate Information" -ForegroundColor Yellow
        Write-Host "Thumbprint: $($certificate.Thumbprint)" -ForegroundColor White
        Write-Host "Save this thumbprint for future use to generate tokens with the same certificate." -ForegroundColor White
    }
    
    # Return token object for pipeline usage
    return $tokenResponse
}
catch {
    Write-Host ""
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure the certificate is uploaded to your Azure AD app registration" -ForegroundColor White
    Write-Host "2. Verify the Client ID and Tenant ID are correct" -ForegroundColor White
    Write-Host "3. Check that the app has the required API permissions" -ForegroundColor White
    Write-Host "4. Run PowerShell as Administrator to create certificates" -ForegroundColor White
    throw
}
